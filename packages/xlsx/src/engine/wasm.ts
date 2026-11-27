// ─── wasm 引擎桥（node:wasi reactor） ──────────────────────────────────
//
// 一次 `call` 一个请求，JSON 进出（ABI 见 rust/xlsx-engine/wasm.rs）：
//   wasi_alloc(len) -> ptr / call(ptr,len) -> i64 / wasi_free(ptr,len)
//
// 三条硬约束：
//   1. **失败只降级**：wasm 缺失或初始化失败 → 返回 null，exec 走 TS 回退路径，
//      并在 RuntimeReport 里声明 `degraded:['engine-init-failed']`。
//   2. **guest 只看得到宿主临时区**：preopen `<tempDir>` → `/tmp`，`TMPDIR=/tmp`
//      （上游 `temp_dir()` / `canonicalize()` 就是按这套假设写的）。
//   3. **实例按目录缓存**：preopen 在实例化时固定，换目录就得重建 —— 所以调用方
//      必须传**稳定**的 `ctx.sessionDir`（宿主按插件提供的会话目录），不能传每次
//      调用都新建的 `ctx.tempDir`，否则每次读都要重开 + 重建索引（O(范围) 失效）。
//      引擎的 workbook session 也活在这个实例里，按 `(ref, version)` 过期由
//      sessions.ts 负责。

import { existsSync, readFileSync } from 'node:fs'
import { mkdir, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

/**
 * 产物路径。正常情况与 `exec.mjs` 同目录（`dist/xlsx-engine.wasm`）；
 * 产物与 `dist/engine.js` 同目录（诊断/测试入口）时退一级到 `../engine/`。
 */
export function wasmArtifactPath(): URL {
  const candidates = ['./xlsx-engine.wasm', '../engine/xlsx-engine.wasm']
  for (const relative of candidates) {
    const url = new URL(relative, import.meta.url)
    if (existsSync(url)) return url
  }
  return new URL(candidates[0], import.meta.url)
}

const REQUIRED_EXPORTS = ['memory', '_initialize', 'call', 'wasi_alloc', 'wasi_free']

export interface WasmEngine {
  readonly tempDir: string
  /** 引擎声明的协议版本（`xlsx_protocol_version` 导出，宿主侧 0 成本断言）。 */
  protocolVersion(): number
  /** 宿主路径 → guest 路径（只允许临时区内的文件）。 */
  guestPath(fileName: string): string
  /** 在临时区里落一个 guest 可见的文件，返回 guest 路径。 */
  writeTemp(fileName: string, content: string | Uint8Array): Promise<string>
  /** 发一个请求；`ok !== true` 时抛带 code 的错误。 */
  invoke<T = unknown>(request: Record<string, unknown>): Promise<T>
}

export type WasmUnavailableReason =
  | 'module-missing'
  | 'wasi-unavailable'
  | 'init-failed'
  | 'abi-mismatch'

export interface WasmProbe {
  readonly ok: boolean
  readonly reason?: WasmUnavailableReason
  readonly protocolVersion?: number
  readonly exports?: string[]
}

let cached: { key: string; engine: WasmEngine } | null = null

function effectiveTempDir(tempDir: string | undefined): string {
  return tempDir && tempDir.length > 0 ? tempDir : tmpdir()
}

function coded(code: string, message: string): Error & { code: string } {
  const error = new Error(message) as Error & { code: string }
  error.code = code
  return error
}

/**
 * 加载（或复用）wasm 引擎。任何失败都返回 null —— 调用方据此回退，
 * 不要在这里抛：wasm 可用性是能力位，不是错误。
 */
export async function loadWasmEngine(tempDir: string | undefined): Promise<WasmEngine | null> {
  return (await loadWasmEngineWithReason(tempDir)).engine
}

let lastFailure: WasmUnavailableReason | null = null

/** 降级原因（诊断用）：区分"没有产物"和"宿主没有 node:wasi"。 */
export function wasmUnavailableReason(): WasmUnavailableReason | null {
  return lastFailure
}

async function loadWasmEngineWithReason(
  tempDir: string | undefined,
): Promise<{ engine: WasmEngine | null }> {
  const dir = effectiveTempDir(tempDir)
  if (cached?.key === dir) {
    lastFailure = null
    return { engine: cached.engine }
  }
  if (cached?.key !== dir) cached = null

  let bytes: Uint8Array<ArrayBuffer>
  try {
    // 拷进自有 ArrayBuffer：Node 的 Buffer 是 ArrayBufferLike（可能是
    // SharedArrayBuffer），不满足 WebAssembly.BufferSource 的类型约束。
    const file = readFileSync(wasmArtifactPath())
    bytes = new Uint8Array(new ArrayBuffer(file.byteLength))
    bytes.set(file)
  } catch {
    lastFailure = 'module-missing'
    return { engine: null }
  }

  let WASI: typeof import('node:wasi').WASI
  try {
    ;({ WASI } = await import('node:wasi'))
  } catch {
    // 例如宿主的 Electron/Node 没有 WASI（需要实测）：这不是"插件坏掉"，
    // 只是执行底座降级，`degraded` 里会写清原因。
    lastFailure = 'wasi-unavailable'
    return { engine: null }
  }

  try {
    const wasi = new WASI({
      version: 'preview1',
      preopens: { '/tmp': dir },
      env: { TMPDIR: '/tmp' },
      returnOnExit: false,
    })
    // compile + instantiate(module)：imports 走 wasi.wasiImport；reactor 用
    // `initialize`（不是 `start` —— 那是 command module 的语义）。
    const module = await WebAssembly.compile(bytes)
    const instance = await WebAssembly.instantiate(module, {
      wasi_snapshot_preview1: wasi.wasiImport,
    })
    wasi.initialize(instance)
    const exported = instance.exports as unknown as Record<string, unknown>
    for (const name of REQUIRED_EXPORTS)
      if (exported[name] === undefined) {
        lastFailure = 'abi-mismatch'
        return { engine: null }
      }

    const memory = exported['memory'] as WebAssembly.Memory
    const alloc = exported['wasi_alloc'] as (length: number) => number
    const free = exported['wasi_free'] as (ptr: number, length: number) => void
    const call = exported['call'] as (ptr: number, length: number) => bigint
    const encoder = new TextEncoder()
    const decoder = new TextDecoder()

    const protocolVersionFn = exported['xlsx_protocol_version'] as
      | (() => number)
      | undefined
    const engine: WasmEngine = {
      tempDir: dir,
      protocolVersion: () => Number(protocolVersionFn?.() ?? 0),
      guestPath: (fileName) => `/tmp/${fileName}`,
      async writeTemp(fileName, content) {
        const hostPath = join(dir, fileName)
        await mkdir(dir, { recursive: true })
        await writeFile(hostPath, content)
        return `/tmp/${fileName}`
      },
      async invoke<T>(request: Record<string, unknown>): Promise<T> {
        const payload = encoder.encode(JSON.stringify(request))
        const ptr = alloc(payload.length)
        if (!ptr) throw coded('engine.unavailable', 'wasi_alloc 返回 0')
        const view = () => new Uint8Array(memory.buffer)
        if (ptr + payload.length > memory.buffer.byteLength)
          throw coded('engine.unavailable', 'wasm 线性内存不足')
        view().set(payload, ptr)
        let returned: bigint
        try {
          returned = call(ptr, payload.length)
        } finally {
          free(ptr, payload.length)
        }
        const outPtr = Number(returned >> 32n)
        const outLen = Number(returned & 0xffffffffn)
        if (outPtr + outLen > memory.buffer.byteLength)
          throw coded('engine.unavailable', 'wasm 返回的响应越界')
        const text = decoder.decode(view().slice(outPtr, outPtr + outLen))
        if (outLen > 0) free(outPtr, outLen)
        let parsed: { ok?: boolean; result?: unknown; error?: { code?: string; message?: string } }
        try {
          parsed = JSON.parse(text)
        } catch {
          throw coded('command.invalid', `wasm 引擎返回了非 JSON 响应：${text.slice(0, 120)}`)
        }
        if (parsed.ok !== true)
          throw coded(
            parsed.error?.code ? `engine.${parsed.error.code}` : 'command.invalid',
            parsed.error?.message ?? 'wasm 引擎拒绝了该请求',
          )
        return parsed.result as T
      },
    }
    cached = { key: dir, engine }
    lastFailure = null
    return { engine }
  } catch {
    lastFailure = 'init-failed'
    return { engine: null }
  }
}

/** 探活：加载 + 协议版本断言，不抛异常。 */
export async function probeWasm(tempDir: string | undefined): Promise<WasmProbe> {
  const engine = await loadWasmEngine(tempDir)
  if (!engine)
    return { ok: false, reason: lastFailure ?? 'init-failed' }
  // 直接调导出读协议版本：比发一个请求便宜，而且失败能区分"ABI 不对"。
  const protocolVersion = engine.protocolVersion()
  if (protocolVersion !== 1)
    return { ok: false, reason: 'abi-mismatch', protocolVersion }
  return { ok: true, protocolVersion }
}

/** 测试/诊断用：丢掉缓存的实例。 */
export function resetWasmEngine(): void {
  cached = null
}
