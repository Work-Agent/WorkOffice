// ─── wasm 工作簿会话（按 (ref, version) 缓存） ────────────────────────
//
// `ReadRange` 需要引擎侧的 session，而 session 是「打开一个路径」的产物。
// 所以这里做三件事：
//
//   1. **把权威字节落成临时文件**（每个 (ref, version) 一次）；
//   2. **单飞 open**：首次读要建索引（实测 60 MB≈1.9 s、300 MB≈15 s），
//      同一 key 并发只允许一次 open，后来者等同一个 Promise；
//   3. **LRU 2 + 版本失效**：版本变（保存后）就 close 并删临时文件。
//
// 会话只是**派生缓存**（主线 I2）：丢了能从字节重建，绝不作为正确性依据。
// 因此这里不做 `commitModel` / `abortModel`，插件清单仍是 `model: 'none'`。

import { rm } from 'node:fs/promises'
import { join } from 'node:path'

import { coded } from './abi'
import type { WasmEngine } from './wasm'

export interface EngineSession {
  readonly ref: string
  readonly version: string
  readonly sessionId: string
  readonly sheets: ReadonlyArray<{ id: string; name: string; rowCount?: number; columnCount?: number }>
  /** `open` 的原样元数据（页面重建 file 模型用）。 */
  readonly metadata: unknown
  /** 本次调用是否复用了已有会话（false = 刚 open）。 */
  readonly cached: boolean
}

interface Entry extends EngineSession {
  readonly fileName: string
  lastUsed: number
}

/** 默认 2 个模型（对照主线 §4.4 的 2 × 512 MB 预算）。 */
const MAX_SESSIONS = 2

const entries = new Map<string, Entry>()
const opening = new Map<string, Promise<Entry>>()
let sequence = 0

const keyOf = (ref: string, version: string) => `${ref}\u0000${version}`

function fileNameOf(ref: string, version: string): string {
  const safeRef = ref.replace(/[^A-Za-z0-9._-]+/g, '_').slice(-60)
  return `session-${safeRef}-${version.slice(0, 12)}.xlsx`
}

/**
 * 取（或建立）会话。`bytes` 是**惰性**的：命中缓存时不会读整份文件。
 */
export async function ensureSession(
  engine: WasmEngine,
  ref: string,
  version: string,
  bytes: () => Promise<Uint8Array>,
): Promise<EngineSession> {
  const key = keyOf(ref, version)
  const hit = entries.get(key)
  if (hit) {
    hit.lastUsed = ++sequence
    return { ...hit, cached: true }
  }
  const pending = opening.get(key)
  if (pending) return { ...(await pending), cached: true }

  const created = openSession(engine, ref, version, bytes)
    .then((entry) => {
      entries.set(key, entry)
      void evict(engine, key)
      return entry
    })
    .finally(() => opening.delete(key))
  opening.set(key, created)
  return created
}

async function openSession(
  engine: WasmEngine,
  ref: string,
  version: string,
  bytes: () => Promise<Uint8Array>,
): Promise<Entry> {
  const fileName = fileNameOf(ref, version)
  const baseline = await bytes()
  await engine.writeTemp(fileName, baseline)
  const metadata = await engine.invoke<{
    sessionId: string
    sheets?: Array<{ id: string; name: string; rowCount?: number; columnCount?: number }>
  }>({
    version: 1,
    requestId: `open-${version.slice(0, 8)}`,
    command: 'open',
    path: engine.guestPath(fileName),
  })
  if (!metadata?.sessionId)
    throw coded('engine.invalid_request', 'wasm open 没有返回 sessionId')
  return {
    ref,
    version,
    sessionId: metadata.sessionId,
    metadata,
    sheets: (metadata.sheets ?? []).map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      ...(sheet.rowCount === undefined ? {} : { rowCount: sheet.rowCount }),
      ...(sheet.columnCount === undefined ? {} : { columnCount: sheet.columnCount }),
    })),
    fileName,
    lastUsed: ++sequence,
    cached: false,
  }
}

async function evict(engine: WasmEngine, keepKey: string): Promise<void> {
  while (entries.size > MAX_SESSIONS) {
    let oldestKey: string | null = null
    let oldest = Number.POSITIVE_INFINITY
    for (const [key, entry] of entries) {
      if (key === keepKey) continue
      if (entry.lastUsed < oldest) {
        oldest = entry.lastUsed
        oldestKey = key
      }
    }
    if (oldestKey === null) return
    const entry = entries.get(oldestKey)
    entries.delete(oldestKey)
    if (!entry) return
    await closeEntry(engine, entry)
  }
}

async function closeEntry(engine: WasmEngine, entry: Entry): Promise<void> {
  try {
    await engine.invoke({
      version: 1,
      requestId: `close-${entry.version.slice(0, 8)}`,
      command: 'close',
      sessionId: entry.sessionId,
    })
  } catch {
    /* 关闭失败不影响正确性：会话只是派生缓存 */
  }
  await rm(join(engine.tempDir, entry.fileName), { force: true }).catch(() => {})
}

/** 版本失效：保存成功后调用（或由调用方按 version 自然换 key）。 */
export async function closeSessionsFor(engine: WasmEngine, ref: string): Promise<void> {
  for (const [key, entry] of [...entries]) {
    if (entry.ref !== ref) continue
    entries.delete(key)
    await closeEntry(engine, entry)
  }
}

/** 诊断/测试：当前缓存的会话（不暴露可变结构）。 */
export function sessionKeys(): string[] {
  return [...entries.keys()]
}

/** 测试用：清空缓存（不调用 close，避免测试依赖引擎状态）。 */
export function resetSessions(): void {
  entries.clear()
  opening.clear()
}
