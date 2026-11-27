#!/usr/bin/env node
// ─── 固定化 xlsx wasm 引擎：下载 + 校验，不再编译 Rust ─────────────────
//
// 设计约束：
//   1. **唯一真源是 `wasm.lock.json`**：url + sha256 + size 都在仓库里，
//      产物换版本必须走一次 `--update-lock`，不会出现"构建机器不同、
//      wasm 也不同"的漂移。
//   2. **下载优先、内容寻址缓存**：先看 `.cache/<sha256>.wasm`，命中就不联网；
//      CI 冷缓存时下一次 7 MB，可接受（比 cargo 冷编译快两个数量级）。
//   3. **校验不过就当没有**：sha256 / 导出符号 / 体积任一不符直接失败，
//      绝不把来路不明的 wasm 落进 dist。
//   4. **默认不阻断本地开发**：找不到产物时只警告并退出 0，让页面回退到
//      「读整份字节」模式；CI 用 `--require` 或 `XLSX_REQUIRE_WASM=1` 才 hard fail。
//
// 用法：
//   node scripts/fetch-wasm.mjs                    # 按 lock 取产物（有缓存则不联网）
//   node scripts/fetch-wasm.mjs --require          # 取不到 → exit 1（CI）
//   node scripts/fetch-wasm.mjs --refresh          # 忽略缓存重新下载
//   node scripts/fetch-wasm.mjs --from <url|path>  # 临时换源（http/https/file/本地路径）
//   node scripts/fetch-wasm.mjs --file <path> --update-lock
//                                                  # 把本地一份 wasm 固定进 lock（重新锚定）
//
// 环境变量：
//   XLSX_ENGINE_WASM_URL   覆盖 lock 里的 url（镜像 / 私有 CDN）
//   XLSX_REQUIRE_WASM=1    等价于 --require
//
// 产物：
//   dist/xlsx-engine.wasm    随包分发（pack.mjs 会做 sha 自证）
//   dist/xlsx-engine.json    { target, features, size, sha256, exports, protocolVersion, source }

import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { mkdir, readFile, rename, stat, writeFile } from 'node:fs/promises'
import { dirname, isAbsolute, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const lockPath = join(root, 'wasm.lock.json')
const cacheDir = join(root, '.cache')
const distDir = join(root, 'dist')
const outWasm = join(distDir, 'xlsx-engine.wasm')
const outJson = join(distDir, 'xlsx-engine.json')

/** reactor 必须有的导出（与 scripts/build-wasm.mjs 保持一致）。 */
const REQUIRED_EXPORTS = ['memory', '_initialize', 'call', 'wasi_alloc', 'wasi_free']
const FORBIDDEN_EXPORTS = ['_start']
const MAX_BYTES = 12 * 1024 * 1024

const argv = process.argv.slice(2)
const flag = (name) => argv.includes(name)
const value = (name, fallback = '') => {
  const index = argv.indexOf(name)
  const next = index >= 0 ? argv[index + 1] : undefined
  return next === undefined || next.startsWith('--') ? fallback : next
}

const requireWasm = flag('--require') || process.env['XLSX_REQUIRE_WASM'] === '1'
const refresh = flag('--refresh')
const quiet = flag('--quiet')
const updateLock = flag('--update-lock')
const fromArg = value('--from')
const fileArg = value('--file')

const log = (...args) => {
  if (!quiet) console.log(...args)
}
const warn = (...args) => console.warn(...args)

class Missing extends Error {}

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex')
}

async function readLock() {
  if (!existsSync(lockPath)) return { version: 1 }
  try {
    return JSON.parse(await readFile(lockPath, 'utf8'))
  } catch (error) {
    throw new Missing(`wasm.lock.json 解析失败：${error instanceof Error ? error.message : error}`)
  }
}

/** 校验 wasm：体积、导出符号、（可选）sha256。返回 manifest 需要的字段。 */
function inspect(bytes, expected) {
  if (bytes.byteLength === 0) throw new Missing('产物为空')
  if (bytes.byteLength > MAX_BYTES)
    throw new Missing(`体积 ${(bytes.byteLength / 1048576).toFixed(2)} MB 超过预算 ${MAX_BYTES / 1048576} MB`)

  if (expected.sha256) {
    const actual = sha256(bytes)
    if (actual !== expected.sha256)
      throw new Missing(
        `sha256 不一致：期望 ${expected.sha256.slice(0, 16)}… 实际 ${actual.slice(0, 16)}…` +
          '（产物换版本请用 --file <wasm> --update-lock 重新锚定）',
      )
  }
  if (expected.size && bytes.byteLength !== expected.size)
    throw new Missing(`体积与 lock 不符：期望 ${expected.size} 实际 ${bytes.byteLength}`)

  let module
  try {
    module = new WebAssembly.Module(bytes)
  } catch (error) {
    throw new Missing(`不是合法 wasm：${error instanceof Error ? error.message : error}`)
  }
  const exports = WebAssembly.Module.exports(module).map((entry) => entry.name)
  for (const name of REQUIRED_EXPORTS)
    if (!exports.includes(name)) throw new Missing(`缺少导出：${name}（现有 ${exports.join(',')}）`)
  for (const name of FORBIDDEN_EXPORTS)
    if (exports.includes(name)) throw new Missing(`含 ${name}：command module 会被 wasi.start 一次性锁死`)

  return { exports, sha256: sha256(bytes), size: bytes.byteLength }
}

async function readLocal(path) {
  const file = path.startsWith('file://') ? fileURLToPath(path) : isAbsolute(path) ? path : resolve(root, path)
  const info = await stat(file).catch(() => null)
  if (!info?.isFile()) throw new Missing(`本地文件不存在：${file}`)
  return new Uint8Array(await readFile(file))
}

async function download(url) {
  const attempts = 3
  let lastError = ''
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const response = await fetch(url, {
        redirect: 'follow',
        signal: AbortSignal.timeout(120_000),
        headers: { 'user-agent': 'workstudio-plugin-xlsx/fetch-wasm' },
      })
      if (!response.ok) {
        lastError = `HTTP ${response.status} ${response.statusText}`
        // 4xx 重试没意义（404 = 还没上传产物）
        if (response.status >= 400 && response.status < 500) break
      } else {
        return new Uint8Array(await response.arrayBuffer())
      }
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error)
    }
    if (attempt < attempts) {
      log(`  第 ${attempt} 次下载失败（${lastError}），${attempt + 1}s 后重试…`)
      await new Promise((done) => setTimeout(done, attempt * 1000))
    }
  }
  throw new Missing(`下载失败：${url}（${lastError}）`)
}

/** 按优先级找一个可用的字节来源；都不行时抛 Missing。 */
async function resolveBytes(lock) {
  const pinned = { sha256: lock.sha256, size: lock.size }

  // 1) 显式本地文件（重新锚定 / 离线环境）
  if (fileArg) {
    log(`  源  本地文件 ${fileArg}`)
    return { bytes: await readLocal(fileArg), source: { type: 'file', path: fileArg } }
  }

  // 2) 内容寻址缓存：命中就完全不联网
  if (!refresh && lock.sha256) {
    const cached = join(cacheDir, `${lock.sha256}.wasm`)
    if (existsSync(cached)) {
      log(`  源  缓存 ${cached}`)
      return { bytes: new Uint8Array(await readFile(cached)), source: { type: 'cache', path: '.cache' } }
    }
  }

  // 3) 显式换源（--from / XLSX_ENGINE_WASM_URL）
  const override = fromArg || process.env['XLSX_ENGINE_WASM_URL'] || ''
  if (override) {
    log(`  源  ${fromArg ? '--from' : 'XLSX_ENGINE_WASM_URL'} ${override}`)
    const bytes = /^https?:\/\//i.test(override) ? await download(override) : await readLocal(override)
    return { bytes, source: { type: 'url', url: override } }
  }

  // 4) lock 里的固定 url
  if (lock.url) {
    log(`  源  wasm.lock.json ${lock.url}`)
    return { bytes: await download(lock.url), source: { type: 'url', url: lock.url, pinned: 'wasm.lock.json' } }
  }

  // 5) 走 npm 依赖分发（如果以后把产物发成包）
  const packaged = join(root, 'node_modules/@workstudio/xlsx-engine-wasm/xlsx-engine.wasm')
  if (existsSync(packaged)) {
    log(`  源  npm 包 ${packaged}`)
    return { bytes: new Uint8Array(await readFile(packaged)), source: { type: 'npm' } }
  }

  throw new Missing(
    '没有可用的产物：wasm.lock.json 缺 url，也没有 --file / --from / 缓存 / npm 包。' +
      '首次锚定：node scripts/fetch-wasm.mjs --file <xlsx-engine.wasm> --update-lock',
  )
}

async function main() {
  const lock = await readLock()

  if (updateLock) {
    if (!fileArg) throw new Missing('--update-lock 需要配合 --file <本地 wasm>')
    const bytes = await readLocal(fileArg)
    const info = inspect(bytes, {})
    const next = {
      ...lock,
      version: 1,
      target: lock.target ?? 'wasm32-wasip1',
      features: lock.features ?? ['single-thread'],
      builtAt: new Date().toISOString(),
      sha256: info.sha256,
      size: info.size,
    }
    await writeFile(lockPath, `${JSON.stringify(next, null, 2)}\n`)
    log(`  锁定 ${fileArg}`)
    log(`       sha256 ${info.sha256}`)
    log(`       size   ${info.size}`)
    log(`       → ${lockPath}`)
    return publish(bytes, info, { type: 'file', path: fileArg }, next)
  }

  const resolved = await resolveBytes(lock)
  const info = inspect(resolved.bytes, { sha256: lock.sha256, size: lock.size })
  if (resolved.source.type !== 'cache') {
    await mkdir(cacheDir, { recursive: true })
    await writeFile(join(cacheDir, `${info.sha256}.wasm`), resolved.bytes)
  }
  return publish(resolved.bytes, info, resolved.source, lock)
}

async function publish(bytes, info, source, lock) {
  const manifest = {
    target: lock.target ?? 'wasm32-wasip1',
    features: lock.features ?? ['single-thread'],
    size: info.size,
    sha256: info.sha256,
    exports: info.exports,
    builtAt: lock.builtAt ?? new Date().toISOString(),
    protocolVersion: info.exports.includes('xlsx_protocol_version') ? 'runtime' : 'missing',
    source,
  }

  await mkdir(distDir, { recursive: true })
  const tmpWasm = `${outWasm}.tmp`
  await writeFile(tmpWasm, bytes)
  await rename(tmpWasm, outWasm)
  await writeFile(outJson, `${JSON.stringify(manifest, null, 2)}\n`)

  log(`  ok  ${outWasm}`)
  log(
    `      ${(info.size / 1048576).toFixed(2)} MB · sha256 ${info.sha256.slice(0, 16)}… · 来源 ${source.type}`,
  )
  log(`      exports ${info.exports.join(',')}`)
}

main().catch(async (error) => {
  const message = error instanceof Error ? error.message : String(error)
  if (error instanceof Missing) {
    // 显式指定了源（--file / --from / XLSX_ENGINE_WASM_URL）还拿不到，就没必要静默降级：
    // 人的意图很明确，失败要看得见。
    const explicit = Boolean(fileArg || fromArg || process.env['XLSX_ENGINE_WASM_URL'])
    if (!requireWasm && !explicit) {
      if (!quiet) {
        warn(`skip  wasm 引擎不可用：${message}`)
        warn('      页面将回退到「读整份字节」模式（CI 请用 --require / XLSX_REQUIRE_WASM=1）')
      }
      return
    }
    const annotation = process.env['GITHUB_ACTIONS'] === 'true' ? '::error::' : ''
    console.error(`${annotation}wasm 引擎不可用：${message}`)
    process.exitCode = 1
    return
  }
  console.error(`✗ fetch-wasm 失败：${message}`)
  process.exitCode = 1
})
