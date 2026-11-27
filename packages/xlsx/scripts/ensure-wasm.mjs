#!/usr/bin/env node
// ─── xlsx wasm 引擎就位（下载优先，编不过就编译） ──────────────────────
//
// 下载与编译两条路共用同一个校验（sha256 / 导出符号 / 体积），产物落在
// `dist/xlsx-engine.wasm` + `dist/xlsx-engine.json`，后续 build / verify /
// test / pack 都只看这两个文件。
//
//   XLSX_ENGINE_WASM=fetch  只用 wasm.lock.json 固定的产物（下载 / 缓存 / --file）
//   XLSX_ENGINE_WASM=build  只用 Rust 源码编译（编译器版本见 wasm.lock.json 的 compile.rustc）
//   XLSX_ENGINE_WASM=auto   先 fetch，拿不到再 build（默认；CI 就是这个）
//   XLSX_ENGINE_WASM=off    不处理，页面回退「读整份字节」模式
//
//   XLSX_REQUIRE_WASM=1 / --require  两条路都失败 → exit 1（不允许静默降级）
//
// 为什么 CI 里要能编译：上游 genoffice 只发 Rust 源码和桌面原生产物，不发 wasm；
// 真正能在页面里跑的是「上游引擎 + packages/xlsx/rust/xlsx-engine/wasm.rs 这层
// reactor ABI」，所以产物必须由我们自己产生——**一次**即可：编译结果按 lock 固定
// 后，后续跑 fetch 就够了，编译只是保底路径。

import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const argv = process.argv.slice(2)
const requireWasm = argv.includes('--require') || process.env['XLSX_REQUIRE_WASM'] === '1'
const mode = process.env['XLSX_ENGINE_WASM'] ?? 'auto'
const artifact = join(root, 'dist/xlsx-engine.wasm')
const artifactManifest = join(root, 'dist/xlsx-engine.json')
const lockPath = join(root, 'wasm.lock.json')

const log = (...args) => console.log(...args)

/** dist 里已经有一份自洽的产物？返回它的 manifest（sha 对不上就当没有）。 */
function settled() {
  if (!existsSync(artifact) || !existsSync(artifactManifest)) return null
  try {
    const manifest = JSON.parse(readFileSync(artifactManifest, 'utf8'))
    const sha = createHash('sha256').update(readFileSync(artifact)).digest('hex')
    if (sha !== manifest.sha256) return null
    const lock = existsSync(lockPath) ? JSON.parse(readFileSync(lockPath, 'utf8')) : {}
    return { manifest, sha, lock }
  } catch {
    return null
  }
}

function run(script, args) {
  const result = spawnSync(process.execPath, [join(root, 'scripts', script), ...args], {
    stdio: 'inherit',
  })
  return result.status === 0
}

function done(via) {
  log(`  wasm 引擎就位（${via}）：${artifact}`)
}

function main() {
  if (mode === 'off' || argv.includes('--no-wasm')) {
    log('  skip  XLSX_ENGINE_WASM=off：页面会回退「读整份字节」模式')
    return
  }
  if (!['auto', 'fetch', 'build'].includes(mode)) {
    throw new Error(`XLSX_ENGINE_WASM=${mode} 不认识（auto | fetch | build | off）`)
  }

  // 已就位且与固定版本一致 → 什么都不做（CI 里先跑过编译回退时，`pnpm build` 不该再编一遍）。
  const already = settled()
  const matchesPin = already ? (already.lock.sha256 ? already.sha === already.lock.sha256 : true) : false

  if (already && mode !== 'build' && matchesPin)
    return done(`已就位（${already.manifest.source?.type ?? 'unknown'}，与固定版本一致）`)

  if (mode !== 'build') {
    // 下载路径：命中内容寻址缓存时不联网，所以 auto 的日常成本≈一次 sha256。
    const ok = run('fetch-wasm.mjs', mode === 'auto' ? ['--quiet'] : requireWasm ? ['--require'] : [])
    if (ok && existsSync(artifact)) return done('fetch')
    if (mode === 'fetch')
      throw new Error('固定产物不可用（wasm.lock.json 的 url / 缓存 / --file 都没拿到）')
    if (already) {
      // 拿不到固定版本，但手上这份是自洽的（sha 与自己的 manifest 对得上）：先沿用，不白编 4 分钟。
      if (already.lock.sha256 && !matchesPin)
        log('  warn  现有产物与 wasm.lock.json 的固定版本不一致，先沿用（fetch --refresh 对齐）')
      return done(`已就位（${already.manifest.source?.type ?? 'unknown'}）`)
    }
    log('  fetch 不可用，回退到 Rust 源码编译')
  }

  const ok = run('build-wasm.mjs', requireWasm ? ['--require'] : [])
  if (ok && existsSync(artifact)) return done('build')
  throw new Error('WASM_UNAVAILABLE')
}

try {
  main()
} catch (error) {
  const message = error instanceof Error ? error.message : String(error)
  if (message === 'WASM_UNAVAILABLE') {
    if (!requireWasm) {
      console.warn('  warn  wasm 引擎不可用：页面将回退到「读整份字节」模式')
      process.exit(0)
    }
    const annotation = process.env['GITHUB_ACTIONS'] === 'true' ? '::error::' : '✗ '
    console.error(`${annotation}wasm 引擎缺失：下载和编译都没成功（见上面的日志）`)
    process.exit(1)
  }
  console.error(`✗ ensure-wasm 失败：${message}`)
  process.exit(1)
}
