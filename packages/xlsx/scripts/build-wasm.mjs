#!/usr/bin/env node
// ─── 把上游 Rust 引擎编成一份 wasm reactor（P1 基线） ─────────────────
//
// 设计约束（都来自评审结论，不要"优化"掉）：
//
//   1. **不污染 submodule**：源码先复制到 build/xlsx-engine-wasm/，所有改动
//      都发生在副本上。上游升级只需重跑，不需要 rebase 一堆 patch。
//   2. **补丁是锚点断言，不是模糊替换**：上游一改，构建立刻失败并指出缺哪个
//      锚点（失败要好过悄悄编出一个缺 ABI 的产物）。
//   3. **只编 lib**（`--lib`）：`src/main.rs` 是 stdin/stdout 服务，wasm 用不到，
//      也就绕开了它的 `mpsc` / reader 线程。
//   4. **产物要自证**：exports 必须含 `call/wasi_alloc/wasi_free/memory`，不得含
//      `_start`（那是 command module，`wasi.start` 只能跑一次）。
//
// 用法：
//   node scripts/build-wasm.mjs                        # 源码不在 → skip（exit 0）
//   node scripts/build-wasm.mjs --source <dir>          # 指定源码目录
//   node scripts/build-wasm.mjs --require               # 源码/cargo 缺失 → exit 1
//   node scripts/build-wasm.mjs --install-target        # rustup target add wasm32-wasip1
//   node scripts/build-wasm.mjs --json                  # 机读输出
//
// 产物：
//   dist/engine/xlsx-engine.wasm   单文件、平台无关
//   dist/engine/xlsx-engine.json   { target, features, size, sha256, protocolVersion }

import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { gzipSync } from 'node:zlib'
import { cp, mkdir, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const buildDir = join(root, 'build/xlsx-engine-wasm')
const target = 'wasm32-wasip1'
const features = ['single-thread']
const wasmModuleName = 'xlsx_sidecar.wasm'
// 实测（1.99.1 / 2024 edition / LTO）：raw ≈ 8.1 MB → gzip ≈ 2 MB。
// 分发走 zip，所以两个上限都断言，raw 只防跑偏。
const MAX_WASM_BYTES = 12 * 1024 * 1024
const MAX_WASM_GZIP_BYTES = 4 * 1024 * 1024
const REQUIRED_EXPORTS = ['memory', 'call', 'wasi_alloc', 'wasi_free', '_initialize']
const FORBIDDEN_EXPORTS = ['_start']

const argv = process.argv.slice(2)
const json = argv.includes('--json')
const requireAll = argv.includes('--require')
const installTarget = argv.includes('--install-target')
const argValue = (name, fallback = '') => {
  const index = argv.indexOf(name)
  const value = index >= 0 ? argv[index + 1] : undefined
  return value === undefined || value.startsWith('--') ? fallback : value
}

function log(message) {
  if (!json) console.log(message)
}

class Skip extends Error {}
class Fail extends Error {}
/** 源码存在但版本不满足 wasm 要求：默认降级为 skip，`--require` 时才算失败。 */
class Stale extends Fail {}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    encoding: 'utf8',
    stdio: options.quiet ? 'pipe' : 'inherit',
    ...options,
  })
  if (result.error) throw new Fail(`${command} 无法执行：${result.error.message}`)
  return result
}

// ─── 1. 源码定位 ───────────────────────────────────────────────────────

function sourceCandidates() {
  const explicit = argValue('--source')
  return [
    explicit,
    process.env['XLSX_PRO_ENGINE_SOURCE'],
    join(root, '../../vendor/genoffice/apps/sheets/native/xlsx-engine'),
    join(root, 'vendor/genoffice/apps/sheets/native/xlsx-engine'),
  ].filter((entry) => typeof entry === 'string' && entry.length > 0)
}

async function findSource() {
  for (const candidate of sourceCandidates()) {
    const manifest = join(candidate, 'Cargo.toml')
    if (existsSync(manifest)) return resolve(candidate)
  }
  return null
}

// ─── 2. 工具链 ─────────────────────────────────────────────────────────

function toolchain() {
  const cargo = run('cargo', ['--version'], { quiet: true })
  if (cargo.status !== 0) throw new Fail('cargo 不可用（需要 Rust 工具链）')
  const rustc = run('rustc', ['--version'], { quiet: true })
  const version = (cargo.stdout ?? '').match(/(\d+)\.(\d+)/)
  const major = Number(version?.[1] ?? 0)
  const minor = Number(version?.[2] ?? 0)
  // edition 2024 需要 1.85+；`#[unsafe(no_mangle)]` 同一条线
  if (major < 1 || (major === 1 && minor < 85))
    throw new Fail(`Rust ${major}.${minor} 太旧：edition 2024 需要 1.85+（${(rustc.stdout ?? '').trim()}）`)
  return { cargo: (cargo.stdout ?? '').trim(), rustc: (rustc.stdout ?? '').trim() }
}

function ensureTarget() {
  const listed = run('rustup', ['target', 'list', '--installed'], { quiet: true })
  if (listed.status !== 0) throw new Fail('rustup 不可用')
  if ((listed.stdout ?? '').split('\n').some((line) => line.trim() === target)) return
  if (!installTarget) throw new Fail(`缺少 target：${target}（跑 --install-target 或 rustup target add ${target}）`)
  log(`  安装 target ${target} …`)
  const added = run('rustup', ['target', 'add', target], { quiet: true })
  if (added.status !== 0) throw new Fail(`rustup target add ${target} 失败：${added.stderr ?? ''}`)
}

// ─── 2.5 引擎依赖预检（wasm 要求纯 Rust 依赖） ────────────────────────
//
// 实测数据：ironcalc 0.7.1 → zip 0.6.6 → bzip2 0.4 → bzip2-sys（C 源码），
// wasm32-wasip1 下要 clang 才能编；而 ironcalc 0.9.1 / zip 9 的依赖树是纯 Rust，
// 一条 `cargo build --target wasm32-wasip1` 直接过。
//
// 所以这里提前判：版本不够就报清楚，别让使用者对着 "failed to find tool clang" 发呆。
const PURE_RUST_FLOOR = { zip: 9, ironcalc: 9, requires: 'cfb' }

async function preflightEngineDeps(source) {
  const cargoToml = await readFile(join(source, 'Cargo.toml'), 'utf8')
  const zipVersion = Number((cargoToml.match(/^zip\s*=\s*\{[^}]*version\s*=\s*"(\d+)/m) ?? [])[1] ?? 0)
  const ironcalcVersion = Number(
    (cargoToml.match(/^ironcalc\s*=\s*"(\d+)\.(\d+)/m) ?? [])[2] ?? 0,
  )
  const hasCfb = /^cfb\s*=/m.test(cargoToml)
  const problems = []
  if (zipVersion < PURE_RUST_FLOOR.zip) problems.push(`zip ${zipVersion || '?'} < ${PURE_RUST_FLOOR.zip}`)
  if (ironcalcVersion < PURE_RUST_FLOOR.ironcalc)
    problems.push(`ironcalc 0.${ironcalcVersion || '?'} < 0.${PURE_RUST_FLOOR.ironcalc}`)
  if (!hasCfb) problems.push('缺少 cfb（没有 .xls 支持，通常意味着源码太旧）')
  if (problems.length === 0) return
  throw new Stale(
    [
      `引擎源码太旧，wasm 编不过：${problems.join('；')}`,
      '',
      '  原因：ironcalc 0.7 依赖 zip 0.6，而 zip 0.6 默认带 bzip2 → bzip2-sys（C 源码），',
      '  wasm32-wasip1 下需要 clang。ironcalc 0.9 / zip 9 / cfb 的依赖树才是纯 Rust。',
      '',
      '  可用源码（任选其一）：',
      '    · 把 vendor/genoffice 的 apps/sheets/native/xlsx-engine 升到 ≥ f6c2358c',
      '      （注意：整包升级会牵动 fork 的 i18n，见 docs/xlsx-migration.md §2.3）',
      '    · 或指向任意满足上述版本的引擎 checkout：',
      '        XLSX_PRO_ENGINE_SOURCE=/path/to/xlsx-engine node scripts/build-wasm.mjs',
    ].join('\n'),
  )
}

// ─── 3. 复制 + 覆盖（锚点断言） ────────────────────────────────────────

async function copySource(source) {
  await mkdir(buildDir, { recursive: true })
  const KEEP = new Set(['target']) // cargo 增量缓存：留着，重复构建从 4 分钟降到秒级
  const SKIP = new Set(['target', '.git', 'tests'])
  for (const entry of await readdir(buildDir, { withFileTypes: true })) {
    if (KEEP.has(entry.name)) continue
    await rm(join(buildDir, entry.name), { recursive: true, force: true })
  }
  for (const entry of await readdir(source, { withFileTypes: true })) {
    if (SKIP.has(entry.name)) continue
    await cp(join(source, entry.name), join(buildDir, entry.name), { recursive: true })
  }
  log(`  源码副本 ${buildDir}`)
}

/** 锚点断言：不存在就报出确切的文件与期望文本。 */
function assertContains(label, text, anchor) {
  if (!text.includes(anchor)) throw new Fail(`${label} 缺少锚点：${JSON.stringify(anchor.slice(0, 60))}`)
}

async function applyOverlays() {
  // (a) Cargo.toml：加 [lib] crate-type —— 默认只有 rlib，产不出 cdylib。
  const cargoPath = join(buildDir, 'Cargo.toml')
  let cargoToml = await readFile(cargoPath, 'utf8')
  assertContains('Cargo.toml', cargoToml, '[package]')
  assertContains('Cargo.toml', cargoToml, 'single-thread = []')
  // (a2) release profile：wasm 分发体积敏感（原始 8.1 MB → gzip 2 MB）
  if (!/^\[profile\.release\]/m.test(cargoToml)) {
    cargoToml +=
      `\n# 由 scripts/build-wasm.mjs 注入：wasm 体积/速度取舍\n` +
      `[profile.release]\n` +
      `lto = true\n` +
      `codegen-units = 1\n` +
      `panic = "abort"\n` +
      `strip = true\n`
    await writeFile(cargoPath, cargoToml)
    log('  overlay Cargo.toml: [profile.release] lto/cgu=1/panic=abort/strip')
  }

  if (!/^\[lib\]/m.test(cargoToml)) {
    cargoToml = `# 由 scripts/build-wasm.mjs 注入：wasm reactor 需要 cdylib\n[lib]\ncrate-type = ["rlib", "cdylib"]\n\n${cargoToml}`
    await writeFile(cargoPath, cargoToml)
    log('  overlay Cargo.toml: [lib] crate-type += cdylib')
  }

  // (b) lib.rs：挂上 wasm 模块。
  const libPath = join(buildDir, 'src/lib.rs')
  let libRs = await readFile(libPath, 'utf8')
  assertContains('src/lib.rs', libRs, 'pub mod protocol;')
  assertContains('src/lib.rs', libRs, 'pub mod find;')
  if (!/mod wasm;/.test(libRs)) {
    libRs += `\n// 由 scripts/build-wasm.mjs 注入：WASI reactor ABI（见 rust/xlsx-engine/wasm.rs）\n#[cfg(target_os = "wasi")]\nmod wasm;\n`
    await writeFile(libPath, libRs)
    log('  overlay src/lib.rs: + mod wasm')
  }

  // (c) src/wasm.rs：仓库内的 ABI 层。
  const abi = await readFile(join(root, 'rust/xlsx-engine/wasm.rs'), 'utf8')
  assertContains('rust/xlsx-engine/wasm.rs', abi, 'pub unsafe extern "C" fn call')
  assertContains('rust/xlsx-engine/wasm.rs', abi, 'Engine::new')
  const libProtocol = await readFile(join(buildDir, 'src/protocol.rs'), 'utf8')
  assertContains('src/protocol.rs', libProtocol, 'pub fn handle_line')
  assertContains('src/protocol.rs', libProtocol, 'pub const PROTOCOL_VERSION')
  await writeFile(join(buildDir, 'src/wasm.rs'), abi)
  log('  overlay src/wasm.rs: `node:wasi` ABI')
}

// ─── 4. 编译 ───────────────────────────────────────────────────────────

async function compile() {
  const started = Date.now()
  const result = spawnSync(
    'cargo',
    [
      'build',
      '--release',
      '--lib',
      '--target',
      target,
      ...features.flatMap((feature) => ['--features', feature]),
      '--manifest-path',
      join(buildDir, 'Cargo.toml'),
      '--target-dir',
      join(buildDir, 'target'),
    ],
    { cwd: buildDir, stdio: 'inherit', env: { ...process.env, CARGO_TERM_COLOR: 'never' } },
  )
  if (result.status !== 0) throw new Fail(`cargo build 失败（exit ${result.status}）`)
  return Date.now() - started
}

// ─── 5. 产物自证 + 发布 ────────────────────────────────────────────────

async function publish() {
  const artifact = join(buildDir, 'target', target, 'release', wasmModuleName)
  if (!existsSync(artifact)) throw new Fail(`找不到产物：${artifact}`)
  const bytes = await readFile(artifact)
  const info = await stat(artifact)
  if (info.size > MAX_WASM_BYTES)
    throw new Fail(`wasm 体积 ${(info.size / 1048576).toFixed(2)} MB 超过 ${MAX_WASM_BYTES / 1048576} MB 预算`)

  const module = new WebAssembly.Module(bytes)
  const exports = WebAssembly.Module.exports(module).map((entry) => entry.name)
  for (const name of REQUIRED_EXPORTS)
    if (!exports.includes(name)) throw new Fail(`产物缺少导出：${name}（现有：${exports.join(',')}）`)
  for (const name of FORBIDDEN_EXPORTS)
    if (exports.includes(name))
      throw new Fail(
        `产物含 ${name}：这是 command module（wasi.start 只能跑一次），reactor 不能有它`,
      )

  const gzipped = gzipSync(bytes, { level: 9 }).byteLength
  if (gzipped > MAX_WASM_GZIP_BYTES)
    throw new Fail(
      `wasm gzip ${(gzipped / 1048576).toFixed(2)} MB 超过 ${MAX_WASM_GZIP_BYTES / 1048576} MB 预算`,
    )
  const sha256 = createHash('sha256').update(bytes).digest('hex')
  const outFile = resolve(argValue('--out', join(root, 'dist/xlsx-engine.wasm')))
  await mkdir(dirname(outFile), { recursive: true })
  await writeFile(outFile, bytes)
  const manifest = {
    target,
    features,
    size: info.size,
    sha256,
    exports,
    builtAt: new Date().toISOString(),
    // 桥启动时会调用它做版本断言；缺失说明 wasm.rs 没生效
    protocolVersion: exports.includes('xlsx_protocol_version') ? 'runtime' : 'missing',
  }
  await writeFile(`${outFile.replace(/\.wasm$/, '')}.json`, `${JSON.stringify(manifest, null, 2)}\n`)
  return { outFile, size: info.size, gzipped, sha256, exports }
}

// ─── 入口 ──────────────────────────────────────────────────────────────

async function main() {
  const source = await findSource()
  if (!source) {
    const message = '未找到 Rust 源码（apps/sheets/native/xlsx-engine 未 vendor）——跳过 wasm 构建'
    if (requireAll) throw new Fail(message)
    log(`skip  ${message}`)
    if (json) console.log(JSON.stringify({ ok: true, skipped: true }, null, 2))
    return
  }

  log(`# build-wasm (${target}, features=${features.join('+')})`)
  log(`  source ${source}`)
  try {
    await preflightEngineDeps(source)
  } catch (error) {
    // 版本不满足时：默认 skip（不然所有人的 `pnpm build` 都会红），
    // 但把原因与两条出路打出来；`--require` 下才算失败。
    if (!(error instanceof Stale) || requireAll) throw error
    log(`skip  ${error.message.split('\n')[0]}`)
    log('      原因与出路见 docs/xlsx-migration.md §2.3；本次 exec 走 TS 回退底座')
    if (json) console.log(JSON.stringify({ ok: true, skipped: true, reason: 'engine-source-too-old' }, null, 2))
    return
  }
  const tool = toolchain()
  log(`  ${tool.cargo} · ${tool.rustc}`)
  ensureTarget()

  await copySource(source)
  await applyOverlays()
  const compileMs = await compile()
  const published = await publish()

  log('')
  log(`  ok   ${published.outFile}`)
  log(
    `       ${(published.size / 1048576).toFixed(2)} MB raw · ${(published.gzipped / 1048576).toFixed(2)} MB gzip · sha256 ${published.sha256.slice(0, 16)}… · 编译 ${(compileMs / 1000).toFixed(1)} s`,
  )
  log(`       exports ${published.exports.join(',')}`)
  if (json)
    console.log(
      JSON.stringify(
        { ok: true, source, compileMs, outFile: published.outFile, size: published.size, sha256: published.sha256, exports: published.exports },
        null,
        2,
      ),
    )
}

main().catch((error) => {
  if (error instanceof Skip) {
    log(`skip  ${error.message}`)
    if (json) console.log(JSON.stringify({ ok: true, skipped: true, reason: error.message }, null, 2))
    return
  }
  const message = error instanceof Error ? error.message : String(error)
  console.error(`\n✗ build-wasm 失败：${message}`)
  if (json) console.log(JSON.stringify({ ok: false, error: message }, null, 2))
  process.exitCode = 1
})
