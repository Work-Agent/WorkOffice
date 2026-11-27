/**
 * 装载前自检：
 *   1. 清单（id / claims / pages.content / exports['./workstudio']）
 *   2. commands.json ↔ dist/engine.js 的动态命令一一对齐
 *   3. 单文件范围守卫（目录 / 非 file:// / 错误扩展名）
 *   4. 页面通道的只读 op 与租约令牌
 *   5. 源码里不得再出现旧 URA 面（window.extension / ura.page / agent-request）
 */

import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'))
const manifest = pkg.workstudio ?? {}

const checks = []
function check(name, run) {
  checks.push({ name, run })
}

check('清单字段', () => {
  assert.equal(manifest.id, 'xlsx')
  assert.ok(manifest.label, '缺少 label')
  assert.ok(manifest.pages?.content, '缺少 pages.content')
  assert.ok(existsSync(join(root, manifest.pages.content)), `页面不存在：${manifest.pages.content}`)
  const entry = pkg.exports?.['./workstudio']
  assert.ok(entry, "缺少 exports['./workstudio']")
  assert.ok(existsSync(join(root, entry)), `入口不存在：${entry}`)
  assert.ok(existsSync(join(root, 'commands.json')), '缺少 commands.json')
  assert.deepEqual(
    manifest.claims?.map((claim) => claim.extensions).flat().sort(),
    ['xlsm', 'xlsx', 'xltx'],
  )
})

const module = await import(pathToFileURL(join(root, 'dist/engine.js')).href)
const plugin = module.workstudio
assert.ok(plugin, 'dist/engine.js 未导出 workstudio')

// 引擎命令依赖插件作用域（wasm 会话落在 scope.tempDir）；按宿主顺序先激活。
const { mkdtemp, rm } = await import('node:fs/promises')
const { tmpdir } = await import('node:os')
const scopeDir = await mkdtemp(join(tmpdir(), 'xlsx-verify-'))
plugin.lifecycle?.activate?.({
  pluginId: 'xlsx',
  tempDir: scopeDir,
  dataDir: scopeDir,
  files: { resolveSave: () => join(scopeDir, 'unused.xlsx') },
})

check('命令对齐', () => {
  const declared = JSON.parse(readFileSync(join(root, 'commands.json'), 'utf8')).commands
  const dynamic = Object.keys(plugin.commands).sort()
  assert.deepEqual(declared.map((item) => item.name).sort(), dynamic)
  for (const item of declared) {
    assert.deepEqual(
      Object.keys(item.args ?? {}).sort(),
      Object.keys(plugin.commands[item.name].args ?? {}).sort(),
      `${item.name} 的 args 与静态描述不一致`,
    )
    assert.ok(item.name.startsWith(`${manifest.id}.`), `命令前缀错误：${item.name}`)
  }
})

check('宿主参数校验兼容', () => {
  // 宿主按 `typeof` 校验声明过的参数：结构化/二进制载荷一律不能声明，
  // 否则页面传对象（或 Uint8Array）会得到「参数类型非法」。
  const declared = new Map(
    JSON.parse(readFileSync(join(root, 'commands.json'), 'utf8')).commands.map((entry) => [
      entry.name,
      entry.args ?? {},
    ]),
  )
  for (const [name, key] of [
    ['xlsx.page', 'input'],
    ['xlsx.apply', 'commands'],
    ['xlsx.create', 'sheets'],
    ['xlsx.create', 'rows'],
    ['xlsx.create', 'csv'],
  ]) {
    assert.ok(!(key in (declared.get(name) ?? {})), `${name} 不应声明 ${key}`)
  }
  for (const entry of declared.values())
    for (const [key, spec] of Object.entries(entry))
      assert.ok(spec && typeof spec === 'object', `参数描述无效：${key}`)
})

check('单文件守卫', async () => {
  const adapter = plugin.resources.xlsx
  assert.deepEqual(await adapter.ls('file:///w'), [])
  assert.deepEqual(await adapter.find('file:///w', 'x'), [])
  await assert.rejects(() => adapter.read('https://example.com/a.xlsx'), /file:|\/\/|单个文件/)
  await assert.rejects(() => adapter.read(pathToFileURL(root).href), /单个文件|扩展名|不存在/)
  await assert.rejects(() => adapter.read(pathToFileURL(join(root, 'package.json')).href), /扩展名|单个文件|只支持/)
})

check('页面通道只读 op', async () => {
  const page = plugin.commands['xlsx.page']
  assert.ok(page, '缺少 xlsx.page')
  assert.equal(page.expose?.agent, false)
  const ref = pathToFileURL(join(root, 'assets/seed.xlsx')).href
  const call = (op, input = {}) =>
    page.run({ ref, pageId: 'verify', op, input: JSON.stringify(input) }, {})

  const status = await call('status')
  assert.ok(status.size > 0)
  assert.match(status.hash, /^[0-9a-f]{64}$/)

  const meta = await call('meta')
  assert.ok(meta.sheets.length >= 1)
  const sheetId = meta.sheets[0].id

  if (meta.source === 'wasm') {
    const range = await call('range', {
      sheetId,
      range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 },
    })
    assert.equal(range.rows.length, 3)
    assert.ok(Array.isArray(range.cells))
  } else {
    await assert.rejects(() => call('range', { sheetId, range: 'A1:C3' }), /wasm|unsupported/)
  }

  const attached = await call('attach', { pageId: 'verify' })
  assert.match(attached.token, /^[0-9a-f-]{36}$/)
  const again = await page
    .run({ ref, pageId: 'other', op: 'attach', input: '{}' }, {})
    .then(() => null, (error) => error)
  assert.ok(again, '第二个 pageId 未带 takeover 时应被拒')

  await assert.rejects(
    () => page.run({ ref, pageId: 'verify', op: 'report', input: JSON.stringify({ token: 'bad' }) }, {}),
    /令牌|denied|token/i,
  )
  await call('report', { token: attached.token, docRev: 1 })
  await call('detach', { token: attached.token })
})

check('wasm 产物', () => {
  const wasm = join(root, 'dist/xlsx-engine.wasm')
  if (!existsSync(wasm)) {
    assert.ok(
      process.env['XLSX_REQUIRE_WASM'] !== '1',
      '缺少 dist/xlsx-engine.wasm：XLSX_REQUIRE_WASM=1 时必须先 node scripts/fetch-wasm.mjs',
    )
    console.log('  · 跳过：dist/xlsx-engine.wasm 未构建（node scripts/fetch-wasm.mjs）')
    return
  }
  const manifest = JSON.parse(readFileSync(join(root, 'dist/xlsx-engine.json'), 'utf8'))
  const bytes = readFileSync(wasm)
  assert.equal(manifest.size, bytes.byteLength)
  assert.match(manifest.sha256, /^[0-9a-f]{64}$/)
  assert.equal(
    manifest.sha256,
    createHash('sha256').update(bytes).digest('hex'),
    'wasm 与 dist/xlsx-engine.json 的 sha256 不一致（node scripts/fetch-wasm.mjs --refresh）',
  )
  // 固定版本：产物必须就是 wasm.lock.json 里那一份，避免别人换了个 wasm 混进包。
  // 例外：本次是自己编的（source=built），字节随编译器浮动，只校验自洽性。
  const lockPath = join(root, 'wasm.lock.json')
  if (existsSync(lockPath) && manifest.source?.type !== 'built') {
    const lock = JSON.parse(readFileSync(lockPath, 'utf8'))
    if (lock.sha256) assert.equal(manifest.sha256, lock.sha256, 'wasm 与 wasm.lock.json 的固定 sha256 不一致')
    if (lock.size) assert.equal(manifest.size, lock.size, 'wasm 与 wasm.lock.json 的固定体积不一致')
  }
  for (const name of ['memory', '_initialize', 'call', 'wasi_alloc', 'wasi_free'])
    assert.ok(manifest.exports.includes(name), `缺少导出：${name}`)
})

check('无旧 URA 面', () => {
  const files = []
  const walk = (dir) => {
    for (const name of readdirSync(dir)) {
      const path = join(dir, name)
      if (statSync(path).isDirectory()) walk(path)
      else if (/\.(ts|tsx|js|mjs)$/.test(name)) files.push(path)
    }
  }
  walk(join(root, 'src'))
  const forbidden = [
    ['window.extension', /window\.extension/],
    ['ura.page', /ura\.page/],
    ['agent-request', /agent-request/],
    ['registerCodec', /registerCodec/],
  ]
  for (const file of files) {
    const text = readFileSync(file, 'utf8')
    for (const [name, pattern] of forbidden) {
      assert.ok(!pattern.test(text), `${file.replace(root, '.')} 仍含 ${name}`)
    }
  }
})

let failed = 0
for (const { name, run } of checks) {
  try {
    await run()
    console.log(`✓ ${name}`)
  } catch (error) {
    failed += 1
    console.error(`✗ ${name}\n  ${error instanceof Error ? error.message : String(error)}`)
  }
}
await rm(scopeDir, { recursive: true, force: true }).catch(() => undefined)
if (failed > 0) {
  console.error(`\n${failed} 项未通过`)
  process.exitCode = 1
} else {
  console.log('\n全部通过')
}
