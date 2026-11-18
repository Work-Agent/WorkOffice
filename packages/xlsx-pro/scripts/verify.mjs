#!/usr/bin/env node
// 装载前自检：清单 / 产物 / 工具声明 / 执行点 ABI / 一次真实往返。

import { readFile, stat } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { strFromU8, unzipSync } from 'fflate'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const problems = []
const warnings = []
const fail = (message) => problems.push(message)
const warn = (message) => warnings.push(message)

const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))
const ws = pkg.workstudio ?? {}

if (ws.id !== 'xlsx-pro') fail(`workstudio.id 必须是 xlsx-pro（当前 ${ws.id}）`)
if (ws.apiVersion !== 2) fail(`apiVersion 必须是 2（当前 ${ws.apiVersion}）`)
if (ws.write !== 'plan') fail(`write 必须是 plan（当前 ${ws.write}）`)
if (!ws.engine?.exec) fail('write: plan 必须声明 engine.exec')
if (Boolean(ws.engine?.plan) !== Boolean(ws.engine?.exec)) fail('engine.plan 与 engine.exec 必须同时出现')
if (ws.content === 'pulled' && !ws.page) fail("content: 'pulled' 必须声明 page")
if (ws.model === 'resident' && ws.write !== 'plan') fail("model: 'resident' 需要 write: 'plan'")

const ura = pkg.exports?.['./ura']
if (!ura) fail('exports["./ura"] 缺失')
for (const [name, file] of Object.entries({ ura, page: ws.page, ...(ws.engine ?? {}) })) {
  if (typeof file !== 'string') {
    fail(`${name}: 入口不是字符串`)
    continue
  }
  const info = await stat(join(root, file)).catch(() => null)
  if (!info?.isFile()) fail(`${name}: 文件不存在 ${file}`)
}
{
  const info = await stat(join(root, 'dist/seed.xlsx')).catch(() => null)
  if (!info?.isFile()) fail('dist/seed.xlsx 缺失（新建模板）')
}
for (const skill of ws.skills ?? []) {
  if (!skill?.name || !skill?.dir) {
    fail('skills[] 缺少 name / dir')
    continue
  }
  const info = await stat(join(root, skill.dir)).catch(() => null)
  if (!info?.isDirectory()) fail(`skills: 目录不存在 ${skill.dir}`)
}

const load = (file) => import(pathToFileURL(join(root, file)).href)
const { tools } = await load(ws.engine.tools)
if (!Array.isArray(tools) || tools.length === 0) fail('engine.tools 没有导出 tools 数组')
const NAME = /^[a-z][a-z0-9._-]*$/i
const seen = new Set()
for (const decl of tools ?? []) {
  const at = decl?.name ?? '(匿名)'
  if (!decl?.name || !NAME.test(decl.name)) fail(`工具名非法：${at}`)
  if (!decl?.label) fail(`工具缺少 label：${at}`)
  if (!decl?.description) fail(`工具缺少 description：${at}`)
  if (!Array.isArray(decl?.kinds) || decl.kinds.length === 0) fail(`工具缺少 kinds：${at}`)
  if (!decl?.parameters || typeof decl.parameters !== 'object') fail(`工具缺少 parameters：${at}`)
  if (typeof decl?.readOnly !== 'boolean') fail(`工具缺少 readOnly：${at}`)
  if (seen.has(decl?.name)) fail(`工具名重复：${at}`)
  seen.add(decl?.name)
}
for (const name of ['xlsx_guide', 'xlsx_read', 'xlsx_create', 'xlsx_apply'])
  if (!seen.has(name)) fail(`四工具面缺少 ${name}`)
if (seen.size !== 4) fail(`四工具面应只有 4 条声明（实际 ${seen.size}）`)
if ((tools ?? []).filter((d) => d.readOnly).length !== 2) fail('四工具面应恰好 2 个只读工具')
if ((tools ?? []).filter((d) => !d.readOnly).length !== 2) fail('四工具面应恰好 2 个写工具')

const planMod = await load(ws.engine.plan)
const execMod = await load(ws.engine.exec)
const readMod = await load(ws.engine.read)
const runtimeMod = await load(ws.engine.runtime)
for (const [name, mod, exports] of [
  ['plan', planMod, ['plan']],
  ['exec', execMod, ['exec']],
  ['read', readMod, ['project', 'query']],
  ['runtime', runtimeMod, ['runtime']],
]) {
  for (const key of exports)
    if (typeof mod[key] !== 'function') fail(`执行点 ${name} 缺少导出 ${key}`)
}

const report = await runtimeMod.runtime({
  readBytes: async () => ({ bytes: new Uint8Array(), version: '' }),
  tempDir: '',
  importEpoch: 0,
  budget: { maxBytes: 1, deadlineMs: 1 },
})
if (report?.write !== 'plan') fail(`runtime.write 应为 plan（当前 ${report?.write}）`)
if (report?.features?.undo !== 'page') fail('runtime.features.undo 应为 page')
if (report?.features?.tokens?.inPage !== 'docRev+structureEpoch')
  fail('runtime.features.tokens.inPage 应为 docRev+structureEpoch')

const seed = new Uint8Array(await readFile(join(root, 'dist/seed.xlsx')))
const ctx = {
  readBytes: async () => ({ bytes: seed, version: 'v1' }),
  tempDir: '',
  importEpoch: 0,
  budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 10_000 },
}
const REF = 'verify://fixture.xlsx'

const planned = await planMod.plan(
  {
    tool: 'xlsx_apply',
    params: {
      ref: REF,
      commands: [
        {
          op: 'set_range',
          sheet: 'Sheet1',
          start: 'A1',
          values: [
            ['姓名', '金额'],
            ['张三', 1200],
          ],
        },
      ],
    },
    ref: REF,
    baseVersion: 'v1',
    bytes: seed,
  },
  ctx,
)
if (planned?.plan?.kind !== 'save') fail(`xlsx_apply 没有产出 save 计划：${planned?.plan?.kind}`)
const produced = await execMod.exec(
  { ref: REF, baseVersion: 'v1', plan: planned.plan, bytes: seed },
  ctx,
)
if (!(produced?.bytes instanceof Uint8Array)) fail('exec 没有产出字节')
else {
  const before = unzipSync(seed)
  const after = unzipSync(produced.bytes)
  const sheet = strFromU8(after['xl/worksheets/sheet1.xml'])
  if (!sheet.includes('张三') || !sheet.includes('1200')) fail('写回结果缺少新单元格')
  for (const name of Object.keys(before)) {
    // sheet1.xml 是本次写入目标；workbook.xml 因 calcPr@fullCalcOnLoad 必然变化
    // （设计稿 §6.1：这是 touchedEntries 的合法成员，不是保真泄漏）。
    if (name === 'xl/worksheets/sheet1.xml' || name === 'xl/workbook.xml') continue
    const same = before[name] && after[name] && Buffer.from(before[name]).equals(Buffer.from(after[name]))
    if (!same) warn(`未涉及条目被重写：${name}`)
  }
}

const guide = await planMod.plan(
  { tool: 'xlsx_guide', params: { ref: REF }, ref: REF, bytes: seed },
  ctx,
)
if (!guide?.summary?.includes('insert_rows')) fail('xlsx_guide 没有返回 structure 目录')
for (const group of ['cells', 'structure', 'sheets', 'objects'])
  if (!guide?.summary?.includes(group)) fail(`xlsx_guide 缺少分组 ${group}`)

const read = await planMod.plan(
  { tool: 'xlsx_read', params: { ref: REF, what: 'context' }, ref: REF, bytes: seed },
  ctx,
)
if (!read?.summary?.includes('Sheet1')) fail('xlsx_read 没有返回工作表投影')

const blocked = await planMod
  .plan(
    {
      tool: 'xlsx_apply',
      params: { ref: REF, commands: [{ op: 'add_chart', sheet: 'Sheet1' }] },
      ref: REF,
      bytes: seed,
    },
    ctx,
  )
  .then(() => null)
  .catch((error) => error)
if (blocked?.code !== 'command.page-unavailable')
  fail(`无页面写 objects 组应回 command.page-unavailable（当前 ${blocked?.code ?? '未报错'}）`)

const created = await planMod.plan(
  {
    tool: 'xlsx_create',
    params: { ref: 'verify://new.xlsx', sheets: [{ name: '数据', rows: [['a', 1], ['b', 2]] }] },
    ref: 'verify://new.xlsx',
    missing: true,
  },
  ctx,
)
if (created?.plan?.kind !== 'new') fail(`xlsx_create 没有产出 new 计划：${created?.plan?.kind}`)
const newBytes = await execMod.exec(
  { ref: 'verify://new.xlsx', baseVersion: '', plan: created.plan, missing: true },
  ctx,
)
if (!(newBytes?.bytes instanceof Uint8Array)) fail('xlsx_create 的 exec 没有产出字节')
else {
  const readNew = await planMod.plan(
    {
      tool: 'xlsx_read',
      params: { ref: 'verify://new.xlsx', what: 'range', sheet: '数据', range: 'A1:B2' },
      ref: 'verify://new.xlsx',
      bytes: newBytes.bytes,
    },
    ctx,
  )
  if (!readNew?.summary?.includes('a') || !readNew?.summary?.includes('2'))
    fail('新建工作簿的内容不可读')
}

{
  // I17 溯源：页面 fork 与 wasm 引擎必须来自同一个 vendor commit。
  // FORK.md 是唯一事实源；vendor 不存在（例如已打包的插件）时跳过。
  const forkFile = join(root, 'src/editor/FORK.md')
  const vendorDir = join(root, '../../vendor/genoffice')
  if (existsSync(forkFile) && existsSync(join(vendorDir, '.git'))) {
    const recorded = (await readFile(forkFile, 'utf8')).match(/^vendorCommit:\s*([0-9a-f]{7,40})/m)?.[1]
    if (!recorded) {
      fail('src/editor/FORK.md 缺少 vendorCommit（I17 溯源）')
    } else {
      const { execFileSync } = await import('node:child_process')
      const head = execFileSync('git', ['-C', vendorDir, 'rev-parse', 'HEAD'], { encoding: 'utf8' }).trim()
      if (!head.startsWith(recorded))
        fail(
          `页面 fork 记录的 vendorCommit=${recorded}，但 vendor HEAD=${head.slice(0, 12)}` +
            '（升级 vendor 后请同步 src/editor/FORK.md，见 §I17）',
        )
    }
  }
}

{
  // wasm 引擎产物：存在就断言"能加载 + 导出面 + 清单自洽"，不存在只 warn
  // （源码未 vendor 时允许走 TS 回退路径，这不算坏产物）。
  const wasmPath = join(root, 'dist/engine/xlsx-engine.wasm')
  if (!existsSync(wasmPath)) {
    warn('未构建 wasm 引擎（dist/engine/xlsx-engine.wasm）——exec 走 TS 回退路径')
  } else {
    const bytes = await readFile(wasmPath)
    const manifestPath = join(root, 'dist/engine/xlsx-engine.json')
    if (!existsSync(manifestPath)) {
      fail('缺少 dist/engine/xlsx-engine.json（构建脚本没跑完？）')
    } else {
      const manifest = JSON.parse(await readFile(manifestPath, 'utf8'))
      const sha256 = createHash('sha256').update(bytes).digest('hex')
      if (manifest.sha256 !== sha256) fail('xlsx-engine.wasm 与清单 sha256 不一致')
      if (manifest.size !== bytes.byteLength)
        fail(`xlsx-engine.wasm 体积与清单不一致（${bytes.byteLength} vs ${manifest.size}）`)
      if (manifest.target !== 'wasm32-wasip1') fail(`wasm target 应为 wasm32-wasip1（当前 ${manifest.target}）`)
    }
    let exports = []
    try {
      exports = WebAssembly.Module.exports(new WebAssembly.Module(bytes)).map((entry) => entry.name)
    } catch (error) {
      fail(`xlsx-engine.wasm 不是合法模块：${error instanceof Error ? error.message : String(error)}`)
    }
    for (const name of ['memory', '_initialize', 'call', 'wasi_alloc', 'wasi_free'])
      if (!exports.includes(name)) fail(`xlsx-engine.wasm 缺少导出：${name}`)
    if (exports.includes('_start')) fail('xlsx-engine.wasm 是 command module（含 _start），reactor 不能有它')
  }
}

for (const message of warnings) console.warn(`warn  ${message}`)
if (problems.length > 0) {
  for (const message of problems) console.error(`error ${message}`)
  console.error(`\nverify failed: ${problems.length} 个问题`)
  process.exit(1)
}
console.log(
  `verify ok · ${pkg.name}@${pkg.version} · tools=${seen.size} · ` +
    `read=${(tools ?? []).filter((d) => d.readOnly).length} · ` +
    `write=${(tools ?? []).filter((d) => !d.readOnly).length}`,
)
