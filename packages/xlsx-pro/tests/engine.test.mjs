import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { strFromU8, unzipSync } from 'fflate'

const { tools } = await import('../dist/engine/tools.mjs')
const planMod = await import('../dist/engine/plan.mjs')
const execMod = await import('../dist/engine/exec.mjs')
const runtimeMod = await import('../dist/engine/runtime.mjs')
const readMod = await import('../dist/engine/read.mjs')

const SEED = new Uint8Array(readFileSync(new URL('../dist/seed.xlsx', import.meta.url)))
const REF = 'verify://fixture.xlsx'
const BUDGET = { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 }
const ctxFor = (bytes = SEED) => ({
  readBytes: async () => ({ bytes, version: 'v1' }),
  tempDir: '',
  importEpoch: 0,
  budget: BUDGET,
})

async function apply(commands, bytes = SEED) {
  const planned = await planMod.plan(
    { tool: 'xlsx_apply', params: { ref: REF, commands }, ref: REF, baseVersion: 'v1', bytes },
    ctxFor(bytes),
  )
  assert.equal(planned.plan.kind, 'save', planned.summary)
  const outcome = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan: planned.plan, bytes },
    ctxFor(bytes),
  )
  assert.ok(outcome.bytes instanceof Uint8Array)
  return outcome
}

async function read(bytes, params) {
  return planMod.plan({ tool: 'xlsx_read', params: { ref: REF, ...params }, ref: REF, bytes }, ctxFor(bytes))
}

test('工具声明：4 个工具 / 2 读 2 写', () => {
  assert.deepEqual(tools.map((tool) => tool.name), ['xlsx_guide', 'xlsx_read', 'xlsx_create', 'xlsx_apply'])
  assert.equal(tools.filter((tool) => tool.readOnly).length, 2)
  assert.equal(tools.find((tool) => tool.name === 'xlsx_apply').requiresBaseVersion, true)
  assert.equal(tools.find((tool) => tool.name === 'xlsx_create').requiresBaseVersion, false)
})

test('guide：分组 / 单条 / 未知 op', async () => {
  const top = await planMod.plan({ tool: 'xlsx_guide', params: {}, ref: REF, bytes: SEED }, ctxFor())
  for (const group of ['cells', 'structure', 'sheets', 'objects']) assert.match(top.summary, new RegExp(group))
  const cells = await planMod.plan({ tool: 'xlsx_guide', params: { group: 'cells' }, ref: REF, bytes: SEED }, ctxFor())
  assert.match(cells.summary, /set_range/)
  const one = await planMod.plan({ tool: 'xlsx_guide', params: { op: 'insert_rows' }, ref: REF, bytes: SEED }, ctxFor())
  assert.match(one.summary, /schema:/)
  assert.match(one.summary, /example:/)
  const unknown = await planMod.plan({ tool: 'xlsx_guide', params: { op: 'nope' }, ref: REF, bytes: SEED }, ctxFor())
  assert.match(unknown.summary, /未知 op/)
})

test('xlsx_apply：set_range → save → exec → 读回', async () => {
  const outcome = await apply([
    { op: 'set_range', sheet: 'Sheet1', start: 'A1', values: [['姓名', '金额'], ['张三', 1200]] },
  ])
  const readBack = await read(outcome.bytes, { what: 'range', sheet: 'Sheet1', range: 'A1:B2' })
  assert.match(readBack.summary, /姓名/)
  assert.match(readBack.summary, /张三/)
  assert.equal(outcome.result.applied, 1)
  assert.equal(outcome.result.sync.kind, 'manifest')
})

test('结构命令与后续 cell 编辑按顺序生效', async () => {
  const outcome = await apply([
    { op: 'set_range', sheet: 'Sheet1', start: 'A1', values: [['h1', 'h2'], [1, 2], [3, 4]] },
    { op: 'insert_rows', sheet: 'Sheet1', row: 2, count: 1 },
    { op: 'set_cell', sheet: 'Sheet1', address: 'A2', value: 'new' },
  ])
  const readBack = await read(outcome.bytes, { what: 'range', sheet: 'Sheet1', range: 'A1:B5' })
  assert.match(readBack.summary, /new/)
  assert.match(readBack.summary, /h1/)
  assert.match(readBack.summary, /3/)
})

test('find_replace / fill_range / sort_range 读取同批前序写入', async () => {
  const replaced = await apply([
    { op: 'set_range', sheet: 'Sheet1', start: 'A1', values: [['a', 'b'], ['c', 'd']] },
    { op: 'find_replace', sheet: 'Sheet1', range: 'A1:B2', find: 'a', replace: 'z' },
  ])
  const replacedRead = await read(replaced.bytes, { what: 'range', sheet: 'Sheet1', range: 'A1:B2' })
  assert.match(replacedRead.summary, /z/)

  const filled = await apply([
    { op: 'set_range', sheet: 'Sheet1', start: 'A1', values: [[1], [2]] },
    { op: 'fill_range', sheet: 'Sheet1', source: 'A1:A2', target: 'A1:A6' },
  ])
  const filledRead = await read(filled.bytes, { what: 'range', sheet: 'Sheet1', range: 'A1:A6' })
  assert.match(filledRead.summary, /1/)
  assert.match(filledRead.summary, /2/)

  const sorted = await apply([
    { op: 'set_range', sheet: 'Sheet1', start: 'A1', values: [[3], [1], [2]] },
    { op: 'sort_range', sheet: 'Sheet1', range: 'A1:A3', byColumn: 'A', order: 'asc' },
  ])
  const sortedRead = await read(sorted.bytes, { what: 'range', sheet: 'Sheet1', range: 'A1:A3' })
  assert.match(sortedRead.summary, /1\t1/)
  assert.match(sortedRead.summary, /3\t3/)
})

test('rename_sheet 与 cell 编辑可同批', async () => {
  const outcome = await apply([
    { op: 'rename_sheet', sheet: 'Sheet1', name: '数据' },
    { op: 'set_range', sheet: '数据', start: 'A1', values: [['ok']] },
  ])
  const context = await read(outcome.bytes, { what: 'context' })
  assert.match(context.summary, /数据/)
  const range = await read(outcome.bytes, { what: 'range', sheet: '数据', range: 'A1' })
  assert.match(range.summary, /ok/)
})

test('add_sheet + set_range', async () => {
  const outcome = await apply([
    { op: 'add_sheet', name: '新表' },
    { op: 'set_range', sheet: '新表', start: 'A1', values: [['hello']] },
  ])
  const range = await read(outcome.bytes, { what: 'range', sheet: '新表', range: 'A1' })
  assert.match(range.summary, /hello/)
})

test('format_range 写入 styles.xml', async () => {
  const outcome = await apply([
    { op: 'set_range', sheet: 'Sheet1', start: 'A1', values: [['h']] },
    { op: 'format_range', sheet: 'Sheet1', range: 'A1', format: { bold: true, fillColor: '#DDEBF7' } },
  ])
  const files = unzipSync(outcome.bytes)
  const styles = strFromU8(files['xl/styles.xml'])
  assert.match(styles, /<b\/>/)
  assert.match(styles, /DDEBF7/)
})

test('objects 组无页面 → command.page-unavailable；未知 op → reject', async () => {
  await assert.rejects(
    planMod.plan(
      { tool: 'xlsx_apply', params: { ref: REF, commands: [{ op: 'add_chart', sheet: 'Sheet1' }] }, ref: REF, bytes: SEED },
      ctxFor(),
    ),
    (error) => error.code === 'command.page-unavailable',
  )
  const unknown = await planMod.plan(
    { tool: 'xlsx_apply', params: { ref: REF, commands: [{ op: 'nothing' }] }, ref: REF, bytes: SEED },
    ctxFor(),
  )
  assert.equal(unknown.plan.kind, 'reject')
  assert.equal(unknown.plan.rejected[0].code, 'command.not-found')
})

test('xlsx_create → exec → 读回', async () => {
  const created = await planMod.plan(
    { tool: 'xlsx_create', params: { ref: 'verify://new.xlsx', sheets: [{ name: '数据', rows: [['a', 1], ['b', 2]] }] }, ref: 'verify://new.xlsx', missing: true },
    ctxFor(),
  )
  assert.equal(created.plan.kind, 'new')
  const outcome = await execMod.exec(
    { ref: 'verify://new.xlsx', baseVersion: '', plan: created.plan, missing: true },
    ctxFor(),
  )
  const range = await read(outcome.bytes, { what: 'range', sheet: '数据', range: 'A1:B2' })
  assert.match(range.summary, /a/)
  assert.match(range.summary, /2/)
})

test('runtime：plan / undo=page / tokens', async () => {
  const report = await runtimeMod.runtime(ctxFor())
  assert.equal(report.write, 'plan')
  assert.equal(report.features.undo, 'page')
  // xlsx 的 A1 在值编辑下稳定，结构改动才失效 → docRev+structureEpoch（不用 epoch）。
  assert.equal(report.features.tokens.inPage, 'docRev+structureEpoch')
  assert.equal(report.features.tokens.commit, 'version')
})

test('只读工具返回 text 计划，不是 noop', async () => {
  const guide = await planMod.plan({ tool: 'xlsx_guide', params: {}, ref: REF, bytes: SEED }, ctxFor())
  assert.equal(guide.plan.kind, 'text')
  const contextRead = await read(SEED, { what: 'context' })
  assert.equal(contextRead.plan.kind, 'text')
  const outcome = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan: contextRead.plan, bytes: SEED },
    ctxFor(),
  )
  assert.equal(outcome.bytes, undefined)
  assert.match(String(outcome.result.audit.text), /sheets=/)
})

test('guide：op 级坐标层级（L0 / L1 / L2）', async () => {
  const l2 = await planMod.plan({ tool: 'xlsx_guide', params: { op: 'set_cell' }, ref: REF, bytes: SEED }, ctxFor())
  assert.match(l2.summary, /坐标层级：L2/)
  const l1 = await planMod.plan({ tool: 'xlsx_guide', params: { op: 'sort_range' }, ref: REF, bytes: SEED }, ctxFor())
  assert.match(l1.summary, /坐标层级：L1/)
  const l0 = await planMod.plan({ tool: 'xlsx_guide', params: { op: 'add_sheet' }, ref: REF, bytes: SEED }, ctxFor())
  assert.match(l0.summary, /坐标层级：L0/)
})

test('整批级失败用 index -1（不得冒充第一条命令）', async () => {
  const commands = Array.from({ length: 201 }, () => ({
    op: 'set_cell',
    sheet: 'Sheet1',
    address: 'A1',
    value: 1,
  }))
  const rejected = await planMod.plan(
    { tool: 'xlsx_apply', params: { ref: REF, commands }, ref: REF, bytes: SEED },
    ctxFor(),
  )
  assert.equal(rejected.plan.kind, 'reject')
  assert.equal(rejected.plan.rejected[0].index, -1)
  assert.equal(rejected.plan.rejected[0].code, 'command.too-large')
})

test('read 模块导出 project / query / parseWorkbook', () => {
  assert.equal(typeof readMod.project, 'function')
  assert.equal(typeof readMod.query, 'function')
  assert.equal(typeof readMod.parseWorkbook, 'function')
})
