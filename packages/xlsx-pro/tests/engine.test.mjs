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
  // 命令目录是静态的：不因页面开着而占用一次页面往返（否则拿到无意义的投影）。
  assert.equal(tools.find((tool) => tool.name === 'xlsx_guide').pageRead, false)
  assert.equal(tools.find((tool) => tool.name === 'xlsx_read').pageRead, undefined)
})

test('xlsx_apply 声明里暴露令牌参数（structureEpoch / docRev）', () => {
  const apply = tools.find((tool) => tool.name === 'xlsx_apply')
  const properties = apply.parameters.properties
  assert.ok(properties.structureEpoch, '缺 structureEpoch 参数：模型无法把读到的令牌带回去')
  assert.ok(properties.docRev, '缺 docRev 参数')
})

test('xlsx_read：无页面读也回基线版本（否则 apply 永远缺 version）', async () => {
  const out = await read(SEED, { what: 'context' })
  assert.match(out.summary, /^version=v1\n# workbook\.xlsx/)
  assert.match(out.plan.text, /^version=v1\n/)
  const sheets = await read(SEED, { what: 'sheets' })
  assert.match(sheets.summary, /^version=v1\n/)
  const found = await read(SEED, { what: 'find', query: 'a' })
  assert.match(found.summary, /^version=v1\n/)
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

/// 解析产物里某个单元格的“实际生效样式”：把 s= 索引翻成 cellXfs 里的那条 xf。
/// 索引值本身不可比（同一样式可能对应多个 xf），但内容必须一致。
function resolvedStyle(sheetXml, stylesXml, ref) {
  const cellXfs = (stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/) || [, ''])[1]
  const xfs = [...cellXfs.matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((m) => m[0])
  const tag = (sheetXml.match(new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>`)) || [''])[0]
  const index = Number((tag.match(/\bs="(\d+)"/) || [])[1])
  const xf = xfs[Number.isFinite(index) ? index : 0] ?? ''
  // 去掉 xfId / 属性顺序差异，只留“字体 / 填充 / 边框 / 数字格式 / 对齐”事实。
  return xf
    .replace(/\s+xfId="\d+"/g, '')
    .replace(/\s+apply\w+="0"/g, '')
    .replace(/\s+/g, ' ')
    .trim()
}

test('copy_range：样式 / 合并 / 行高一起复制（"按现有格式补一块"）', async () => {
  // 1) 造一个“带格式的周报块”：标题加粗填充 + 内容区合并换行 + 行高
  const first = await apply([
    {
      op: 'set_range',
      sheet: 'Sheet1',
      start: 'A1',
      values: [['2026/11/16~2026/11/20周报'], ['智能体开发：本周完成……'], ['', '', '']],
    },
    { op: 'format_range', sheet: 'Sheet1', range: 'A1:C1', format: { bold: true, fontSize: 14, fillColor: '#DDEBF7' } },
    { op: 'format_range', sheet: 'Sheet1', range: 'A2:C3', format: { wrapText: true, verticalAlign: 'top' } },
    { op: 'merge_cells', sheet: 'Sheet1', range: 'A2:C3' },
    { op: 'set_row_height', sheet: 'Sheet1', row: 1, heightPoints: 30 },
  ])
  // 2) 整块向下复制（+4 行）
  const copied = await apply([{ op: 'copy_range', sheet: 'Sheet1', source: 'A1:C3', target: 'A5' }], first.bytes)
  const sheetXml = strFromU8(unzipSync(copied.bytes)['xl/worksheets/sheet1.xml'])
  const stylesXml = strFromU8(unzipSync(copied.bytes)['xl/styles.xml'])
  // 样式：标题、内容、填充格（C3 是空值但有样式）都要跟着走
  for (const [source, target] of [['A1', 'A5'], ['A2', 'A6'], ['C3', 'C7']]) {
    assert.equal(
      resolvedStyle(sheetXml, stylesXml, target),
      resolvedStyle(sheetXml, stylesXml, source),
      `${source} 的格式没有复制到 ${target}`,
    )
  }
  // 合并区按偏移复制
  assert.match(sheetXml, /<mergeCell ref="A2:C3"\/>/)
  assert.match(sheetXml, /<mergeCell ref="A6:C7"\/>/)
  // 行高按源行对应复制（px → pt 往返不产生偏差）
  const rows = Object.fromEntries([...sheetXml.matchAll(/<row r="(\d+)"([^>]*)>/g)].map((m) => [m[1], m[2]]))
  assert.match(rows['1'], /ht="30"/)
  assert.match(rows['5'], /ht="30"/)
})

test('delete_sheet：删表不算“改动未涉及的 ZIP 条目”（I1 漂移检查）', async () => {
  // 先建第二张表，再把它删掉：删除必须能落盘，而不是被漂移检查挡住。
  const added = await apply([{ op: 'add_sheet', name: '探针表' }])
  const removed = await apply([{ op: 'delete_sheet', sheet: '探针表' }], added.bytes)
  const names = [...strFromU8(unzipSync(removed.bytes)['xl/workbook.xml']).matchAll(/<sheet\b[^>]*\bname="([^"]+)"/g)]
    .map((m) => m[1])
  assert.ok(!names.includes('探针表'), `探针表没有被删掉：${names.join('/')}`)
})
