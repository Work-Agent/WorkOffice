// 页面保存请求的两条路径必须产出同一份包：
//   A. `plan { kind:'bytes' }`  页面自己调 gateway + JSZip（回退路径）
//   B. `plan { kind:'save'  }`  请求交给执行点，由 exec 调 gateway + JSZip
//
// 两者共用 `planWorkbookSave` 的映射，所以任何「只有一条路径传了某组参数」
// 的偏差都会在这里表现为包内容不一致，而不是用户文件里少了一张图。

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { strFromU8, unzipSync } from 'fflate'

const execMod = await import('../dist/engine/exec.mjs')
const saveMod = await import('../dist/shared/page-save.mjs')

const SEED = new Uint8Array(readFileSync(new URL('../dist/seed.xlsx', import.meta.url)))
const REF = 'verify://parity.xlsx'
const BUDGET = { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 }
const ctx = {
  readBytes: async () => ({ bytes: SEED, version: 'v1' }),
  tempDir: '',
  importEpoch: 0,
  budget: BUDGET,
}
const FILE = { name: 'fixture.xlsx', sheets: [{ id: 'sheet-1', name: 'Sheet1' }] }

/** App 侧保存请求的骨架（WorkbookSaveRequest 的必填组）。 */
function pageRequest(overrides = {}) {
  return {
    edits: [],
    bulkConstantFills: [],
    structuralOps: [],
    chartEdits: [],
    visualEdits: [],
    visualAdditions: [],
    cellImageAdditions: [],
    tableAdditions: [],
    tableEdits: [],
    pivotAdditions: [],
    sheetOps: [],
    sheetOrder: [],
    filterStates: [],
    hyperlinkEdits: [],
    cfStates: [],
    dvStates: [],
    pageSetupStates: [],
    noteStates: [],
    pivotCacheRefreshPaths: [],
    pivotRefreshUpdates: [],
    sheetProtections: [],
    sparklineAdditions: [],
    formulaValues: [],
    definedNamesState: null,
    themeState: null,
    workbookProtectionState: null,
    protectedRangeStates: [],
    ...overrides,
  }
}

const cellsRequest = pageRequest({
  edits: [
    { sheetId: 'sheet-1', row: 1, column: 0, writeValue: true, value: '张三' },
    { sheetId: 'sheet-1', row: 1, column: 1, writeValue: true, value: 1200 },
  ],
  hyperlinkEdits: [
    { sheetId: 'sheet-1', row: 2, column: 0, target: 'https://example.com/xlsx-pro' },
  ],
  sheetOps: [{ kind: 'set-sheet-tab-color', sheetId: 'sheet-1', color: '#FF6600' }],
})

async function execSave(plan, bytes) {
  const outcome = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan, bytes: bytes ?? SEED },
    ctx,
  )
  assert.ok(outcome.bytes instanceof Uint8Array)
  return outcome
}

function entriesOf(bytes) {
  const archive = unzipSync(bytes)
  const map = new Map()
  for (const [name, content] of Object.entries(archive)) map.set(name, content)
  return map
}

function assertSameEntries(a, b, message) {
  const left = entriesOf(a)
  const right = entriesOf(b)
  assert.deepEqual([...left.keys()].sort(), [...right.keys()].sort(), `${message}：条目集合不同`)
  for (const [name, content] of left) {
    assert.deepEqual(
      [...content],
      [...right.get(name)],
      `${message}：条目 ${name} 内容不同`,
    )
  }
}

test('page-save 映射：附加组必须全部落到请求上', () => {
  const plan = saveMod.planWorkbookSave(cellsRequest, FILE)
  assert.equal(plan.edits.length, 2)
  assert.equal(plan.hyperlinkEdits.length, 1)
  assert.deepEqual(plan.tabColorStates, [{ sheetName: 'Sheet1', color: '#FF6600' }])
})

test('kind:save 与 kind:bytes 产出同一份包', async () => {
  const page = await saveMod.saveWorkbookEdits(SEED, cellsRequest, FILE)
  const viaBytes = page.bytes
  const viaSave = (await execSave({ kind: 'save', request: saveMod.planWorkbookSave(cellsRequest, FILE) })).bytes
  assertSameEntries(viaBytes, viaSave, 'save vs bytes')
})

test('kind:save 不再丢弃页面专属字段（超链接 / 标签色）', async () => {
  const plan = saveMod.planWorkbookSave(
    pageRequest({
      edits: [{ sheetId: 'sheet-1', row: 0, column: 0, writeValue: true, value: 'a' }],
      hyperlinkEdits: [
        { sheetId: 'sheet-1', row: 0, column: 0, target: 'https://example.com/a' },
      ],
      sheetOps: [{ kind: 'set-sheet-tab-color', sheetId: 'sheet-1', color: '#00AA55' }],
    }),
    FILE,
  )
  assert.equal(plan.hyperlinkEdits?.length, 1)
  assert.equal(plan.tabColorStates?.length, 1)

  const outcome = await execSave({ kind: 'save', request: plan })
  const entries = entriesOf(outcome.bytes)
  const sheet = strFromU8(entries.get('xl/worksheets/sheet1.xml'))
  assert.match(sheet, /<hyperlink\b/)
  assert.match(sheet, /tabColor/)
  const rels = [...entries.keys()].find((name) => name.includes('worksheets/_rels'))
  assert.ok(rels, '超链接必须写入 worksheet rels')
  assert.match(strFromU8(entries.get(rels)), /example\.com/)
  assert.deepEqual(outcome.result.sync?.kind, 'manifest')
})
