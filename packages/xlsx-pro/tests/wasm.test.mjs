// wasm 执行底座（P3）：recalc + save_archive。
//
// 断言的是能力位与两条不变量：
//   1. 有产物时 `runtime().exec === 'wasm'`，`recalc === true`；
//   2. 未改条目的**压缩流**逐字节不变（I1）；
//   3. 改公式后保存，缓存 `<v>` 由引擎算出（I4），不是旧值、也不是空白。
//
// 产物不在（未 vendor / 未构建）时整组 skip —— CI 里 wasm 是可选的。

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'
import { strFromU8, unzipSync } from 'fflate'

const WASM = new URL('../dist/engine/xlsx-engine.wasm', import.meta.url)
const enabled = existsSync(WASM)

const planMod = await import('../dist/engine/plan.mjs')
const execMod = await import('../dist/engine/exec.mjs')
const runtimeMod = await import('../dist/engine/runtime.mjs')
const zipRaw = await import('../dist/shared/zip-raw.mjs')

const SEED = new Uint8Array(readFileSync(new URL('../dist/seed.xlsx', import.meta.url)))
const REF = 'verify://wasm.xlsx'
const ctx = () => ({
  readBytes: async () => ({ bytes: SEED, version: 'v1' }),
  tempDir: '',
  importEpoch: 0,
  budget: { maxBytes: 256 * 1024 * 1024, deadlineMs: 60_000 },
})

async function applyAndExec(commands, bytes = SEED) {
  const planned = await planMod.plan(
    { tool: 'xlsx_apply', params: { ref: REF, commands }, ref: REF, baseVersion: 'v1', bytes },
    ctx(),
  )
  assert.equal(planned.plan.kind, 'save', planned.summary)
  const outcome = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan: planned.plan, bytes },
    ctx(),
  )
  assert.ok(outcome.bytes instanceof Uint8Array, 'exec 没有产出字节')
  return outcome
}

test('runtime：wasm 产物在时 exec=wasm / recalc=true', { skip: !enabled }, async () => {
  const report = await runtimeMod.runtime(ctx())
  assert.equal(report.exec, 'wasm')
  assert.equal(report.features.recalc, true)
  assert.equal(report.probe?.engine, 'wasm')
  assert.deepEqual(report.degraded, ['cancel-unavailable'])
  assert.equal(report.features.tokens.inPage, 'docRev+structureEpoch')
})

test('save_archive：未改条目压缩流逐字节不变（I1）', { skip: !enabled }, async () => {
  const outcome = await applyAndExec([
    { op: 'set_cell', sheet: 'Sheet1', address: 'A1', value: 'hello' },
    { op: 'set_cell', sheet: 'Sheet1', address: 'B1', value: 42 },
  ])
  const touched = outcome.result.audit.touchedEntries
  const drifted = zipRaw.driftedEntries(SEED, outcome.bytes, touched)
  assert.deepEqual(drifted, [], `未改条目被重写：${drifted.join(',')}`)
  assert.equal(outcome.result.audit.engine, 'wasm')
})

test('recalc_cells：改公式后缓存 <v> 由引擎算出（I4）', { skip: !enabled }, async () => {
  const outcome = await applyAndExec([
    { op: 'set_cell', sheet: 'Sheet1', address: 'A1', value: 10 },
    { op: 'set_cell', sheet: 'Sheet1', address: 'A2', value: 20 },
    { op: 'set_formula', sheet: 'Sheet1', address: 'A3', formula: '=SUM(A1:A2)' },
  ])
  const sheet = strFromU8(unzipSync(outcome.bytes)['xl/worksheets/sheet1.xml'])
  const a3 = /<c r="A3"[^>]*>(.*?)<\/c>/s.exec(sheet)?.[1] ?? ''
  assert.match(a3, /<f>SUM\(A1:A2\)<\/f>/, `A3 缺少公式：${a3}`)
  assert.match(a3, /<v>30<\/v>/, `A3 缓存值不对：${a3}`)

  // 依赖单元格：改输入后，另一条保存要把旧缓存刷成新值
  const next = await applyAndExec(
    [{ op: 'set_cell', sheet: 'Sheet1', address: 'A1', value: 100 }],
    outcome.bytes,
  )
  const sheet2 = strFromU8(unzipSync(next.bytes)['xl/worksheets/sheet1.xml'])
  const a3b = /<c r="A3"[^>]*>(.*?)<\/c>/s.exec(sheet2)?.[1] ?? ''
  assert.match(a3b, /<v>120<\/v>/, `A1 改动后 A3 缓存未刷新：${a3b}`)
  assert.ok(
    (next.result.notices ?? []).length === 0 || Array.isArray(next.result.notices),
    'notices 形状',
  )
})

test('save 计划与 bytes 路径在 wasm 下仍逐条目一致', { skip: !enabled }, async () => {
  const pageSave = await import('../dist/shared/page-save.mjs')
  const file = { name: 'fixture.xlsx', sheets: [{ id: 'sheet-1', name: 'Sheet1' }] }
  const request = {
    edits: [{ sheetId: 'sheet-1', row: 0, column: 0, writeValue: true, value: 'x' }],
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
  }
  const viaBytes = (await pageSave.saveWorkbookEdits(SEED, request, file)).bytes
  const viaSave = (
    await execMod.exec(
      { ref: REF, baseVersion: 'v1', plan: { kind: 'save', request: pageSave.planWorkbookSave(request, file) }, bytes: SEED },
      ctx(),
    )
  ).bytes
  const left = unzipSync(viaBytes)
  const right = unzipSync(viaSave)
  assert.deepEqual(Object.keys(left).sort(), Object.keys(right).sort())
  for (const name of Object.keys(left))
    assert.deepEqual([...left[name]], [...right[name]], `条目内容不同：${name}`)
})
