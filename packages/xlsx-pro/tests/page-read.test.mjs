// 页面 invoke 的数据源（P4a）：结构化范围读取 + 会话复用。
//
// 断言三件事：
//   1. `xlsx_read what:'cells'` 返回**引擎原样**的 RangeResult（键集合与 fork 的
//      readWorkbookRange 一致，页面只透传不翻译）；
//   2. 会话按 (ref, version) 复用：第二次 cached=true、sessionId 不变；
//   3. 版本变化 → 新会话；非法参数是显式错误，不是空结果。

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'

const enabled = existsSync(new URL('../dist/engine/xlsx-engine.wasm', import.meta.url))
const planMod = await import('../dist/engine/plan.mjs')
const execMod = await import('../dist/engine/exec.mjs')

const SEED = new Uint8Array(readFileSync(new URL('../dist/seed.xlsx', import.meta.url)))
const REF = 'verify://cells.xlsx'
const ctxFor = (bytes) => ({
  readBytes: async () => ({ bytes, version: 'v1' }),
  // 稳定会话目录（宿主按插件提供）；tempDir 是每次调用的中间产物
  sessionDir: '',
  tempDir: '',
  importEpoch: 0,
  budget: { maxBytes: 256 * 1024 * 1024, deadlineMs: 60_000 },
})

async function fixtureBytes() {
  const planned = await planMod.plan(
    {
      tool: 'xlsx_apply',
      params: {
        ref: REF,
        commands: [
          { op: 'set_range', sheet: 'Sheet1', start: 'A1', values: [['hello', 42]] },
          { op: 'set_formula', sheet: 'Sheet1', address: 'A2', formula: '=B1*2' },
        ],
      },
      ref: REF,
      baseVersion: 'v1',
      bytes: SEED,
    },
    ctxFor(SEED),
  )
  assert.equal(planned.plan.kind, 'save', planned.summary)
  const outcome = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan: planned.plan, bytes: SEED },
    ctxFor(SEED),
  )
  assert.ok(outcome.bytes instanceof Uint8Array)
  return outcome.bytes
}

const readCells = (bytes, params, baseVersion = 'v1') =>
  planMod.plan(
    { tool: 'xlsx_read', params: { what: 'cells', ...params }, ref: REF, baseVersion, bytes },
    ctxFor(bytes),
  )

test('结构化范围读取返回引擎原样的 RangeResult', { skip: !enabled }, async () => {
  const bytes = await fixtureBytes()
  const planned = await readCells(bytes, { sheetId: 'sheet-1', range: 'A1:B2' })
  assert.equal(planned.plan.kind, 'read')
  const data = planned.plan.data
  for (const key of [
    'cells',
    'rows',
    'merges',
    'hyperlinks',
    'indexedThroughRow',
    'indexingComplete',
  ])
    assert.ok(key in data, `RangeResult 缺少 ${key}`)

  const cells = new Map(data.cells.map((cell) => [`${cell.row}:${cell.column}`, cell]))
  assert.equal(cells.get('0:0')?.value, 'hello')
  assert.equal(cells.get('0:1')?.value, 42)
  assert.match(String(cells.get('1:0')?.formula ?? ''), /B1\*2/)
  assert.equal(cells.get('1:0')?.value, 84, '公式缓存值应由引擎给出')

  // 引擎与页面共用 sheet id 命名（`sheet-<sheetId>`），不需要翻译
  assert.equal(planned.plan.audit?.cached, false)
})

test('会话按 (ref, version) 复用；版本变化换会话', { skip: !enabled }, async () => {
  const bytes = await fixtureBytes()
  const first = await readCells(bytes, { sheetId: 'sheet-1', range: 'A1:B1' })
  const second = await readCells(bytes, { sheetId: 'sheet-1', range: 'A1:B1' })
  assert.equal(second.plan.audit?.cached, true, '第二次应命中会话')
  assert.equal(second.plan.audit?.sessionId, first.plan.audit?.sessionId)

  const nextVersion = await readCells(bytes, { sheetId: 'sheet-1', range: 'A1:B1' }, 'v2')
  assert.equal(nextVersion.plan.audit?.cached, false)
  assert.notEqual(nextVersion.plan.audit?.sessionId, first.plan.audit?.sessionId)
})

test('工作表可按名字寻址；非法参数是显式错误', { skip: !enabled }, async () => {
  const bytes = await fixtureBytes()
  const byName = await readCells(bytes, { sheet: 'Sheet1', range: 'A1:B1' })
  assert.equal(byName.plan.kind, 'read')

  await assert.rejects(
    readCells(bytes, { sheetId: '没有这张表', range: 'A1' }),
    (error) => error.code === 'command.not-found',
  )
  // parseRange 会把 B2:A1 归一化成 A1:B2，所以用一个真正非法的串
  await assert.rejects(
    readCells(bytes, { sheetId: 'sheet-1', range: 'not-a-range' }),
    (error) => error.code === 'command.invalid',
  )
})

test("meta：返回 open 的原样元数据（页面重建 file 模型用）", { skip: !enabled }, async () => {
  const bytes = await fixtureBytes()
  // 独立 ref：会话缓存是模块级的，前面的用例已经开过 REF 的会话
  const metaRef = 'verify://meta.xlsx'
  const meta = await planMod.plan(
    { tool: 'xlsx_read', params: { what: 'meta' }, ref: metaRef, baseVersion: 'm1', bytes },
    ctxFor(bytes),
  )
  assert.equal(meta.plan.kind, 'read')
  const data = meta.plan.data
  for (const key of ['sheets', 'styles', 'definedNames', 'visuals', 'activeTab', 'entryCount'])
    assert.ok(key in data, `元数据缺少 ${key}`)
  assert.ok(Array.isArray(data.styles) && data.styles.length > 0, 'styles 应为 cellXfs 表')
  assert.equal(data.sheets[0].id, 'sheet-1')
  assert.equal(meta.plan.audit?.cached, false)
  const again = await planMod.plan(
    {
      tool: 'xlsx_read',
      params: { what: 'cells', sheetId: 'sheet-1', range: 'A1' },
      ref: metaRef,
      baseVersion: 'm1',
      bytes,
    },
    ctxFor(bytes),
  )
  assert.equal(again.plan.audit?.cached, true, 'meta 之后的范围读取必须命中同一会话')
  assert.equal(again.plan.audit?.sessionId, meta.plan.audit?.sessionId)
})

test('exec 不产字节，只把结构化结果放进 audit.data', { skip: !enabled }, async () => {
  const bytes = await fixtureBytes()
  const planned = await readCells(bytes, { sheetId: 'sheet-1', range: 'A1' })
  const outcome = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan: planned.plan, bytes },
    ctxFor(bytes),
  )
  assert.equal(outcome.bytes, undefined)
  assert.ok(outcome.result.audit?.data, 'audit.data 应带上范围结果')
  assert.equal(outcome.result.audit?.sessionId, planned.plan.audit?.sessionId)
})
