// 保存路径判定必须是纯函数：页面侧没有「失败再回退」的余地，
// 半途换路径会让同一批编辑被应用两次（save 计划不是内容幂等的）。

import { test } from 'node:test'
import assert from 'node:assert/strict'

const { shouldUseSavePlan, planJsonBytes, DEFAULT_MAX_PLAN_BYTES } = await import(
  '../dist/shared/save-plan-budget.mjs'
)

const base = { edits: [{ sheetName: 'Sheet1', row: 0, column: 0, writeValue: true, cell: { value: 1 } }], structuralOps: [] }

test('默认走 save 计划', () => {
  const decision = shouldUseSavePlan(base, { mode: 'save' })
  assert.equal(decision.usePlan, true)
  assert.equal(decision.reason, 'ok')
  assert.ok(decision.planBytes > 0)
})

test('save-as / recovery 必须走字节路径（目标不是当前 ref）', () => {
  for (const mode of ['save-as', 'recovery']) {
    const decision = shouldUseSavePlan(base, { mode })
    assert.equal(decision.usePlan, false, mode)
    assert.equal(decision.reason, 'save-as')
  }
})

test('含内联媒体走字节路径', () => {
  assert.equal(
    shouldUseSavePlan(
      { ...base, cellImageAdditions: [{ sheetName: 'Sheet1', dataUrl: 'data:image/png;base64,AAA' }] },
      { mode: 'save' },
    ).reason,
    'inline-media',
  )
  assert.equal(
    shouldUseSavePlan(
      { ...base, visualAdditions: [{ sheetName: 'Sheet1', image: { dataUrl: 'x' } }] },
      { mode: 'save' },
    ).reason,
    'inline-media',
  )
})

test('计划体积超预算走字节路径', () => {
  const decision = shouldUseSavePlan(base, { mode: 'save', maxPlanBytes: 1 })
  assert.equal(decision.usePlan, false)
  assert.equal(decision.reason, 'too-large')
  assert.equal(shouldUseSavePlan(base, {}).planBytes, planJsonBytes(base))
  assert.equal(DEFAULT_MAX_PLAN_BYTES, 2 * 1024 * 1024)
})
