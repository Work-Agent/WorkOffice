// ─── 令牌归一（回归：apply 一直被 stale-index 拒掉的根因） ─────────────
//
// 校验点只认 `docRev` / `structureEpoch`；宿主历史上从没透传，工具入参里
// 又叫 `epoch` / `revision`。名字对不上 → 含 A1 坐标的批次 100% 被拒。

import { test } from 'node:test'
import assert from 'node:assert/strict'

const { resolveTokens, tokenNeeds } = await import('../src/shared/op-tiers.ts')

const batch = (params) => [{ tool: 'xlsx_apply', params }]

test('令牌：消息顶层优先（宿主透传）', () => {
  const tokens = resolveTokens({
    message: { docRev: 3, structureEpoch: 7 },
    commands: batch({ epoch: 1, revision: 2, structureEpoch: 9 }),
  })
  assert.deepEqual(tokens, { docRev: 3, structureEpoch: 7 })
})

test('令牌：入参旧名 epoch / revision 能被接住', () => {
  const tokens = resolveTokens({ commands: batch({ epoch: 5, revision: 4 }) })
  assert.deepEqual(tokens, { docRev: 4, structureEpoch: 5 })
})

test('令牌：入参新名 structureEpoch / docRev 生效', () => {
  const tokens = resolveTokens({ commands: batch({ structureEpoch: 11, docRev: 12 }) })
  assert.deepEqual(tokens, { docRev: 12, structureEpoch: 11 })
})

test('令牌：一层 commands 里的 op 参数也能读取（扁平形状）', () => {
  const tokens = resolveTokens({
    commands: [{ op: 'set_range', epoch: 2 }, { op: 'copy_range', revision: 6 }],
  })
  assert.deepEqual(tokens, { docRev: 6, structureEpoch: 2 })
})

test('令牌：取不到就不返回（宁可报 stale-index，也不伪造）', () => {
  assert.deepEqual(resolveTokens({ commands: batch({}) }), {})
  assert.deepEqual(resolveTokens({}), {})
})

test('令牌：未注水 0 是合法值（页面初始态就是 0）', () => {
  const tokens = resolveTokens({
    message: { docRev: 0, structureEpoch: 0 },
  })
  assert.deepEqual(tokens, { docRev: 0, structureEpoch: 0 })
})

test('令牌：按 op 层级判断需要哪一把', () => {
  assert.deepEqual(tokenNeeds([{ op: 'add_sheet' }]), { docRev: false, structureEpoch: false })
  assert.deepEqual(tokenNeeds([{ op: 'set_range' }]), { docRev: false, structureEpoch: true })
  assert.deepEqual(tokenNeeds([{ op: 'copy_range' }]), { docRev: true, structureEpoch: false })
})
