// ─── op 坐标层级（设计稿 §6.4） ────────────────────────────────────────
//
// 令牌判定是 op 级，应用是批级原子的：
//   L0 全局语义     无坐标，按名称/全局状态重新解析      → 不校验令牌
//   L1 内容选择器   需要「读到的内容仍未变」              → docRev
//   L2 绝对坐标     A1 地址 / 结构行号，值编辑下稳定      → structureEpoch
//
// 页面代理与无页面路径必须用同一张表，否则同一批命令在两条路径上
// 得到不同的新鲜度判定。

export type OpTier = 'L0' | 'L1' | 'L2'

/** 读内容型：会读取源区域的值，内容变了结果就不一样。 */
const L1_OPS: ReadonlySet<string> = new Set([
  'find_replace',
  'fill_range',
  'copy_range',
  'move_range',
  'sort_range',
  'convert_to_values',
])

/** 绝对坐标型：A1 地址只在结构改动下失效（值编辑不影响）。 */
const L2_OPS: ReadonlySet<string> = new Set([
  'set_cell',
  'set_formula',
  'clear_cell',
  'set_range',
  'clear_range',
  'format_range',
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
  'merge_cells',
  'unmerge_cells',
  'set_row_height',
  'set_col_width',
  'set_rows_hidden',
  'set_cols_hidden',
])

export function opTier(op: string): OpTier {
  if (L1_OPS.has(op)) return 'L1'
  if (L2_OPS.has(op)) return 'L2'
  return 'L0'
}

export interface OpTokenNeeds {
  readonly docRev: boolean
  readonly structureEpoch: boolean
}

/** 一批命令需要哪把令牌（L1 → docRev，L2 → structureEpoch）。 */
export function tokenNeeds(ops: Iterable<unknown>): OpTokenNeeds {
  let docRev = false
  let structureEpoch = false
  for (const entry of ops) {
    const op = String((entry as { op?: unknown } | null | undefined)?.op ?? '')
    const tier = opTier(op)
    if (tier === 'L1') docRev = true
    else if (tier === 'L2') structureEpoch = true
    if (docRev && structureEpoch) break
  }
  return { docRev, structureEpoch }
}

const asNumber = (value: unknown): number | undefined =>
  typeof value === 'number' && Number.isFinite(value) ? value : undefined

/**
 * 一批命令实际可用的令牌：消息顶层（宿主透传）→ 命令入参（旧名兼容）。
 *
 * 校验点只认 `docRev` / `structureEpoch`，而工具入参历史上叫 `revision` /
 * `epoch`。两边名字对不上时，含 A1 坐标的批次会 100% 被 `command.stale-index`
 * 拒掉（"必须带上读到上下文时的 structureEpoch"），所以在这里显式归一。
 */
export function resolveTokens(sources: {
  readonly message?: { readonly docRev?: number; readonly structureEpoch?: number }
  readonly commands?: Iterable<unknown>
}): { docRev?: number; structureEpoch?: number } {
  let docRev = asNumber(sources.message?.docRev)
  let structureEpoch = asNumber(sources.message?.structureEpoch)
  for (const entry of sources.commands ?? []) {
    if (docRev !== undefined && structureEpoch !== undefined) break
    // 写批命令的形状是 { tool, params: { …, epoch / revision / docRev / structureEpoch } }。
    const params = ((entry as { params?: unknown } | null | undefined)?.params ?? entry) as
      | Record<string, unknown>
      | null
      | undefined
    if (!params || typeof params !== 'object') continue
    docRev ??= asNumber(params['docRev']) ?? asNumber(params['revision'])
    structureEpoch ??= asNumber(params['structureEpoch']) ?? asNumber(params['epoch'])
  }
  return {
    ...(docRev === undefined ? {} : { docRev }),
    ...(structureEpoch === undefined ? {} : { structureEpoch }),
  }
}
