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
