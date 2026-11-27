// ─── 命令规范化（页面与执行点同源） ──────────────────────────────────
//
// 上游目录把 `setChart` 标成 not-ai-callable；本插件对模型暴露 `editChart`
// 高层字段，统一在进入 runTxn 前映射成 `setChart`。

export interface PptxOp {
  op: string
  target?: { slide?: number | string; el?: string; part?: string }
  [key: string]: unknown
}

export interface NormalizedOps {
  ops: PptxOp[]
  warnings: string[]
}

export function normalizeOps(raw: unknown): NormalizedOps {
  if (!Array.isArray(raw)) return { ops: [], warnings: ['commands 不是数组'] }
  const ops: PptxOp[] = []
  const warnings: string[] = []
  for (const item of raw) {
    if (!item || typeof item !== 'object') {
      ops.push({ op: '' })
      continue
    }
    const op = { ...(item as PptxOp) }
    if (op.op === 'editChart') {
      const { op: _drop, target, ...patch } = op
      ops.push({ op: 'setChart', target, patch })
      continue
    }
    ops.push(op)
  }
  return { ops, warnings }
}

/** 这些 op 需要页面 UI / 二进制 / 剪贴板状态；无页面时整批拒绝。 */
export const PAGE_ONLY_OPS = new Set([
  'addPicture',
  'replacePicture',
  'setImageFill',
  'addMedia',
  'addNarration',
  'clearNarration',
  'addModel3d',
  'setModel3dView',
  'pasteElements',
  'pasteSlide',
  'insertSlidePptx',
  'setAnimations',
])

/** 上游注册但不对模型直接暴露的低层 op。 */
export const HIDDEN_OPS = new Set(['setChart'])

export function opNameOf(value: unknown): string {
  return String((value as { op?: unknown } | null)?.op ?? '')
}
