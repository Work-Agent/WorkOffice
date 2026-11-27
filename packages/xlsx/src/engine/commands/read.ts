// ─── xlsx.read：单文件只读投影 ───────────────────────────────────
//
// P0 只读磁盘；P3 接入页面租约后，有页面时改问 live 模型（同一条命令、同一形状）。

import { assertXlsxFile } from '../ref'
import { loadView, projectRead, type ReadWhat } from '../read'
import { CommandError, invalid } from '../errors'
import { enqueue, leaseOf } from '../state'

const WHATS: ReadWhat[] = ['context', 'sheets', 'range', 'formulas', 'find', 'audit']

export interface ReadResult {
  ref: string
  what: ReadWhat
  hash: string
  size: number
  text: string
}

export async function read(input: Record<string, unknown>): Promise<ReadResult> {
  const ref = String(input.ref ?? '')
  const what = String(input.what ?? 'context') as ReadWhat
  if (!WHATS.includes(what))
    invalid(`what 非法：${what}（可用：${WHATS.join(' / ')}）`)

  const file = await assertXlsxFile(ref)

  // 有页面：读 live 模型（含未保存修改）。页面没实现这个 what 时会回 unsupported。
  const lease = leaseOf(ref)
  if (lease) {
    const payload = await enqueue(ref, { type: 'read', what, params: input })
    if (payload.error) throw new CommandError(payload.error.code, payload.error.message)
    const data = (payload.data ?? {}) as { text?: unknown }
    const body = typeof data.text === 'string' ? data.text : JSON.stringify(payload.data)
    return {
      ref,
      what,
      hash: lease.hash,
      size: file.size,
      text: `hash=${lease.hash} size=${file.size} docRev=${lease.docRev} structureEpoch=${lease.structureEpoch}\n${body}`,
    }
  }

  const entry = await loadView(file)
  const projection = await projectRead(entry, file, what, input)
  return {
    ref,
    what,
    hash: entry.hash,
    size: entry.size,
    // 首行令牌：写命令的唯一基线来源（P3 起 apply 用它做磁盘 CAS）。
    text: `hash=${entry.hash} size=${entry.size}\n${projection.text}`,
  }
}
