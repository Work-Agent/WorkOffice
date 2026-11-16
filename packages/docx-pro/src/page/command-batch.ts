// ─── docx_apply：混合命令批（页面路径） ────────────────────────────────
//
// 三组命令按**原顺序**执行：
//   blocks  → PM 直接编辑（校验就在它自己的位置上做，因为块索引会被前序命令移动）
//   editor  → 连续的一段 op 合成一次 executeOps（一个事务、一次撤销）
//   actions → executeTool 的专用动作（HTML 片段 / 图表 / 修订）
//
// 失败语义：整批回滚。执行前抓一份 PM JSON 快照，任何一条被拒就恢复——这样
// 「批内一条被拒 = 整批不落盘」对模型成立（回执带 index 与 code）。

import type { Editor } from '@tiptap/core'
import { BLANK_BULLET_NUM_ID, BLANK_ORDERED_NUM_ID } from '@genoffice/docx-engine'
import type { FileActionContext } from '../editor/file-actions'
import { executeTool } from '../editor/ai/tools'
import { executeOps } from '../editor/ai/ops'
import { findNumId, type NumIds } from '../editor/ai/protocol'
import { groupOf } from '../tools/catalog'
import {
  BlockCommandError,
  MAX_COMMANDS,
  planBlockCommand,
} from './block-commands'

interface Command {
  op?: unknown
  [key: string]: unknown
}

export interface BatchResult {
  applied: number
  mutated: boolean
  summary: string
}

export function numIdsFor(ctx: FileActionContext): NumIds {
  const blocks = ctx.doc?.parsed.blocks ?? []
  const isBlank = ctx.doc?.isBlank === true
  return {
    bullet: findNumId(blocks, 'bullet') ?? (isBlank ? BLANK_BULLET_NUM_ID : null),
    ordered: findNumId(blocks, 'ordered') ?? (isBlank ? BLANK_ORDERED_NUM_ID : null),
  }
}

export async function applyCommandBatch(
  editor: Editor,
  raw: unknown,
  ctx: FileActionContext,
): Promise<BatchResult> {
  if (!Array.isArray(raw)) throw new BlockCommandError('command.invalid', 'commands 必须是数组')
  if (raw.length === 0) return { applied: 0, mutated: false, summary: '没有命令' }
  if (raw.length > MAX_COMMANDS)
    throw new BlockCommandError(
      'command.too-large',
      `批上限 ${MAX_COMMANDS} 条（收到 ${raw.length}）`,
    )

  const commands = raw as Command[]
  // 先做纯粹的「名字对不对」检查：拼错 op 不该走到半批。
  const groups = commands.map((command, index) => {
    const op = String(command.op ?? '')
    const group = groupOf(op)
    if (!group)
      throw new BlockCommandError(
        'command.not-found',
        `commands[${index}]：未知 op「${op || '(空)'}」（用 docx_guide 取可用 op）`,
      )
    return group
  })

  const snapshot = editor.getJSON()
  const numIds = numIdsFor(ctx)
  let applied = 0
  let mutated = false
  let pending: Array<Record<string, unknown>> = []

  const flushOps = (): void => {
    if (pending.length === 0) return
    const batch = pending
    pending = []
    const outcome = executeOps(editor, batch)
    if (!outcome.ok)
      throw new BlockCommandError(
        'command.invalid',
        `op 被拒绝：${outcome.error ?? '未知原因'}`,
      )
    applied += batch.length
    mutated = true
  }

  try {
    for (const [index, command] of commands.entries()) {
      const group = groups[index]

      if (group === 'editor') {
        const { op: _op, ...fields } = command
        pending.push({ op: String(command.op), ...fields })
        continue
      }

      // 顺序不能乱：先落掉累积的 op，再执行块命令 / 动作。
      flushOps()

      if (group === 'actions') {
        const outcome = await executeTool(
          editor,
          { id: 'docx-apply', name: String(command.op), input: { ...command } } as never,
          numIds,
        )
        if (outcome.isError)
          throw new BlockCommandError(
            'command.invalid',
            `commands[${index}]（${String(command.op)}）：${outcome.output}`,
          )
        applied += 1
        mutated ||= outcome.mutated
        continue
      }

      const planned = planBlockCommand(editor, command, index)
      planned.run()
      applied += 1
      mutated = true
    }
    flushOps()
  } catch (error) {
    // 整批回滚（快照是执行前的 PM JSON）。
    try {
      editor.commands.setContent(snapshot)
    } catch {
      /* 回滚失败就让错误原样上抛，页面会在下次保存时写入当前模型 */
    }
    throw error
  }

  return { applied, mutated, summary: `applied=${applied}` }
}
