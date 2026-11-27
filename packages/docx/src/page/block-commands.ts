// ─── 块命令 → 编辑器事务（页面路径） ───────────────────────────────────
//
// 与执行点的 `shared/ops.ts` 是同一套命令词汇（docx_guide 的 blocks 组），
// 这里把它映射到 ProseMirror：
//   - setText / insertParagraph 直接用 PM（保留段落属性与节点类型）
//   - 其余走编辑器的 op 注册表（executeOps：一条 = 一个事务 = 一次撤销）
//
// 拆成「plan → run」两步是为了配合 `docx_apply` 的顺序语义：块索引会随结构
// 命令变化，所以每条命令必须在它自己的位置上校验，不能整批预先校验。

import type { Editor } from '@tiptap/core'
import { executeOps } from '../editor/ai/ops'
import { blockRangePositions } from '../editor/ai/doc-utils'

export class BlockCommandError extends Error {
  readonly code: string
  constructor(code: string, message: string) {
    super(message)
    this.name = 'BlockCommandError'
    this.code = code
  }
}

export interface BlockCommandResult {
  /** 成功应用的命令条数（回执 applied 的取值）。 */
  applied: number
  /** 给模型的摘要（页面代理不把它带回模型，只进调用详情）。 */
  summary: string
  /** 是否为写操作（决定要不要提交）。 */
  mutated: boolean
}

export interface PlannedBlockCommand {
  /** 校验已通过，执行这一步。 */
  run(): void
}

interface Command {
  op?: unknown
  [key: string]: unknown
}

export const MAX_COMMANDS = 200

function intOf(value: unknown, field: string): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0)
    throw new BlockCommandError('command.invalid', `${field} 必须是非负整数`)
  return value
}

function nodeAt(editor: Editor, index: number) {
  const doc = editor.state.doc
  if (index < 0 || index >= doc.childCount)
    throw new BlockCommandError(
      'block.not-found',
      `块索引越界：${index}（文档共 ${doc.childCount} 个顶层块）`,
    )
  return doc.child(index)
}

/** 校验一条块命令，返回可执行的闭包（不修改文档）。 */
export function planBlockCommand(
  editor: Editor,
  command: Command,
  at: number,
): PlannedBlockCommand {
  const op = String(command.op ?? '')
  const where = `commands[${at}]`

  switch (op) {
    case 'setText': {
      const index = intOf(command.index, `${where}.index`)
      const text = typeof command.text === 'string' ? command.text : null
      if (text === null)
        throw new BlockCommandError('command.invalid', `${where}.text 必须是字符串`)
      const node = nodeAt(editor, index)
      if (node.type.name === 'table' || node.type.name === 'image')
        throw new BlockCommandError(
          'block.not-paragraph',
          `块 ${index} 是 ${node.type.name}，不能整块改文本`,
        )
      return {
        run: () => {
          const { from, to } = blockRangePositions(editor, index, index)
          editor.view.dispatch(editor.state.tr.insertText(text, from + 1, to - 1))
        },
      }
    }

    case 'insertParagraph': {
      const text = typeof command.text === 'string' ? command.text : ''
      const heading = Number(command.heading ?? 0)
      if (!Number.isInteger(heading) || heading < 0 || heading > 6)
        throw new BlockCommandError('command.invalid', `${where}.heading 必须是 0..6`)
      const after =
        command.after === null || command.after === undefined
          ? null
          : intOf(command.after, `${where}.after`)
      if (after !== null) nodeAt(editor, after)
      return {
        run: () => {
          const schema = editor.schema
          const type = heading > 0 ? schema.nodes['docHeading'] : schema.nodes['docParagraph']
          if (!type)
            throw new BlockCommandError('command.unsupported', '编辑器缺少段落节点类型')
          const node = type.create(
            heading > 0 ? { level: heading } : null,
            text ? schema.text(text) : null,
          )
          const pos = after === null ? 0 : blockRangePositions(editor, after, after).to
          editor.chain().insertContentAt(pos, node).run()
        },
      }
    }

    case 'deleteBlock': {
      const index = intOf(command.index, `${where}.index`)
      nodeAt(editor, index)
      return opsCommand(editor, [{ op: 'deleteBlocks', target: { blockIndexes: [index] } }])
    }

    case 'moveBlock': {
      const index = intOf(command.index, `${where}.index`)
      nodeAt(editor, index)
      const to =
        command.to === null || command.to === undefined
          ? null
          : intOf(command.to, `${where}.to`)
      const last = editor.state.doc.childCount - 1
      return opsCommand(editor, [
        {
          op: 'moveBlocks',
          blockIndexes: [index],
          afterBlockIndex: to === null ? -1 : Math.min(to, last),
        },
      ])
    }

    case 'setHeadingLevel': {
      const index = intOf(command.index, `${where}.index`)
      nodeAt(editor, index)
      const level = Number(command.level ?? 0)
      if (!Number.isInteger(level) || level < 0 || level > 6)
        throw new BlockCommandError('command.invalid', `${where}.level 必须是 0..6`)
      return opsCommand(editor, [
        { op: 'setHeadingLevel', target: { blockIndexes: [index] }, level },
      ])
    }

    case 'setTextStyle': {
      const index = intOf(command.index, `${where}.index`)
      nodeAt(editor, index)
      const target = { blockIndexes: [index] }
      const font: Record<string, unknown> = { op: 'setFont', target: { ...target } }
      for (const key of ['bold', 'italic', 'underline', 'color'] as const)
        if (command[key] !== undefined) font[key] = command[key]
      if (command.size !== undefined) font['fontSize'] = command.size
      const batch: Array<Record<string, unknown>> = []
      if (Object.keys(font).length > 2) batch.push(font)
      if (command.align !== undefined)
        batch.push({ op: 'setParagraphFormat', target: { ...target }, align: command.align })
      if (batch.length === 0)
        throw new BlockCommandError('command.invalid', `${where} 没有可应用的样式字段`)
      return opsCommand(editor, batch)
    }

    case 'findReplace': {
      const find = typeof command.find === 'string' ? command.find : ''
      const replace = typeof command.replace === 'string' ? command.replace : ''
      if (!find) throw new BlockCommandError('command.invalid', `${where}.find 不能为空`)
      return opsCommand(editor, [
        {
          op: 'findReplace',
          find,
          replace,
          ...(command.matchCase === true ? { matchCase: true } : {}),
        },
      ])
    }

    case 'check':
      throw new BlockCommandError(
        'command.unsupported',
        'check 是只读诊断，请用 docx_read({ what: "audit" })',
      )

    default:
      throw new BlockCommandError(
        'command.not-found',
        `未知块命令：${op || '(空)'}（用 docx_guide 取可用 op）`,
      )
  }
}

function opsCommand(
  editor: Editor,
  ops: Array<Record<string, unknown>>,
): PlannedBlockCommand {
  return {
    run: (): void => {
      const outcome = executeOps(editor, ops)
      if (!outcome.ok)
        throw new BlockCommandError('command.invalid', outcome.error ?? '命令被拒绝')
    },
  }
}

/** 块命令批（page 侧的旧入口；顺序校验 + 顺序执行）。 */
export function applyBlockCommands(
  editor: Editor,
  raw: unknown,
): BlockCommandResult {
  if (!Array.isArray(raw)) throw new BlockCommandError('command.invalid', 'commands 必须是数组')
  if (raw.length === 0) return { applied: 0, summary: '没有命令', mutated: false }
  if (raw.length > MAX_COMMANDS)
    throw new BlockCommandError(
      'command.too-large',
      `批上限 ${MAX_COMMANDS} 条（收到 ${raw.length}）`,
    )
  for (const [index, command] of (raw as Command[]).entries()) {
    const planned = planBlockCommand(editor, command, index)
    planned.run()
  }
  return {
    applied: raw.length,
    summary: `applied=${raw.length}（blocks）`,
    mutated: true,
  }
}
