// ─── 统一命令目录（docx_guide 的唯一来源） ─────────────────────────────
//
// 三组命令，同一套 `docx_apply` 提交：
//
//   blocks  (8)  块级结构 / 文本 / 格式 —— 有页面映射成 PM 事务，无页面走 XML 切片
//   editor  (23) 编辑器 op 注册表 —— 需页面（apply_ops 的词汇）
//   actions (10) 专用动作（HTML 片段 / 图表 / 修订）—— 需页面
//
// 目录 = 可用面：没接线的动作不进目录（UNAVAILABLE_ACTIONS 只用于文档说明）。

import { GROUPS, OP_INDEX } from '../shared/catalog'
import { EDITOR_OPS } from './op-catalog.generated'
import { ACTION_COMMANDS, UNAVAILABLE_ACTIONS } from './commands.generated'

export const TARGET_SCHEMA = {
  type: 'object',
  description:
    '要改哪些块：nodeType / headingLevel / containsText / blockIndexes（顶层块下标）/ scope / range，至少给一个条件',
  properties: {
    nodeType: {
      type: 'string',
      enum: ['docHeading', 'docParagraph', 'docListItem', 'image', 'table'],
    },
    headingLevel: { type: 'integer', minimum: 1, maximum: 6 },
    containsText: { type: 'string' },
    matchCase: { type: 'boolean' },
    blockIndexes: { type: 'array', items: { type: 'integer', minimum: 0 } },
    scope: { type: 'string', enum: ['selection', 'document'] },
    range: {
      type: 'object',
      properties: { from: { type: 'number' }, to: { type: 'number' } },
      required: ['from', 'to'],
    },
  },
} as const

export const BLOCK_OP_NAMES = new Set(OP_INDEX.keys())
export const EDITOR_OP_NAMES = new Set(EDITOR_OPS.map((spec) => spec.name))
export const ACTION_NAMES = new Set(ACTION_COMMANDS.map((action) => action.name))

export const OP_COUNTS = {
  blocks: BLOCK_OP_NAMES.size,
  editor: EDITOR_OP_NAMES.size,
  actions: ACTION_NAMES.size,
  total: BLOCK_OP_NAMES.size + EDITOR_OP_NAMES.size + ACTION_NAMES.size,
}

export type OpGroup = 'blocks' | 'editor' | 'actions'

export function groupOf(op: string): OpGroup | null {
  if (BLOCK_OP_NAMES.has(op)) return 'blocks'
  if (EDITOR_OP_NAMES.has(op)) return 'editor'
  if (ACTION_NAMES.has(op)) return 'actions'
  return null
}

export interface GuideInput {
  version?: string
  epoch?: number
  group?: string
  op?: string
}

export interface GuideResult {
  /** 回给模型的文本（宿主的工具摘要只用字符串）。 */
  summary: string
  /** 结构化版本（进调用详情）。 */
  data: unknown
}

const short = (version?: string): string =>
  version ? `version=${version.length > 12 ? `${version.slice(0, 12)}…` : version}` : ''

const oneLine = (text: string, max = 160): string => {
  const flat = text.replace(/\s+/g, ' ').trim()
  return flat.length > max ? `${flat.slice(0, max)}…` : flat
}

/** 三级发现：无参 → 分组；group → 该组全部 schema；op → 单条。 */
export function buildCommandGuide(input: GuideInput): GuideResult {
  const header = [
    '# docx 命令目录',
    short(input.version),
    input.epoch !== undefined ? `epoch=${input.epoch}` : '',
    `ops=${OP_COUNTS.total} (blocks=${OP_COUNTS.blocks} editor=${OP_COUNTS.editor} actions=${OP_COUNTS.actions})`,
  ]
    .filter(Boolean)
    .join(' ')

  if (input.op) return opGuide(header, input.op)

  if (input.group) {
    const group = input.group as OpGroup
    if (group === 'blocks') return blockGroupGuide(header)
    if (group === 'editor') return editorGroupGuide(header)
    if (group === 'actions') return actionGroupGuide(header)
    // 兼容块命令的原始子分组：blocks / text / format / check
    const legacy = GROUPS.find((item) => item.name === input.group)
    if (legacy) return legacyGroupGuide(header, legacy.name)
    return {
      summary: `${header}\n未知分组：${input.group}（可用：blocks / editor / actions）`,
      data: { error: 'group.not-found' },
    }
  }

  const summary = [
    header,
    '',
    '## blocks · 块命令（需页面；8）',
    GROUPS.flatMap((group) => group.ops)
      .map((spec) => `- ${spec.op} (${spec.level}) ${spec.summary}`)
      .join('\n'),
    '',
    '## editor · 编辑器 op（需打开页面；23）',
    EDITOR_OPS.map((spec) => `- ${spec.name} target=${spec.target}`).join('\n'),
    '',
    '## actions · 专用动作（需打开页面；10）',
    ACTION_COMMANDS.map((action) => `- ${action.name} — ${oneLine(action.description, 90)}`).join('\n'),
    '',
    UNAVAILABLE_ACTIONS.length > 0
      ? `未接线（不进目录）：${UNAVAILABLE_ACTIONS.join(', ')}`
      : '',
    '',
    '写入只允许在已打开的文档页执行；没有页面时 apply 直接拒绝。',
    `用法：command docx.apply，input.commands 是 JSON 字符串，例如 ${JSON.stringify([{ op: 'setText', index: 3, text: '新标题' }])}`,
    '下一步：guide({ group }) 取该组 schema；guide({ op }) 取单条。写前先 dryRun:true。',
  ]
    .filter((line) => line !== '')
    .join('\n')

  return {
    summary,
    data: {
      ops: OP_COUNTS,
      blocks: [...BLOCK_OP_NAMES],
      editor: [...EDITOR_OP_NAMES],
      actions: [...ACTION_NAMES],
      unavailable: [...UNAVAILABLE_ACTIONS],
    },
  }
}

function opGuide(header: string, op: string): GuideResult {
  const block = OP_INDEX.get(op)
  if (block) {
    return {
      summary: [
        header,
        `## ${block.spec.op} · blocks · ${block.group.name}`,
        block.spec.summary,
        `level=${block.spec.level}（位置型：改完重新读，不要复用旧坐标）`,
        'schema:',
        JSON.stringify(block.spec.schema, null, 2),
        'example:',
        JSON.stringify(block.spec.example, null, 2),
      ].join('\n'),
      data: { op: block.spec, group: 'blocks' },
    }
  }

  const editor = EDITOR_OPS.find((spec) => spec.name === op)
  if (editor) {
    return {
      summary: [
        header,
        `## ${op} · editor`,
        'signature:',
        editor.signature,
        `target: ${editor.target}`,
        `fields: ${editor.keys.join(', ') || '（无）'}`,
        'example:',
        JSON.stringify(exampleOf(op), null, 2),
        '说明：一批 op 在页面上是一个事务、一次撤销；页面未打开时 apply 直接拒绝。',
      ].join('\n'),
      data: { op: editor, group: 'editor' },
    }
  }

  const action = ACTION_COMMANDS.find((item) => item.name === op)
  if (action) {
    return {
      summary: [
        header,
        `## ${op} · actions`,
        action.description,
        'parameters:',
        JSON.stringify(action.parameters, null, 2),
        '说明：动作按命令顺序执行，需要打开页面；结果只回 applied，需要内容时用 docx.read。',
      ].join('\n'),
      data: { op: { name: action.name }, group: 'actions' },
    }
  }

  return {
    summary: `${header}\n未知 op：${op}（blocks / editor / actions 见无参调用）`,
    data: { error: 'op.not-found' },
  }
}

function blockGroupGuide(header: string): GuideResult {
  const lines = [header, '## blocks · 块命令（需页面）']
  for (const group of GROUPS) {
    lines.push(`### ${group.name} · ${group.summary}`)
    for (const spec of group.ops) {
      lines.push(`- ${spec.op} (${spec.level}) ${spec.summary}`)
      lines.push(`  schema: ${JSON.stringify(spec.schema)}`)
      lines.push(`  example: ${JSON.stringify(spec.example)}`)
    }
  }
  return { summary: lines.join('\n'), data: { group: 'blocks' } }
}

function legacyGroupGuide(header: string, name: string): GuideResult {
  const group = GROUPS.find((item) => item.name === name)!
  const lines = [header, `## ${group.name} · ${group.summary}`]
  for (const spec of group.ops) {
    lines.push(`### ${spec.op} (${spec.level})`, spec.summary, JSON.stringify(spec.schema))
  }
  return { summary: lines.join('\n'), data: { group } }
}

function editorGroupGuide(header: string): GuideResult {
  const lines = [
    header,
    '## editor · 编辑器 op（需页面；target 必填则必须给 target）',
  ]
  for (const spec of EDITOR_OPS) {
    lines.push(`- ${spec.name} target=${spec.target}`)
    lines.push(`  ${spec.signature}`)
  }
  lines.push('', '单条示例：guide({ op: "setFont" })')
  return { summary: lines.join('\n'), data: { group: 'editor' } }
}

function actionGroupGuide(header: string): GuideResult {
  const lines = [header, '## actions · 专用动作（需页面）']
  for (const action of ACTION_COMMANDS) {
    lines.push(`- ${action.name} — ${oneLine(action.description, 120)}`)
    lines.push(`  parameters: ${JSON.stringify(action.parameters)}`)
  }
  return { summary: lines.join('\n'), data: { group: 'actions' } }
}

function exampleOf(name: string): Record<string, unknown> {
  const target = { blockIndexes: [0] }
  switch (name) {
    case 'setFont':
      return { op: 'setFont', target, bold: true }
    case 'setMatchedFont':
      return { op: 'setMatchedFont', text: '旧文字', bold: true }
    case 'setParagraphFormat':
      return { op: 'setParagraphFormat', target, lineSpacing: 1.5 }
    case 'setHeadingLevel':
      return { op: 'setHeadingLevel', target, level: 2 }
    case 'findReplace':
      return { op: 'findReplace', find: '旧', replace: '新' }
    case 'deleteBlocks':
      return { op: 'deleteBlocks', target }
    case 'moveBlocks':
      return { op: 'moveBlocks', blockIndexes: [2], afterBlockIndex: 5 }
    case 'insertToc':
      return { op: 'insertToc', afterBlockIndex: 2 }
    case 'updateFields':
      return { op: 'updateFields' }
    default:
      return { op: name, target }
  }
}
