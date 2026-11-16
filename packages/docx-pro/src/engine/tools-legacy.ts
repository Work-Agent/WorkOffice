// ─── engine.tools：主进程只声明（不导出执行函数） ──────────────────────
//
// 契约 §6.1 / §7.2：执行体由宿主生成——只读 → 执行点 plan；写 → 有页面转发
// 页面代理，无页面走权威层锁内 plan → exec → write。
//
// 声明分三类：
//   1. 文档编辑器的原生工具（READ_TOOLS / WRITE_TOOLS，由 AGENT_TOOLS 机械转换）
//   2. 无页面也能跑的块命令（docx_edit）与命令目录（docx_guide）
//   3. 只读诊断（docx_check）
//
// 这个模块在**主进程**求值：只允许纯数据，不得 import 编辑器模块。

import { GENERATED_TOOLS, KINDS, READ_TOOLS, WRITE_TOOLS } from '../tools/declarations.generated'
import type { ToolDecl } from '../tools/types'

const REF = { type: 'string', description: '资源 id，如 file:///w/报告.docx' }
const VERSION = {
  type: 'string',
  description: '无页面写必填；有页面时由页面会话持有，可省略',
}
const EPOCH = {
  type: 'number',
  description: 'agent 读到的基线代数；位置型命令在页面存在时用于 stale-index 判定',
}

const guide: ToolDecl = {
  name: 'docx_guide',
  label: 'Word 命令目录',
  kinds: KINDS,
  readOnly: true,
  description:
    '取 .docx 的命令词汇。无参：分组 + op 名；group：该组 schema；op：单条 schema 与示例。' +
    '两个来源：block 组（docx_edit 的块命令，无页面也能跑）与 editor 组（apply_ops 的 op，需打开页面）。' +
    '写命令前先调它，不要凭记忆写 op。',
  tokens: 'docRev+epoch',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      group: {
        type: 'string',
        description: '分组名：blocks / text / format / check / editor',
      },
      op: { type: 'string', description: '单个 op 名' },
    },
    required: ['ref'],
  },
}

const check: ToolDecl = {
  name: 'docx_check',
  label: 'Word 结构诊断',
  kinds: KINDS,
  readOnly: true,
  description:
    '只读诊断：块 / 段落 / 表格 / 标题 / 空段落计数与正文规模。不产字节、不落盘，适合在写之前确认文档结构。',
  parameters: {
    type: 'object',
    properties: { ref: REF },
    required: ['ref'],
  },
}

const create: ToolDecl = {
  name: 'docx_create',
  label: 'Word 新建',
  kinds: KINDS,
  readOnly: false,
  requiresBaseVersion: false,
  description:
    '从零新建 .docx。ref 是尚不存在的目标资源 id；content 是 Markdown 全文。' +
    'content 为空时创建空白文档。支持标题、段落、列表、表格、粗体/斜体/删除线、代码块、引用。' +
    '目标已存在时拒绝；修改已有文档请用 docx_apply。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      content: {
        type: 'string',
        description: 'Markdown 文档内容；空字符串 = 空白文档',
      },
    },
    required: ['ref', 'content'],
  },
}

const edit: ToolDecl = {
  name: 'docx_edit',
  label: 'Word 块命令批',
  kinds: KINDS,
  readOnly: false,
  requiresBaseVersion: true,
  maxItems: 200,
  tokens: 'docRev+epoch',
  description:
    '原子执行一批块级命令（上限 200）。无页面时由执行点按 XML 切片写回（未改块逐字节不变）；' +
    '有页面时同样的命令映射到编辑器事务。一条被拒则整批不落盘，回执给出 index 与原因。' +
    '只需要改文本 / 标题 / 格式 / 结构时优先用它；需要 run 级、批注、图片、页眉页脚等编辑器能力时用对应工具。' +
    'index 用 read_blocks / get_document_context 得到的块索引；改完重新读，不要复用旧坐标。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      commands: {
        type: 'array',
        description: '块命令批；op 见 docx_guide 的 blocks / text / format 组',
        items: { type: 'object' },
      },
      version: VERSION,
      epoch: EPOCH,
      dryRun: { type: 'boolean', description: '只校验，不落盘' },
    },
    required: ['ref', 'commands'],
  },
}

export const tools: ToolDecl[] = [
  guide,
  check,
  create,
  edit,
  ...READ_TOOLS,
  ...WRITE_TOOLS,
]

export { GENERATED_TOOLS, READ_TOOLS, WRITE_TOOLS }
export default tools
