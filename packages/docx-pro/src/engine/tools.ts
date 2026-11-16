// ─── engine.tools：主进程只声明（不导出执行函数） ──────────────────────
//
// 契约 §6.1 / §7.2：执行体由宿主生成——只读 → 执行点 plan；写 → 有页面转发
// 页面代理，无页面走权威层锁内 plan → exec → write。
//
// 四个工具：一张目录 + 一个读 + 一个新建 + 一个批量写。
// 55 条命令（blocks 8 / editor 23 / actions 10）全在 `docx_apply` 的 commands
// 里，schema 与说明按需用 `docx_guide` 取，不常驻在工具列表。
//
// 这个模块在**主进程**求值：只允许纯数据，不得 import 编辑器模块。

import type { ToolDecl } from '../tools/types'

const KINDS = ['docx', 'docm', 'dotx']

const REF = { type: 'string', description: '资源 id，如 file:///w/报告.docx' }

const guide: ToolDecl = {
  name: 'docx_guide',
  label: 'Word 命令目录',
  kinds: KINDS,
  readOnly: true,
  description:
    '取 .docx 的命令词汇。无参：三组 op 名与数量；group：该组全部 schema；op：单条 schema 与示例。' +
    '三组：blocks（块命令，无页面也能跑）、editor（格式/表格/域等 op，需打开页面）、actions（HTML 片段、图表、修订，需打开页面）。' +
    '写命令前先调它，不要凭记忆写 op。',
  tokens: 'docRev+epoch',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      group: { type: 'string', enum: ['blocks', 'editor', 'actions'] },
      op: { type: 'string', description: '单个 op 名' },
    },
    required: ['ref'],
  },
}

const read: ToolDecl = {
  name: 'docx_read',
  label: 'Word 读取',
  kinds: KINDS,
  readOnly: true,
  description:
    '读文档：块列表与统计（context）、块范围全文（blocks）、批注（comments）、脚注尾注（notes）、' +
    '修订（revisions）、样式（styles）、结构诊断（audit）。只读，不产字节。' +
    '块索引就是 docx_apply 的 index；改完重新读，不要复用旧坐标。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      what: {
        type: 'string',
        enum: ['context', 'blocks', 'comments', 'notes', 'revisions', 'styles', 'audit'],
        description: '要读什么（默认 context）',
      },
      startBlockIndex: { type: 'integer', description: 'blocks：起始块索引（含）' },
      endBlockIndex: { type: 'integer', description: 'blocks：结束块索引（含）' },
      offset: { type: 'integer', description: 'blocks：字符偏移，用于分段读长块' },
    },
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

const apply: ToolDecl = {
  name: 'docx_apply',
  label: 'Word 命令批',
  kinds: KINDS,
  readOnly: false,
  requiresBaseVersion: true,
  maxItems: 200,
  tokens: 'docRev+epoch',
  description:
    '按顺序执行一批命令（上限 200，op 见 docx_guide）。' +
    '块命令（setText / insertParagraph / setHeadingLevel / setTextStyle / deleteBlock / moveBlock / findReplace）' +
    '不需要页面；editor 与 actions 组需要文档页面打开，否则整批回 command.page-unavailable。' +
    '被拒时回执给出 index 与 code，按 code 处理（不要换通道绕）。' +
    '禁止用 write/edit 改 .docx。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      commands: {
        type: 'array',
        description: '命令批：{ op, ...字段 }；op 与字段用 docx_guide 取',
        items: { type: 'object' },
      },
      version: {
        type: 'string',
        description: '无页面写必填；有页面时由页面会话持有，可省略',
      },
      epoch: {
        type: 'number',
        description: 'agent 读到的基线代数；位置型命令在页面存在时用于 stale-index 判定',
      },
      dryRun: { type: 'boolean', description: '只校验，不落盘' },
    },
    required: ['ref', 'commands'],
  },
}

export const tools: ToolDecl[] = [guide, read, create, apply]

export default tools
