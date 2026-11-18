// ─── engine.tools：主进程只声明（不导出执行函数） ──────────────────────
//
// 4 个工具：一张目录 + 一个读 + 一个新建 + 一个批量写。
// 目录里的 63 条命令（上游 62 条 AI 可调用 + editChart）按需用 pptx_guide 取。

import type { ToolDecl } from '../tools/types'

const KINDS = ['pptx', 'pptm', 'potx']
const REF = { type: 'string', description: '资源 id，如 file:///w/2026年度汇报.pptx' }

const guide: ToolDecl = {
  name: 'pptx_guide',
  label: 'PPT 命令目录',
  kinds: KINDS,
  readOnly: true,
  description:
    '取 .pptx 的命令词汇。无参：分组与数量；group：该组全部 schema；op：单条 schema 与示例。' +
    'text / element / table / slide / deck / chart 多数无页面可执行；图片/媒体/剪贴板/3D/旁白需页面。' +
    '写命令前先调它，不要凭记忆写 op。',
  tokens: 'docRev',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      group: { type: 'string', enum: ['text', 'element', 'insert', 'table', 'slide', 'deck', 'chart'] },
      op: { type: 'string', description: '单个 op 名' },
    },
    required: ['ref'],
  },
}

const read: ToolDecl = {
  name: 'pptx_read',
  label: 'PPT 读取',
  kinds: KINDS,
  readOnly: true,
  description:
    '读演示文稿：概览与页列表（context / outline / slides）、单页元素（slide）、备注（notes）、' +
    '批注（comments）、图表（charts）、表格（tables）、主题（theme）、结构诊断（audit）、' +
    '全文查找（find）。只读，不产字节。target.slide 用 s_* 或 0-based 索引，target.el 用 e_*。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      what: {
        type: 'string',
        enum: ['context', 'outline', 'slides', 'slide', 'notes', 'comments', 'charts', 'tables', 'theme', 'audit', 'find'],
        description: '要读什么（默认 context）',
      },
      slide: { type: 'integer', description: 'slide：0-based 页索引' },
      query: { type: 'string', description: 'find：查询文本' },
      limit: { type: 'integer', description: 'find：上限' },
    },
    required: ['ref'],
  },
}

const create: ToolDecl = {
  name: 'pptx_create',
  label: 'PPT 新建',
  kinds: KINDS,
  readOnly: false,
  requiresBaseVersion: false,
  description:
    '从零新建 .pptx。ref 是尚不存在的目标资源 id；slides 是 [{title,body?,bullets?,notes?,background?}]，' +
    '也可以用 Markdown content（# 页标题、--- 分页、- 项目符号、notes: 备注）。' +
    '空内容 = 空白演示文稿。目标已存在时拒绝；修改已有演示文稿请用 pptx_apply。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      slides: {
        type: 'array',
        description: '页规格列表',
        items: {
          type: 'object',
          properties: {
            title: { type: 'string' },
            body: { oneOf: [{ type: 'string' }, { type: 'array', items: { type: 'string' } }] },
            bullets: { type: 'array', items: { type: 'string' } },
            notes: { type: 'string' },
            background: { type: 'string', description: '#RRGGBB' },
          },
        },
      },
      content: { type: 'string', description: '轻量 Markdown' },
      theme: { type: 'object', description: 'applyTheme 的 {name,colors,majorFont?,minorFont?}' },
      size: { oneOf: [{ type: 'string', enum: ['16:9', '4:3'] }, { type: 'object' }] },
    },
    required: ['ref'],
  },
}

const apply: ToolDecl = {
  name: 'pptx_apply',
  label: 'PPT 命令批',
  kinds: KINDS,
  readOnly: false,
  requiresBaseVersion: true,
  maxItems: 200,
  tokens: 'docRev',
  description:
    '按顺序执行一批命令（上限 200，op 见 pptx_guide）。' +
    '多数 text / element / insert / table / slide / deck / chart 命令不需要页面；' +
    '图片/媒体/剪贴板/3D/旁白需要页面，否则整批回 command.page-unavailable。' +
    '默认 atomic：一条被拒整批回滚；isolation:"per_op" 可逐条保留。' +
    '被拒时按回执的 index + code 修那一条后重发。禁止用 write/edit 改 .pptx。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      commands: {
        type: 'array',
        description: '命令批：{ op, target?, ...字段 }；op 与字段用 pptx_guide 取',
        items: { type: 'object' },
      },
      version: {
        type: 'string',
        description: '无页面写必填；有页面时由页面会话持有，可省略',
      },
      docRev: {
        type: 'string',
        description: 'agent 读到的页面模型代数；页面存在时用于 stale-index 判定',
      },
      epoch: {
        type: 'number',
        description: '宿主透传的基线代数；pptx 不消费（保留兼容）',
      },
      isolation: {
        type: 'string',
        enum: ['atomic', 'per_op'],
        description: '默认 atomic；per_op 逐条保留成功项',
      },
      dryRun: { type: 'boolean', description: '只校验，不落盘' },
    },
    required: ['ref', 'commands'],
  },
}

export const tools: ToolDecl[] = [guide, read, create, apply]

export default tools
