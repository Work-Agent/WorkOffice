// ─── engine.tools：主进程只声明（不导出执行函数） ──────────────────────
//
// 四个工具：一张目录 + 一个读 + 一个新建 + 一个批量写。
// cells / structure / sheets 无页面可执行；objects 需打开页面。

import type { ToolDecl } from '../tools/types'

const KINDS = ['xlsx', 'xlsm', 'xltx']
const REF = { type: 'string', description: '资源 id，如 file:///w/预算.xlsx' }

const guide: ToolDecl = {
  name: 'xlsx_guide',
  label: 'Excel 命令目录',
  kinds: KINDS,
  readOnly: true,
  description:
    '取 .xlsx 的命令词汇。无参：四组 op 名与数量；group：该组全部 schema；op：单条 schema 与示例。' +
    'cells / structure / sheets 无页面可执行；objects（图表/表格/透视表/条件格式等）需要打开页面。' +
    '写命令前先调它，不要凭记忆写 op。',
  tokens: 'docRev+structureEpoch',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      group: { type: 'string', enum: ['cells', 'structure', 'sheets', 'objects'] },
      op: { type: 'string', description: '单个 op 名' },
    },
    required: ['ref'],
  },
}

const read: ToolDecl = {
  name: 'xlsx_read',
  label: 'Excel 读取',
  kinds: KINDS,
  readOnly: true,
  description:
    '读工作簿：概览与工作表清单（context）、工作表内容预览（sheets）、区域值（range）、' +
    '公式列表（formulas）、查找（find）、包结构诊断（audit）。只读，不产字节。' +
    '有页面时读页面 live 模型（含未保存修改）；无页面时读落盘字节。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      what: {
        type: 'string',
        enum: ['context', 'sheets', 'range', 'formulas', 'find', 'audit', 'cells', 'meta'],
        description:
          '要读什么（默认 context）。cells / meta 是页面 invoke 用的结构化读（需要 wasm 引擎）',
      },
      sheet: { type: 'string', description: '工作表名称（range / sheets / formulas 可用）' },
      range: { type: 'string', description: 'A1 范围，如 A1:F30' },
      start: { type: 'string', description: 'range 的左上角' },
      end: { type: 'string', description: 'range 的右下角' },
      query: { type: 'string', description: 'find 的查询文本' },
      regex: { type: 'boolean', description: 'find 是否按正则' },
      matchCase: { type: 'boolean' },
      limit: { type: 'integer', description: 'find / formulas 上限' },
    },
    required: ['ref'],
  },
}

const create: ToolDecl = {
  name: 'xlsx_create',
  label: 'Excel 新建',
  kinds: KINDS,
  readOnly: false,
  requiresBaseVersion: false,
  description:
    '从零新建 .xlsx。ref 是尚不存在的目标资源 id。' +
    'sheets 是 [{ name, rows }]，rows 是二维数组（字符串以 = 开头写成公式）；' +
    '也可以用 rows 建单表，或用 csv 文本建单表。空内容 = 空白工作簿。' +
    '目标已存在时拒绝；修改已有工作簿请用 xlsx_apply。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      sheets: {
        type: 'array',
        description: '工作表列表',
        items: {
          type: 'object',
          properties: {
            name: { type: 'string' },
            rows: { type: 'array', items: { type: 'array', items: {} } },
          },
          required: ['name'],
        },
      },
      name: { type: 'string', description: '单表名称（用 rows / csv 时）' },
      rows: { type: 'array', description: '单表二维数组', items: { type: 'array', items: {} } },
      csv: { type: 'string', description: '单表 CSV 文本' },
    },
    required: ['ref'],
  },
}

const apply: ToolDecl = {
  name: 'xlsx_apply',
  label: 'Excel 命令批',
  kinds: KINDS,
  readOnly: false,
  requiresBaseVersion: true,
  maxItems: 200,
  tokens: 'docRev+structureEpoch',
  description:
    '按顺序执行一批命令（上限 200，op 见 xlsx_guide）。' +
    'cells / structure / sheets 组不需要页面；objects 组需要工作簿页面打开，否则整批回 command.page-unavailable。' +
    '有页面时命令进页面会话：含 A1 坐标的 op 需带 structureEpoch，读内容型 op（fill/copy/sort/find）需带 docRev。' +
    '被拒时回执给出 index 与 code，按 code 处理（不要换通道绕）。' +
    '禁止用 write/edit 改 .xlsx。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      commands: {
        type: 'array',
        description: '命令批：{ op, sheet, ...字段 }；op 与字段用 xlsx_guide 取',
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
      revision: {
        type: 'number',
        description: '页面内 Univer revision；页面代理用它判断读到的是不是同一份内容',
      },
      dryRun: { type: 'boolean', description: '只校验，不落盘' },
    },
    required: ['ref', 'commands'],
  },
}

export const tools: ToolDecl[] = [guide, read, create, apply]

export default tools
