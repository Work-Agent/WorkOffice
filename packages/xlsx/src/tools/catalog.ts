// ─── 统一命令目录（xlsx_guide 的唯一来源） ─────────────────────────────
//
// cells / structure / sheets 是无页面可执行组；objects 必须打开页面。
// 目录就是能力面：不在目录里的 op 不存在，模型不会猜。

import { opTier } from '../shared/op-tiers'

export type OpGroup = 'cells' | 'structure' | 'sheets' | 'objects'

export interface OpSpec {
  op: string
  group: OpGroup
  headless: boolean
  summary: string
  schema: unknown
  example: unknown
}

const S = (type: string, description?: string) =>
  description ? { type, description } : { type }

const sheet = S('string', '工作表名称（不要用会话 id）')
const range = S('string', 'A1 范围，如 A1:C10')

const cellOps: OpSpec[] = [
  {
    op: 'set_cell',
    group: 'cells',
    headless: true,
    summary: '写一个单元格的值；字符串以 = 开头时写成公式',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'address', 'value'],
      properties: { op: S('string'), sheet, address: S('string', '如 B2'), value: {} },
    },
    example: { op: 'set_cell', sheet: 'Sheet1', address: 'B2', value: 1200 },
  },
  {
    op: 'set_formula',
    group: 'cells',
    headless: true,
    summary: '写一个公式（必须带 =）',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'address', 'formula'],
      properties: { op: S('string'), sheet, address: S('string'), formula: S('string') },
    },
    example: { op: 'set_formula', sheet: 'Sheet1', address: 'C2', formula: '=SUM(B2:B4)' },
  },
  {
    op: 'clear_cell',
    group: 'cells',
    headless: true,
    summary: '清空一个单元格的内容',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'address'],
      properties: { op: S('string'), sheet, address: S('string') },
    },
    example: { op: 'clear_cell', sheet: 'Sheet1', address: 'C2' },
  },
  {
    op: 'set_range',
    group: 'cells',
    headless: true,
    summary: '按矩形二维数组写一段区域；null 清空该格',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'start', 'values'],
      properties: {
        op: S('string'),
        sheet,
        start: S('string', '左上角，如 A1'),
        values: { type: 'array', items: { type: 'array', items: {} } },
      },
    },
    example: { op: 'set_range', sheet: 'Sheet1', start: 'A1', values: [['姓名', '金额'], ['张三', 1200]] },
  },
  {
    op: 'clear_range',
    group: 'cells',
    headless: true,
    summary: '清空一段区域的内容',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'range'],
      properties: { op: S('string'), sheet, range },
    },
    example: { op: 'clear_range', sheet: 'Sheet1', range: 'A2:C10' },
  },
  {
    op: 'format_range',
    group: 'cells',
    headless: true,
    summary: '设置区域格式（粗体、颜色、数字格式、对齐等）',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'range', 'format'],
      properties: {
        op: S('string'),
        sheet,
        range,
        format: {
          type: 'object',
          properties: {
            bold: S('boolean'), italic: S('boolean'), underline: S('boolean'),
            strikethrough: S('boolean'), fontFamily: S('string'), fontSize: S('number'),
            fontColor: S('string'), fillColor: S('string'), numberFormat: S('string'),
            horizontalAlign: S('string'), verticalAlign: S('string'),
            wrapText: S('boolean'), indent: S('number'),
          },
        },
      },
    },
    example: { op: 'format_range', sheet: 'Sheet1', range: 'A1:B1', format: { bold: true, fillColor: '#DDEBF7' } },
  },
  {
    op: 'find_replace',
    group: 'cells',
    headless: true,
    summary: '在范围内查找并替换文本',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'range', 'find', 'replace'],
      properties: {
        op: S('string'), sheet, range, find: S('string'), replace: S('string'),
        matchCase: S('boolean'), wholeCell: S('boolean'),
      },
    },
    example: { op: 'find_replace', sheet: 'Sheet1', range: 'A1:C20', find: '旧值', replace: '新值' },
  },
  {
    op: 'fill_range',
    group: 'cells',
    headless: true,
    summary: '把源区域平铺/填充到目标区域（公式相对引用会偏移）',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'source', 'target'],
      properties: { op: S('string'), sheet, source: S('string'), target: S('string') },
    },
    example: { op: 'fill_range', sheet: 'Sheet1', source: 'B2', target: 'B2:B10' },
  },
  {
    op: 'copy_range',
    group: 'cells',
    headless: true,
    summary: '复制一块区域到目标位置：值/公式 + 样式 + 合并区 + 行高（公式相对引用会偏移）',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'source', 'target'],
      properties: { op: S('string'), sheet, source: S('string'), target: S('string') },
    },
    example: { op: 'copy_range', sheet: 'Sheet1', source: 'A1:B3', target: 'D1' },
  },
  {
    op: 'move_range',
    group: 'cells',
    headless: true,
    summary: '剪切一块区域到目标位置并清空源区域（样式/合并/行高跟着走）',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'source', 'target'],
      properties: { op: S('string'), sheet, source: S('string'), target: S('string') },
    },
    example: { op: 'move_range', sheet: 'Sheet1', source: 'A1:B3', target: 'D1' },
  },
  {
    op: 'sort_range',
    group: 'cells',
    headless: true,
    summary: '按某列对区域排序（可含表头）',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'range', 'byColumn', 'order'],
      properties: {
        op: S('string'), sheet, range, byColumn: S('string', '列字母，如 B'),
        order: { type: 'string', enum: ['asc', 'desc'] }, hasHeader: S('boolean'),
      },
    },
    example: { op: 'sort_range', sheet: 'Sheet1', range: 'A1:C20', byColumn: 'B', order: 'desc', hasHeader: true },
  },
  {
    op: 'convert_to_values',
    group: 'cells',
    headless: false,
    summary: '把公式冻结成当前计算值（需要页面公式引擎）',
    schema: {
      type: 'object',
      required: ['op', 'sheet', 'range'],
      properties: { op: S('string'), sheet, range },
    },
    example: { op: 'convert_to_values', sheet: 'Sheet1', range: 'C2:C20' },
  },
]

const structureOps: OpSpec[] = [
  {
    op: 'insert_rows', group: 'structure', headless: true,
    summary: '在指定行之前插入若干行',
    schema: { type: 'object', required: ['op', 'sheet', 'row', 'count'], properties: { op: S('string'), sheet, row: S('number', '1-based'), count: S('number') } },
    example: { op: 'insert_rows', sheet: 'Sheet1', row: 3, count: 1 },
  },
  {
    op: 'delete_rows', group: 'structure', headless: true,
    summary: '删除若干行',
    schema: { type: 'object', required: ['op', 'sheet', 'row', 'count'], properties: { op: S('string'), sheet, row: S('number'), count: S('number') } },
    example: { op: 'delete_rows', sheet: 'Sheet1', row: 3, count: 1 },
  },
  {
    op: 'insert_cols', group: 'structure', headless: true,
    summary: '在指定列之前插入若干列',
    schema: { type: 'object', required: ['op', 'sheet', 'column', 'count'], properties: { op: S('string'), sheet, column: S('string'), count: S('number') } },
    example: { op: 'insert_cols', sheet: 'Sheet1', column: 'C', count: 1 },
  },
  {
    op: 'delete_cols', group: 'structure', headless: true,
    summary: '删除若干列',
    schema: { type: 'object', required: ['op', 'sheet', 'column', 'count'], properties: { op: S('string'), sheet, column: S('string'), count: S('number') } },
    example: { op: 'delete_cols', sheet: 'Sheet1', column: 'C', count: 1 },
  },
  {
    op: 'merge_cells', group: 'structure', headless: true,
    summary: '合并单元格区域',
    schema: { type: 'object', required: ['op', 'sheet', 'range'], properties: { op: S('string'), sheet, range } },
    example: { op: 'merge_cells', sheet: 'Sheet1', range: 'A1:C1' },
  },
  {
    op: 'unmerge_cells', group: 'structure', headless: true,
    summary: '取消合并单元格区域',
    schema: { type: 'object', required: ['op', 'sheet', 'range'], properties: { op: S('string'), sheet, range } },
    example: { op: 'unmerge_cells', sheet: 'Sheet1', range: 'A1:C1' },
  },
  {
    op: 'set_row_height', group: 'structure', headless: true,
    summary: '设置行高（磅）',
    schema: { type: 'object', required: ['op', 'sheet', 'row', 'heightPoints'], properties: { op: S('string'), sheet, row: S('number'), count: S('number'), heightPoints: S('number') } },
    example: { op: 'set_row_height', sheet: 'Sheet1', row: 1, heightPoints: 22 },
  },
  {
    op: 'set_col_width', group: 'structure', headless: true,
    summary: '设置列宽（像素）',
    schema: { type: 'object', required: ['op', 'sheet', 'column', 'widthPx'], properties: { op: S('string'), sheet, column: S('string'), count: S('number'), widthPx: S('number') } },
    example: { op: 'set_col_width', sheet: 'Sheet1', column: 'A', widthPx: 140 },
  },
  {
    op: 'set_rows_hidden', group: 'structure', headless: true,
    summary: '隐藏或取消隐藏若干行',
    schema: { type: 'object', required: ['op', 'sheet', 'row', 'hidden'], properties: { op: S('string'), sheet, row: S('number'), count: S('number'), hidden: S('boolean') } },
    example: { op: 'set_rows_hidden', sheet: 'Sheet1', row: 3, count: 2, hidden: true },
  },
  {
    op: 'set_cols_hidden', group: 'structure', headless: true,
    summary: '隐藏或取消隐藏若干列',
    schema: { type: 'object', required: ['op', 'sheet', 'column', 'hidden'], properties: { op: S('string'), sheet, column: S('string'), count: S('number'), hidden: S('boolean') } },
    example: { op: 'set_cols_hidden', sheet: 'Sheet1', column: 'C', count: 1, hidden: true },
  },
]

const sheetOps: OpSpec[] = [
  {
    op: 'add_sheet', group: 'sheets', headless: true,
    summary: '新增工作表',
    schema: { type: 'object', required: ['op', 'name'], properties: { op: S('string'), name: S('string'), rows: S('number'), columns: S('number') } },
    example: { op: 'add_sheet', name: '明细' },
  },
  {
    op: 'rename_sheet', group: 'sheets', headless: true,
    summary: '重命名工作表',
    schema: { type: 'object', required: ['op', 'sheet', 'name'], properties: { op: S('string'), sheet, name: S('string') } },
    example: { op: 'rename_sheet', sheet: 'Sheet1', name: '明细' },
  },
  {
    op: 'delete_sheet', group: 'sheets', headless: true,
    summary: '删除工作表（最后一张不可删）',
    schema: { type: 'object', required: ['op', 'sheet'], properties: { op: S('string'), sheet } },
    example: { op: 'delete_sheet', sheet: 'Sheet2' },
  },
  {
    op: 'duplicate_sheet', group: 'sheets', headless: true,
    summary: '复制工作表',
    schema: { type: 'object', required: ['op', 'sheet'], properties: { op: S('string'), sheet, name: S('string') } },
    example: { op: 'duplicate_sheet', sheet: 'Sheet1', name: 'Sheet1 副本' },
  },
  {
    op: 'move_sheet', group: 'sheets', headless: true,
    summary: '移动工作表到指定标签位置（1-based）',
    schema: { type: 'object', required: ['op', 'sheet', 'position'], properties: { op: S('string'), sheet, position: S('number') } },
    example: { op: 'move_sheet', sheet: '明细', position: 1 },
  },
  {
    op: 'set_sheet_hidden', group: 'sheets', headless: true,
    summary: '隐藏或取消隐藏工作表',
    schema: { type: 'object', required: ['op', 'sheet', 'hidden'], properties: { op: S('string'), sheet, hidden: S('boolean') } },
    example: { op: 'set_sheet_hidden', sheet: 'Sheet2', hidden: true },
  },
]

const objectGroups: Array<[string, string]> = [
  ['add_chart', '插入图表'], ['edit_chart', '编辑图表'], ['delete_visual', '删除图表/形状/图片'],
  ['add_shape', '插入形状'], ['edit_shape', '编辑形状'], ['add_image', '插入图片'],
  ['add_table', '插入表格'], ['add_table_row', '表格加行'], ['add_table_column', '表格加列'],
  ['delete_table_row', '表格删行'], ['delete_table_column', '表格删列'], ['delete_table', '删除表格'],
  ['add_pivot', '插入透视表'], ['refresh_pivot', '刷新透视表'], ['add_sparkline', '插入迷你图'],
  ['add_conditional_format', '条件格式'], ['clear_conditional_formats', '清除条件格式'],
  ['set_data_validation', '数据验证'], ['set_note', '单元格批注'], ['set_filter', '设置筛选'],
  ['clear_filter', '清除筛选'], ['set_filter_criteria', '筛选条件'], ['protect_sheet', '保护工作表'],
  ['set_page_setup', '页面设置'], ['set_freeze', '冻结窗格'], ['add_defined_name', '定义名称'],
  ['delete_defined_name', '删除定义名称'], ['set_hyperlink', '设置超链接'],
]

const objectOps: OpSpec[] = objectGroups.map(([op, summary]) => ({
  op,
  group: 'objects',
  headless: false,
  summary,
  schema: { type: 'object', required: ['op'], properties: { op: S('string') } },
  example: { op },
}))

export const OPS: OpSpec[] = [...cellOps, ...structureOps, ...sheetOps, ...objectOps]
const INDEX = new Map(OPS.map((spec) => [spec.op, spec]))

export const OP_COUNTS = {
  cells: cellOps.length,
  structure: structureOps.length,
  sheets: sheetOps.length,
  objects: objectOps.length,
  headless: OPS.filter((spec) => spec.headless).length,
  total: OPS.length,
}

export function groupOf(op: string): OpGroup | null {
  return INDEX.get(op)?.group ?? null
}

export function specOf(op: string): OpSpec | null {
  return INDEX.get(op) ?? null
}

export function isHeadlessOp(op: string): boolean {
  return INDEX.get(op)?.headless === true
}

const oneLine = (text: string, max = 140): string => {
  const flat = text.replace(/\s+/g, ' ').trim()
  return flat.length > max ? `${flat.slice(0, max)}…` : flat
}

export function buildGuide(input: {
  group?: string
  op?: string
  version?: string
  epoch?: number
}): { summary: string; data: unknown } {
  const header = [
    '# xlsx 命令目录',
    input.version ? `version=${input.version.slice(0, 12)}` : '',
    input.epoch !== undefined ? `epoch=${input.epoch}` : '',
    `ops=${OP_COUNTS.total} (cells=${OP_COUNTS.cells} structure=${OP_COUNTS.structure} sheets=${OP_COUNTS.sheets} objects=${OP_COUNTS.objects}; headless=${OP_COUNTS.headless})`,
  ].filter(Boolean).join(' ')

  if (input.op) {
    const spec = INDEX.get(input.op)
    if (!spec)
      return {
        summary: `${header}\n未知 op：${input.op}（无参调用看分组）`,
        data: { error: 'op.not-found' },
      }
    return {
      summary: [
        header,
        `## ${spec.op} · ${spec.group}`,
        spec.summary,
        `无页面：${spec.headless ? '可执行' : '不可执行（需打开页面）'}`,
        `坐标层级：${opTier(spec.op)}（L1 → 校验 docRev；L2 → 校验 structureEpoch；L0 → 不校验）`,
        'schema:',
        JSON.stringify(spec.schema, null, 2),
        'example:',
        JSON.stringify(spec.example, null, 2),
      ].join('\n'),
      data: { op: spec },
    }
  }

  if (input.group) {
    const group = input.group as OpGroup
    if (!['cells', 'structure', 'sheets', 'objects'].includes(group))
      return {
        summary: `${header}\n未知分组：${input.group}（可用：cells / structure / sheets / objects）`,
        data: { error: 'group.not-found' },
      }
    const lines = [header, `## ${group}（${group === 'objects' ? '需页面' : '无页面可执行'}）`]
    for (const spec of OPS.filter((item) => item.group === group)) {
      lines.push(`### ${spec.op} · ${opTier(spec.op)} · ${spec.summary}`)
      lines.push(`schema: ${JSON.stringify(spec.schema)}`)
      lines.push(`example: ${JSON.stringify(spec.example)}`)
    }
    return { summary: lines.join('\n'), data: { group } }
  }

  const lines = [header, '']
  for (const group of ['cells', 'structure', 'sheets'] as const) {
    lines.push(`## ${group} · 无页面可执行（${OPS.filter((s) => s.group === group).length}）`)
    lines.push(OPS.filter((s) => s.group === group).map((s) => `- ${s.op} — ${oneLine(s.summary, 90)}`).join('\n'))
    lines.push('')
  }
  lines.push(`## objects · 需要打开页面（${objectOps.length}）`)
  lines.push(objectOps.map((s) => `- ${s.op} — ${oneLine(s.summary, 90)}`).join('\n'))
  lines.push('')
  lines.push('只读：xlsx.read { what: context | sheets | range | formulas | find | audit }。')
  lines.push('坐标层级：L1 op 需带 docRev，L2 op 需带 structureEpoch（见 guide group / op）。')
  lines.push('写：xlsx.apply({ ref, commands: [{ op: "set_range", sheet: "Sheet1", start: "A1", values: [[...]] }] })；先 dryRun:true。')
  lines.push('新建：xlsx.create({ ref, sheets: [{ name, rows }] })；ref 必须不存在。')
  return {
    summary: lines.join('\n'),
    data: {
      ops: OP_COUNTS,
      headless: OPS.filter((s) => s.headless).map((s) => s.op),
      pageOnly: OPS.filter((s) => !s.headless).map((s) => s.op),
    },
  }
}
