// ─── 单元格 / 工作表结构（页面与引擎共用） ────────────────────────────
//
// 只描述解析结果，不带宿主身份；与 GenOffice 的 WorksheetState 保持同形。

export type CellValue = string | number | boolean | null | undefined

export interface CellStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontColor?: string
  fillColor?: string
  fontSize?: number
  fontFamily?: string
  align?: 'left' | 'center' | 'right'
  vertical?: 'top' | 'center' | 'bottom'
  wrap?: boolean
  numberFormat?: string
}

export interface CellLike {
  value: CellValue
  formula?: string | undefined
  rawValue?: CellValue
  style?: CellStyle
}

export interface SheetLike {
  id: string
  name: string
  /** ZIP 内的 worksheet 路径（写回用）。 */
  path?: string
  cells: Readonly<Record<string, CellLike>>
  gridRows?: number
  gridColumns?: number
  styles?: Readonly<Record<string, unknown>>
  /** 每个有显式样式的地址（含空值单元格）。 */
  formats?: Readonly<Record<string, CellStyle>>
  /** 合并区域，A1:C1 形式。 */
  merges?: readonly string[]
  /** 列宽（px，按 A/B/C 键）。 */
  colWidths?: Readonly<Record<string, number>>
  /** 行高（px，按 1/2/3 键）。 */
  rowHeights?: Readonly<Record<string, number>>
}
