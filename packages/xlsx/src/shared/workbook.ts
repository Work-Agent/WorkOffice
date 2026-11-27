// ─── 字节侧工作簿读取（页面 / 执行点同源） ─────────────────────────────
//
// 引擎不依赖 Univer；只借 xlsx-gateway 的 readBasicWorkbook 读单元格值与
// 工作表清单。样式、图表、透视表等复杂能力由 gateway 保存侧处理，读侧只
// 做模型需要的投影。

import {
  readBasicWorkbook,
  inventoryXlsx,
  type ImportedXlsx,
  type PackageEntry,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import {
  columnIndex,
  columnLabel,
  formatAddress,
  parseAddress,
  parseRange,
  type RangeBounds,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import type {
  CellScalar,
  CellState,
  WorksheetState,
} from '@genoffice/xlsx-gateway/domain/workbook.types'
import { parseWorkbook as parseSheetLikeWorkbook } from './xlsx-parse'
import type { CellStyle } from './sheet-types'

export interface WorkbookView {
  snapshot: ImportedXlsx['snapshot']
  sheetNamesById: ImportedXlsx['sheetNamesById']
  sheets: WorksheetState[]
  byName: Map<string, WorksheetState>
  /**
   * 无页面侧的显式格式（工作表名 → 地址 → workbook DSL 风格的样式增补）。
   *
   * `copy_range` 要能“沿用现有格式”，而 `CellState` 不带样式：解析时顺手
   * 把 cellXfs 归一化成 gateway 的样式形状，复制时一并下发。
   */
  formats: Map<string, Record<string, Record<string, unknown>>>
  /** 合并区（工作表名 → `A2:C8` 列表），`copy_range` 一并偏移复制。 */
  merges: Map<string, readonly string[]>
  /** 行高（工作表名 → 行号 → px），复制时换算回磅一并下发。 */
  heights: Map<string, Readonly<Record<string, number>>>
}

/**
 * `CellStyle`（解析出的 OOXML 样式）→ gateway `WorkbookStyleEdit`（严格 schema）。
 *
 * 只映射两边都存在的属性：字体/填充/边框以外的主题色、富文本不在这里猜。
 */
function stylePatchOf(style: CellStyle | undefined): Record<string, unknown> | null {
  if (!style) return null
  const patch: Record<string, unknown> = {}
  if (style.bold) patch.bold = true
  if (style.italic) patch.italic = true
  if (style.underline) patch.underline = true
  if (style.fontFamily) patch.fontFamily = style.fontFamily
  if (style.fontSize) patch.fontSize = style.fontSize
  if (style.fontColor) patch.fontColor = style.fontColor
  if (style.fillColor) patch.fillColor = style.fillColor
  if (style.numberFormat) patch.numberFormat = style.numberFormat
  if (style.align) patch.horizontalAlignment = style.align
  if (style.vertical) patch.verticalAlignment = style.vertical
  if (style.wrap) patch.wrapText = true
  return Object.keys(patch).length > 0 ? patch : null
}

export class XlsxParseError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'XlsxParseError'
  }
}

export async function parseWorkbook(bytes: Uint8Array): Promise<WorkbookView> {
  if (!(bytes instanceof Uint8Array) || bytes.byteLength === 0)
    throw new XlsxParseError('工作簿字节为空')
  try {
    // 页面 codec 与执行点共用同一份解析器：对命名空间前缀、非常规 rels
    // Target 更宽容，避免“有些表格打不开”。
    const parsed = parseSheetLikeWorkbook(bytes)
    const sheets: WorksheetState[] = parsed.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      cells: Object.fromEntries(
        Object.entries(sheet.cells).map(([address, cell]) => [
          address,
          {
            value: (cell.value ?? null) as CellScalar,
            ...(cell.formula ? { formula: cell.formula } : {}),
          } satisfies CellState,
        ]),
      ),
    }))
    if (sheets.length === 0) throw new Error('no readable worksheets')
    const sheetNamesById = Object.fromEntries(sheets.map((sheet) => [sheet.id, sheet.name]))
    const snapshot = { revision: 0, sheets }
    const formats = new Map<string, Record<string, Record<string, unknown>>>()
    const merges = new Map<string, readonly string[]>()
    const heights = new Map<string, Readonly<Record<string, number>>>()
    for (const sheet of parsed) {
      const map: Record<string, Record<string, unknown>> = {}
      // `formats` 覆盖“空值但带样式”的填充格（cells 里没有它们）。
      const styled = sheet.formats ?? {}
      for (const [address, style] of Object.entries(styled)) {
        const patch = stylePatchOf(style)
        if (patch) map[address] = patch
      }
      for (const [address, cell] of Object.entries(sheet.cells)) {
        if (map[address]) continue
        const patch = stylePatchOf(cell.style)
        if (patch) map[address] = patch
      }
      formats.set(sheet.name, map)
      if (sheet.merges?.length) merges.set(sheet.name, sheet.merges)
      if (sheet.rowHeights && Object.keys(sheet.rowHeights).length > 0)
        heights.set(sheet.name, sheet.rowHeights)
    }
    return {
      snapshot,
      sheetNamesById,
      sheets,
      byName: new Map(sheets.map((sheet) => [sheet.name, sheet])),
      formats,
      merges,
      heights,
    }
  } catch (parseError) {
    try {
      const imported = await readBasicWorkbook(Buffer.from(bytes))
      const byName = new Map(imported.snapshot.sheets.map((sheet) => [sheet.name, sheet]))
      return {
        snapshot: imported.snapshot,
        sheetNamesById: imported.sheetNamesById,
        sheets: [...imported.snapshot.sheets],
        byName,
        // gateway 回退路径没有样式/合并投影：copy_range 只能复制值（已在 summary 里说明）。
        formats: new Map(),
        merges: new Map(),
        heights: new Map(),
      }
    } catch (gatewayError) {
      throw new XlsxParseError(
        `不是有效的 XLSX 包：${parseError instanceof Error ? parseError.message : String(parseError)}`,
      )
    }
  }
}

export function isXlsxPackage(bytes: Uint8Array): boolean {
  // ZIP 本地文件头 + workbook 条目；完整校验交给 parseWorkbook。
  if (bytes.byteLength < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) return false
  return true
}

export async function packageEntries(bytes: Uint8Array): Promise<PackageEntry[]> {
  try {
    return [...(await inventoryXlsx(Buffer.from(bytes)))]
  } catch {
    return []
  }
}

export function usedRange(sheet: WorksheetState): { startRow: number; startColumn: number; endRow: number; endColumn: number; rows: number; columns: number } | null {
  let startRow = Number.POSITIVE_INFINITY
  let startColumn = Number.POSITIVE_INFINITY
  let endRow = -1
  let endColumn = -1
  for (const address of Object.keys(sheet.cells)) {
    let parsed
    try {
      parsed = parseAddress(address)
    } catch {
      continue
    }
    startRow = Math.min(startRow, parsed.row)
    startColumn = Math.min(startColumn, parsed.column)
    endRow = Math.max(endRow, parsed.row)
    endColumn = Math.max(endColumn, parsed.column)
  }
  if (endRow < 0) return null
  return {
    startRow,
    startColumn,
    endRow,
    endColumn,
    rows: endRow - startRow + 1,
    columns: endColumn - startColumn + 1,
  }
}

export function a1Of(row: number, column: number): string {
  return formatAddress(row, column)
}

export function boundsOf(text: string): RangeBounds {
  return parseRange(text)
}

export function columnOf(label: string): number {
  return columnIndex(label)
}

export function columnNameOf(index: number): string {
  return columnLabel(index)
}

export function cellAt(view: WorkbookView, sheetName: string, address: string): CellState | null {
  const sheet = view.byName.get(sheetName)
  if (!sheet) return null
  return sheet.cells[address.toUpperCase()] ?? null
}

export function cellReader(view: WorkbookView) {
  return (address: string, sheetId: string): CellState => {
    const sheet = view.byName.get(sheetId) ?? view.sheets.find((item) => item.id === sheetId)
    return sheet?.cells[address] ?? { value: null }
  }
}

export function sheetNameOf(view: WorkbookView, sheetIdOrName: string): string {
  return view.sheetNamesById[sheetIdOrName] ?? sheetIdOrName
}

export function cellText(cell: CellState | null | undefined): string {
  if (!cell) return ''
  if (cell.formula) return cell.formula
  if (cell.value === null || cell.value === undefined) return ''
  return typeof cell.value === 'boolean' ? (cell.value ? 'TRUE' : 'FALSE') : String(cell.value)
}

export function typeOfValue(value: CellScalar | undefined): string {
  if (value === null || value === undefined) return 'empty'
  if (typeof value === 'number') return 'number'
  if (typeof value === 'boolean') return 'boolean'
  return 'text'
}

export interface RangeRead {
  sheet: string
  range: string
  rows: Array<{ row: number; cells: Array<{ address: string; value: CellScalar; formula?: string; text: string }> }>
  truncated: boolean
}

export function readRange(
  view: WorkbookView,
  sheetName: string,
  rangeText: string,
  maxCells = 2_000,
): RangeRead {
  const sheet = view.byName.get(sheetName)
  if (!sheet) throw new XlsxParseError(`工作表不存在：${sheetName}`)
  const bounds = parseRange(rangeText)
  const rows: RangeRead['rows'] = []
  let count = 0
  let truncated = false
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    const cells: RangeRead['rows'][number]['cells'] = []
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      if (count >= maxCells) {
        truncated = true
        break
      }
      count += 1
      const address = formatAddress(row, column)
      const cell = sheet.cells[address] ?? { value: null }
      cells.push({
        address,
        value: cell.value,
        ...(cell.formula ? { formula: cell.formula } : {}),
        text: cellText(cell),
      })
    }
    rows.push({ row: row + 1, cells })
    if (truncated) break
  }
  return { sheet: sheetName, range: rangeText, rows, truncated }
}

export interface WorkbookContext {
  name: string
  sheetCount: number
  sheets: Array<{
    id: string
    name: string
    rows: number
    columns: number
    usedRange: string | null
    cellCount: number
  }>
}

export function contextOf(view: WorkbookView, name: string): WorkbookContext {
  return {
    name,
    sheetCount: view.sheets.length,
    sheets: view.sheets.map((sheet) => {
      const range = usedRange(sheet)
      return {
        id: sheet.id,
        name: sheet.name,
        rows: range?.rows ?? 0,
        columns: range?.columns ?? 0,
        usedRange: range
          ? `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`
          : null,
        cellCount: Object.keys(sheet.cells).length,
      }
    }),
  }
}

export function contextText(view: WorkbookView, name: string): string {
  const ctx = contextOf(view, name)
  const lines = [`# ${name}`, `sheets=${ctx.sheetCount}`]
  for (const sheet of ctx.sheets) {
    lines.push(
      `- ${sheet.name} id=${sheet.id} used=${sheet.usedRange ?? '(empty)'} cells=${sheet.cellCount} ` +
        `rows=${sheet.rows} columns=${sheet.columns}`,
    )
  }
  return lines.join('\n')
}

export function sheetText(view: WorkbookView, sheetName: string, maxRows = 60, maxChars = 12_000): string {
  const sheet = view.byName.get(sheetName)
  if (!sheet) throw new XlsxParseError(`工作表不存在：${sheetName}`)
  const range = usedRange(sheet)
  const lines = [`# ${sheetName}`, `used=${range ? `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}` : '(empty)'}`]
  if (!range) return lines.join('\n')
  const maxRow = Math.min(range.endRow, range.startRow + maxRows - 1)
  for (let row = range.startRow; row <= maxRow; row += 1) {
    const values: string[] = []
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const address = formatAddress(row, column)
      const cell = sheet.cells[address]
      values.push(cellText(cell).replace(/\t/g, ' ').replace(/\n/g, ' '))
    }
    const line = `${row + 1}\t${values.join('\t')}`
    if (lines.join('\n').length + line.length > maxChars) {
      lines.push('…（已截断）')
      break
    }
    lines.push(line)
  }
  if (maxRow < range.endRow) lines.push(`…（还有 ${range.endRow - maxRow} 行）`)
  return lines.join('\n')
}

export interface FindHit {
  sheet: string
  address: string
  value: CellScalar
  formula?: string
  text: string
}

export function findCells(
  view: WorkbookView,
  query: string,
  options: { sheet?: string; regex?: boolean; matchCase?: boolean; limit?: number } = {},
): { hits: FindHit[]; truncated: boolean } {
  const limit = Math.max(1, Math.min(2_000, options.limit ?? 200))
  const needle = options.matchCase ? query : query.toLowerCase()
  let matcher: ((text: string) => boolean) | null = null
  if (options.regex) {
    try {
      const re = new RegExp(query, options.matchCase ? '' : 'i')
      matcher = (text) => re.test(text)
    } catch {
      matcher = null
    }
  }
  const hits: FindHit[] = []
  const sheets = options.sheet ? [view.byName.get(options.sheet)].filter(Boolean) : view.sheets
  for (const sheet of sheets as WorksheetState[]) {
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const text = cellText(cell)
      if (!text) continue
      const matched = matcher
        ? matcher(text)
        : (options.matchCase ? text : text.toLowerCase()).includes(needle)
      if (!matched) continue
      hits.push({
        sheet: sheet.name,
        address,
        value: cell.value,
        ...(cell.formula ? { formula: cell.formula } : {}),
        text,
      })
      if (hits.length >= limit) return { hits, truncated: true }
    }
  }
  return { hits, truncated: false }
}

export async function auditText(bytes: Uint8Array, view: WorkbookView): Promise<string> {
  const entries = await packageEntries(bytes)
  const parts = new Set(entries.map((entry) => entry.path))
  const formulas = view.sheets.reduce(
    (sum, sheet) => sum + Object.values(sheet.cells).filter((cell) => cell.formula).length,
    0,
  )
  const sheets = view.sheets.map((sheet) => `${sheet.name}(${Object.keys(sheet.cells).length})`).join(', ')
  return [
    '# XLSX 诊断',
    `parts=${entries.length}`,
    `sheets=${view.sheets.length}${sheets ? `: ${sheets}` : ''}`,
    `formulaCells=${formulas}`,
    `hasSharedStrings=${parts.has('xl/sharedStrings.xml')}`,
    `hasStyles=${parts.has('xl/styles.xml')}`,
    `hasCalcChain=${parts.has('xl/calcChain.xml')}`,
    `hasWorkbook=${parts.has('xl/workbook.xml')}`,
    `hasContentTypes=${parts.has('[Content_Types].xml')}`,
    `bytes=${bytes.byteLength}`,
    formulas > 0 ? 'note=无页面路径不重算公式；缓存值可能保持旧值' : '',
  ].filter(Boolean).join('\n')
}
