// ─── 无页面命令 → gateway 保存请求 ─────────────────────────────────────
//
// 这里不维护第二份工作簿模型：每条命令都基于基线快照展开成 gateway 的
// CellEdit / SheetStructuralOps / SheetEditPlan。预览和执行共用这一份结果。

import { offsetFormulaRefs } from '@genoffice/xlsx-gateway/domain/formula-shift'
import {
  columnIndex,
  formatAddress,
  parseAddress,
  parseRange,
  rangeCellCount,
  type RangeBounds,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import type {
  CellFormatPatch,
  WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { CellScalar, CellState } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type {
  CellEdit,
  SheetStructuralOps,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { StructuralOp } from '@genoffice/xlsx-gateway/gateway/xlsx-structure'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import type { XlsxSaveRequest } from './types'
import { cellText, parseWorkbook, type WorkbookView } from './workbook'

const MAX_EXPANDED_CELLS = 2_000
const MAX_RANGE_CELLS = 200_000

export class HeadlessOpError extends Error {
  constructor(
    readonly code: string,
    message: string,
  ) {
    super(message)
    this.name = 'HeadlessOpError'
  }
}

function invalid(message: string): never {
  throw new HeadlessOpError('command.invalid', message)
}

function unsupported(message: string): never {
  throw new HeadlessOpError('command.unsupported', message)
}

function sheetOf(view: WorkbookView, name: string) {
  const sheet = view.byName.get(name)
  if (!sheet) invalid(`工作表不存在：${name}（先 xlsx_read what=context 取工作表名）`)
  return sheet
}

function checkRangeSize(bounds: RangeBounds, what: string): void {
  const count = rangeCellCount(bounds)
  if (count > MAX_RANGE_CELLS)
    invalid(`${what} 覆盖 ${count} 个单元格，超过 ${MAX_RANGE_CELLS}；拆成多次`)
}

function checkExpand(count: number, what: string): void {
  if (count > MAX_EXPANDED_CELLS)
    invalid(`${what} 展开为 ${count} 个单元格，超过 ${MAX_EXPANDED_CELLS}；拆成多次`)
}

function valueOrFormula(value: CellScalar): { value: CellScalar; formula?: string } {
  if (typeof value === 'string' && value.startsWith('=')) return { value: null, formula: value }
  return { value }
}

function cellEdit(
  sheetName: string,
  address: string,
  cell: CellState,
  style?: Record<string, unknown> | undefined,
  writeValue = true,
  styleReset = false,
): CellEdit {
  const parsed = parseAddress(address)
  return {
    sheetName,
    row: parsed.row,
    column: parsed.column,
    writeValue,
    cell,
    ...(style && Object.keys(style).length > 0 ? { style: style as never } : {}),
    ...(styleReset ? { styleReset: true } : {}),
  }
}

function styleOf(patch: CellFormatPatch): Record<string, unknown> {
  const style: Record<string, unknown> = {}
  if (patch.bold !== undefined && patch.bold !== null) style.bold = patch.bold
  if (patch.italic !== undefined && patch.italic !== null) style.italic = patch.italic
  if (patch.underline !== undefined && patch.underline !== null) style.underline = patch.underline
  if (patch.strikethrough !== undefined && patch.strikethrough !== null)
    style.strikethrough = patch.strikethrough
  if (patch.fontFamily !== undefined && patch.fontFamily !== null) style.fontFamily = patch.fontFamily
  if (patch.fontSize !== undefined && patch.fontSize !== null) style.fontSize = patch.fontSize
  if (patch.fontColor !== undefined) style.fontColor = patch.fontColor
  if (patch.fillColor !== undefined) style.fillColor = patch.fillColor
  if (patch.fill !== undefined) style.fill = patch.fill
  if (patch.numberFormat !== undefined) style.numberFormat = patch.numberFormat
  if (patch.horizontalAlign !== undefined && patch.horizontalAlign !== null)
    style.horizontalAlignment = patch.horizontalAlign
  if (patch.verticalAlign !== undefined && patch.verticalAlign !== null)
    style.verticalAlignment = patch.verticalAlign
  if (patch.wrapText !== undefined && patch.wrapText !== null) style.wrapText = patch.wrapText
  if (patch.textRotation !== undefined && patch.textRotation !== null) {
    const rotation = patch.textRotation
    style.textRotation = rotation === 'vertical' ? 255 : rotation < 0 ? rotation + 90 : rotation
  }
  if (patch.indent !== undefined && patch.indent !== null) style.indent = patch.indent
  if (patch.border) {
    const border = { style: 'thin' as const, ...(patch.border.color === undefined ? {} : { color: patch.border.color }) }
    if (patch.border.type === 'all' || patch.border.type === 'top') style.borderTop = border
    if (patch.border.type === 'all' || patch.border.type === 'bottom') style.borderBottom = border
    if (patch.border.type === 'all' || patch.border.type === 'left') style.borderLeft = border
    if (patch.border.type === 'all' || patch.border.type === 'right') style.borderRight = border
    if (patch.border.type === 'none') {
      style.borderTop = null
      style.borderBottom = null
      style.borderLeft = null
      style.borderRight = null
    }
  }
  return style
}

function emptyRequest(): XlsxSaveRequest {
  return { edits: [], structuralOps: [] }
}

function addStructural(
  request: XlsxSaveRequest,
  sheetName: string,
  op: StructuralOp,
): void {
  const entry = request.structuralOps.find((item) => item.sheetName === sheetName)
  if (entry) {
    request.structuralOps[request.structuralOps.indexOf(entry)] = {
      sheetName,
      ops: [...entry.ops, op],
    }
  } else {
    request.structuralOps.push({ sheetName, ops: [op] })
  }
}

function areaOf(bounds: RangeBounds) {
  return {
    startRow: bounds.startRow,
    endRow: bounds.endRow,
    startColumn: bounds.startColumn,
    endColumn: bounds.endColumn,
  }
}

function copyBlock(
  request: XlsxSaveRequest,
  sheetName: string,
  source: RangeBounds,
  targetStartRow: number,
  targetStartColumn: number,
  readCell: (sheetName: string, address: string) => CellState,
  sourceSheetName = sheetName,
): number {
  let count = 0
  for (let row = source.startRow; row <= source.endRow; row += 1) {
    for (let column = source.startColumn; column <= source.endColumn; column += 1) {
      const address = formatAddress(row, column)
      const cell = readCell(sourceSheetName, address)
      const targetRow = targetStartRow + (row - source.startRow)
      const targetColumn = targetStartColumn + (column - source.startColumn)
      const targetAddress = formatAddress(targetRow, targetColumn)
      const next: CellState = cell.formula
        ? { value: null, formula: offsetFormulaRefs(cell.formula, targetRow - row, targetColumn - column) }
        : { value: cell.value }
      request.edits.push(cellEdit(sheetName, targetAddress, next))
      count += 1
    }
  }
  checkExpand(count, 'copy_range')
  return count
}

export function buildSaveRequest(view: WorkbookView, operations: readonly WorkbookOperation[]): {
  request: XlsxSaveRequest
  changedCount: number
  warnings: string[]
} {
  const request = emptyRequest()
  const warnings: string[] = []
  const overlay = new Map<string, Map<string, CellState>>()
  const addedSheets = new Set<string>()
  const order = view.sheets.map((sheet) => sheet.name)
  const originalByCurrent = new Map(order.map((name) => [name, name]))
  const finalNameByOriginal = new Map<string, string>()
  const renames: Array<{ sheetName: string; newName: string }> = []
  const additions: Array<{ name: string; sourceSheetName?: string }> = []
  const removals: string[] = []
  const hiddenChanges: Array<{ sheetName: string; hidden: boolean }> = []
  let orderChanged = false
  let edited = 0

  const currentName = (name: string): string => name

  const readCell = (sheetName: string, address: string): CellState => {
    const overlaid = overlay.get(sheetName)?.get(address)
    if (overlaid !== undefined) return overlaid
    const sheet = view.byName.get(sheetName)
    if (sheet) return sheet.cells[address] ?? { value: null }
    if (addedSheets.has(sheetName)) return { value: null }
    return sheetOf(view, sheetName).cells[address] ?? { value: null }
  }

  const writeCell = (
    sheetName: string,
    address: string,
    cell: CellState,
    style?: Record<string, unknown>,
    writeValue = true,
  ): void => {
    if (writeValue) {
      const cells = overlay.get(sheetName) ?? new Map<string, CellState>()
      cells.set(address, cell)
      overlay.set(sheetName, cells)
    }
    request.edits.push(
      cellEdit(sheetName, address, cell, style, writeValue),
    )
  }

  const shiftPending = (
    sheetName: string,
    axis: 'row' | 'column',
    index: number,
    count: number,
    remove: boolean,
  ): void => {
    const nextEdits: CellEdit[] = []
    for (const edit of request.edits) {
      if (edit.sheetName !== sheetName) {
        nextEdits.push(edit)
        continue
      }
      const position = axis === 'row' ? edit.row : edit.column
      if (remove) {
        if (position >= index && position < index + count) continue
        if (position >= index + count) {
          nextEdits.push(
            axis === 'row'
              ? { ...edit, row: position - count }
              : { ...edit, column: position - count },
          )
          continue
        }
        nextEdits.push(edit)
        continue
      }
      if (position >= index) {
        nextEdits.push(
          axis === 'row'
            ? { ...edit, row: position + count }
            : { ...edit, column: position + count },
        )
        continue
      }
      nextEdits.push(edit)
    }
    request.edits = nextEdits
    for (const [overlaySheet, cells] of overlay) {
      if (overlaySheet !== sheetName) continue
      const shifted = new Map<string, CellState>()
      for (const [address, cell] of cells) {
        let parsed
        try {
          parsed = parseAddress(address)
        } catch {
          continue
        }
        const position = axis === 'row' ? parsed.row : parsed.column
        if (remove && position >= index && position < index + count) continue
        const nextPosition =
          position >= index + (remove ? count : 0)
            ? position + (remove ? -count : count)
            : position
        if (nextPosition < 0) continue
        const nextAddress =
          axis === 'row'
            ? formatAddress(nextPosition, parsed.column)
            : formatAddress(parsed.row, nextPosition)
        shifted.set(nextAddress, cell)
      }
      overlay.set(overlaySheet, shifted)
    }
  }

  for (const operation of operations) {
    switch (operation.op) {
      case 'set_cell': {
        const { value, formula } = valueOrFormula(operation.value)
        writeCell(operation.sheetId, operation.address, formula ? { value: null, formula } : { value })
        edited += 1
        break
      }
      case 'set_formula': {
        writeCell(operation.sheetId, operation.address, { value: null, formula: operation.formula })
        edited += 1
        break
      }
      case 'clear_cell': {
        writeCell(operation.sheetId, operation.address, { value: null })
        edited += 1
        break
      }
      case 'set_range': {
        const start = operation.start ? parseAddress(operation.start) : parseRange(operation.range!)
        const startRow = 'row' in start ? start.row : start.startRow
        const startColumn = 'column' in start ? start.column : start.startColumn
        let count = 0
        operation.values.forEach((rowValues, rowOffset) => {
          rowValues.forEach((value, columnOffset) => {
            const address = formatAddress(startRow + rowOffset, startColumn + columnOffset)
            const { value: cellValue, formula } = valueOrFormula(value)
            writeCell(operation.sheetId, address, formula ? { value: null, formula } : { value: cellValue })
            count += 1
          })
        })
        checkExpand(count, 'set_range')
        edited += count
        break
      }
      case 'clear_range': {
        const bounds = parseRange(operation.range)
        checkRangeSize(bounds, 'clear_range')
        let count = 0
        for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
          for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
            writeCell(operation.sheetId, formatAddress(row, column), { value: null })
            count += 1
          }
        }
        checkExpand(count, 'clear_range')
        edited += count
        break
      }
      case 'format_range': {
        const bounds = parseRange(operation.range)
        checkRangeSize(bounds, 'format_range')
        let count = 0
        for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
          for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
            writeCell(
              operation.sheetId,
              formatAddress(row, column),
              { value: null },
              styleOf(operation.format),
              false,
            )
            count += 1
          }
        }
        checkExpand(count, 'format_range')
        edited += count
        break
      }
      case 'find_replace': {
        const bounds = parseRange(operation.range)
        checkRangeSize(bounds, 'find_replace')
        const find = operation.matchCase ? operation.find : operation.find.toLowerCase()
        let count = 0
        for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
          for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
            const address = formatAddress(row, column)
            const cell = readCell(operation.sheetId, address)
            if (!cell || cell.formula) continue
            const text = cellText(cell)
            if (!text) continue
            const haystack = operation.matchCase ? text : text.toLowerCase()
            const matched = operation.wholeCell ? haystack === find : haystack.includes(find)
            if (!matched) continue
            const next = operation.wholeCell
              ? operation.replace
              : operation.matchCase
                ? text.split(operation.find).join(operation.replace)
                : text.replace(new RegExp(operation.find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), () => operation.replace)
            writeCell(operation.sheetId, address, { value: next })
            count += 1
          }
        }
        edited += count
        break
      }
      case 'fill_range': {
        const source = parseRange(operation.source)
        const target = parseRange(operation.target)
        const sourceRows = source.endRow - source.startRow + 1
        const sourceCols = source.endColumn - source.startColumn + 1
        const targetRows = target.endRow - target.startRow + 1
        const targetCols = target.endColumn - target.startColumn + 1
        if (targetRows % sourceRows !== 0 || targetCols % sourceCols !== 0)
          invalid('fill_range 的目标行列数必须是源区域的整数倍')
        let count = 0
        for (let row = target.startRow; row <= target.endRow; row += 1) {
          for (let column = target.startColumn; column <= target.endColumn; column += 1) {
            const sourceRow = source.startRow + ((row - target.startRow) % sourceRows)
            const sourceColumn = source.startColumn + ((column - target.startColumn) % sourceCols)
            const cell = readCell(operation.sourceSheetId ?? operation.sheetId, formatAddress(sourceRow, sourceColumn))
            const next: CellState = cell.formula
              ? { value: null, formula: offsetFormulaRefs(cell.formula, row - sourceRow, column - sourceColumn) }
              : { value: cell.value }
            writeCell(operation.sheetId, formatAddress(row, column), next)
            count += 1
          }
        }
        checkExpand(count, 'fill_range')
        edited += count
        break
      }
      case 'copy_range': {
        const source = parseRange(operation.source)
        const targetText = operation.target.includes(':') ? operation.target.split(':')[0]! : operation.target
        const target = parseAddress(targetText)
        checkRangeSize(source, 'copy_range')
        edited += copyBlock(request, operation.sheetId, source, target.row, target.column, readCell, operation.sourceSheetId ?? operation.sheetId)
        break
      }
      case 'move_range': {
        const source = parseRange(operation.source)
        const targetText = operation.target.includes(':') ? operation.target.split(':')[0]! : operation.target
        const target = parseAddress(targetText)
        checkRangeSize(source, 'move_range')
        const sourceRows = source.endRow - source.startRow + 1
        const sourceCols = source.endColumn - source.startColumn + 1
        const overlap =
          target.row <= source.endRow &&
          target.row + sourceRows - 1 >= source.startRow &&
          target.column <= source.endColumn &&
          target.column + sourceCols - 1 >= source.startColumn
        if (overlap) unsupported('move_range 的源区域与目标区域重叠；请拆成 copy_range + clear_range')
        edited += copyBlock(request, operation.sheetId, source, target.row, target.column, readCell, operation.sourceSheetId ?? operation.sheetId)
        const clearSheet = operation.sourceSheetId ?? operation.sheetId
        for (let row = source.startRow; row <= source.endRow; row += 1) {
          for (let column = source.startColumn; column <= source.endColumn; column += 1) {
            writeCell(clearSheet, formatAddress(row, column), { value: null })
            edited += 1
          }
        }
        break
      }
      case 'sort_range': {
        const bounds = parseRange(operation.range)
        checkRangeSize(bounds, 'sort_range')
        const column = columnIndex(operation.byColumn)
        const startRow = operation.hasHeader ? bounds.startRow + 1 : bounds.startRow
        const rows: Array<{ row: number; cells: CellState[] }> = []
        for (let row = startRow; row <= bounds.endRow; row += 1) {
          const cells: CellState[] = []
          for (let col = bounds.startColumn; col <= bounds.endColumn; col += 1) {
            cells.push(readCell(operation.sheetId, formatAddress(row, col)))
          }
          rows.push({ row, cells })
        }
        const keyIndex = column - bounds.startColumn
        if (keyIndex < 0 || keyIndex >= bounds.endColumn - bounds.startColumn + 1)
          invalid(`sort_range 的 byColumn ${operation.byColumn} 不在范围 ${operation.range} 内`)
        rows.sort((left, right) => {
          const a = left.cells[keyIndex]?.value
          const b = right.cells[keyIndex]?.value
          const an = typeof a === 'number' ? a : Number.NaN
          const bn = typeof b === 'number' ? b : Number.NaN
          const cmp = Number.isFinite(an) && Number.isFinite(bn)
            ? an - bn
            : String(a ?? '').localeCompare(String(b ?? ''), 'zh-CN')
          return operation.order === 'desc' ? -cmp : cmp
        })
        rows.forEach((entry, offset) => {
          entry.cells.forEach((cell, cellOffset) => {
            const address = formatAddress(startRow + offset, bounds.startColumn + cellOffset)
            writeCell(operation.sheetId, address, cell.formula ? { value: null, formula: cell.formula } : { value: cell.value })
            edited += 1
          })
        })
        break
      }
      case 'insert_rows':
        shiftPending(operation.sheetId, 'row', operation.row - 1, operation.count, false)
        addStructural(request, operation.sheetId, { kind: 'insert-rows', index: operation.row - 1, count: operation.count })
        edited += 1
        break
      case 'delete_rows':
        shiftPending(operation.sheetId, 'row', operation.row - 1, operation.count, true)
        addStructural(request, operation.sheetId, { kind: 'remove-rows', index: operation.row - 1, count: operation.count })
        edited += 1
        break
      case 'insert_cols':
        shiftPending(operation.sheetId, 'column', columnIndex(operation.column), operation.count, false)
        addStructural(request, operation.sheetId, { kind: 'insert-cols', index: columnIndex(operation.column), count: operation.count })
        edited += 1
        break
      case 'delete_cols':
        shiftPending(operation.sheetId, 'column', columnIndex(operation.column), operation.count, true)
        addStructural(request, operation.sheetId, { kind: 'remove-cols', index: columnIndex(operation.column), count: operation.count })
        edited += 1
        break
      case 'merge_cells':
        addStructural(request, operation.sheetId, { kind: 'merge-cells', range: areaOf(parseRange(operation.range)) })
        edited += 1
        break
      case 'unmerge_cells':
        addStructural(request, operation.sheetId, { kind: 'unmerge-cells', range: areaOf(parseRange(operation.range)) })
        edited += 1
        break
      case 'set_row_height':
        addStructural(request, operation.sheetId, {
          kind: 'set-row-size',
          start: operation.row - 1,
          end: operation.row - 1 + (operation.count ?? 1) - 1,
          size: operation.heightPoints,
        })
        edited += 1
        break
      case 'set_col_width':
        addStructural(request, operation.sheetId, {
          kind: 'set-col-size',
          start: columnIndex(operation.column),
          end: columnIndex(operation.column) + (operation.count ?? 1) - 1,
          size: Math.max(0.1, (operation.widthPx - 5) / 7),
        })
        edited += 1
        break
      case 'set_rows_hidden':
        addStructural(request, operation.sheetId, {
          kind: 'set-rows-hidden',
          start: operation.row - 1,
          end: operation.row - 1 + (operation.count ?? 1) - 1,
          hidden: operation.hidden,
        })
        edited += 1
        break
      case 'set_cols_hidden':
        addStructural(request, operation.sheetId, {
          kind: 'set-cols-hidden',
          start: columnIndex(operation.column),
          end: columnIndex(operation.column) + (operation.count ?? 1) - 1,
          hidden: operation.hidden,
        })
        edited += 1
        break
      case 'add_sheet': {
        const name = operation.name.trim()
        additions.push({ name })
        addedSheets.add(name)
        order.push(name)
        edited += 1
        break
      }
      case 'rename_sheet': {
        const current = currentName(operation.sheetId)
        const nextName = operation.name.trim()
        if (addedSheets.has(current)) {
          const addition = additions.find((entry) => entry.name === current)
          if (addition) addition.name = nextName
          const index = order.indexOf(current)
          if (index >= 0) order[index] = nextName
          addedSheets.delete(current)
          addedSheets.add(nextName)
          originalByCurrent.set(nextName, nextName)
          edited += 1
          break
        }
        const original = originalByCurrent.get(current) ?? current
        renames.push({ sheetName: original, newName: nextName })
        finalNameByOriginal.set(original, nextName)
        const index = order.indexOf(current)
        if (index >= 0) order[index] = nextName
        originalByCurrent.set(nextName, original)
        edited += 1
        break
      }
      case 'delete_sheet': {
        const current = currentName(operation.sheetId)
        if (addedSheets.has(current)) {
          const index = additions.findIndex((entry) => entry.name === current)
          if (index >= 0) additions.splice(index, 1)
          addedSheets.delete(current)
          const orderIndex = order.indexOf(current)
          if (orderIndex >= 0) order.splice(orderIndex, 1)
          overlay.delete(operation.sheetId)
          request.edits = request.edits.filter((edit) => edit.sheetName !== operation.sheetId)
          edited += 1
          break
        }
        const original = originalByCurrent.get(current) ?? current
        removals.push(original)
        overlay.delete(operation.sheetId)
        request.edits = request.edits.filter((edit) => edit.sheetName !== operation.sheetId)
        const index = order.indexOf(current)
        if (index >= 0) order.splice(index, 1)
        edited += 1
        break
      }
      case 'duplicate_sheet': {
        const current = currentName(operation.sheetId)
        const original = originalByCurrent.get(current) ?? current
        const name = operation.name?.trim() ?? `${current} 副本`
        additions.push({ name, sourceSheetName: original })
        addedSheets.add(name)
        order.push(name)
        edited += 1
        break
      }
      case 'move_sheet': {
        const current = currentName(operation.sheetId)
        const index = order.indexOf(current)
        if (index >= 0) {
          order.splice(index, 1)
          order.splice(Math.max(0, Math.min(order.length, operation.position - 1)), 0, current)
        }
        orderChanged = true
        edited += 1
        break
      }
      case 'set_sheet_hidden': {
        const current = currentName(operation.sheetId)
        hiddenChanges.push({ sheetName: originalByCurrent.get(current) ?? current, hidden: operation.hidden })
        edited += 1
        break
      }
      default:
        unsupported(`op ${operation.op} 在无页面路径暂不支持；请打开页面`)
    }
  }

  // gateway 先处理 cell/structural 编辑，最后才应用 sheet plan；因此编辑
  // 目标要用**原始**工作表名（新增表用新名字）。这样 rename + 编辑可同批。
  request.edits = request.edits.map((edit) => ({
    ...edit,
    sheetName: originalByCurrent.get(edit.sheetName) ?? edit.sheetName,
  }))
  request.structuralOps = request.structuralOps.map((entry) => ({
    ...entry,
    sheetName: originalByCurrent.get(entry.sheetName) ?? entry.sheetName,
  }))

  if (renames.length > 0 || additions.length > 0 || removals.length > 0 || hiddenChanges.length > 0 || orderChanged) {
    request.sheetPlan = {
      renames,
      additions,
      removals,
      order,
      ...(hiddenChanges.length > 0 ? { hiddenChanges } : {}),
      ...(orderChanged ? { orderChanged: true } : {}),
    }
  }

  if (request.edits.length === 0 && request.structuralOps.length === 0 && !request.sheetPlan)
    warnings.push('noop: 命令没有产生可保存的修改')

  return { request, changedCount: edited, warnings }
}

export async function saveRequestFromCommands(
  bytes: Uint8Array,
  operations: readonly WorkbookOperation[],
): Promise<{ request: XlsxSaveRequest; changedCount: number; warnings: string[] }> {
  const view = await parseWorkbook(bytes)
  return buildSaveRequest(view, operations)
}
