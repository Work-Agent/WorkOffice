// ─── GenOffice 保存请求 → 页面内 xlsx-gateway ────────────────────────
//
// App 的 WorkbookSaveRequest 与 GenOffice 主进程的 saveWorkbookViaSidecar
// 同构。这里把 sheetId 映射成 sheetName，然后调用纯 TS 的
// planCellEditsToXlsx + assembleWithJsZip，避免依赖 Electron main / native sidecar。

import { Buffer } from 'buffer'
import {
  assembleWithJsZip,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { XlsxSaveRequest } from '../shared/types'

function resolveSheetName(
  sheetId: string,
  file: any,
  added: Map<string, string>,
): string {
  const name = added.get(sheetId) ?? file.sheets.find((sheet: any) => sheet.id === sheetId)?.name
  if (!name) throw new Error(`Unknown worksheet ${sheetId}`)
  return name
}

/// 把 App 的保存请求映射成 gateway 的具名参数（sheet 一律用名称）。
///
/// 页面路径（`kind:'save'`，请求进执行点）与回退路径（`kind:'bytes'`，页面自己
/// 产字节）**共用这一份映射**：任何字段只在一条路径上被映射，都会变成静默丢数据。
export function planWorkbookSave(request: any, file: any): XlsxSaveRequest {
  const addedSheetNames = new Map<string, string>()
  const duplicateSources = new Map<string, string>()
  const renames: Array<{ sheetName: string; newName: string }> = []
  const removals: string[] = []
  const hiddenChanges: Array<{ sheetName: string; hidden: boolean }> = []
  const tabColorStates: Array<{ sheetName: string; color: string | null }> = []
  let orderChanged = false

  for (const op of request.sheetOps ?? []) {
    if (op.kind === 'add-sheet') {
      addedSheetNames.set(op.sheetId, op.name)
      continue
    }
    if (op.kind === 'duplicate-sheet') {
      const sourceName = resolveSheetName(op.sourceSheetId, file, addedSheetNames)
      addedSheetNames.set(op.sheetId, op.name)
      duplicateSources.set(op.sheetId, sourceName)
      continue
    }
    if (op.kind === 'reorder-sheets') {
      orderChanged = true
      continue
    }
    const sheetName = resolveSheetName(op.sheetId, file, addedSheetNames)
    if (op.kind === 'rename-sheet') renames.push({ sheetName, newName: op.newName })
    else if (op.kind === 'set-sheet-hidden') hiddenChanges.push({ sheetName, hidden: op.hidden })
    else if (op.kind === 'set-sheet-tab-color')
      tabColorStates.push({ sheetName, color: op.color })
    else removals.push(sheetName)
  }

  const renameByOriginal = new Map(renames.map((rename) => [rename.sheetName, rename.newName]))
  let sheetPlan: any
  if ((request.sheetOps ?? []).some((op: any) => op.kind !== 'set-sheet-tab-color')) {
    sheetPlan = {
      renames,
      additions: [...addedSheetNames].map(([sheetId, name]) => ({
        name,
        sourceSheetName: duplicateSources.get(sheetId),
      })),
      removals,
      hiddenChanges,
      orderChanged,
      order: (request.sheetOrder ?? file.sheets.map((sheet: any) => sheet.id)).map(
        (sheetId: string) => {
          const original = resolveSheetName(sheetId, file, addedSheetNames)
          return addedSheetNames.has(sheetId)
            ? original
            : (renameByOriginal.get(original) ?? original)
        },
      ),
    }
  }

  const edits = (request.edits ?? []).map((edit: any) => ({
    sheetName: resolveSheetName(edit.sheetId, file, addedSheetNames),
    row: edit.row,
    column: edit.column,
    writeValue: edit.writeValue,
    cell: { value: edit.value, formula: edit.formula, rawValue: edit.rawValue },
    style: edit.style,
    rich: edit.rich,
    styleReset: edit.styleReset,
  }))

  const structuralOpsBySheet = new Map<string, any[]>()
  for (const op of request.structuralOps ?? []) {
    const sheetName = resolveSheetName(op.sheetId, file, addedSheetNames)
    const list = structuralOpsBySheet.get(sheetName) ?? []
    list.push({ ...op, sheetId: undefined })
    structuralOpsBySheet.set(sheetName, list)
  }
  const structuralOps = [...structuralOpsBySheet].map(([sheetName, ops]) => ({
    sheetName,
    ops,
  }))

  const bulkConstantFills = (request.bulkConstantFills ?? []).map((fill: any) => {
    const { sheetId, ...rest } = fill
    return { sheetName: resolveSheetName(sheetId, file, addedSheetNames), ...rest }
  })

  const formulaValuesBySheet = new Map<string, any[]>()
  for (const cell of request.formulaValues ?? []) {
    const sheetName = resolveSheetName(cell.sheetId, file, addedSheetNames)
    const list = formulaValuesBySheet.get(sheetName) ?? []
    list.push({ row: cell.row, column: cell.column, value: cell.value })
    formulaValuesBySheet.set(sheetName, list)
  }
  const formulaValues = [...formulaValuesBySheet].map(([sheetName, cells]) => ({
    sheetName,
    cells,
  }))

  const filterStates = (request.filterStates ?? []).map((state: any) => ({
    sheetName: resolveSheetName(state.sheetId, file, addedSheetNames),
    filter: state.filter,
    hiddenRows: state.hiddenRows,
    visibilityRange: state.visibilityRange,
  }))

  const linksBySheet = new Map<string, any[]>()
  for (const link of request.hyperlinkEdits ?? []) {
    const sheetName = resolveSheetName(link.sheetId, file, addedSheetNames)
    const list = linksBySheet.get(sheetName) ?? []
    list.push({ row: link.row, column: link.column, target: link.target })
    linksBySheet.set(sheetName, list)
  }
  const hyperlinkEdits = [...linksBySheet].map(([sheetName, links]) => ({ sheetName, edits: links }))

  const cfStates = (request.cfStates ?? []).map((state: any) => ({
    sheetName: resolveSheetName(state.sheetId, file, addedSheetNames),
    rules: state.rules,
  }))
  const dvStates = (request.dvStates ?? []).map((state: any) => ({
    sheetName: resolveSheetName(state.sheetId, file, addedSheetNames),
    rules: state.rules,
  }))
  const sheetProtections = (request.sheetProtections ?? []).map(({ sheetId, ...state }: any) => ({
    sheetName: resolveSheetName(sheetId, file, addedSheetNames),
    ...state,
  }))
  const protectedRangeStates = (request.protectedRangeStates ?? []).map((state: any) => ({
    sheetName: resolveSheetName(state.sheetId, file, addedSheetNames),
    ranges: state.ranges,
  }))
  const pageSetupStates = (request.pageSetupStates ?? []).map(({ sheetId, ...state }: any) => ({
    sheetName: resolveSheetName(sheetId, file, addedSheetNames),
    ...state,
  }))
  const noteStates = (request.noteStates ?? []).map(({ sheetId, notes }: any) => ({
    sheetName: resolveSheetName(sheetId, file, addedSheetNames),
    notes,
  }))
  const visualAdditions = (request.visualAdditions ?? []).map((addition: any) => ({
    sheetName: resolveSheetName(addition.sheetId, file, addedSheetNames),
    anchor: addition.anchor,
    chart: addition.chart,
    shape: addition.shape,
    image: addition.image,
  }))
  const cellImageAdditions = (request.cellImageAdditions ?? []).map(
    ({ sheetId, ...addition }: any) => ({
      sheetName: resolveSheetName(sheetId, file, addedSheetNames),
      ...addition,
    }),
  )
  const tableAdditions = (request.tableAdditions ?? []).map((table: any) => ({
    sheetName: resolveSheetName(table.sheetId, file, addedSheetNames),
    area: table.area,
    name: table.name,
    columnNames: table.columnNames,
    style: table.style,
    bandedRows: table.bandedRows,
    options: {
      headerRow: table.headerRow,
      totalsRow: table.totalsRow,
      firstColumn: table.firstColumn,
      lastColumn: table.lastColumn,
      bandedColumns: table.bandedColumns,
      filterButton: table.filterButton,
    },
  }))
  const tableEdits = (request.tableEdits ?? []).map(({ sheetId, ...edit }: any) => ({
    ...edit,
    sheetName: resolveSheetName(sheetId, file, addedSheetNames),
  }))
  const sparklineAdditions = (request.sparklineAdditions ?? []).map(
    ({ sheetId, ...group }: any) => ({
      sheetName: resolveSheetName(sheetId, file, addedSheetNames),
      ...group,
    }),
  )
  const pivotAdditions = (request.pivotAdditions ?? []).map((pivot: any) => ({
    sheetName: resolveSheetName(pivot.sheetId, file, addedSheetNames),
    sourceSheetName: resolveSheetName(pivot.sourceSheetId, file, addedSheetNames),
    sourceArea: pivot.sourceArea,
    location: pivot.location,
    name: pivot.name,
    fieldNames: pivot.fieldNames,
    rowFieldIndices: pivot.rowFieldIndices,
    columnFieldIndex: pivot.columnFieldIndex,
    pageFieldIndices: pivot.pageFieldIndices,
    pageLevelItems: pivot.pageLevelItems,
    pageItems: pivot.pageItems,
    rowItems: pivot.rowItems,
    rowLevelItems: pivot.rowLevelItems,
    rowLines: pivot.rowLines,
    columnItems: pivot.columnItems,
    columnFieldIndices: pivot.columnFieldIndices,
    colLevelItems: pivot.colLevelItems,
    colLines: pivot.colLines,
    groupings: pivot.groupings,
    filters: pivot.filters,
    rowHiddenItems: pivot.rowHiddenItems,
    colHiddenItems: pivot.colHiddenItems,
    values: pivot.values,
  }))
  const pivotRefreshUpdates = (request.pivotRefreshUpdates ?? []).map((update: any) => ({
    cachePath: update.cachePath,
    sheetName: resolveSheetName(update.sheetId, file, addedSheetNames),
    newOutputRef: update.newOutputRef,
    relayout: update.relayout,
  }))

  return {
    edits,
    structuralOps,
    ...(sheetPlan === undefined ? {} : { sheetPlan }),
    ...(request.chartEdits?.length ? { chartEdits: request.chartEdits } : {}),
    ...(filterStates.length > 0 ? { filterStates } : {}),
    ...(hyperlinkEdits.length > 0 ? { hyperlinkEdits } : {}),
    ...(cfStates.length > 0 ? { cfStates } : {}),
    ...(dvStates.length > 0 ? { dvStates } : {}),
    ...(sheetProtections.length > 0 ? { sheetProtections } : {}),
    ...(request.definedNamesState ? { definedNamesState: request.definedNamesState } : {}),
    ...(visualAdditions.length > 0 ? { visualAdditions } : {}),
    ...(pageSetupStates.length > 0 ? { pageSetupStates } : {}),
    ...(noteStates.length > 0 ? { noteStates } : {}),
    ...(tableAdditions.length > 0 ? { tableAdditions } : {}),
    ...(pivotAdditions.length > 0 ? { pivotAdditions } : {}),
    ...(request.pivotCacheRefreshPaths?.length
      ? { pivotCacheRefreshPaths: request.pivotCacheRefreshPaths }
      : {}),
    ...(pivotRefreshUpdates.length > 0 ? { pivotRefreshUpdates } : {}),
    ...(request.visualEdits?.length ? { visualEdits: request.visualEdits } : {}),
    ...(sparklineAdditions.length > 0 ? { sparklineAdditions } : {}),
    ...(formulaValues.length > 0 ? { formulaValues } : {}),
    ...(request.themeState ? { themeState: request.themeState } : {}),
    ...(request.workbookProtectionState
      ? { workbookProtectionState: request.workbookProtectionState }
      : {}),
    ...(protectedRangeStates.length > 0 ? { protectedRangeStates } : {}),
    ...(bulkConstantFills.length > 0 ? { bulkConstantFills } : {}),
    ...(tabColorStates.length > 0 ? { tabColorStates } : {}),
    ...(tableEdits.length > 0 ? { tableEdits } : {}),
    ...(cellImageAdditions.length > 0 ? { cellImageAdditions } : {}),
  }
}

/// 回退路径：页面自己产整份字节（`plan { kind:'bytes' }`）。
export async function saveWorkbookEdits(
  source: Uint8Array,
  request: any,
  file: any,
): Promise<{ bytes: Uint8Array; file: any; touchedEntries: string[] }> {
  const plan = planWorkbookSave(request, file)
  const sourceBuffer = Buffer.from(source)
  const mutation = await assembleWithJsZip(
    sourceBuffer,
    await planCellEditsToXlsx(
      await createBufferEntrySource(sourceBuffer),
      plan.edits,
      plan.structuralOps,
      plan.chartEdits ?? [],
      plan.sheetPlan,
      plan.filterStates ?? [],
      plan.hyperlinkEdits ?? [],
      plan.cfStates ?? [],
      plan.dvStates ?? [],
      plan.sheetProtections ?? [],
      plan.definedNamesState ?? null,
      plan.visualAdditions ?? [],
      plan.pageSetupStates ?? [],
      plan.noteStates ?? [],
      plan.tableAdditions ?? [],
      plan.pivotAdditions ?? [],
      plan.pivotCacheRefreshPaths ?? [],
      plan.pivotRefreshUpdates ?? [],
      plan.visualEdits ?? [],
      plan.sparklineAdditions ?? [],
      plan.formulaValues ?? [],
      plan.themeState ?? null,
      plan.workbookProtectionState ?? null,
      plan.protectedRangeStates ?? [],
      plan.bulkConstantFills ?? [],
      plan.tabColorStates ?? [],
      plan.tableEdits ?? [],
      plan.cellImageAdditions ?? [],
    ),
  )

  const nextFile = {
    ...file,
    name: request.mode === 'save-as' && request.targetPath ? file.name : file.name,
    fileBytes: mutation.buffer.length,
  }
  return {
    bytes: new Uint8Array(mutation.buffer),
    file: nextFile,
    touchedEntries: [...mutation.touchedEntries],
  }
}
