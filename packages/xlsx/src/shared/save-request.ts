// ─── 页面保存请求 → gateway 参数的具名形状 ─────────────────────────────
//
// 页面用 `planWorkbookSave` 产出这个形状，引擎用 `planCellEditsToXlsx` 消费它。
// 页面能产的每一组都必须在这里有字段，否则「页面能改」会变成「保存丢数据」。

import type {
  BulkConstantFill,
  CellEdit,
  SheetCfState,
  SheetCellImageAddition,
  SheetDvState,
  SheetFormulaValues,
  SheetHyperlinkEdits,
  SheetNoteState,
  SheetPivotAddition,
  SheetProtectedRangesState,
  SheetProtectionState,
  SheetStructuralOps,
  SheetTabColorState,
  SheetTableAddition,
  SheetTableEdit,
  SheetVisualAddition,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { DefinedNamesState } from '@genoffice/xlsx-gateway/gateway/xlsx-defined-names'
import type { SheetFilterState } from '@genoffice/xlsx-gateway/gateway/xlsx-filter'
import type { SheetPageSetupState } from '@genoffice/xlsx-gateway/gateway/xlsx-page-setup'
import type { PivotRefreshUpdate } from '@genoffice/xlsx-gateway/gateway/xlsx-pivot-expand'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import type { SheetSparklineAddition } from '@genoffice/xlsx-gateway/gateway/xlsx-sparkline'
import type { WorkbookThemeState } from '@genoffice/xlsx-gateway/gateway/xlsx-theme'
import type { WorkbookChartEdit, WorkbookVisualEdit } from '@genoffice/xlsx-gateway/shared/edit-schemas'
import type { CellScalar } from '@genoffice/xlsx-gateway/domain/workbook.types'

export interface XlsxSaveRequest {
  edits: CellEdit[]
  structuralOps: SheetStructuralOps[]
  sheetPlan?: SheetEditPlan

  chartEdits?: WorkbookChartEdit[]
  filterStates?: SheetFilterState[]
  hyperlinkEdits?: SheetHyperlinkEdits[]
  cfStates?: SheetCfState[]
  dvStates?: SheetDvState[]
  sheetProtections?: SheetProtectionState[]
  definedNamesState?: DefinedNamesState | null
  visualAdditions?: SheetVisualAddition[]
  pageSetupStates?: SheetPageSetupState[]
  noteStates?: SheetNoteState[]
  tableAdditions?: SheetTableAddition[]
  pivotAdditions?: SheetPivotAddition[]
  pivotCacheRefreshPaths?: string[]
  pivotRefreshUpdates?: PivotRefreshUpdate[]
  visualEdits?: WorkbookVisualEdit[]
  sparklineAdditions?: SheetSparklineAddition[]
  formulaValues?: SheetFormulaValues[]
  themeState?: WorkbookThemeState | null
  workbookProtectionState?: { readonly lockStructure: boolean } | null
  protectedRangeStates?: SheetProtectedRangesState[]
  bulkConstantFills?: BulkConstantFill[]
  tabColorStates?: SheetTabColorState[]
  tableEdits?: SheetTableEdit[]
  cellImageAdditions?: SheetCellImageAddition[]
}
