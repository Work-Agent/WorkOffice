// ─── XLSX-PRO 内部类型：计划形状 + 保存请求 ───────────────────────────
//
// 计划词汇与 docx-pro 对齐：
//   bytes  页面完整序列化出的整份字节（页面路径的回退通道）
//   save   保存请求（页面与无页面共用；exec 调 xlsx-gateway 产字节）
//   new    新建工作簿（exec 用 seed.xlsx + 初始 sheet/cell）
//   text / noop / check / reject 只读结果、空操作、诊断与被拒批
//
// `XlsxSaveRequest` 是 `planCellEditsToXlsx` 参数的具名版本（sheet 一律用名称）：
// 页面能产的每一组都要能表达，否则切到 `kind:'save'` 就是在静默丢数据。

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

export interface Rejection {
  index: number
  code: string
  message: string
}

export class CommandRejected extends Error {
  constructor(
    readonly index: number,
    readonly code: string,
    message: string,
  ) {
    super(message)
    this.name = 'CommandRejected'
  }
}

/**
 * 保存请求（sheet 一律用名称）。
 *
 * 前 3 个是两条路径共用的核心集（无页面只产出这些）；其后各组是页面能产、
 * headless 不产的，exec 一路透传给 `planCellEditsToXlsx`。
 */
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

export interface NewSheetInput {
  name: string
  rows?: CellScalar[][]
}

export type XlsxProPlan =
  | { kind: 'bytes'; bytes: Uint8Array; applied?: number }
  | {
      kind: 'save'
      request: XlsxSaveRequest
      changedCount: number
      applied?: number
      dryRun?: boolean
      warnings?: string[]
    }
  | { kind: 'new'; sheets: NewSheetInput[]; applied?: number; dryRun?: boolean }
  | { kind: 'noop'; applied?: number; summary?: string }
  | { kind: 'text'; text: string }
  // 结构化只读结果（页面 invoke 用）：原样透传引擎的 RangeResult 等形状。
  | { kind: 'read'; data: unknown; audit?: { sessionId?: string; cached?: boolean } }
  | { kind: 'check'; applied?: number; audit?: unknown }
  | { kind: 'reject'; applied?: number; rejected: Rejection[] }
