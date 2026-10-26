/**
 * Shared workbook reader/search types.
 *
 * These used to live in the AI tool schema module. The reader, search, and
 * status-bar code still need them after the AI layer is removed, so they are
 * kept here without any agent/runtime dependency.
 */
import type { RangeBounds } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { CellScalar } from '@genoffice/xlsx-gateway/domain/workbook.types'

export interface SheetRef {
  readonly id: string
  readonly name: string
  readonly rows?: number
  readonly columns?: number
  readonly readOnlyOversized?: boolean
}

export interface ChartRef {
  readonly path: string
  readonly title: string
  readonly types: string
  readonly sheetId: string
}

export interface FrozenSelection {
  readonly a1: string
  readonly sheetId: string
  readonly columns?: readonly string[]
}

export interface ActiveSheetInfo {
  readonly mode: 'demo' | 'lazy' | 'none'
  readonly streaming?: boolean | undefined
  readonly sheetId: string
  readonly sheetName: string
  readonly revision?: number
  readonly knownAddresses: readonly string[]
  readonly loadedRange?: string | undefined
  readonly sheets: readonly SheetRef[]
  readonly selection?: string | undefined
  readonly selectionFrozen?: boolean | undefined
  readonly selectionColumns?: readonly string[] | undefined
  readonly merges?: readonly string[] | undefined
  readonly charts?: readonly ChartRef[] | undefined
}

export interface FindCellsOptions {
  readonly query: string
  readonly regex: boolean
  readonly lookIn: 'values' | 'formulas' | 'both'
  readonly errorsOnly: boolean
  readonly sheetId?: string | undefined
  readonly maxResults: number
}

export interface FindCellsMatch {
  readonly sheetName: string
  readonly address: string
  readonly value: CellScalar
  readonly formula?: string | undefined
}

export interface FindCellsOutcome {
  readonly matches: readonly FindCellsMatch[]
  readonly truncated: boolean
  readonly incompleteSheets: readonly string[]
  readonly error?: string
}

export interface SelectRangeOutcome {
  readonly ok: boolean
  readonly sheetName?: string
  readonly error?: string
}

export type { RangeBounds }
