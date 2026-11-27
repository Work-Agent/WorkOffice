// ─── 保存前重算（wasm `recalc_cells`） ────────────────────────────────
//
// 顺序不能反（上游 §14.3.5）：先拿到 stale 缓存值与新写公式的计算结果，
// 再和本次编辑一起交给 gateway 组包；反过来会把旧 `<v>` 写回文件。
// 重算失败不阻断保存，只在 warnings 里说明（Excel 打开时会全量重算）。

import { randomUUID } from 'node:crypto'
import { readFile, rm } from 'node:fs/promises'
import { join } from 'node:path'

import type { CellEdit, FormulaCachedValue, SheetFormulaValues } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'

import type { XlsxSaveRequest } from '../shared/save-request'
import { getScope } from './scope'
import { loadWasmEngine } from './wasm'

interface RecalcCell {
  sheet: string
  row: number
  column: number
  formatted: string
  number?: number
  boolean?: boolean
  isError: boolean
  isFormula: boolean
}

function recalcInputOf(edit: CellEdit): string {
  const cell = edit.cell as { value?: unknown; formula?: string } | undefined
  if (cell?.formula) return cell.formula
  if (cell?.value === null || cell?.value === undefined) return ''
  return String(cell.value)
}

function cachedValueOf(cell: RecalcCell): FormulaCachedValue {
  if (cell.isError) return { error: cell.formatted }
  if (typeof cell.number === 'number') return cell.number
  if (typeof cell.boolean === 'boolean') return cell.boolean
  return cell.formatted
}

export interface RecalcOutcome {
  formulaValues: SheetFormulaValues[] | null
  notices: string[]
}

export async function recalcFormulaValues(
  baseline: Uint8Array,
  request: XlsxSaveRequest,
): Promise<RecalcOutcome> {
  const engine = await loadWasmEngine(getScope().tempDir)
  if (!engine) return { formulaValues: null, notices: [] }

  const notices: string[] = []
  const runId = randomUUID().slice(0, 8)
  const fileName = `recalc-${runId}.xlsx`
  const guest = await engine.writeTemp(fileName, baseline)
  try {
    const result = await engine.invoke<{
      cells?: RecalcCell[]
      stale?: RecalcCell[]
      staleUnrefreshed?: unknown[]
      auditSkipped?: string
    }>({
      version: 1,
      requestId: `recalc-${runId}`,
      command: 'recalc_cells',
      path: guest,
      edits: request.edits
        .filter((edit) => edit.writeValue)
        .map((edit) => ({
          sheet: edit.sheetName,
          row: edit.row,
          column: edit.column,
          input: recalcInputOf(edit),
        })),
      // 新写的公式单元格在文件里没有旧 `<v>`，`stale` 审计抓不到它 —— 显式
      // 读回来拿引擎算出的值（1×1 range、按编辑顺序、上限 2000 条）。
      reads: request.edits
        .filter(
          (edit) => edit.writeValue && typeof (edit.cell as { formula?: string } | undefined)?.formula === 'string',
        )
        .slice(0, 2000)
        .map((edit) => ({
          sheet: edit.sheetName,
          range: {
            startRow: edit.row,
            endRow: edit.row,
            startColumn: edit.column,
            endColumn: edit.column,
          },
        })),
      reportStale: true,
    })
    if (result.auditSkipped) notices.push(`recalc-audit-skipped:${String(result.auditSkipped)}`)
    if ((result.staleUnrefreshed?.length ?? 0) > 0)
      notices.push(`recalc-unrefreshed:${result.staleUnrefreshed!.length}`)

    const merged = new Map<string, RecalcCell>()
    const keyOf = (cell: RecalcCell) => `${cell.sheet}\u0000${cell.row}\u0000${cell.column}`
    for (const cell of result.cells ?? []) if (cell.isFormula) merged.set(keyOf(cell), cell)
    for (const cell of result.stale ?? []) merged.set(keyOf(cell), cell)
    const stale = [...merged.values()]
    if (stale.length === 0) return { formulaValues: null, notices }

    const bySheet = new Map<string, Array<{ row: number; column: number; value: FormulaCachedValue }>>()
    for (const cell of stale) {
      const list = bySheet.get(cell.sheet) ?? []
      list.push({ row: cell.row, column: cell.column, value: cachedValueOf(cell) })
      bySheet.set(cell.sheet, list)
    }
    return {
      formulaValues: [...bySheet].map(([sheetName, cells]) => ({ sheetName, cells })),
      notices,
    }
  } finally {
    await rm(join(engine.tempDir, fileName), { force: true }).catch(() => undefined)
  }
}

export async function baselineOf(file: { path: string }): Promise<Uint8Array> {
  return new Uint8Array(await readFile(file.path))
}
