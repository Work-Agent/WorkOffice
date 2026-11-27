// ─── 保存请求 → XLSX 字节（引擎侧唯一产字节的地方） ────────────────────
//
// 参数顺序必须与 `planCellEditsToXlsx` 一致：页面能填满的每一组都要透传，
// 少传一组就是静默丢数据（旧包装丢字段的教训）。
// P1 走 JSZip 重打包；wasm `save_archive` 在 P2 接入，调用点不变。

import { unzipSync } from 'fflate'
import {
  assembleWithJsZip,
  assertOnlyTouchedEntriesChanged,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { CellEdit, SheetStructuralOps } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'

import type { XlsxSaveRequest } from '../shared/save-request'
import { invalid } from './errors'

export type SavePlan = Awaited<ReturnType<typeof planCellEditsToXlsx>>

/** 只算「改什么」：wasm `save_archive` 与 JSZip 共用同一份计划。 */
export async function planSaveRequest(
  baseline: Uint8Array,
  request: XlsxSaveRequest,
): Promise<SavePlan> {
  try {
    return await planCellEditsToXlsx(
      await createBufferEntrySource(Buffer.from(baseline)),
      request.edits as CellEdit[],
      request.structuralOps as SheetStructuralOps[],
      request.chartEdits ?? [],
      request.sheetPlan as SheetEditPlan | undefined,
      request.filterStates ?? [],
      request.hyperlinkEdits ?? [],
      request.cfStates ?? [],
      request.dvStates ?? [],
      request.sheetProtections ?? [],
      request.definedNamesState ?? null,
      request.visualAdditions ?? [],
      request.pageSetupStates ?? [],
      request.noteStates ?? [],
      request.tableAdditions ?? [],
      request.pivotAdditions ?? [],
      request.pivotCacheRefreshPaths ?? [],
      request.pivotRefreshUpdates ?? [],
      request.visualEdits ?? [],
      request.sparklineAdditions ?? [],
      request.formulaValues ?? [],
      request.themeState ?? null,
      request.workbookProtectionState ?? null,
      request.protectedRangeStates ?? [],
      request.bulkConstantFills ?? [],
      request.tabColorStates ?? [],
      request.tableEdits ?? [],
      request.cellImageAdditions ?? [],
    )
  } catch (error) {
    invalid(`保存请求无法应用：${error instanceof Error ? error.message : String(error)}`)
  }
}

/** JSZip 重打包（wasm 不可用时的回退）。 */
export async function bytesOfPlan(
  baseline: Uint8Array,
  plan: SavePlan,
): Promise<{ bytes: Uint8Array; touched: string[] }> {
  try {
    const mutation = await assembleWithJsZip(Buffer.from(baseline), plan)
    assertOnlyTouchedEntriesChanged(mutation)
    return {
      // mutation 是 `{ buffer, touchedEntries, … }`；明确拷一次，
      // 避免把 Buffer 的共享池底层 ArrayBuffer 带出去。
      bytes: new Uint8Array(mutation.buffer),
      touched: [...(mutation.touchedEntries ?? plan.touchedEntries)],
    }
  } catch (error) {
    invalid(`保存产物无法组包：${error instanceof Error ? error.message : String(error)}`)
  }
}

export async function bytesOfSaveRequest(
  baseline: Uint8Array,
  request: XlsxSaveRequest,
): Promise<{ bytes: Uint8Array; touched: string[] }> {
  return bytesOfPlan(baseline, await planSaveRequest(baseline, request))
}

/** 拒绝空包 / 非 ZIP / 缺 workbook part 的字节。 */
export function assertXlsxPackage(bytes: Uint8Array): void {
  if (bytes.byteLength === 0) invalid('拒绝写入空字节')
  let archive: Record<string, Uint8Array>
  try {
    archive = unzipSync(bytes, {
      filter: (file) =>
        file.name === '[Content_Types].xml' || file.name === 'xl/workbook.xml',
    })
  } catch {
    invalid('不是有效的 XLSX ZIP 包')
  }
  if (!archive['[Content_Types].xml'] || !archive['xl/workbook.xml'])
    invalid('XLSX 包缺少 [Content_Types].xml 或 xl/workbook.xml')
}
