// ─── engine.exec：计划 → XLSX 字节（唯一字节产生者） ───────────────────
//
// 两种计划来源，一个字节出口：
//   1. save  无页面路径的高层保存请求 → xlsx-gateway planCellEditsToXlsx
//            + assembleWithJsZip；只重写被命令涉及的 part。
//   2. bytes 页面完整序列化出的整份字节；exec 校验包结构 / 体积后原样输出。
//   3. new   新建工作簿：seed.xlsx + 初始 sheet/cell。

import { createHash, randomUUID } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { readFile, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { unzipSync } from 'fflate'
import {
  assembleWithJsZip,
  assertOnlyTouchedEntriesChanged,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { CellEdit, SheetStructuralOps } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { FormulaCachedValue, SheetFormulaValues } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { NewSheetInput, XlsxSaveRequest } from '../shared/types'
import { contextOf, isXlsxPackage, parseWorkbook } from '../shared/workbook'
import { driftedEntries } from '../shared/zip-raw'
import { coded, messageOf, type EngineCtx, type ExecOutcome, type ExecRequest } from './abi'
import { loadWasmEngine, type WasmEngine } from './wasm'

export async function exec(req: ExecRequest, ctx: EngineCtx): Promise<ExecOutcome> {
  if (req.command === 'undo' || req.command === 'redo')
    throw coded('command.unsupported', 'xlsx 的撤销是页面模型历史；无页面路径不提供 undo')

  const plan = req.plan
  if (!plan || typeof plan !== 'object') throw coded('command.invalid', '计划不是对象')
  const kind = (plan as { kind?: unknown }).kind

  if (kind === 'noop')
    return {
      result: { applied: Number((plan as { applied?: number }).applied ?? 0), sync: { kind: 'version' } },
    }

  if (kind === 'text')
    return {
      result: {
        applied: 0,
        audit: { text: String((plan as { text?: unknown }).text ?? '') },
        sync: { kind: 'version' },
      },
    }

  if (kind === 'read')
    return {
      result: {
        applied: 0,
        audit: {
          data: (plan as { data?: unknown }).data,
          ...((plan as { audit?: Record<string, unknown> }).audit ?? {}),
        },
        sync: { kind: 'version' },
      },
    }

  if (kind === 'check')
    return {
      result: {
        applied: 0,
        ...((plan as { audit?: unknown }).audit !== undefined ? { audit: (plan as { audit?: unknown }).audit } : {}),
        sync: { kind: 'version' },
      },
    }

  if (kind === 'reject')
    return {
      result: {
        applied: 0,
        rejected: ((plan as { rejected?: unknown[] }).rejected ?? []) as never,
        sync: { kind: 'version' },
      },
    }

  if (kind === 'bytes') return execBytes(req, plan as { bytes?: unknown; applied?: number }, ctx)
  if (kind === 'new') return execNew(req, plan as { sheets?: NewSheetInput[]; applied?: number; dryRun?: boolean }, ctx)
  if (kind === 'save')
    return execSave(
      req,
      plan as {
        request?: XlsxSaveRequest
        changedCount?: number
        applied?: number
        dryRun?: boolean
        warnings?: string[]
      },
      ctx,
    )

  throw coded('command.invalid', `未知计划类型：${String(kind)}`)
}

function guardOutput(bytes: Uint8Array, ctx: EngineCtx): void {
  if (bytes.byteLength === 0) throw coded('command.invalid', 'exec 产出了空字节')
  if (bytes.byteLength > ctx.budget.maxBytes)
    throw coded('command.too-large', `输出 ${bytes.byteLength} 字节超过预算 ${ctx.budget.maxBytes}`)
}

function asBytes(value: unknown): Uint8Array | null {
  if (value instanceof Uint8Array) return value
  if (Array.isArray(value)) return new Uint8Array(value as number[])
  if (value instanceof ArrayBuffer) return new Uint8Array(value)
  return null
}

async function execBytes(
  req: ExecRequest,
  plan: { bytes?: unknown; applied?: number },
  ctx: EngineCtx,
): Promise<ExecOutcome> {
  const bytes = asBytes(plan.bytes)
  if (!bytes) throw coded('command.invalid', 'bytes 计划缺少字节')
  if (bytes.byteLength === 0) throw coded('command.invalid', '页面提交了空字节')
  if (!isXlsxPackage(bytes)) throw coded('command.invalid', '页面提交的不是有效的 XLSX 包')
  // 页面已经用同一套 xlsx-gateway 产出了合法包；这里只做结构校验，避免
  // headless 解析器对合法但“不常见”的工作簿（图表页、特殊关系路径）误拒。
  let archive: Record<string, Uint8Array>
  try {
    archive = unzipSync(bytes, {
      filter: (file) =>
        file.name === '[Content_Types].xml' || file.name === 'xl/workbook.xml',
    })
  } catch {
    throw coded('command.invalid', '页面提交的不是有效的 XLSX ZIP 包')
  }
  if (!archive['[Content_Types].xml'] || !archive['xl/workbook.xml'])
    throw coded('command.invalid', '页面提交的 XLSX 包缺少必需 part')
  if (req.dryRun) return { result: { applied: plan.applied ?? 1, sync: { kind: 'version' } } }
  guardOutput(bytes, ctx)
  return {
    bytes,
    result: {
      applied: plan.applied ?? 1,
      sync: { kind: 'version' },
      audit: { source: 'page', sha256: sha256Hex(bytes) },
    },
  }
}

async function execSave(
  req: ExecRequest,
  plan: { request?: XlsxSaveRequest; changedCount?: number; applied?: number; dryRun?: boolean; warnings?: string[] },
  ctx: EngineCtx,
): Promise<ExecOutcome> {
  if (!plan.request) throw coded('command.invalid', 'save 计划缺少 request')
  const baseline = req.bytes ?? (await ctx.readBytes()).bytes
  if (!isXlsxPackage(baseline)) throw coded('command.invalid', '基线不是有效的 XLSX 包')
  return applySaveRequest(req, baseline, plan.request, {
    changedCount: plan.changedCount ?? 0,
    applied: plan.applied ?? 0,
    dryRun: req.dryRun || plan.dryRun === true,
    warnings: plan.warnings ?? [],
    ctx,
  })
}

async function applySaveRequest(
  req: ExecRequest,
  baseline: Uint8Array,
  request: XlsxSaveRequest,
  options: {
    changedCount: number
    applied: number
    dryRun: boolean
    warnings: string[]
    ctx: EngineCtx
  },
): Promise<ExecOutcome> {
  // 有 wasm 引擎时：先重算（只为拿到 stale 缓存值），再把 stale `<v>` 与本次
  // 编辑一起交给 gateway 组包 —— 顺序照上游 §14.3.5，反了就会把旧缓存写回文件。
  const engine = options.dryRun
    ? null
    : await loadWasmEngine(options.ctx.sessionDir ?? options.ctx.tempDir)
  const warnings = [...options.warnings]
  const notices: string[] = []
  let formulaValues = request.formulaValues ?? []
  if (engine && request.edits.some((edit) => edit.writeValue)) {
    const recalculated = await recalcFormulaValues(engine, baseline, request, notices).catch(
      (error: unknown) => {
        // 重算失败不阻断保存：文件仍然可写，只是缓存值可能陈旧（gateway 会打
        // calcPr@fullCalcOnLoad 让 Excel 打开时全量重算）。
        warnings.push(`recalc-failed:${messageOf(error).slice(0, 160)}`)
        return null
      },
    )
    if (recalculated) formulaValues = recalculated
  }

  const mutationPlan = await planGateway(baseline, request, formulaValues)
  const touched = [...mutationPlan.touchedEntries]

  if (options.dryRun)
    return {
      result: {
        applied: options.applied,
        ...(warnings.length > 0 ? { warnings } : {}),
        ...(notices.length > 0 ? { notices } : {}),
        audit: {
          changedCount: options.changedCount,
          edits: request.edits.length,
          touchedEntries: touched,
          dryRun: true,
        },
        sync: { kind: 'version' },
      },
    }

  let bytes: Uint8Array
  if (engine) {
    const written = await saveViaWasm(engine, baseline, mutationPlan, options.ctx)
    bytes = written.bytes
    // I1：未改条目的**压缩流**逐字节不变（比解压后 hash 强）
    if (written.drifted.length > 0)
      throw coded(
        'command.invalid',
        `保存改动了未涉及的 ZIP 条目：${written.drifted.slice(0, 5).join(', ')}`,
      )
  } else {
    try {
      const mutation = await assembleWithJsZip(Buffer.from(baseline), mutationPlan)
      bytes = new Uint8Array(mutation.buffer)
      assertOnlyTouchedEntriesChanged(mutation)
    } catch (error) {
      throw coded('command.invalid', messageOf(error))
    }
  }
  guardOutput(bytes, options.ctx)

  let view
  try {
    view = await parseWorkbook(bytes)
  } catch (error) {
    throw coded('command.invalid', `保存产出的 XLSX 非法：${messageOf(error)}`)
  }
  const context = contextOf(view, req.ref)
  return {
    bytes,
    result: {
      applied: options.applied,
      ...(warnings.length > 0 ? { warnings } : {}),
      ...(notices.length > 0 ? { notices } : {}),
      audit: {
        changedCount: options.changedCount,
        edits: request.edits.length,
        touchedEntries: touched,
        sha256: sha256Hex(bytes),
        engine: engine ? 'wasm' : 'js',
        recalculated: formulaValues.length > 0,
      },
      sync: {
        kind: 'manifest',
        sheets: context.sheets,
        sha256: sha256Hex(bytes),
        // 新体积：页面据此更新自己的 size 提示，避免"保存后文件变大"被当成错误
        size: bytes.byteLength,
      },
    },
  }
}

/// gateway 的"改什么"（语义层）。位置参数顺序必须与 `planCellEditsToXlsx`
/// 一致：页面能填满这些组，无页面只填前几项，少传一组就是静默丢数据。
async function planGateway(
  baseline: Uint8Array,
  request: XlsxSaveRequest,
  formulaValues: readonly SheetFormulaValues[],
) {
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
      formulaValues,
      request.themeState ?? null,
      request.workbookProtectionState ?? null,
      request.protectedRangeStates ?? [],
      request.bulkConstantFills ?? [],
      request.tabColorStates ?? [],
      request.tableEdits ?? [],
      request.cellImageAdditions ?? [],
    )
  } catch (error) {
    throw coded('command.invalid', messageOf(error))
  }
}

/// `recalc_cells`：拿 stale 缓存值，转成 gateway 的 `formulaValues`。
async function recalcFormulaValues(
  engine: WasmEngine,
  baseline: Uint8Array,
  request: XlsxSaveRequest,
  notices: string[],
): Promise<SheetFormulaValues[] | null> {
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
      // 新写的公式单元格在文件里没有旧 `<v>`，`stale` 审计抓不到它 —— 必须
      // 显式 read 回来拿引擎算出的值。1x1 range、按编辑顺序、上限 2000 条。
      reads: request.edits
        .filter(
          (edit) =>
            edit.writeValue &&
            typeof (edit.cell as { formula?: string } | undefined)?.formula === 'string',
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
    if (result.auditSkipped)
      notices.push(`recalc-audit-skipped:${String(result.auditSkipped)}`)
    if ((result.staleUnrefreshed?.length ?? 0) > 0)
      notices.push(`recalc-unrefreshed:${result.staleUnrefreshed!.length}`)
    // cells = 本次 read 回来的值（含新写的公式）；stale = 文件里缓存与引擎不一致
    // 的依赖单元格。两者合并去重，stale 后写覆盖（它对应文件里的旧缓存）。
    const merged = new Map<string, RecalcCell>()
    const keyOf = (cell: RecalcCell) => `${cell.sheet}\u0000${cell.row}\u0000${cell.column}`
    for (const cell of result.cells ?? []) if (cell.isFormula) merged.set(keyOf(cell), cell)
    for (const cell of result.stale ?? []) merged.set(keyOf(cell), cell)
    const stale = [...merged.values()]
    if (stale.length === 0) return null
    const bySheet = new Map<string, Array<{ row: number; column: number; value: FormulaCachedValue }>>()
    for (const cell of stale) {
      const list = bySheet.get(cell.sheet) ?? []
      list.push({ row: cell.row, column: cell.column, value: cachedValueOf(cell) })
      bySheet.set(cell.sheet, list)
    }
    return [...bySheet].map(([sheetName, cells]) => ({ sheetName, cells }))
  } finally {
    await rm(join(engine.tempDir, fileName), { force: true }).catch(() => {})
  }
}

function recalcInputOf(edit: CellEdit): string {
  const cell = edit.cell as { value?: unknown; formula?: string } | undefined
  if (cell?.formula) return cell.formula
  if (cell?.value === null || cell?.value === undefined) return ''
  return String(cell.value)
}

/// `recalc_cells` 的单格结果（对应 recalc.rs 的 `RecalcCell`）。
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

function cachedValueOf(cell: RecalcCell): FormulaCachedValue {
  if (cell.isError) return { error: cell.formatted }
  if (typeof cell.number === 'number') return cell.number
  if (typeof cell.boolean === 'boolean') return cell.boolean
  return cell.formatted
}

/// `save_archive`：替换/新增的 entry 以**临时文件**形式交给引擎（`EntryContent`
/// 是 `{name, contentPath}`），未改条目原样拷贝，不再整包重压缩。
async function saveViaWasm(
  engine: WasmEngine,
  baseline: Uint8Array,
  plan: Awaited<ReturnType<typeof planGateway>>,
  ctx: EngineCtx,
): Promise<{ bytes: Uint8Array; drifted: string[] }> {
  const runId = randomUUID().slice(0, 8)
  const written: string[] = []
  const tempName = (suffix: string) => {
    const name = `save-${runId}-${suffix}`
    written.push(name)
    return name
  }
  try {
    const sourceName = tempName('base.xlsx')
    await engine.writeTemp(sourceName, baseline)
    const replacements: Array<{ name: string; contentPath: string }> = []
    let index = 0
    for (const [name, content] of plan.replaced) {
      const fileName = tempName(`r${index++}.entry`)
      replacements.push({ name, contentPath: await engine.writeTemp(fileName, content) })
    }
    const additions: Array<{ name: string; contentPath: string }> = []
    for (const [name, content] of plan.added) {
      const fileName = tempName(`a${index++}.entry`)
      additions.push({ name, contentPath: await engine.writeTemp(fileName, content) })
    }
    for (const [name, content] of plan.addedBinary) {
      const fileName = tempName(`b${index++}.entry`)
      additions.push({ name, contentPath: await engine.writeTemp(fileName, content) })
    }
    const targetName = tempName('out.xlsx')
    await engine.invoke({
      version: 1,
      requestId: `save-${runId}`,
      command: 'save_archive',
      sourcePath: engine.guestPath(sourceName),
      targetPath: engine.guestPath(targetName),
      replacements,
      removals: [...plan.removedEntries],
      additions,
      streamPatches: [],
    })
    const bytes = new Uint8Array(await readFile(join(engine.tempDir, targetName)))
    guardOutput(bytes, ctx)
    // 被删掉的条目（`delete_sheet` 的 worksheet part）本来就应当从包里消失；
    // 不把它们算进“涉及条目”，I1 漂移检查会把删除误判成“保存改了无关的 ZIP 条目”。
    return {
      bytes,
      drifted: driftedEntries(baseline, bytes, [
        ...plan.touchedEntries,
        ...plan.removedEntries,
      ]),
    }
  } catch (error) {
    if (error && typeof error === 'object' && 'code' in error) throw error
    throw coded('command.invalid', `wasm save_archive 失败：${messageOf(error)}`)
  } finally {
    await Promise.all(
      written.map((name) => rm(join(engine.tempDir, name), { force: true }).catch(() => undefined)),
    )
  }
}

let SEED: Uint8Array | null = null

function seedBytes(): Uint8Array {
  SEED ??= new Uint8Array(readFileSync(new URL('../seed.xlsx', import.meta.url)))
  return SEED
}

async function execNew(
  req: ExecRequest,
  plan: { sheets?: NewSheetInput[]; applied?: number; dryRun?: boolean },
  ctx: EngineCtx,
): Promise<ExecOutcome> {
  const sheets = plan.sheets && plan.sheets.length > 0 ? plan.sheets : [{ name: 'Sheet1', rows: [] }]
  const seed = seedBytes()
  const first = sheets[0]!
  const request: XlsxSaveRequest = { edits: [], structuralOps: [] }
  const renames = first.name === 'Sheet1' ? [] : [{ sheetName: 'Sheet1', newName: first.name }]
  const additions = sheets.slice(1).map((sheet) => ({ name: sheet.name }))
  const order = [first.name, ...sheets.slice(1).map((sheet) => sheet.name)]
  if (renames.length > 0 || additions.length > 0) {
    request.sheetPlan = {
      renames,
      additions,
      removals: [],
      order,
      orderChanged: additions.length > 0,
    }
  }
  for (const sheet of sheets) {
    const targetSheet = sheet === first ? 'Sheet1' : sheet.name
    const rows = sheet.rows ?? []
    rows.forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        const address = formatAddress(rowIndex, columnIndex)
        if (typeof value === 'string' && value.startsWith('=')) {
          request.edits.push(cellEdit(targetSheet, address, { value: null, formula: value }))
        } else {
          request.edits.push(cellEdit(targetSheet, address, { value: value ?? null }))
        }
      })
    })
  }
  if (sheets.length === 1 && sheets[0]!.name === 'Sheet1' && (sheets[0]!.rows?.length ?? 0) === 0) {
    if (plan.dryRun || req.dryRun)
      return { result: { applied: 0, sync: { kind: 'version' } } }
    guardOutput(seed, ctx)
    return {
      bytes: seed,
      result: { applied: 0, sync: { kind: 'version' }, audit: { source: 'seed', sha256: sha256Hex(seed) } },
    }
  }
  return applySaveRequest(req, seed, request, {
    changedCount: request.edits.length,
    applied: plan.applied ?? sheets.length,
    dryRun: req.dryRun || plan.dryRun === true,
    warnings: [],
    ctx,
  })
}

function cellEdit(sheetName: string, address: string, cell: { value: string | number | boolean | null; formula?: string }): CellEdit {
  const match = /^([A-Z]+)(\d+)$/.exec(address)
  if (!match) throw coded('command.invalid', `非法地址：${address}`)
  let column = 0
  for (const char of match[1]!) column = column * 26 + char.charCodeAt(0) - 64
  return {
    sheetName,
    row: Number(match[2]) - 1,
    column: column - 1,
    writeValue: true,
    cell,
  }
}

function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

export { seedBytes }
