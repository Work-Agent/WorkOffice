// ─── 结构化范围读取（页面 invoke 的数据源） ────────────────────────────
//
// 页面调 `ura.page.invoke(sessionId, 'xlsx_read', { what:'cells', sheetId, range })`，
// 宿主把只读工具交到执行点，这里用 wasm 的 `read_range` 回答：
//
//   - O(范围)，不是 O(整包)：首次调用建索引（秒级），之后实测 1 ms；
//   - 返回**原样**的 `RangeResult`（cells / rows / merges / … / indexedThroughRow），
//     与雪 fork 里 `readWorkbookRange` 期望的形状逐键一致 —— 页面只做透传，
//     不做字段翻译（翻译就是第二套语义，迟早漂）。
//
// TS 回退底座没有索引，无法提供这个形状：明确 `command.unsupported`，
// 让页面回退到自己的 parse 路径，而不是返回空结果。

import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'

import { coded, messageOf, type EngineCtx, type PlanRequest } from './abi'
import { ensureSession } from './sessions'
import { loadWasmEngine } from './wasm'

type AnyRecord = Record<string, unknown>

export interface PageMetaRead {
  readonly data: unknown
  readonly sessionId: string
  readonly cached: boolean
}

export interface PageRangeRead {
  readonly data: unknown
  readonly sessionId: string
  readonly cached: boolean
  readonly sheetId: string
}

export async function readRangeForPage(
  req: PlanRequest,
  params: AnyRecord,
  ctx: EngineCtx,
): Promise<PageRangeRead> {
  const engine = await loadWasmEngine(ctx.sessionDir ?? ctx.tempDir)
  if (!engine)
    throw coded(
      'command.unsupported',
      '结构化范围读取需要 wasm 引擎（当前为 TS 回退底座）；页面应回退到本地解析',
    )

  const wanted = String(params['sheetId'] ?? params['sheet'] ?? '')
  if (!wanted) throw coded('command.invalid', '缺少 sheet / sheetId')
  const bounds = rangeOf(params)

  // 版本优先用调用方给的（宿主已知），否则才读字节（命中缓存时不会读整份文件）。
  const version = req.baseVersion ?? (await ctx.readBytes()).version
  const session = await ensureSession(engine, req.ref, version, async () =>
    req.bytes ?? (await ctx.readBytes()).bytes,
  )

  const sheetId =
    session.sheets.find((sheet) => sheet.id === wanted)?.id ??
    session.sheets.find((sheet) => sheet.name === wanted)?.id
  if (!sheetId)
    throw coded(
      'command.not-found',
      `工作表不存在：${wanted}（可用：${session.sheets.map((sheet) => sheet.name).join(' / ')}）`,
    )

  const data = await engine
    .invoke({
      version: 1,
      requestId: `range-${bounds.startRow}-${bounds.endRow}`,
      command: 'read_range',
      sessionId: session.sessionId,
      sheetId,
      range: bounds,
    })
    .catch((error: unknown) => {
      throw coded('command.invalid', `read_range 失败：${messageOf(error)}`)
    })

  return { data, sessionId: session.sessionId, cached: session.cached, sheetId }
}

/**
 * 工作簿元数据（`open` 的原样返回）：`sheets / styles / definedNames / visuals /
 * activeTab / normalFontName …`。页面用它重建 file 模型，从而不再解析整包。
 */
export async function readMetaForPage(req: PlanRequest, ctx: EngineCtx): Promise<PageMetaRead> {
  const engine = await loadWasmEngine(ctx.sessionDir ?? ctx.tempDir)
  if (!engine)
    throw coded('command.unsupported', '元数据读取需要 wasm 引擎（当前为 TS 回退底座）')
  const version = req.baseVersion ?? (await ctx.readBytes()).version
  const session = await ensureSession(engine, req.ref, version, async () =>
    req.bytes ?? (await ctx.readBytes()).bytes,
  )
  // ensureSession 已经把 open 的元数据缓存在会话里，这里复用同一份，不再多开一次。
  return { data: session.metadata, sessionId: session.sessionId, cached: session.cached }
}

function rangeOf(params: AnyRecord): {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
} {
  const direct = params['range']
  if (direct && typeof direct === 'object') {
    const record = direct as AnyRecord
    const bounds = {
      startRow: Number(record['startRow'] ?? record['start_row'] ?? 0),
      endRow: Number(record['endRow'] ?? record['end_row'] ?? 0),
      startColumn: Number(record['startColumn'] ?? record['start_column'] ?? 0),
      endColumn: Number(record['endColumn'] ?? record['end_column'] ?? 0),
    }
    if (!isValid(bounds)) throw coded('command.invalid', 'range 边界非法（或顺序颠倒）')
    return bounds
  }
  const text = String(direct ?? `${params['start'] ?? 'A1'}:${params['end'] ?? 'A1'}`)
  try {
    return parseRange(text)
  } catch (error) {
    throw coded('command.invalid', `range 非法：${text}（${messageOf(error)}）`)
  }
}

function isValid(bounds: {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}): boolean {
  const values = [bounds.startRow, bounds.endRow, bounds.startColumn, bounds.endColumn]
  if (values.some((value) => !Number.isInteger(value) || value < 0)) return false
  return bounds.startRow <= bounds.endRow && bounds.startColumn <= bounds.endColumn
}
