// ─── wasm 结构化读：meta（open 原样元数据）+ range（read_range） ─────
//
// 页面流式打开时只调这两个 op：一次 `open` 拿工作表清单/样式/列宽/合并，
// 之后按可视区 `read_range`。会话按 (ref, hash) 缓存，保存后 hash 变即换会话。

import { readFile } from 'node:fs/promises'

import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import { CommandError } from './errors'
import { getScope } from './scope'
import { ensureSession } from './sessions'
import { loadWasmEngine } from './wasm'

const workDir = (): string => getScope().tempDir

export interface WasmMeta {
  source: 'wasm'
  meta: unknown
  sessionId: string
  cached: boolean
}

export async function wasmMeta(file: { ref: string; path: string }, hash: string): Promise<WasmMeta | null> {
  const engine = await loadWasmEngine(workDir())
  if (!engine) return null
  const session = await ensureSession(engine, file.ref, hash, () => readFile(file.path).then((b) => new Uint8Array(b)))
  return { source: 'wasm', meta: session.metadata, sessionId: session.sessionId, cached: session.cached }
}

export async function wasmRange(
  file: { ref: string; path: string },
  hash: string,
  input: { sheetId?: string; sheet?: string; range?: string; start?: string; end?: string },
): Promise<unknown> {
  const engine = await loadWasmEngine(workDir())
  if (!engine)
    throw new CommandError(
      'command.unsupported',
      '结构化范围读取需要 wasm 引擎（当前为 TS 底座）；页面应回退到整份字节',
    )

  const battery = await ensureSession(engine, file.ref, hash, () => readFile(file.path).then((b) => new Uint8Array(b)))
  const wanted = String(input.sheetId ?? input.sheet ?? '')
  if (!wanted) throw new CommandError('command.invalid', '缺少 sheet / sheetId')
  const sheetId =
    battery.sheets.find((sheet) => sheet.id === wanted)?.id ??
    battery.sheets.find((sheet) => sheet.name === wanted)?.id
  if (!sheetId)
    throw new CommandError(
      'command.not-found',
      `工作表不存在：${wanted}（可用：${battery.sheets.map((sheet) => sheet.name).join(' / ')}）`,
    )

  const bounds = boundsOf(input)
  return engine
    .invoke({
      version: 1,
      requestId: `range-${Date.now()}`,
      command: 'read_range',
      sessionId: battery.sessionId,
      sheetId,
      range: bounds,
    })
    .catch((error: unknown) => {
      const message = error instanceof Error ? error.message : String(error)
      // 空表（`<dimension ref="A1"/>`）读 A1:C3 会被引擎判成越界；空白区域
      // 应当渲染成空格子，而不是让编辑器整页失败。
      if (/outside the worksheet/i.test(message)) return emptyRange(bounds)
      throw new CommandError('command.invalid', `read_range 失败：${message}`)
    })
}

/** 空白区域：与 read_range 同形的空结果（编辑器只读这些键）。 */
function emptyRange(bounds: {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}): Record<string, unknown> {
  const rows: Array<{ row: number; hidden: boolean }> = []
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) rows.push({ row, hidden: false })
  return {
    cells: [],
    rows,
    merges: [],
    hyperlinks: [],
    conditionalRules: [],
    autoFilter: null,
    autoFilterColumns: [],
    dataValidations: [],
    sheetProtection: null,
    rowBreaks: [],
    colBreaks: [],
    pageSetup: null,
    protectedRanges: [],
    indexedThroughRow: bounds.endRow + 1,
    indexingComplete: true,
  }
}

function boundsOf(input: {
  range?: unknown
  start?: string
  end?: string
}): { startRow: number; endRow: number; startColumn: number; endColumn: number } {
  // 编辑器传的是 `{startRow,endRow,startColumn,endColumn}`；也接受 A1 文本。
  if (input.range && typeof input.range === 'object') {
    const record = input.range as Record<string, unknown>
    const bounds = {
      startRow: Number(record['startRow'] ?? 0),
      endRow: Number(record['endRow'] ?? 0),
      startColumn: Number(record['startColumn'] ?? 0),
      endColumn: Number(record['endColumn'] ?? 0),
    }
    const values = [bounds.startRow, bounds.endRow, bounds.startColumn, bounds.endColumn]
    if (values.some((value) => !Number.isInteger(value) || value < 0))
      throw new CommandError('command.invalid', 'range 边界非法')
    if (bounds.startRow > bounds.endRow || bounds.startColumn > bounds.endColumn)
      throw new CommandError('command.invalid', 'range 边界顺序颠倒')
    return bounds
  }
  const text =
    typeof input.range === 'string' && input.range
      ? input.range
      : `${input.start ?? 'A1'}:${input.end ?? input.start ?? 'A1'}`
  try {
    return parseRange(text)
  } catch (error) {
    throw new CommandError(
      'command.invalid',
      `range 非法：${text}（${error instanceof Error ? error.message : String(error)}）`,
    )
  }
}
