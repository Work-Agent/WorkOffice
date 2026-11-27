// ─── 字节侧读投影（命令的唯一读实现） ─────────────────────────────────
//
// P0 用 TS 解析（`shared/workbook`）；wasm 结构化读在 P1 接入，
// 缓存形状不变：按 ref 存一份，哈希变了才重解析。

import { readFile } from 'node:fs/promises'

import {
  auditText,
  contextOf,
  contextText,
  findCells,
  parseWorkbook,
  readRange,
  sheetText,
  usedRange,
  type RangeRead,
  type WorkbookContext,
  type WorkbookView,
} from '../shared/workbook'
import { hashOfFile, sha256Hex, rememberHash } from './hash'
import { wasmMeta } from './wasm-read'
import { invalid, notFound } from './errors'
import type { XlsxFile } from './ref'
import { tierOf } from './ref'

export interface LoadedView {
  ref: string
  hash: string
  size: number
  view: WorkbookView
}

/** 只留最近 2 个工作簿（与 wasm 会话预算一致）。 */
const MAX_VIEWS = 2
const views = new Map<string, LoadedView>()

export async function loadView(file: XlsxFile): Promise<LoadedView> {
  const hash = await hashOfFile(file)
  const hit = views.get(file.ref)
  if (hit && hit.hash === hash) {
    views.delete(file.ref)
    views.set(file.ref, hit)
    return hit
  }
  const bytes = new Uint8Array(await readFile(file.path))
  const actual = sha256Hex(bytes)
  rememberHash(file.path, file.size, file.mtime, actual)
  let view: WorkbookView
  try {
    view = await parseWorkbook(bytes)
  } catch (error) {
    invalid(`不是有效的 XLSX 包：${error instanceof Error ? error.message : String(error)}`)
  }
  const entry: LoadedView = { ref: file.ref, hash: actual, size: bytes.byteLength, view }
  views.set(file.ref, entry)
  while (views.size > MAX_VIEWS) {
    const oldest = views.keys().next().value
    if (!oldest) break
    views.delete(oldest)
  }
  return entry
}

/** 页面 `meta`：wasm 可用时给 open 的完整元数据，否则给薄清单（页面回退整份字节）。 */
export async function pageMeta(file: XlsxFile, hash: string): Promise<Record<string, unknown>> {
  const wasm = await wasmMeta(file, hash)
  if (wasm) {
    const meta = (wasm.meta ?? {}) as Record<string, unknown>
    return {
      ref: file.ref,
      name: file.name,
      size: file.size,
      hash,
      tier: tierOf(file.size),
      source: 'wasm',
      ...meta,
    }
  }
  const entry = await loadView(file)
  const context = contextOf(entry.view, file.name)
  return {
    ref: file.ref,
    name: file.name,
    size: file.size,
    hash: entry.hash,
    tier: tierOf(file.size),
    source: 'ts',
    sheets: context.sheets,
    active: context.sheets[0]?.name ?? '',
  }
}

export function resetViews(): void {
  views.clear()
}

/** 页面通道的 `meta`：工作表清单 + 已用区域（几 KB，不传单元格）。 */
export function metaOf(entry: LoadedView, name: string): {
  ref: string
  name: string
  size: number
  hash: string
  tier: string
  sheets: WorkbookContext['sheets']
  active: string
} {
  const context = contextOf(entry.view, name)
  return {
    ref: entry.ref,
    name,
    size: entry.size,
    hash: entry.hash,
    tier: tierOf(entry.size),
    sheets: context.sheets,
    active: context.sheets[0]?.name ?? '',
  }
}

function sheetNameOf(view: WorkbookView, wanted: string): string {
  if (view.byName.has(wanted)) return wanted
  const byId = view.sheetNamesById[wanted]
  if (byId) return byId
  const lower = wanted.toLowerCase()
  const match = view.sheets.find((sheet) => sheet.name.toLowerCase() === lower)
  if (match) return match.name
  return notFound(`工作表不存在：${wanted}（可用：${view.sheets.map((s) => s.name).join(' / ')}）`)
}

export function rangeOf(
  entry: LoadedView,
  input: { sheetId?: string; sheet?: string; range?: string; start?: string; end?: string },
  maxCells = 5_000,
): RangeRead & { hash: string; size: number; sheetId: string } {
  const wanted = String(input.sheetId ?? input.sheet ?? '')
  if (!wanted) invalid('缺少 sheet / sheetId')
  const sheetName = sheetNameOf(entry.view, wanted)
  const text =
    typeof input.range === 'string' && input.range
      ? input.range
      : `${input.start ?? 'A1'}:${input.end ?? input.start ?? 'A1'}`
  let result: RangeRead
  try {
    result = readRange(entry.view, sheetName, text, maxCells)
  } catch (error) {
    invalid(error instanceof Error ? error.message : String(error))
  }
  const sheetId = entry.view.sheets.find((sheet) => sheet.name === sheetName)?.id ?? sheetName
  return { ...result, hash: entry.hash, size: entry.size, sheetId }
}

export type ReadWhat = 'context' | 'sheets' | 'range' | 'formulas' | 'find' | 'audit'

export interface ReadProjection {
  text: string
  data?: unknown
}

export async function projectRead(
  entry: LoadedView,
  file: XlsxFile,
  what: ReadWhat,
  params: Record<string, unknown>,
): Promise<ReadProjection> {
  if (what === 'context') return { text: contextText(entry.view, file.name) }
  if (what === 'sheets') {
    const parts = entry.view.sheets.map((sheet) => sheetText(entry.view, sheet.name, 20, 6_000))
    return { text: parts.join('\n\n').slice(0, 200_000) }
  }
  if (what === 'range') {
    const result = rangeOf(entry, params as never)
    const lines = [`# ${result.sheet}!${result.range}`]
    for (const row of result.rows) {
      lines.push(`${row.row}\t${row.cells.map((cell) => cell.text.replace(/\t/g, ' ')).join('\t')}`)
    }
    if (result.truncated) lines.push('…（超过上限，已截断）')
    return { text: lines.join('\n') }
  }
  if (what === 'formulas') {
    const limit = Math.max(1, Math.min(2_000, Number(params['limit'] ?? 200)))
    const sheetFilter = typeof params['sheet'] === 'string' ? params['sheet'] : undefined
    const lines = ['# formulas']
    let count = 0
    for (const sheet of entry.view.sheets) {
      if (sheetFilter && sheet.name !== sheetFilter) continue
      for (const [address, cell] of Object.entries(sheet.cells)) {
        if (!cell.formula) continue
        lines.push(`${sheet.name}!${address}\t${cell.formula}`)
        count += 1
        if (count >= limit) {
          lines.push(`…（截断于 ${limit} 条）`)
          return { text: lines.join('\n') }
        }
      }
    }
    if (count === 0) lines.push('(no formulas)')
    return { text: lines.join('\n') }
  }
  if (what === 'find') {
    const query = String(params['query'] ?? params['q'] ?? '')
    if (!query) invalid('缺少 query')
    const result = findCells(entry.view, query, {
      ...(typeof params['sheet'] === 'string' ? { sheet: params['sheet'] } : {}),
      ...(params['regex'] === true ? { regex: true } : {}),
      ...(params['matchCase'] === true ? { matchCase: true } : {}),
      ...(typeof params['limit'] === 'number' ? { limit: params['limit'] } : {}),
    })
    const lines = [`# find ${query}`, `hits=${result.hits.length}${result.truncated ? ' (truncated)' : ''}`]
    for (const hit of result.hits) lines.push(`${hit.sheet}!${hit.address}\t${hit.text}`)
    return { text: lines.join('\n') }
  }
  const bytes = new Uint8Array(await readFile(file.path))
  return { text: await auditText(bytes, entry.view) }
}

export { usedRange }
