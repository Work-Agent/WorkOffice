// ─── engine.read：字节侧只读投影 ───────────────────────────────────────
//
// 只读工具由执行点回答，不依赖 Univer。读的是落盘字节；页面有未保存修改时
// 用页面上下文（agent-request kind=context）才能看到最新内容。

import type { CellState } from '@genoffice/xlsx-gateway/domain/workbook.types'
import { formatAddress, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  auditText,
  cellText,
  contextOf,
  contextText,
  findCells,
  parseWorkbook,
  readRange,
  sheetText,
  usedRange,
  type WorkbookView,
} from '../shared/workbook'
import { HeadlessOpError } from '../shared/headless'
import { coded, messageOf, type EngineCtx, type PlanRequest } from './abi'

const MAX_READ_TEXT = 200_000

export const MAX_READ_CELLS = 2_000

export function parse(view: WorkbookView): WorkbookView {
  return view
}

export function context(view: WorkbookView, name: string): string {
  return contextText(view, name).slice(0, MAX_READ_TEXT)
}

export function sheets(view: WorkbookView): string {
  const parts: string[] = []
  for (const sheet of view.sheets) {
    parts.push(sheetText(view, sheet.name, 20, 6_000))
    if (parts.join('\n\n').length > MAX_READ_TEXT) {
      parts.push('…（已截断）')
      break
    }
  }
  return parts.join('\n\n').slice(0, MAX_READ_TEXT)
}

export function range(view: WorkbookView, params: Record<string, unknown>): string {
  const sheet = String(params['sheet'] ?? params['sheetId'] ?? '')
  if (!sheet) throw new HeadlessOpError('command.invalid', '缺少 sheet')
  const rangeText = rangeTextOf(params)
  const result = readRange(view, sheet, rangeText, MAX_READ_CELLS)
  const lines = [`# ${sheet}!${result.range}`]
  for (const row of result.rows) {
    lines.push(`${row.row}\t${row.cells.map((cell) => cell.text.replace(/\t/g, ' ').replace(/\n/g, ' ')).join('\t')}`)
  }
  if (result.truncated) lines.push(`…（超过 ${MAX_READ_CELLS} 格，已截断）`)
  return lines.join('\n')
}

function rangeTextOf(params: Record<string, unknown>): string {
  if (typeof params['range'] === 'string' && params['range']) return params['range']
  const start = typeof params['start'] === 'string' ? params['start'] : null
  const end = typeof params['end'] === 'string' ? params['end'] : null
  if (start && end) return `${start}:${end}`
  if (start) return start
  throw new HeadlessOpError('command.invalid', '缺少 range / start')
}

export function formulas(view: WorkbookView, params: Record<string, unknown>): string {
  const sheetFilter = typeof params['sheet'] === 'string' ? params['sheet'] : undefined
  const limit = Math.max(1, Math.min(2_000, Number(params['limit'] ?? 200)))
  const lines = ['# formulas']
  let count = 0
  for (const sheet of view.sheets) {
    if (sheetFilter && sheet.name !== sheetFilter) continue
    for (const [address, cell] of Object.entries(sheet.cells)) {
      if (!cell.formula) continue
      lines.push(`${sheet.name}!${address}\t${cell.formula}`)
      count += 1
      if (count >= limit) {
        lines.push(`…（截断于 ${limit} 条）`)
        return lines.join('\n')
      }
    }
  }
  if (count === 0) lines.push('(no formulas)')
  return lines.join('\n')
}

export function find(view: WorkbookView, params: Record<string, unknown>): string {
  const query = String(params['query'] ?? params['q'] ?? '')
  if (!query) throw new HeadlessOpError('command.invalid', '缺少 query')
  const result = findCells(view, query, {
    ...(typeof params['sheet'] === 'string' ? { sheet: params['sheet'] } : {}),
    ...(params['regex'] === true ? { regex: true } : {}),
    ...(params['matchCase'] === true ? { matchCase: true } : {}),
    ...(typeof params['limit'] === 'number' ? { limit: params['limit'] } : {}),
  })
  const lines = [`# find ${query}`, `hits=${result.hits.length}${result.truncated ? ' (truncated)' : ''}`]
  for (const hit of result.hits)
    lines.push(`${hit.sheet}!${hit.address}\t${hit.text.replace(/\t/g, ' ').replace(/\n/g, ' ')}`)
  return lines.join('\n')
}

export async function audit(bytes: Uint8Array, view: WorkbookView): Promise<string> {
  return auditText(bytes, view)
}

export async function project(req: PlanRequest, ctx: EngineCtx): Promise<string> {
  const bytes = req.bytes ?? (await ctx.readBytes()).bytes
  let view: WorkbookView
  try {
    view = await parseWorkbook(bytes)
  } catch (error) {
    throw coded('command.invalid', messageOf(error))
  }
  const params = (req.params ?? {}) as Record<string, unknown>
  return context(view, String(params['name'] ?? 'workbook.xlsx'))
}

export async function query(req: PlanRequest, ctx: EngineCtx): Promise<unknown> {
  const bytes = req.bytes ?? (await ctx.readBytes()).bytes
  let view: WorkbookView
  try {
    view = await parseWorkbook(bytes)
  } catch (error) {
    throw coded('command.invalid', messageOf(error))
  }
  const params = (req.params ?? {}) as Record<string, unknown>
  return contextOf(view, String(params['name'] ?? 'workbook.xlsx'))
}

export { cellText, formatAddress, parseRange, parseWorkbook, usedRange, type CellState }
