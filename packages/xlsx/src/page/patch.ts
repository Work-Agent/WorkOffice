// ─── XLSX 单元格级写回 ──────────────────────────────────────────────
//
// 只改 worksheet 的 <c> 节点：保留样式、合并、图表等其它部件，
// 新值用 inlineStr / 数字 / 布尔 / <f> 表达；不重写 sharedStrings。
// 这是页面自有的最小编辑写回，不依赖 GenOffice。

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import type { SheetLike } from '../ura-types'

export interface CellEdit {
  sheetId: string
  address: string
  value: string
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function columnIndex(address: string): number {
  const match = /^([A-Z]+)/.exec(address)
  if (!match) return 0
  let value = 0
  for (const ch of match[1]!) value = value * 26 + (ch.charCodeAt(0) - 64)
  return value
}

function rowNumber(address: string): number {
  const match = /(\d+)$/.exec(address)
  return match ? Number(match[1]) : 0
}

function cellXml(address: string, raw: string, attrs: string): string {
  const style = /\bs="[^"]*"/.exec(attrs)?.[0]
  const styleAttr = style ? ` ${style}` : ''
  const value = raw
  if (value === '') return `<c r="${address}"${styleAttr}/>`
  if (value.startsWith('=')) {
    return `<c r="${address}"${styleAttr}><f>${escapeXml(value.slice(1))}</f></c>`
  }
  if (/^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(value)) {
    return `<c r="${address}"${styleAttr}><v>${value}</v></c>`
  }
  if (/^(true|false)$/i.test(value)) {
    return `<c r="${address}" t="b"${styleAttr}><v>${value.toLowerCase() === 'true' ? 1 : 0}</v></c>`
  }
  return `<c r="${address}" t="inlineStr"${styleAttr}><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`
}

function insertCell(xml: string, address: string, cell: string): string {
  const row = rowNumber(address)
  const col = columnIndex(address)
  const rowPattern = new RegExp(
    `<row\\b([^>]*?\\br="${row}"[^>]*?)(?:\\/>|>([\\s\\S]*?)<\\/row>)`,
  )
  const match = rowPattern.exec(xml)
  if (match) {
    const attrs = match[1] ?? ''
    const body = match[2]
    if (body === undefined) {
      return (
        xml.slice(0, match.index) +
        `<row${attrs}>${cell}</row>` +
        xml.slice(match.index + match[0].length)
      )
    }
    let insertAt = body.length
    const cellPattern = /<c\b[^>]*?\br="([A-Z]+\d+)"[^>]*>/g
    let cellMatch: RegExpExecArray | null
    while ((cellMatch = cellPattern.exec(body)) !== null) {
      if (columnIndex(cellMatch[1]!) > col) {
        insertAt = cellMatch.index
        break
      }
    }
    const nextBody = body.slice(0, insertAt) + cell + body.slice(insertAt)
    return (
      xml.slice(0, match.index) +
      `<row${attrs}>${nextBody}</row>` +
      xml.slice(match.index + match[0].length)
    )
  }

  // No row yet: insert before the first later row, else before </sheetData>.
  const rowStart = /<row\b/gi
  let candidate: RegExpExecArray | null
  let insertAt = -1
  while ((candidate = rowStart.exec(xml)) !== null) {
    const tagEnd = xml.indexOf('>', candidate.index)
    const tag = xml.slice(candidate.index, tagEnd)
    const r = /\br="(\d+)"/.exec(tag)
    if (r && Number(r[1]) > row) {
      insertAt = candidate.index
      break
    }
  }
  if (insertAt < 0) insertAt = xml.indexOf('</sheetData>')
  if (insertAt < 0) return xml
  return `${xml.slice(0, insertAt)}<row r="${row}">${cell}</row>${xml.slice(insertAt)}`
}

function patchSheetXml(xml: string, edits: CellEdit[]): string {
  let next = xml
  const ordered = [...edits].sort((a, b) => {
    const rowDiff = rowNumber(b.address) - rowNumber(a.address)
    return rowDiff || columnIndex(b.address) - columnIndex(a.address)
  })
  for (const edit of ordered) {
    const address = edit.address
    const pattern = new RegExp(
      `<c\\b([^>]*?)\\br="${escapeRegExp(address)}"([^>]*?)(?:\\/>|>[\\s\\S]*?<\\/c>)`,
    )
    const match = pattern.exec(next)
    const attrs = match ? `${match[1] ?? ''} ${match[2] ?? ''}` : ''
    const cell = cellXml(address, edit.value, attrs)
    if (match) {
      next = next.slice(0, match.index) + cell + next.slice(match.index + match[0].length)
    } else {
      next = insertCell(next, address, cell)
    }
  }
  return next
}

export function patchWorkbook(
  bytes: Uint8Array,
  edits: CellEdit[],
  sheets: SheetLike[],
): Uint8Array {
  if (edits.length === 0) return bytes
  const files = unzipSync(bytes)
  const bySheet = new Map<string, CellEdit[]>()
  for (const edit of edits) {
    const list = bySheet.get(edit.sheetId) ?? []
    list.push(edit)
    bySheet.set(edit.sheetId, list)
  }
  for (const [sheetId, list] of bySheet) {
    const sheet = sheets.find((item) => item.id === sheetId)
    if (!sheet?.path) continue
    const xmlBytes = files[sheet.path]
    if (!xmlBytes) continue
    files[sheet.path] = strToU8(patchSheetXml(strFromU8(xmlBytes), list))
  }
  return zipSync(files, { level: 6 })
}
