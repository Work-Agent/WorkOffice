// ─── XLSX codec（WorkStudio 自有） ───────────────────────────────────
//
// codec 不依赖 GenOffice：用 fflate + fast-xml-parser 读 OOXML 包，
// 生成 sheet/cell 快照供 projection 与 agent 文本使用。页面只消费快照。

import { strFromU8, unzipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import { registerCodec, registerConsumer } from './sdk'
import type {
  AgentProjection,
  CellLike,
  CellStyle,
  Codec,
  Consumer,
  DecodedDoc,
  EditorProjection,
  Patch,
  RawContent,
  SheetLike,
} from './ura-types'

const DOC_TYPE = 'ooxml-xlsx@1' as const

const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
}

interface XlsxDoc {
  bytes: string
  name: string
  sheets: SheetLike[]
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  isArray: (name) =>
    ['sheet', 'Relationship', 'si', 'row', 'c', 'col', 'mergeCell'].includes(name),
})

function attr(node: unknown, name: string): string {
  if (!node || typeof node !== 'object') return ''
  const record = node as Record<string, unknown>
  const direct = record[`@_${name}`]
  if (typeof direct === 'string') return direct
  for (const [key, value] of Object.entries(record)) {
    if (key === `@_${name}` || key.endsWith(`:${name}`)) {
      if (typeof value === 'string') return value
    }
  }
  return ''
}

function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return []
  return Array.isArray(value) ? value : [value]
}

function decodeXml(text: string): string {
  return text
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) =>
      String.fromCodePoint(Number.parseInt(hex, 16)),
    )
    .replace(/&#(\d+);/g, (_, dec: string) =>
      String.fromCodePoint(Number.parseInt(dec, 10)),
    )
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

function textOfNode(node: unknown): string {
  if (typeof node === 'string') return decodeXml(node)
  if (typeof node === 'number' || typeof node === 'boolean') return String(node)
  if (!node || typeof node !== 'object') return ''
  const record = node as Record<string, unknown>
  if (record['#text'] !== undefined) return decodeXml(String(record['#text']))
  if (record.t !== undefined) return textOfNode(record.t)
  if (record.r !== undefined) {
    return asArray(record.r)
      .map((run) => textOfNode((run as Record<string, unknown>)?.t))
      .join('')
  }
  return ''
}

function base64Of(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString('base64')
}

function nameOf(ref: string): string {
  try {
    const url = new URL(ref)
    const name = url.pathname.slice(url.pathname.lastIndexOf('/') + 1)
    return decodeURIComponent(name) || 'workbook.xlsx'
  } catch {
    const name = ref.slice(ref.lastIndexOf('/') + 1)
    return name || 'workbook.xlsx'
  }
}

function argb(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined
  const hex = value.replace(/^#/, '')
  if (hex.length === 8) return `#${hex.slice(2)}`
  if (hex.length === 6) return `#${hex}`
  return undefined
}

function parseStyles(xml: string | null): CellStyle[] {
  if (!xml) return []
  const parsed = parser.parse(xml) as {
    styleSheet?: {
      numFmts?: { numFmt?: unknown[] }
      fonts?: { font?: unknown[] }
      fills?: { fill?: unknown[] }
      cellXfs?: { xf?: unknown[] }
    }
  }
  const sheet = parsed.styleSheet
  if (!sheet) return []
  const custom = new Map<number, string>()
  for (const item of asArray(sheet.numFmts?.numFmt)) {
    const record = item as Record<string, unknown>
    const id = Number(attr(record, 'numFmtId'))
    const code = attr(record, 'formatCode')
    if (Number.isFinite(id) && code) custom.set(id, decodeXml(code))
  }
  const fonts = asArray(sheet.fonts?.font).map((item) => {
    const record = item as Record<string, unknown>
    const colorNode = record.color as Record<string, unknown> | undefined
    return {
      bold: record.b !== undefined,
      italic: record.i !== undefined,
      underline: record.u !== undefined,
      fontSize: Number(attr(record.sz, 'val')) || undefined,
      fontFamily: attr(record.name, 'val') || undefined,
      fontColor: argb(colorNode?.['@_rgb']) ?? argb(colorNode?.['@_val']),
    } satisfies CellStyle
  })
  const fills = asArray(sheet.fills?.fill).map((item) => {
    const record = item as Record<string, unknown>
    const pattern = record.patternFill as Record<string, unknown> | undefined
    const fg = pattern?.fgColor as Record<string, unknown> | undefined
    const bg = pattern?.bgColor as Record<string, unknown> | undefined
    return argb(fg?.['@_rgb']) ?? argb(fg?.['@_val']) ?? argb(bg?.['@_rgb']) ?? undefined
  })
  return asArray(sheet.cellXfs?.xf).map((item) => {
    const record = item as Record<string, unknown>
    const font = fonts[Number(attr(record, 'fontId'))] ?? {}
    const fill = fills[Number(attr(record, 'fillId'))]
    const numFmtId = Number(attr(record, 'numFmtId'))
    const numberFormat = custom.get(numFmtId) ?? BUILTIN_FORMATS[numFmtId]
    const alignment = record.alignment as Record<string, unknown> | undefined
    const horizontal = attr(alignment, 'horizontal')
    return {
      ...font,
      ...(fill ? { fillColor: fill } : {}),
      ...(numberFormat && numberFormat !== 'General' ? { numberFormat } : {}),
      ...(horizontal === 'center' || horizontal === 'right' || horizontal === 'left'
        ? { align: horizontal }
        : {}),
      ...(attr(alignment, 'wrapText') === '1' ? { wrap: true } : {}),
    } satisfies CellStyle
  })
}

function parseWorkbook(bytes: Uint8Array): SheetLike[] {
  let files: Record<string, Uint8Array>
  try {
    files = unzipSync(bytes)
  } catch {
    throw new Error('不是有效的 XLSX 包（ZIP 解压失败）')
  }
  const read = (path: string): string | null => {
    const file = files[path] ?? files[path.replace(/^xl\//, '')]
    return file ? strFromU8(file) : null
  }

  const workbookXml = read('xl/workbook.xml')
  if (!workbookXml) throw new Error('XLSX 缺少 xl/workbook.xml')
  const relsXml = read('xl/_rels/workbook.xml.rels')
  const sharedXml = read('xl/sharedStrings.xml')
  const styles = parseStyles(read('xl/styles.xml'))
  const workbook = parser.parse(workbookXml) as {
    workbook?: { sheets?: { sheet?: Array<Record<string, unknown>> } }
  }
  const rels = relsXml
    ? (parser.parse(relsXml) as {
        Relationships?: { Relationship?: Array<Record<string, unknown>> }
      })
    : {}

  const relTargets = new Map<string, string>()
  for (const rel of asArray(rels.Relationships?.Relationship)) {
    const id = attr(rel, 'Id')
    const target = attr(rel, 'Target')
    if (!id || !target) continue
    const normalized = target.startsWith('/')
      ? target.slice(1)
      : target.startsWith('xl/')
        ? target
        : `xl/${target}`
    relTargets.set(id, normalized)
  }

  const sharedStrings: string[] = []
  if (sharedXml) {
    const sst = parser.parse(sharedXml) as { sst?: { si?: unknown[] } }
    for (const si of asArray(sst.sst?.si)) sharedStrings.push(textOfNode(si))
  }

  const sheets: SheetLike[] = []
  for (const sheet of asArray(workbook.workbook?.sheets?.sheet)) {
    const name = attr(sheet, 'name')
    const sheetId = attr(sheet, 'sheetId') || String(sheets.length + 1)
    const relId = attr(sheet, 'r:id') || attr(sheet, 'id')
    const path = relTargets.get(relId)
    if (!name || !path) continue
    const sheetXml = read(path)
    if (!sheetXml) continue
    const parsed = parser.parse(sheetXml) as {
      worksheet?: {
        sheetData?: { row?: unknown[] }
        cols?: { col?: unknown[] }
        mergeCells?: { mergeCell?: unknown[] }
      }
    }
    const cells: Record<string, CellLike> = {}
    const colWidths: Record<string, number> = {}
    const rowHeights: Record<string, number> = {}
    const merges: string[] = []
    for (const row of asArray(parsed.worksheet?.sheetData?.row)) {
      const rowRecord = row as Record<string, unknown>
      const rowNumber = Number(attr(rowRecord, 'r')) || 0
      const heightPt = Number(attr(rowRecord, 'ht'))
      if (rowNumber && Number.isFinite(heightPt) && heightPt > 0) {
        rowHeights[String(rowNumber)] = Math.round((heightPt * 96) / 72)
      }
      let columnIndex = 0
      for (const cellNode of asArray(rowRecord.c)) {
        const cell = cellNode as Record<string, unknown>
        const address = attr(cell, 'r') || `${columnLabel(columnIndex)}${rowNumber || 1}`
        columnIndex = parseColumnIndex(address) || columnIndex + 1
        const type = attr(cell, 't')
        const formula = cell.f !== undefined ? decodeXml(String(cell.f)) : undefined
        const rawValue = cell.v !== undefined ? decodeXml(String(cell.v)) : undefined
        let value: string | number | boolean | null = null
        if (type === 's' && rawValue !== undefined) {
          value = sharedStrings[Number(rawValue)] ?? ''
        } else if (type === 'inlineStr') {
          value = textOfNode(cell.is)
        } else if (type === 'b') {
          value = rawValue === '1'
        } else if (type === 'str' || type === 'e') {
          value = rawValue ?? null
        } else if (rawValue !== undefined) {
          const numeric = Number(rawValue)
          value = Number.isFinite(numeric) ? numeric : rawValue
        }
        if (value === null && formula === undefined) continue
        const styleIndex = Number(attr(cell, 's'))
        const style = Number.isFinite(styleIndex) ? styles[styleIndex] : undefined
        cells[address] = {
          value,
          ...(formula !== undefined ? { formula: `=${formula}` } : {}),
          ...(style && Object.keys(style).length ? { style } : {}),
        }
      }
    }
    for (const col of asArray(parsed.worksheet?.cols?.col)) {
      const colRecord = col as Record<string, unknown>
      const min = Number(attr(colRecord, 'min'))
      const max = Number(attr(colRecord, 'max')) || min
      const width = Number(attr(colRecord, 'width'))
      if (!min || !Number.isFinite(width) || width <= 0) continue
      const px = Math.max(32, Math.min(480, Math.round(width * 7 + 5)))
      for (let index = min; index <= max && index < min + 200; index++) {
        colWidths[columnLabel(index - 1)] = px
      }
    }
    for (const merge of asArray(parsed.worksheet?.mergeCells?.mergeCell)) {
      const ref = attr(merge, 'ref')
      if (ref) merges.push(ref)
    }
    sheets.push({
      id: `sheet-${sheetId}`,
      name,
      path,
      cells,
      ...(merges.length ? { merges } : {}),
      ...(Object.keys(colWidths).length ? { colWidths } : {}),
      ...(Object.keys(rowHeights).length ? { rowHeights } : {}),
    })
  }
  if (sheets.length === 0) throw new Error('XLSX 没有可读工作表')
  return sheets
}

function columnLabel(index: number): string {
  let label = ''
  let value = index + 1
  while (value > 0) {
    const rem = (value - 1) % 26
    label = String.fromCharCode(65 + rem) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

function parseColumnIndex(address: string): number {
  const match = /^([A-Z]+)/.exec(address)
  if (!match) return 0
  let value = 0
  for (const ch of match[1]!) value = value * 26 + (ch.charCodeAt(0) - 64)
  return value
}

function cellText(cell: CellLike): string {
  if (cell.formula) return cell.formula
  if (cell.value === null || cell.value === undefined) return ''
  return typeof cell.value === 'boolean'
    ? cell.value
      ? 'TRUE'
      : 'FALSE'
    : String(cell.value)
}

function workbookText(sheets: SheetLike[]): string {
  const parts: string[] = []
  for (const sheet of sheets) {
    const rows = new Map<number, Array<{ col: number; text: string }>>()
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const match = /^([A-Z]+)(\d+)$/.exec(address)
      if (!match) continue
      const row = Number(match[2])
      const text = cellText(cell)
      if (!text) continue
      const list = rows.get(row) ?? []
      list.push({ col: parseColumnIndex(address), text })
      rows.set(row, list)
    }
    const body = [...rows.entries()]
      .sort(([a], [b]) => a - b)
      .map(([, cells]) =>
        cells
          .sort((a, b) => a.col - b.col)
          .map((cell) => cell.text)
          .join('\t'),
      )
      .join('\n')
    parts.push(body ? `## ${sheet.name}\n${body}` : `## ${sheet.name}`)
    if (parts.join('\n').length > 20_000) break
  }
  return parts.join('\n\n').slice(0, 20_000)
}

const codec: Codec = {
  hints: ['ooxml-xlsx'],
  extensions: ['xlsx', 'xlsm', 'xltx'],
  docType: DOC_TYPE,
  patchType: 'bytes',

  decode(raw: RawContent, ref: string) {
    const sheets = parseWorkbook(raw.bytes)
    return {
      docType: DOC_TYPE,
      doc: { bytes: base64Of(raw.bytes), name: nameOf(ref), sheets } satisfies XlsxDoc,
      projection: { bytes: base64Of(raw.bytes), name: nameOf(ref), sheets },
    }
  },

  locate(_decoded, anchor) {
    return { start: anchor.start, end: anchor.end }
  },

  encode(_decoded, change): Patch {
    if (change.type === 'replace-all' && change.content instanceof Uint8Array) {
      return { type: 'replace-all', bytes: change.content }
    }
    throw new Error('XLSX 文档只接受二进制写回')
  },
}

const agentConsumer: Consumer = {
  docType: DOC_TYPE,
  type: 'agent',
  project(decoded): AgentProjection {
    return {
      type: 'text',
      content:
        workbookText((decoded.doc as XlsxDoc).sheets) || '(空 XLSX 工作簿)',
    }
  },
}

const editorConsumer: Consumer = {
  docType: DOC_TYPE,
  type: 'editor',
  surface: 'xlsx-viewer',
  project(decoded, raw, ref): EditorProjection {
    const doc = decoded.doc as XlsxDoc
    return {
      type: 'page',
      surface: 'xlsx-viewer',
      page: '',
      data: {
        bytes: doc.bytes,
        name: doc.name,
        ref,
        version: raw.version,
        sheets: doc.sheets,
      },
    }
  },
}

registerCodec(codec)
registerConsumer(agentConsumer)
registerConsumer(editorConsumer)
