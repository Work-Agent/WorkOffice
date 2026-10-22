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
  Codec,
  Consumer,
  DecodedDoc,
  EditorProjection,
  Patch,
  RawContent,
  SheetLike,
} from './ura-types'

const DOC_TYPE = 'ooxml-xlsx@1' as const

interface XlsxDoc {
  name: string
  sheets: SheetLike[]
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  isArray: (name) => ['sheet', 'Relationship', 'si', 'row', 'c'].includes(name),
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
      worksheet?: { sheetData?: { row?: unknown[] } }
    }
    const cells: Record<string, CellLike> = {}
    for (const row of asArray(parsed.worksheet?.sheetData?.row)) {
      const rowRecord = row as Record<string, unknown>
      const rowNumber = Number(attr(rowRecord, 'r')) || 0
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
        cells[address] = {
          value,
          ...(formula !== undefined ? { formula: `=${formula}` } : {}),
        }
      }
    }
    sheets.push({ id: `sheet-${sheetId}`, name, cells })
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
      doc: { name: nameOf(ref), sheets } satisfies XlsxDoc,
      projection: { name: nameOf(ref), sheets },
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
