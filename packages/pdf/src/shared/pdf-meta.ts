// ─── PDF 结构探测（自研，零依赖） ────────────────────────────────────
//
// 只做「读字典」级别的工作：头部版本、trailer 字典、经典 xref 表、
// /Root → /Pages → /Count、以及 /Info 的标题/作者。
//
// 硬规则（方案 §5.2）：
//   - 不 inflate 内容流、不解析字体、不建 DOM、不 import pdfjs-dist；
//   - 任何异常都降级为「无元数据」，绝不阻断打开；
//   - 只在主进程 codec 里调用，开销是 O(头部 + trailer + 对象数)。

export interface PdfTrailerMeta {
  headerVersion?: string
  xref: 'table' | 'stream' | 'missing'
  size?: number
  encrypted: boolean
  pageCount?: number
  title?: string
  author?: string
  warning?: string
}

const HEADER_SCAN = 1024
const TAIL_SCAN = 64 * 1024
/** xref 表可能很长；超过这个长度就不再解析页数，只保留 trailer 字典。 */
const XREF_SCAN = 4 * 1024 * 1024
const OBJECT_SCAN = 16 * 1024

/** latin1 解码（PDF 语法是字节导向的，不能用 UTF-8 解释结构字符）。 */
function latin1(bytes: Uint8Array, start: number, end: number): string {
  const from = Math.max(0, Math.min(bytes.length, start))
  const to = Math.max(from, Math.min(bytes.length, end))
  let out = ''
  const step = 8192
  for (let i = from; i < to; i += step) {
    const slice = bytes.subarray(i, Math.min(to, i + step))
    out += String.fromCharCode(...slice)
  }
  return out
}

export function hasPdfHeader(bytes: Uint8Array): boolean {
  return latin1(bytes, 0, Math.min(HEADER_SCAN, bytes.length)).includes('%PDF-')
}

function headerVersion(text: string): string | undefined {
  return /%PDF-(\d+\.\d+)/.exec(text)?.[1]
}

interface XrefEntries {
  offsets: Map<number, number>
  trailer: string
}

/** 经典 xref 表：`xref` + 若干 subsection + `trailer << … >>`。 */
function parseXrefTable(text: string): XrefEntries {
  const offsets = new Map<number, number>()
  const lines = text.split(/\r\n|\r|\n/)
  let i = 1 // lines[0] === 'xref'
  while (i < lines.length) {
    const line = lines[i]!.trim()
    if (!line || line === 'trailer') break
    const header = /^(\d+)\s+(\d+)$/.exec(line)
    if (!header) {
      i++
      continue
    }
    let num = Number(header[1])
    const count = Number(header[2])
    i++
    for (let k = 0; k < count && i < lines.length; k++, i++) {
      const entry = /^(\d{1,10})\s+(\d{1,5})\s+([nf])/.exec(lines[i]!.trim())
      if (!entry) break
      if (entry[3] === 'n') offsets.set(num, Number(entry[1]))
      num++
    }
  }
  const trailerAt = text.lastIndexOf('trailer')
  let trailer = trailerAt >= 0 ? text.slice(trailerAt + 'trailer'.length) : ''
  const startAt = trailer.search(/startxref\b/)
  if (startAt >= 0) trailer = trailer.slice(0, startAt)
  return { offsets, trailer }
}

/** 从 `<<` 开始取一个平衡的字典文本（不处理字符串内的 <<，PDF 字典足够）。 */
function takeDict(text: string, from: number): string | undefined {
  const start = text.indexOf('<<', from)
  if (start < 0) return undefined
  let depth = 0
  for (let i = start; i < text.length - 1; i++) {
    const pair = text.slice(i, i + 2)
    if (pair === '<<') {
      depth++
      i++
    } else if (pair === '>>') {
      depth--
      i++
      if (depth === 0) return text.slice(start, i + 1)
    }
  }
  return text.slice(start)
}

function numOf(dict: string, key: string): number | undefined {
  const match = new RegExp(`/${key}\\b\\s*(\\d+)`).exec(dict)
  return match ? Number(match[1]) : undefined
}

function refOf(dict: string, key: string): { num: number; gen: number } | undefined {
  const match = new RegExp(`/${key}\\b\\s*(\\d+)\\s+(\\d+)\\s+R`).exec(dict)
  return match ? { num: Number(match[1]), gen: Number(match[2]) } : undefined
}

/** 读 offset 处的间接对象，返回其中的字典文本。 */
function readObjectDict(
  bytes: Uint8Array,
  offset: number,
): { dict: string; body: string } | undefined {
  if (!Number.isFinite(offset) || offset < 0 || offset >= bytes.length) return undefined
  const text = latin1(bytes, offset, offset + OBJECT_SCAN)
  const header = /^\s*(\d+)\s+(\d+)\s+obj\b/.exec(text)
  if (!header) return undefined
  const dict = takeDict(text, header[0].length)
  if (!dict) return undefined
  return { dict, body: text }
}

function decodePdfBytes(bytes: number[]): string {
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    let out = ''
    for (let i = 2; i + 1 < bytes.length; i += 2) {
      out += String.fromCharCode((bytes[i]! << 8) | bytes[i + 1]!)
    }
    return out
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(new Uint8Array(bytes))
  } catch {
    return String.fromCharCode(...bytes)
  }
}

/** `/Title (…)` 里的字面量字符串（支持转义与嵌套括号）。 */
function literalOf(dict: string, key: string): string | undefined {
  const at = new RegExp(`/${key}\\b\\s*\\(`).exec(dict)
  if (!at) return undefined
  const bytes: number[] = []
  let depth = 1
  for (let i = at.index + at[0].length; i < dict.length; i++) {
    const ch = dict[i]!
    if (ch === '\\') {
      const next = dict[++i]
      if (next === undefined) break
      if (next === 'n') bytes.push(10)
      else if (next === 'r') bytes.push(13)
      else if (next === 't') bytes.push(9)
      else if (next === 'b') bytes.push(8)
      else if (next === 'f') bytes.push(12)
      else if (next === '(' || next === ')' || next === '\\') bytes.push(next.charCodeAt(0))
      else if (/[0-7]/.test(next)) {
        let oct = next
        for (let k = 0; k < 2; k++) {
          const more = dict[i + 1]
          if (!more || !/[0-7]/.test(more)) break
          oct += more
          i++
        }
        bytes.push(Number.parseInt(oct, 8) & 0xff)
      } else if (next === '\r' || next === '\n') {
        // 续行：反斜杠 + 换行不产生字符
      } else bytes.push(next.charCodeAt(0))
      continue
    }
    if (ch === '(') depth++
    else if (ch === ')') {
      depth--
      if (depth === 0) break
    }
    bytes.push(ch.charCodeAt(0))
  }
  return bytes.length > 0 ? decodePdfBytes(bytes) : undefined
}

/** `/Title <FEFF…>` 里的十六进制字符串。 */
function hexStringOf(dict: string, key: string): string | undefined {
  const match = new RegExp(`/${key}\\b\\s*<([0-9A-Fa-f\\s]+)>`).exec(dict)
  if (!match) return undefined
  const hex = match[1]!.replace(/\s+/g, '')
  const bytes: number[] = []
  for (let i = 0; i + 1 < hex.length; i += 2) bytes.push(Number.parseInt(hex.slice(i, i + 2), 16))
  return bytes.length > 0 ? decodePdfBytes(bytes) : undefined
}

function stringOf(dict: string, key: string): string | undefined {
  return literalOf(dict, key) ?? hexStringOf(dict, key)
}

function clean(text: string | undefined): string | undefined {
  const value = text?.replace(/\u0000/g, '').trim()
  return value ? value : undefined
}

function findXrefStart(text: string): number {
  const at = text.lastIndexOf('startxref')
  if (at < 0) return Number.NaN
  const match = /\s*(\d+)/.exec(text.slice(at + 'startxref'.length))
  return match ? Number(match[1]) : Number.NaN
}

export function probePdfMeta(bytes: Uint8Array): PdfTrailerMeta {
  const header = latin1(bytes, 0, Math.min(HEADER_SCAN, bytes.length))
  const version = headerVersion(header)
  if (!version) return { xref: 'missing', encrypted: false, warning: 'not-pdf' }

  const tail = latin1(bytes, Math.max(0, bytes.length - TAIL_SCAN), bytes.length)
  const start = findXrefStart(tail)
  if (!Number.isFinite(start) || start < 0 || start >= bytes.length) {
    return { headerVersion: version, xref: 'missing', encrypted: false, warning: 'xref-missing' }
  }

  const head = latin1(bytes, start, Math.min(bytes.length, start + XREF_SCAN))
  if (!/^\s*xref\b/.test(head)) {
    // xref 流：只读字典（/Size /Root /Info /Encrypt），不 inflate 对象流。
    const obj = readObjectDict(bytes, start)
    if (!obj) {
      return {
        headerVersion: version,
        xref: 'missing',
        encrypted: false,
        warning: 'xref-unreadable',
      }
    }
    const dict = obj.dict
    const size = numOf(dict, 'Size')
    return {
      headerVersion: version,
      xref: 'stream',
      encrypted: /\/Encrypt\b/.test(dict),
      ...(size !== undefined ? { size } : {}),
    }
  }

  const { offsets, trailer } = parseXrefTable(head)
  const size = numOf(trailer, 'Size')
  const root = refOf(trailer, 'Root')
  const info = refOf(trailer, 'Info')
  const infoDict = info ? readObjectDict(bytes, offsets.get(info.num) ?? -1) : undefined

  let pageCount: number | undefined
  const rootDict = root ? readObjectDict(bytes, offsets.get(root.num) ?? -1) : undefined
  const pages = rootDict ? refOf(rootDict.dict, 'Pages') : undefined
  const pagesDict = pages ? readObjectDict(bytes, offsets.get(pages.num) ?? -1) : undefined
  if (pagesDict) {
    const count = numOf(pagesDict.dict, 'Count')
    if (count !== undefined && count >= 0 && count < 1_000_000) pageCount = count
  }

  return {
    headerVersion: version,
    xref: 'table',
    encrypted: /\/Encrypt\b/.test(trailer),
    ...(size !== undefined ? { size } : {}),
    ...(pageCount !== undefined ? { pageCount } : {}),
    ...(infoDict ? titleAuthor(infoDict.dict) : {}),
    ...(offsets.size === 0 ? { warning: 'xref-partial' } : {}),
  }
}

/**
 * xref 流场景不解析流内容，因此拿不到 /Info 对象偏移；标题/作者由
 * ≤ 8 MB 路径的 pdf-lib 补齐，大文件留空（不阻断打开）。
 */
function titleAuthor(dict: string): { title?: string; author?: string } {
  const title = clean(stringOf(dict, 'Title'))
  const author = clean(stringOf(dict, 'Author'))
  return {
    ...(title ? { title } : {}),
    ...(author ? { author } : {}),
  }
}
