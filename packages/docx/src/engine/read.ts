// ─── engine.read：只读投影与只读工具的实现 ─────────────────────────────
//
// 契约 §6.2：read / plan 只返回可序列化投影，不返回渲染模型、不返回大字节。
// 宿主把 readOnly 工具交给执行点（无页面时读磁盘字节）；有页面时页面代理
// 用 buildDocBytes() 序列化 live 文档后调用同一个 readMany()，两边语义一致。
//
// 只解析需要的 zip 条目：document.xml（块）、comments.xml、footnotes.xml、
// endnotes.xml、styles.xml。其余条目（媒体、图表）一律不解。

import { strFromU8, unzipSync } from 'fflate'
import {
  DOCUMENT_PART,
  DocxParseError,
  parseDocxBytes,
  tierFor,
  type ParsedDocx,
} from '../shared/document'
import { projectModel, type ProjectMeta } from '../shared/project'
import { paragraphTextOf, tableSizeOf } from '../shared/xml-ops'
import type { ModelBlock } from '../shared/types'
function coded(code: string, message: string): Error {
  const error = new Error(`${code}: ${message}`) as Error & { code: string }
  error.code = code
  return error
}

export const MAX_READ_TEXT = 200_000

const PARTS = {
  comments: 'word/comments.xml',
  footnotes: 'word/footnotes.xml',
  endnotes: 'word/endnotes.xml',
  styles: 'word/styles.xml',
} as const

export function parse(bytes: Uint8Array): ParsedDocx {
  try {
    return parseDocxBytes(bytes)
  } catch (error) {
    if (error instanceof DocxParseError) throw error
    throw new DocxParseError(
      `不是有效的 DOCX 包：${error instanceof Error ? error.message : String(error)}`,
    )
  }
}

/** 读取单个 zip 条目（不存在返回 null；不整包解压）。 */
export function readPart(bytes: Uint8Array, name: string): string | null {
  let archive: Record<string, Uint8Array>
  try {
    archive = unzipSync(bytes, { filter: (file) => file.name === name })
  } catch {
    return null
  }
  const entry = archive[name]
  return entry ? strFromU8(entry) : null
}

/** 包里的条目名（诊断 / 能力报告用）。 */
export function partNames(bytes: Uint8Array): string[] {
  try {
    const all = unzipSync(bytes, { filter: (file) => file.name.endsWith('.xml') })
    return Object.keys(all)
  } catch {
    return []
  }
}

export function isDocxPackage(bytes: Uint8Array): boolean {
  try {
    const archive = unzipSync(bytes, {
      filter: (file) =>
        file.name === DOCUMENT_PART || file.name === '[Content_Types].xml',
    })
    return Boolean(archive[DOCUMENT_PART]) && Boolean(archive['[Content_Types].xml'])
  } catch {
    return false
  }
}

// ─── 文本工具 ──────────────────────────────────────────────────────────

function unescape(text: string): string {
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

/** 元素内的全部 `w:t` 文本（`<w:tab/>` → \t，`<w:br/>` → \n）。 */
function textOf(xml: string): string {
  let out = ''
  const token =
    /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>|<w:br\s*\/>|<w:cr\s*\/>/g
  for (const match of xml.matchAll(token)) {
    const raw = match[1]
    if (raw !== undefined) out += unescape(raw)
    else if (match[0].startsWith('<w:tab')) out += '\t'
    else out += '\n'
  }
  return out
}

function attr(openTag: string, name: string): string | null {
  const match = new RegExp(`\\b${name}="([^"]*)"`).exec(openTag)
  return match ? unescape(match[1]!) : null
}

function clip(text: string, max = MAX_READ_TEXT): string {
  return text.length > max ? `${text.slice(0, max)}\n…（已截断）` : text
}

function oneLine(text: string, max = 200): string {
  const flat = text.replace(/\s+/g, ' ').trim()
  return flat.length > max ? `${flat.slice(0, max)}…` : flat
}

// ─── 工具：块 ──────────────────────────────────────────────────────────

export function plainText(blocks: readonly ModelBlock[]): string {
  const lines: string[] = []
  for (const block of blocks) {
    if (block.hidden) continue
    if (block.kind === 'p') {
      if (block.text) lines.push(block.text)
      continue
    }
    if (block.kind === 'tbl') {
      const { rows, cols } = tableSizeOf(block.xml)
      lines.push(`[表格 ${rows}x${cols}]`)
      continue
    }
    lines.push(`[${block.kind}]`)
  }
  return lines.join('\n')
}

export interface ReadBlocksInput {
  startBlockIndex?: unknown
  endBlockIndex?: unknown
  offset?: unknown
}

/** `read_blocks`：块范围全文（纯文本，按 offset 分页）。 */
export function readBlocks(
  parsed: ParsedDocx,
  input: ReadBlocksInput,
): string {
  const visible = parsed.blocks.filter((block) => !block.hidden)
  const start = Number(input.startBlockIndex ?? 0)
  const end = Number(input.endBlockIndex ?? start)
  if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end < start)
    throw new Error(`startBlockIndex / endBlockIndex 非法：${String(start)}..${String(end)}`)
  const offset = Math.max(0, Number(input.offset ?? 0) || 0)
  const slice = visible.filter(
    (block) => block.docxIndex !== null && block.docxIndex >= start && block.docxIndex <= end,
  )
  if (slice.length === 0) return `（${start}..${end} 之间没有可读的块）`
  const lines = slice.map((block) => {
    const index = block.docxIndex
    if (block.kind === 'p')
      return `[${index}] ${block.heading > 0 ? `h${block.heading}` : 'p'} ${block.text}`
    if (block.kind === 'tbl') {
      const { rows, cols } = tableSizeOf(block.xml)
      return `[${index}] table ${rows}x${cols}`
    }
    return `[${index}] ${block.kind}`
  })
  const body = clip(lines.join('\n').slice(offset))
  const total = lines.join('\n').length
  return offset > 0 || body.length < total
    ? `${body}\n（offset=${offset}，本段 ${body.length} 字符；继续读用 offset=${offset + body.length}）`
    : body
}

/** `get_document_context`：块列表 + 规模统计（页面选区读取不到，见 SKILL）。 */
export function documentContext(
  parsed: ParsedDocx,
  meta: ProjectMeta = {},
): string {
  const visible = parsed.blocks.filter((block) => !block.hidden)
  const paragraphs = visible.filter((block) => block.kind === 'p')
  const tables = visible.filter((block) => block.kind === 'tbl')
  const headings = paragraphs.filter((block) => block.heading > 0)
  const chars = paragraphs.reduce((sum, block) => sum + block.text.length, 0)
  const projection = projectModel(parsed.blocks, meta)
  return [
    projection,
    '',
    `统计：段落 ${paragraphs.length}，表格 ${tables.length}，标题 ${headings.length}，正文字符 ${chars}`,
    '说明：无页面路径没有「当前选区」；用块索引（[n]）定位。',
  ].join('\n')
}

/** `docx_read` 的全部 what 分发（执行点与页面同源；页面传入 live 序列化字节）。 */
export const READ_WHAT = new Set([
  'context',
  'blocks',
  'comments',
  'notes',
  'revisions',
  'styles',
  'audit',
])

export function readMany(bytes: Uint8Array, params: Record<string, unknown>): string {
  const what = String(params['what'] ?? 'context')
  if (!READ_WHAT.has(what))
    throw coded('command.invalid', `what 非法：${what}（可用：${[...READ_WHAT].join(' / ')}）`)
  if (what === 'comments') return readComments(bytes)
  if (what === 'notes') return readNotes(bytes)
  if (what === 'styles') return readStyles(bytes)
  if (what === 'audit') return audit(parse(bytes))
  if (what === 'revisions') return readRevisions(bytes, parse(bytes))
  if (what === 'blocks') return readBlocks(parse(bytes), params)
  return documentContext(parse(bytes), {
    ...(typeof params['name'] === 'string' ? { name: params['name'] } : {}),
    ...(typeof params['version'] === 'string' ? { version: params['version'] } : {}),
    ...(typeof params['epoch'] === 'number' ? { epoch: params['epoch'] } : {}),
  })
}

// ─── 工具：批注 / 脚注 / 修订 / 样式 ───────────────────────────────────

export function readComments(bytes: Uint8Array): string {
  const xml = readPart(bytes, PARTS.comments)
  if (!xml) return '（文档没有批注）'
  const out: string[] = []
  const pattern = /<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g
  for (const match of xml.matchAll(pattern)) {
    const open = match[1] ?? ''
    const id = attr(open, 'w:id') ?? '?'
    const author = attr(open, 'w:author') ?? '（未知作者）'
    const date = attr(open, 'w:date')
    const text = oneLine(textOf(match[2] ?? ''), 500)
    out.push(`#${id} ${author}${date ? ` · ${date}` : ''}\n  ${text || '（空）'}`)
  }
  return out.length > 0 ? clip(out.join('\n')) : '（文档没有批注）'
}

export function readNotes(bytes: Uint8Array): string {
  const sections: string[] = []
  for (const [label, part] of [
    ['脚注', PARTS.footnotes],
    ['尾注', PARTS.endnotes],
  ] as const) {
    const xml = readPart(bytes, part)
    if (!xml) continue
    const out: string[] = []
    const pattern =
      /<w:(footnote|endnote)\b([^>]*)>([\s\S]*?)<\/w:\1>/g
    for (const match of xml.matchAll(pattern)) {
      const open = match[2] ?? ''
      const type = attr(open, 'w:type')
      if (type === 'separator' || type === 'continuationSeparator') continue
      const id = attr(open, 'w:id') ?? '?'
      const text = oneLine(textOf(match[3] ?? ''), 500)
      if (!text) continue
      out.push(`#${id} ${text}`)
    }
    if (out.length > 0) sections.push(`${label}：\n${out.join('\n')}`)
  }
  return sections.length > 0 ? clip(sections.join('\n\n')) : '（文档没有脚注 / 尾注）'
}

export function readRevisions(bytes: Uint8Array, parsed: ParsedDocx): string {
  const xml = parsed.documentXml
  const inserted = [...xml.matchAll(/<w:ins\b([^>]*)>([\s\S]*?)<\/w:ins>/g)]
  const deleted = [...xml.matchAll(/<w:del\b([^>]*)>([\s\S]*?)<\/w:del>/g)]
  if (inserted.length === 0 && deleted.length === 0)
    return '（文档没有未处理的修订）'
  const lines: string[] = []
  for (const [index, match] of inserted.slice(0, 200).entries()) {
    const author = attr(match[1] ?? '', 'w:author') ?? '（未知作者）'
    lines.push(`+${index + 1} ${author}：${oneLine(textOf(match[2] ?? ''), 300)}`)
  }
  for (const [index, match] of deleted.slice(0, 200).entries()) {
    const author = attr(match[1] ?? '', 'w:author') ?? '（未知作者）'
    lines.push(`-${index + 1} ${author}：${oneLine(textOf(match[2] ?? ''), 300)}`)
  }
  return clip(
    [
      `修订：插入 ${inserted.length}，删除 ${deleted.length}`,
      ...lines,
      '用 accept_changes / reject_changes 处理（需要打开页面）。',
    ].join('\n'),
  )
}

export function readStyles(bytes: Uint8Array): string {
  const xml = readPart(bytes, PARTS.styles)
  if (!xml) return '（文档没有 styles.xml）'
  const out: string[] = []
  const pattern = /<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g
  for (const match of xml.matchAll(pattern)) {
    const open = match[1] ?? ''
    const id = attr(open, 'w:styleId') ?? '?'
    const type = attr(open, 'w:type') ?? 'paragraph'
    const inner = match[2] ?? ''
    const name = /<w:name\b[^>]*w:val="([^"]*)"/.exec(inner)?.[1] ?? id
    out.push(`${id}${name !== id ? ` (${unescape(name)})` : ''} · ${type}`)
  }
  return out.length > 0 ? clip(out.join('\n')) : '（文档没有样式定义）'
}

/** 结构诊断（`docx_check`）。 */
export function audit(parsed: ParsedDocx): string {
  const visible = parsed.blocks.filter((block) => !block.hidden)
  const paragraphs = visible.filter((block) => block.kind === 'p')
  const tables = visible.filter((block) => block.kind === 'tbl')
  const others = visible.filter((block) => block.kind !== 'p' && block.kind !== 'tbl')
  const empty = paragraphs.filter(
    (block) => paragraphTextOf(block.xml).trim().length === 0,
  ).length
  const headings = paragraphs
    .filter((block) => block.heading > 0)
    .slice(0, 50)
    .map((block) => `  h${block.heading} [${block.docxIndex}] ${oneLine(block.text, 80)}`)
  return [
    `blocks=${visible.length} 段落=${paragraphs.length} 表格=${tables.length} 其他=${others.length}`,
    `空段落=${empty} 正文字符=${parsed.textLength} tier=${tierFor(parsed.weight)}`,
    headings.length > 0 ? `标题：\n${headings.join('\n')}` : '标题：（无）',
  ].join('\n')
}
