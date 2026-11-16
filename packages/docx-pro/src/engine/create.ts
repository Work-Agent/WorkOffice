// ─── 新建计划：Markdown → word/document.xml ─────────────────────────────
//
// 新建文档不经过页面：插件自己把 Markdown 转成 OOXML 段落 / 表格，exec
// 再把这段 XML 换进 seed.docx 的 word/document.xml。样式、编号定义来自
// seed（Heading1-6 / ListParagraph / bullet numId=1 / ordered numId=2），
// 所以这里只生成内容，不重新定义样式。

import { marked } from 'marked'

interface Run {
  text: string
  bold?: boolean
  italic?: boolean
  strike?: boolean
  code?: boolean
}

interface ParaOptions {
  styleId?: string
  level?: number
  list?: { kind: 'bullet' | 'ordered'; level: number }
  indentLeft?: number
  align?: 'left' | 'center' | 'right'
}

const XML_DECL =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const DOC_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'

const SECT_PR =
  '<w:sectPr>' +
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"' +
  ' w:header="708" w:footer="708" w:gutter="0"/>' +
  '</w:sectPr>'

const BULLET_NUM_ID = '1'
const ORDERED_NUM_ID = '2'

export function markdownToDocumentXml(markdown: string): string {
  const tokens = marked.lexer(markdown, { gfm: true })
  const body: string[] = []

  for (const token of tokens as unknown as MarkdownToken[]) {
    body.push(...blockXml(token))
  }

  if (body.length === 0) body.push('<w:p/>')

  return (
    XML_DECL +
    `<w:document ${DOC_NS}><w:body>${body.join('')}${SECT_PR}</w:body></w:document>`
  )
}

// ─── 块 ────────────────────────────────────────────────────────────────

interface MarkdownToken {
  type?: string
  text?: string
  tokens?: MarkdownToken[]
  depth?: number
  ordered?: boolean
  items?: MarkdownToken[]
  header?: TableCell[]
  rows?: TableCell[][]
  lang?: string
  [key: string]: unknown
}

interface TableCell {
  text?: string
  tokens?: MarkdownToken[]
  align?: 'left' | 'center' | 'right' | null
}

function blockXml(token: MarkdownToken): string[] {
  switch (token.type) {
    case 'heading':
      return [
        paragraphXml(
          inlineRuns(token.tokens ?? []),
          { level: Math.min(6, Math.max(1, token.depth ?? 1)) },
        ),
      ]
    case 'paragraph':
      return [paragraphXml(inlineRuns(token.tokens ?? []))]
    case 'text':
      return [paragraphXml(inlineRuns(token.tokens ?? [{ type: 'text', text: token.text ?? '' }]))]
    case 'space':
      return []
    case 'list':
      return listXml(token, 0)
    case 'table':
      return [tableXml(token)]
    case 'code':
      return [codeBlockXml(token.text ?? '')]
    case 'blockquote':
      return (token.tokens ?? []).flatMap((child) => blockXml(child)).map((xml) =>
        addIndent(xml, 720),
      )
    case 'hr':
      return []
    case 'html':
      return textBlocks(token.text ?? '')
    default:
      return token.text ? textBlocks(token.text) : []
  }
}

function textBlocks(text: string): string[] {
  return text
    .replace(/\r\n?/g, '\n')
    .split(/\n{2,}/)
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => paragraphXml([{ text: part }]))
}

function listXml(token: MarkdownToken, level: number): string[] {
  const out: string[] = []
  const kind = token.ordered ? 'ordered' : 'bullet'

  for (const item of token.items ?? []) {
    const tokens = item.tokens ?? []
    let wroteParagraph = false
    for (const child of tokens) {
      if (child.type === 'list') {
        out.push(...listXml(child, level + 1))
        continue
      }
      if (child.type === 'code') {
        out.push(codeBlockXml(child.text ?? '', level))
        wroteParagraph = true
        continue
      }
      if (child.type === 'text' || child.type === 'paragraph') {
        const runs = inlineRuns(
          child.tokens ?? [{ type: 'text', text: child.text ?? '' }],
        )
        if (runs.length === 0) continue
        out.push(paragraphXml(runs, { list: { kind, level } }))
        wroteParagraph = true
      }
    }
    if (!wroteParagraph) {
      const runs = inlineRuns(
        item.tokens ?? [{ type: 'text', text: item.text ?? '' }],
      )
      if (runs.length > 0)
        out.push(paragraphXml(runs, { list: { kind, level } }))
    }
  }

  return out
}

function tableXml(token: MarkdownToken): string {
  const rows: TableCell[][] = []
  if (Array.isArray(token.header) && token.header.length > 0)
    rows.push(token.header)
  for (const row of token.rows ?? []) rows.push(row)

  const columns = Math.max(1, ...rows.map((row) => row.length))
  const width = Math.max(1, Math.floor(9360 / columns))

  const grid = `<w:tblGrid>${Array.from(
    { length: columns },
    () => `<w:gridCol w:w="${width}"/>`,
  ).join('')}</w:tblGrid>`

  const border = (name: string) =>
    `<w:${name} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`

  const trs = rows
    .map((row, rowIndex) => {
      const cells: string[] = []
      for (let i = 0; i < columns; i += 1) {
        const cell = row[i]
        const runs = cell
          ? inlineRuns(
              cell.tokens ?? [{ type: 'text', text: cell.text ?? '' }],
            )
          : []
        const isHeader = rowIndex === 0
        const merged = isHeader
          ? runs.map((run) => ({ ...run, bold: true }))
          : runs
        const shade = isHeader
          ? '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
          : ''
        cells.push(
          '<w:tc><w:tcPr>' +
            `<w:tcW w:w="${width}" w:type="dxa"/>` +
            shade +
            '</w:tcPr>' +
            paragraphXml(merged, {
              ...(cell?.align ? { align: cell.align } : {}),
            }) +
            '</w:tc>',
        )
      }
      return `<w:tr>${cells.join('')}</w:tr>`
    })
    .join('')

  return (
    '<w:tbl><w:tblPr>' +
    '<w:tblW w:w="0" w:type="auto"/>' +
    '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
      .map(border)
      .join('') +
    '</w:tblBorders>' +
    '</w:tblPr>' +
    grid +
    trs +
    '</w:tbl>'
  )
}

function codeBlockXml(text: string, level?: number): string {
  const runs: Run[] = []
  const lines = text.replace(/\r\n?/g, '\n').split('\n')
  lines.forEach((line, index) => {
    if (index > 0) runs.push({ text: '\n', code: true })
    if (line) runs.push({ text: line, code: true })
  })
  if (runs.length === 0) runs.push({ text: '', code: true })
  return paragraphXml(runs, {
    ...(level !== undefined ? { indentLeft: 720 + level * 360 } : {}),
  })
}

function addIndent(xml: string, left: number): string {
  if (!xml.startsWith('<w:p')) return xml
  const pPr = `<w:pPr><w:ind w:left="${left}"/></w:pPr>`
  return xml.replace(/^(<w:p[^>]*>)(<w:pPr>[\s\S]*?<\/w:pPr>)?/, (_m, open, existing) => {
    if (existing) {
      const merged = existing.replace(
        /<\/w:pPr>$/,
        `<w:ind w:left="${left}"/></w:pPr>`,
      )
      return open + merged
    }
    return open + pPr
  })
}

// ─── 段落 / run ────────────────────────────────────────────────────────

function paragraphXml(runs: Run[], options: ParaOptions = {}): string {
  const pPr: string[] = []
  if (options.styleId)
    pPr.push(`<w:pStyle w:val="${escapeAttr(options.styleId)}"/>`)
  if (options.level !== undefined)
    pPr.push(`<w:pStyle w:val="Heading${options.level}"/>`)
  if (options.list) {
    pPr.push('<w:pStyle w:val="ListParagraph"/>')
    pPr.push(
      '<w:numPr>' +
        `<w:ilvl w:val="${Math.min(8, Math.max(0, options.list.level))}"/>` +
        `<w:numId w:val="${options.list.kind === 'ordered' ? ORDERED_NUM_ID : BULLET_NUM_ID}"/>` +
        '</w:numPr>',
    )
  }
  if (options.indentLeft !== undefined)
    pPr.push(`<w:ind w:left="${options.indentLeft}"/>`)
  if (options.align) pPr.push(`<w:jc w:val="${options.align}"/>`)

  const runsXml = runs.map(runXml).join('')
  return `<w:p>${pPr.length > 0 ? `<w:pPr>${pPr.join('')}</w:pPr>` : ''}${runsXml}</w:p>`
}

function runXml(run: Run): string {
  if (!run.text) return ''
  const rPr: string[] = []
  if (run.code)
    rPr.push(
      '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>',
    )
  if (run.bold) rPr.push('<w:b/><w:bCs/>')
  if (run.italic) rPr.push('<w:i/><w:iCs/>')
  if (run.strike) rPr.push('<w:strike/>')

  const segments: string[] = []
  let buffer = ''
  const flush = (): void => {
    if (buffer) {
      segments.push(`<w:t xml:space="preserve">${escapeText(buffer)}</w:t>`)
      buffer = ''
    }
  }
  for (const char of run.text) {
    if (char === '\n') {
      flush()
      segments.push('<w:br/>')
    } else if (char === '\t') {
      flush()
      segments.push('<w:tab/>')
    } else {
      buffer += char
    }
  }
  flush()

  return `<w:r>${rPr.length > 0 ? `<w:rPr>${rPr.join('')}</w:rPr>` : ''}${segments.join('')}</w:r>`
}

// ─── 行内 token ────────────────────────────────────────────────────────

function inlineRuns(
  tokens: MarkdownToken[],
  style: Pick<Run, 'bold' | 'italic' | 'strike' | 'code'> = {},
): Run[] {
  const out: Run[] = []
  for (const token of tokens) {
    switch (token.type) {
      case 'text':
      case 'escape':
        if (token.tokens && token.tokens.length > 0)
          out.push(...inlineRuns(token.tokens, style))
        else if (token.text) out.push({ text: token.text, ...style })
        break
      case 'strong':
        out.push(...inlineRuns(token.tokens ?? [], { ...style, bold: true }))
        break
      case 'em':
        out.push(...inlineRuns(token.tokens ?? [], { ...style, italic: true }))
        break
      case 'del':
        out.push(...inlineRuns(token.tokens ?? [], { ...style, strike: true }))
        break
      case 'codespan':
        out.push({ text: token.text ?? '', ...style, code: true })
        break
      case 'br':
        out.push({ text: '\n', ...style })
        break
      case 'link':
        out.push(...inlineRuns(token.tokens ?? [], style))
        break
      case 'image':
        if (token.text) out.push({ text: token.text, ...style })
        break
      case 'html': {
        const text = stripHtml(token.text ?? '')
        if (text) out.push({ text, ...style })
        break
      }
      default:
        if (token.tokens) out.push(...inlineRuns(token.tokens, style))
        else if (token.text) out.push({ text: token.text, ...style })
        break
    }
  }
  return out.filter((run) => run.text.length > 0)
}

function stripHtml(html: string): string {
  return html
    .replace(/<[^>]*>/g, '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
}

// ─── XML 转义 ──────────────────────────────────────────────────────────

function escapeText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

function escapeAttr(text: string): string {
  return escapeText(text).replace(/"/g, '&quot;')
}
