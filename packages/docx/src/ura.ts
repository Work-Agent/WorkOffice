// ─── DOCX codec（WorkStudio 自有） ───────────────────────────────────
//
// codec 不依赖 GenOffice：只做 OOXML 包的最小解析——取出 word/document.xml，
// 抽取段落文本与标题大纲，供 agent 投影使用。页面层仍可使用 GenOffice 的
// 编辑器来渲染/编辑，但那是页面的事，不进入 codec。

import { strFromU8, unzipSync } from 'fflate'
import { registerCodec, registerConsumer } from './sdk'
import type {
  AgentProjection,
  Codec,
  Consumer,
  DecodedDoc,
  EditorProjection,
  Patch,
  RawContent,
} from './ura-types'

const DOC_TYPE = 'ooxml-docx@1' as const

interface DocxDoc {
  /** 原始 OOXML 字节（base64；页面数据面只走可结构化克隆的字符串）。 */
  bytes: string
  text: string
  name: string
}

function base64Of(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString('base64')
}

function nameOf(ref: string): string {
  try {
    const url = new URL(ref)
    const name = url.pathname.slice(url.pathname.lastIndexOf('/') + 1)
    return decodeURIComponent(name) || 'document.docx'
  } catch {
    const name = ref.slice(ref.lastIndexOf('/') + 1)
    return name || 'document.docx'
  }
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

function headingLevel(paragraphXml: string): number {
  const style = /<w:pStyle\b[^>]*\bw:val="([^"]*)"/i.exec(paragraphXml)
  if (style) {
    const match = /heading\s*([1-9])/i.exec(style[1]!)
    if (match) return Number(match[1])
  }
  const outline = /<w:outlineLvl\b[^>]*\bw:val="([0-8])"/i.exec(paragraphXml)
  if (outline) return Number(outline[1]) + 1
  return 0
}

function paragraphText(paragraphXml: string): string {
  let text = ''
  const token =
    /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>|<w:br\s*\/>|<w:cr\s*\/>/g
  for (const match of paragraphXml.matchAll(token)) {
    const raw = match[1]
    if (raw !== undefined) text += decodeXml(raw)
    else if (match[0].startsWith('<w:tab')) text += '\t'
    else text += '\n'
  }
  return text
}

function parseDocumentXml(xml: string): {
  text: string
  outline: Array<{ level: number; title: string }>
} {
  const paragraphs: string[] = []
  const outline: Array<{ level: number; title: string }> = []
  const pattern = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g
  for (const match of xml.matchAll(pattern)) {
    const body = match[1] ?? ''
    const text = paragraphText(body).replace(/[ \t]+$/gm, '').trimEnd()
    if (!text.trim()) continue
    paragraphs.push(text)
    const level = headingLevel(body)
    if (level > 0) outline.push({ level, title: text.replace(/\s+/g, ' ').trim() })
  }
  return { text: paragraphs.join('\n'), outline }
}

const codec: Codec = {
  hints: ['ooxml-docx'],
  extensions: ['docx', 'docm', 'dotx'],
  docType: DOC_TYPE,
  patchType: 'bytes',

  decode(raw: RawContent, ref: string) {
    let archive: Record<string, Uint8Array>
    try {
      archive = unzipSync(raw.bytes, {
        filter: (file) => file.name === 'word/document.xml',
      })
    } catch {
      throw new Error('不是有效的 DOCX 包（ZIP 解压失败）')
    }
    const documentXml = archive['word/document.xml']
    if (!documentXml) throw new Error('DOCX 缺少 word/document.xml')

    const { text, outline } = parseDocumentXml(strFromU8(documentXml))
    const bytes = base64Of(raw.bytes)
    return {
      docType: DOC_TYPE,
      doc: { bytes, text, name: nameOf(ref) } satisfies DocxDoc,
      projection: { bytes, name: nameOf(ref) },
      outline,
    }
  },

  locate(_decoded, anchor) {
    return { start: anchor.start, end: anchor.end }
  },

  encode(_decoded, change): Patch {
    if (change.type === 'replace-all' && change.content instanceof Uint8Array) {
      return { type: 'replace-all', bytes: change.content }
    }
    throw new Error('DOCX 文档只接受二进制写回（页面通过 saveBytes 提交）')
  },
}

const agentConsumer: Consumer = {
  docType: DOC_TYPE,
  type: 'agent',
  project(decoded): AgentProjection {
    const text = (decoded.doc as DocxDoc).text
    return { type: 'text', content: text || '(空 DOCX 文档)' }
  },
}

const editorConsumer: Consumer = {
  docType: DOC_TYPE,
  type: 'editor',
  surface: 'docx-editor',
  project(decoded, raw, ref): EditorProjection {
    const doc = decoded.doc as DocxDoc
    return {
      type: 'page',
      surface: 'docx-editor',
      page: '',
      data: {
        bytes: doc.bytes,
        name: doc.name,
        ref,
        version: raw.version,
      },
    }
  },
}

registerCodec(codec)
registerConsumer(agentConsumer)
registerConsumer(editorConsumer)
