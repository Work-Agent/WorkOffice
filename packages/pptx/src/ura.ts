// ─── PPTX codec（WorkStudio 自有） ───────────────────────────────────
//
// codec 只做 OOXML 包的最小解析：按 slideN.xml 顺序抽文本与标题，
// 供 agent 投影使用。页面层再拿原始字节交给 GenOffice 渲染树。

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

const DOC_TYPE = 'ooxml-pptx@1' as const

interface SlideText {
  index: number
  title: string
  text: string
}

interface PptxDoc {
  bytes: string
  name: string
  slides: SlideText[]
}

function base64Of(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString('base64')
}

function nameOf(ref: string): string {
  try {
    const url = new URL(ref)
    const name = url.pathname.slice(url.pathname.lastIndexOf('/') + 1)
    return decodeURIComponent(name) || 'presentation.pptx'
  } catch {
    const name = ref.slice(ref.lastIndexOf('/') + 1)
    return name || 'presentation.pptx'
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

function slideText(xml: string): string[] {
  const paragraphs: string[] = []
  const paragraphPattern = /<a:p(?:\s[^>]*)?>([\s\S]*?)<\/a:p>/g
  for (const paragraph of xml.matchAll(paragraphPattern)) {
    const body = paragraph[1] ?? ''
    let text = ''
    const token = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>|<a:br\s*\/>/g
    for (const match of body.matchAll(token)) {
      if (match[1] !== undefined) text += decodeXml(match[1])
      else text += '\n'
    }
    const normalized = text.trimEnd()
    if (normalized.trim()) paragraphs.push(normalized)
  }
  return paragraphs
}

function parseSlides(bytes: Uint8Array): SlideText[] {
  let files: Record<string, Uint8Array>
  try {
    files = unzipSync(bytes)
  } catch {
    throw new Error('不是有效的 PPTX 包（ZIP 解压失败）')
  }
  const slides = Object.keys(files)
    .map((path) => {
      const match = /^ppt\/slides\/slide(\d+)\.xml$/.exec(path)
      return match
        ? { path, index: Number(match[1]), file: files[path]! }
        : null
    })
    .filter((entry): entry is { path: string; index: number; file: Uint8Array } => !!entry)
    .sort((a, b) => a.index - b.index)
  if (slides.length === 0) throw new Error('PPTX 没有可读幻灯片')
  return slides.map(({ index, file }) => {
    const paragraphs = slideText(strFromU8(file))
    return {
      index,
      title: paragraphs[0] ?? '',
      text: paragraphs.join('\n'),
    }
  })
}

function slidesText(slides: SlideText[]): string {
  return slides
    .map((slide, i) =>
      slide.text
        ? `## 第 ${i + 1} 页\n${slide.text}`
        : `## 第 ${i + 1} 页`,
    )
    .join('\n\n')
}

const codec: Codec = {
  hints: ['ooxml-pptx'],
  extensions: ['pptx', 'pptm', 'potx'],
  docType: DOC_TYPE,
  patchType: 'bytes',

  decode(raw: RawContent, ref: string) {
    const slides = parseSlides(raw.bytes)
    const bytes = base64Of(raw.bytes)
    return {
      docType: DOC_TYPE,
      doc: { bytes, name: nameOf(ref), slides } satisfies PptxDoc,
      projection: {
        bytes,
        name: nameOf(ref),
        slideCount: slides.length,
      },
      outline: slides.map((slide, i) => ({
        level: 1,
        title: slide.title || `第 ${i + 1} 页`,
      })),
    }
  },

  locate(_decoded, anchor) {
    return { start: anchor.start, end: anchor.end }
  },

  encode(_decoded, change): Patch {
    if (change.type === 'replace-all' && change.content instanceof Uint8Array) {
      return { type: 'replace-all', bytes: change.content }
    }
    throw new Error('PPTX 文档只接受二进制写回')
  },
}

const agentConsumer: Consumer = {
  docType: DOC_TYPE,
  type: 'agent',
  project(decoded): AgentProjection {
    const text = slidesText((decoded.doc as PptxDoc).slides)
    return { type: 'text', content: text || '(空 PPTX 文档)' }
  },
}

const editorConsumer: Consumer = {
  docType: DOC_TYPE,
  type: 'editor',
  surface: 'pptx-viewer',
  project(decoded, raw, ref): EditorProjection {
    const doc = decoded.doc as PptxDoc
    return {
      type: 'page',
      surface: 'pptx-viewer',
      page: '',
      data: {
        bytes: doc.bytes,
        name: doc.name,
        ref,
        version: raw.version,
        slideCount: doc.slides.length,
      },
    }
  },
}

registerCodec(codec)
registerConsumer(agentConsumer)
registerConsumer(editorConsumer)
