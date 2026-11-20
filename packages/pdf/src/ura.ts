// ─── PDF codec / presenter / sidecar ──────────────────────────────────
//
// 只读插件：不声明 write、不注册 store / tools / engine。
// codec 只做「头部校验 + trailer 探测 + ≤8 MB 元数据增强」，
// 不在主进程解析内容流、不启动 PDF.js worker。
//
// sidecar 在主进程求值，因此只拼 codec 已经算好的元数据（O(1) 字符串拼接）。

import { registerCodec, registerPresenter, registerSidecar } from './sdk'
import { MIME } from './ura-types'
import type {
  Codec,
  DecodedDoc,
  Patch,
  Presenter,
  RawContent,
  Sidecar,
  SidecarProvider,
} from './ura-types'
import { formatSize, truncate } from './shared/format'
import { enrichMeta } from './shared/meta'
import { hasPdfHeader, probePdfMeta } from './shared/pdf-meta'
import type { PdfDoc } from './shared/pdf-types'

const CODEC_ID = 'pdf'
const MIME_PDF = MIME.pdf
const SIDECAR_MAX_CHARS = 32_000

interface PdfDecoded {
  doc: PdfDoc
  summary: string
}

function nameOf(ref: string): string {
  try {
    const url = new URL(ref)
    const name = url.pathname.slice(url.pathname.lastIndexOf('/') + 1)
    return decodeURIComponent(name) || 'document.pdf'
  } catch {
    const name = ref.slice(Math.max(ref.lastIndexOf('/'), ref.lastIndexOf('\\')) + 1)
    return name || 'document.pdf'
  }
}

/** sidecar 与 anchorMap 共用的元数据摘要；不是正文。 */
export function summaryOf(doc: PdfDoc): string {
  const lines = [
    `# ${doc.name}`,
    `pages=${doc.pageCount ?? '?'} size=${formatSize(doc.size)}`,
    doc.headerVersion ? `pdf=${doc.headerVersion}` : '',
    doc.title ? `title: ${doc.title}` : '',
    doc.author ? `author: ${doc.author}` : '',
    doc.encrypted ? 'encrypted: true' : '',
    doc.firstPageSize ? `firstPage: ${doc.firstPageSize.width}×${doc.firstPageSize.height} pt` : '',
    doc.metadataWarning ? `metadata: ${doc.metadataWarning}` : '',
  ].filter(Boolean)
  return truncate(lines.join('\n'), SIDECAR_MAX_CHARS)
}

const codec: Codec = {
  id: CODEC_ID,
  keys: [{ transport: '*', mime: MIME_PDF }],
  extensions: ['pdf'],
  patchType: 'bytes',

  async decode(raw: RawContent, ref: string): Promise<DecodedDoc> {
    const bytes = raw.bytes
    if (!hasPdfHeader(bytes)) {
      throw new Error('不是有效的 PDF 文件（缺少 %PDF- 头）')
    }
    const probe = probePdfMeta(bytes)
    const doc: PdfDoc = {
      kind: 'pdf',
      name: nameOf(ref),
      size: bytes.byteLength,
      version: raw.version,
      ...(probe.headerVersion ? { headerVersion: probe.headerVersion } : {}),
      ...(probe.pageCount !== undefined ? { pageCount: probe.pageCount } : {}),
      encrypted: probe.encrypted,
      ...(probe.title ? { title: probe.title } : {}),
      ...(probe.author ? { author: probe.author } : {}),
      ...(probe.warning ? { metadataWarning: probe.warning } : {}),
    }
    const enriched = await enrichMeta(bytes, doc)
    const merged: PdfDoc = {
      ...doc,
      ...(enriched.pageCount !== undefined ? { pageCount: enriched.pageCount } : {}),
      ...(enriched.encrypted !== undefined ? { encrypted: enriched.encrypted } : {}),
      ...(enriched.title ? { title: enriched.title } : {}),
      ...(enriched.author ? { author: enriched.author } : {}),
      ...(enriched.firstPageSize ? { firstPageSize: enriched.firstPageSize } : {}),
      ...(enriched.metadataWarning ? { metadataWarning: enriched.metadataWarning } : {}),
    }
    const decoded: PdfDecoded = { doc: merged, summary: summaryOf(merged) }
    return { doc: decoded, anchorMap: decoded.summary }
  },

  // 只读插件没有锚点语义：locate 只把锚点原样返回（页面不会用它做写回）。
  locate(_decoded: DecodedDoc, anchor) {
    return { start: anchor.start, end: anchor.end }
  },

  encode(): Patch {
    throw new Error('command.forbidden')
  },
}

const presenter: Presenter = {
  key: { transport: '*', mime: MIME_PDF, codec: CODEC_ID },
  type: 'page',
  surface: 'pdf-viewer',

  present(decoded: DecodedDoc, raw: RawContent, ref: string) {
    const { doc } = decoded.doc as PdfDecoded
    return {
      type: 'page',
      surface: 'pdf-viewer',
      page: '',
      data: {
        kind: 'pdf',
        name: doc.name,
        ref,
        size: raw.bytes.byteLength,
        version: raw.version,
        pageCount: doc.pageCount ?? null,
        encrypted: doc.encrypted ?? null,
        title: doc.title ?? null,
      },
    }
  },
}

const sidecar: SidecarProvider = {
  key: { transport: '*', mime: MIME_PDF, codec: CODEC_ID },

  sidecar(decoded: DecodedDoc): Sidecar {
    // 纯字符串拼接：不读 raw.bytes、不 import pdfjs-dist、不做任何 PDF 解析。
    const { summary } = decoded.doc as PdfDecoded
    return {
      type: 'text',
      content: summary,
      ...(summary.length >= SIDECAR_MAX_CHARS ? { truncated: true } : {}),
    }
  },
}

registerCodec(codec)
registerPresenter(presenter)
registerSidecar(sidecar)
