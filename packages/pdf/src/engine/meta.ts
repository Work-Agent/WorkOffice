// ─── 元数据增强与摘要 ────────────────────────────────────────────────
//
// pdf-lib 解析 xref 与对象表是 O(对象数)。主进程打开/读取时单次 ≤ 8 MB，
// 更大文件只保留 trailer 探测结果（页数由页面加载后的 numPages 补齐）。

import { PDFDocument } from 'pdf-lib'
import type { PdfDoc } from '../shared/pdf-types'
import { formatSize, truncate } from '../shared/format'

const PDF_LIB_LIMIT = 8 * 1024 * 1024
const LARGE_FILE = 32 * 1024 * 1024
export const SIDECAR_MAX_CHARS = 32_000

export interface EnrichedMeta {
  pageCount?: number
  encrypted?: boolean
  title?: string
  author?: string
  firstPageSize?: { width: number; height: number }
  metadataWarning?: string
}

export async function enrichMeta(bytes: Uint8Array, base: PdfDoc): Promise<EnrichedMeta> {
  const size = bytes.byteLength
  if (size > PDF_LIB_LIMIT) {
    return {
      ...(base.pageCount !== undefined ? { pageCount: base.pageCount } : {}),
      ...(base.encrypted !== undefined ? { encrypted: base.encrypted } : {}),
      ...(size > LARGE_FILE ? { metadataWarning: 'large-file-metadata-skipped' } : {}),
    }
  }

  try {
    const pdf = await PDFDocument.load(bytes, {
      ignoreEncryption: true,
      updateMetadata: false,
      throwOnInvalidObject: false,
    })
    const pageCount = safe(() => pdf.getPageCount())
    const title = safe(() => clean(pdf.getTitle()))
    const author = safe(() => clean(pdf.getAuthor()))
    const firstPageSize = safe(() => {
      if (pdf.getPageCount() === 0) return undefined
      const size = pdf.getPage(0).getSize()
      return Number.isFinite(size.width) && Number.isFinite(size.height)
        ? { width: size.width, height: size.height }
        : undefined
    })
    return {
      ...(pageCount !== undefined && pageCount >= 0 ? { pageCount } : {}),
      encrypted: base.encrypted === true || safe(() => pdf.isEncrypted) === true,
      ...(title ? { title } : {}),
      ...(author ? { author } : {}),
      ...(firstPageSize ? { firstPageSize } : {}),
    }
  } catch {
    // 加密/损坏/超大对象表都走这里：打开仍然交给 PDF.js 恢复。
    return {
      ...(base.pageCount !== undefined ? { pageCount: base.pageCount } : {}),
      ...(base.encrypted !== undefined ? { encrypted: base.encrypted } : {}),
      metadataWarning: 'metadata-unavailable',
    }
  }
}

/** 资源 read / 页面 context 共用的元数据摘要；不是正文。 */
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

function safe<T>(read: () => T): T | undefined {
  try {
    return read()
  } catch {
    return undefined
  }
}

function clean(value: string | undefined): string | undefined {
  const text = value?.replace(/\u0000/g, '').trim()
  return text ? text : undefined
}
