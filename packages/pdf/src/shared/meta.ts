// ─── 轻量元数据增强（仅 ≤ 8 MB） ─────────────────────────────────────
//
// pdf-lib 解析 xref 与对象表是 O(对象数)。契约要求主进程单次 ≤ 约 50 ms，
// 所以只在 8 MB 以内启用；更大文件只保留 trailer 探测结果（页数由页面
// 加载后的 numPages 补齐）。

import { PDFDocument } from 'pdf-lib'
import type { PdfDoc } from './pdf-types'

const PDF_LIB_LIMIT = 8 * 1024 * 1024
const LARGE_FILE = 32 * 1024 * 1024

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
