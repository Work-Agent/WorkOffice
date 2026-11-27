// ─── PDF 打开：头部校验 + trailer 探测 + 元数据增强 ───────────────────
//
// codec.decode 已不存在；这里就是页面与资源读共用的“decode”。

import { hasPdfHeader, probePdfMeta } from '../shared/pdf-meta'
import type { PdfDoc } from '../shared/pdf-types'
import { enrichMeta } from './meta'
import { MAX_PDF_BYTES, fail } from './ref'

export async function loadDoc(
  ref: string,
  name: string,
  bytes: Uint8Array,
  version: string,
): Promise<PdfDoc> {
  if (bytes.byteLength === 0) fail('command.invalid', 'PDF 文件为空')
  if (bytes.byteLength > MAX_PDF_BYTES) {
    fail('command.too-large', `PDF 超过 ${MAX_PDF_BYTES / 1024 / 1024} MB 上限`)
  }
  if (!hasPdfHeader(bytes)) {
    fail('command.invalid', '不是有效的 PDF 文件（缺少 %PDF- 头）')
  }

  const probe = probePdfMeta(bytes)
  const doc: PdfDoc = {
    kind: 'pdf',
    name,
    size: bytes.byteLength,
    version,
    ...(probe.headerVersion ? { headerVersion: probe.headerVersion } : {}),
    ...(probe.pageCount !== undefined ? { pageCount: probe.pageCount } : {}),
    encrypted: probe.encrypted,
    ...(probe.title ? { title: probe.title } : {}),
    ...(probe.author ? { author: probe.author } : {}),
    ...(probe.warning ? { metadataWarning: probe.warning } : {}),
  }
  const enriched = await enrichMeta(bytes, doc)
  return {
    ...doc,
    ...(enriched.pageCount !== undefined ? { pageCount: enriched.pageCount } : {}),
    ...(enriched.encrypted !== undefined ? { encrypted: enriched.encrypted } : {}),
    ...(enriched.title ? { title: enriched.title } : {}),
    ...(enriched.author ? { author: enriched.author } : {}),
    ...(enriched.firstPageSize ? { firstPageSize: enriched.firstPageSize } : {}),
    ...(enriched.metadataWarning ? { metadataWarning: enriched.metadataWarning } : {}),
  }
}
