// ─── 分块读取 ────────────────────────────────────────────────────────
//
// 页面必须持有整份字节（pdf.chunk 是范围读，但 PDF.js 需要整份 data），
// 所以拼装方式决定峰值：
// 预分配一次、按片直填，避免「先收 chunks 再 concat」的 2× 峰值。

export const CHUNK_BYTES = 1024 * 1024
/** 与方案 §8 对齐：超过 128 MB 拒绝打开，不静默截断。 */
export const MAX_PDF_BYTES = 128 * 1024 * 1024

export type ChunkReader = (range: { start: number; end: number }) => Promise<Uint8Array>

export interface AssembleOptions {
  /** 已知总长度（来自 page.get 的 info.size）；为 0 时按空块终止。 */
  size: number
  chunkBytes?: number
  /** 防御性上限：超过后停止并报错（默认 MAX_PDF_BYTES）。 */
  maxBytes?: number
}

export async function assembleBytes(
  read: ChunkReader,
  { size, chunkBytes = CHUNK_BYTES, maxBytes = MAX_PDF_BYTES }: AssembleOptions,
): Promise<Uint8Array> {
  if (size > maxBytes) {
    throw new Error(
      `PDF 过大：${Math.round(size / 1024 / 1024)} MB（上限 ${Math.round(maxBytes / 1024 / 1024)} MB）`,
    )
  }
  if (size === 0) {
    // size 未知：退化为「按空块终止」的收集路径。
    const chunks: Uint8Array[] = []
    let offset = 0
    for (;;) {
      const chunk = await read({ start: offset, end: offset + chunkBytes })
      if (chunk.byteLength === 0) break
      chunks.push(chunk)
      offset += chunk.byteLength
      if (offset > maxBytes) throw new Error('PDF 分块读取超出大小上限')
    }
    if (chunks.length === 0) throw new Error('页面会话没有返回内容')
    const all = new Uint8Array(offset)
    let at = 0
    for (const chunk of chunks) {
      all.set(chunk, at)
      at += chunk.byteLength
    }
    return all
  }

  const out = new Uint8Array(size)
  let offset = 0
  while (offset < size) {
    const end = Math.min(offset + chunkBytes, size)
    const chunk = await read({ start: offset, end })
    if (chunk.byteLength === 0) break
    out.set(chunk, offset)
    offset += chunk.byteLength
  }
  if (offset !== size) {
    throw new Error(`PDF 字节读取不完整：期望 ${size}，实际 ${offset}`)
  }
  return out
}
