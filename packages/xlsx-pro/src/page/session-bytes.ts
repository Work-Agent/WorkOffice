// ─── 页面会话字节的取用（打开时 / 回退路径按需） ───────────────────────
//
// 页面默认**不再**在打开时读整份字节（P5）：只有两种情况才需要它 ——
//   ① 宿主没有 `ura.page.invoke`（老宿主，只能本地解析）；
//   ② 保存回退到 `plan { kind:'bytes' }`（save-as / 内联媒体 / 计划超预算）。
//
// 因此这里既有启动时的「一次取全」，也有保存前的「按需取一次」，共用同一套
// 完整性校验：块大小、终止条件、声明长度、ZIP 魔数。

import { assembleChunks, ChunkError, DEFAULT_MAX_BYTES, hasZipMagic } from '../shared/chunks'

export const SESSION_CHUNK_BYTES = 1024 * 1024
/** 防死循环的硬上限（真正的上限由 assembleChunks 的字节预算判定）。 */
const MAX_CHUNKS = 4096

export type ByteWindowFetcher = (
  sessionId: string,
  window: { start: number; end: number },
) => Promise<Uint8Array>

/**
 * 取整份会话字节。
 *
 * `size` 是**提示**，不是约束：它是页面打开/上次保存时记下的文件长度，
 * 而文件会在保存后变大（我们写 `<v>`、`calcPr`、重压缩都会改变字节数）。
 * 所以：
 *   - 少读（total < size）→ 截断，报错（不让下游解析器猜）；
 *   - 多读（total > size）→ 文件长大了，以实际字节为准（最多打一条 debug 日志）；
 *   - size 未知（0）→ 以「返回空块」为终止条件。
 */
export async function fetchSessionBytes(
  fetcher: ByteWindowFetcher,
  sessionId: string,
  size = 0,
  chunkBytes: number = SESSION_CHUNK_BYTES,
): Promise<Uint8Array> {
  if (size > DEFAULT_MAX_BYTES)
    throw new ChunkError(
      'command.too-large',
      `工作簿 ${size} 字节超过页面打开上限 ${DEFAULT_MAX_BYTES}`,
    )
  const chunks: Uint8Array[] = []
  let offset = 0
  for (;;) {
    const chunk = await fetcher(sessionId, { start: offset, end: offset + chunkBytes })
    const bytes = chunk instanceof Uint8Array ? chunk : new Uint8Array(chunk as ArrayBufferLike)
    if (bytes.byteLength === 0) break
    chunks.push(bytes)
    offset += bytes.byteLength
    if (size > 0 && offset >= size) break
    if (chunks.length > MAX_CHUNKS) break
  }
  const total = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0)
  if (size > 0 && total < size)
    throw new ChunkError(
      'command.invalid',
      `工作簿分块读取不完整：期望 ${size} 字节，收到 ${total} 字节（${chunks.length} 块）`,
    )
  if (size > 0 && total !== size)
    console.debug(
      `[xlsx-pro] 文件长度与打开时的记录不同（${size} → ${total}），以实际字节为准`,
    )
  const all = assembleChunks(chunks)
  if (!hasZipMagic(all)) throw new ChunkError('command.invalid', '页面会话返回的不是有效的 XLSX 包')
  return all
}
