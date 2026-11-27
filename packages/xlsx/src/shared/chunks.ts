// ─── 分块字节拼装（页面打开 / 重新取字节） ─────────────────────────────
//
// 协议（设计稿 §5.1）：
//   块大小 ≤ 1 MB，end 排他；返回 0 字节即结束；块数上限 1024。
//
// 旧实现把「中途返回 0 字节」也当成正常结束，于是截断的文件会被
// 当成完整字节交给解析器，错误最后报在「不是有效的 XLSX」上。这里把
// 完整性判定挪到拼装处：**长度对不上就是截断**，不再靠下游猜。

export type ChunkErrorCode = 'command.invalid' | 'command.too-large'

export class ChunkError extends Error {
  constructor(
    readonly code: ChunkErrorCode,
    message: string,
  ) {
    super(message)
    this.name = 'ChunkError'
  }
}

export interface ChunkLimits {
  /** 单次打开接受的字节上限（默认 64 MB，与 engine 的 maxFileBytes 对齐）。 */
  readonly maxBytes?: number
  /** 块数上限（默认 1024，≈ 1 GB @1 MB 块）。 */
  readonly maxChunks?: number
}

export const DEFAULT_MAX_BYTES = 64 * 1024 * 1024
export const DEFAULT_MAX_CHUNKS = 1024

/**
 * 拼装分块字节。
 *
 * ⚠️ 严格相等只适合"长度由协议保证"的通道。**读文件不要用严格模式**：
 * 文件在保存后会变大（我们写 `<v>` / `calcPr` / 重压缩都会改字节数），
 * 拿打开时的 size 硬卡长度会把正常保存判成"读错了"。文件读取走
 * `src/page/session-bytes.ts`（只把"少读"当截断）。
 *
 * - `expectedSize > 0` 时长度必须**精确相等**，否则判为截断（不完整 / 超出）。
 * - `expectedSize` 未知（0 / undefined）时只能按块数上限兜底。
 * - 任何一块为空都表示结束（调用方不应再追加），这里只校验总量。
 */
export function assembleChunks(
  chunks: readonly Uint8Array[],
  expectedSize?: number,
  limits: ChunkLimits = {},
): Uint8Array {
  const maxBytes = limits.maxBytes ?? DEFAULT_MAX_BYTES
  const maxChunks = limits.maxChunks ?? DEFAULT_MAX_CHUNKS

  if (chunks.length === 0) throw new ChunkError('command.invalid', '页面会话没有返回内容')
  if (chunks.length > maxChunks)
    throw new ChunkError(
      'command.too-large',
      `分块数超过上限（${chunks.length} > ${maxChunks}）；工作簿过大，无法在页面打开`,
    )

  let total = 0
  for (const chunk of chunks) total += chunk.byteLength
  if (total > maxBytes)
    throw new ChunkError(
      'command.too-large',
      `工作簿 ${total} 字节超过页面打开上限 ${maxBytes}；请用 headless 路径处理`,
    )

  const expected = Number(expectedSize ?? 0)
  if (expected > 0 && total !== expected)
    throw new ChunkError(
      'command.invalid',
      total < expected
        ? `工作簿分块读取不完整：期望 ${expected} 字节，收到 ${total} 字节（${chunks.length} 块）`
        : `工作簿分块读取超出声明大小：期望 ${expected} 字节，收到 ${total} 字节`,
    )

  if (chunks.length === 1) return chunks[0]!
  const all = new Uint8Array(total)
  let at = 0
  for (const chunk of chunks) {
    all.set(chunk, at)
    at += chunk.byteLength
  }
  return all
}

/** ZIP 本地文件头 / 空包结束标记。比 ZIP 之外的解析错误好定位。 */
export function hasZipMagic(bytes: Uint8Array): boolean {
  if (bytes.byteLength < 4) return false
  const [a, b, c, d] = bytes
  if (a !== 0x50 || b !== 0x4b) return false
  // PK\x03\x04 = 本地文件头；PK\x05\x06 = 空归档；PK\x07\x08 = 跨卷
  return (c === 3 && d === 4) || (c === 5 && d === 6) || (c === 7 && d === 8)
}
