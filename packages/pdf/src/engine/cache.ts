// ─── 字节缓存：大文件 page.chunk 的来源 ────────────────────────────
//
// pdf.load 为了算 hash 已经整份读过一次；大文件把这份字节留在主进程，
// page.chunk 直接切片，避免每块都重新读盘。小文件字节已经发给页面，
// 不再缓存。

interface Entry {
  hash: string
  bytes: Uint8Array
  at: number
}

export class BytesCache {
  private readonly entries = new Map<string, Entry>()

  constructor(
    private readonly maxEntries = 1,
    private readonly maxBytes = 160 * 1024 * 1024,
  ) {}

  get(ref: string): Entry | null {
    return this.entries.get(ref) ?? null
  }

  put(ref: string, hash: string, bytes: Uint8Array): void {
    if (bytes.byteLength > this.maxBytes) return
    this.entries.delete(ref)
    this.entries.set(ref, { hash, bytes, at: Date.now() })
    while (this.entries.size > this.maxEntries) {
      const first = this.entries.keys().next().value
      if (first === undefined) break
      this.entries.delete(first)
    }
  }

  delete(ref: string): void {
    this.entries.delete(ref)
  }

  clear(): void {
    this.entries.clear()
  }
}

export const bytesCache = new BytesCache()
