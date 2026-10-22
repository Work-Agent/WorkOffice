// 浏览器侧的最小 node:crypto 替身。只覆盖 pptx-engine 在只读打开时
// 实际调用的 createHash('sha256').update(...).digest('hex') 与 randomUUID。
// 哈希仅作为打开包的缓存键，不参与文件内容校验。

function bytesOf(data: unknown): Uint8Array {
  if (typeof data === 'string') return new TextEncoder().encode(data)
  if (data instanceof Uint8Array) return data
  if (data instanceof ArrayBuffer) return new Uint8Array(data)
  if (ArrayBuffer.isView(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength)
  }
  return new TextEncoder().encode(String(data))
}

export function createHash(_algorithm: string) {
  let hash = 2166136261
  let size = 0
  return {
    update(data: unknown) {
      for (const byte of bytesOf(data)) {
        hash ^= byte
        hash = Math.imul(hash, 16777619)
        size += 1
      }
      return this
    },
    digest(encoding?: string) {
      const hex = (hash >>> 0).toString(16).padStart(8, '0')
      const out = `${hex}${size.toString(16)}`
      return encoding === 'hex' ? out : out
    },
  }
}

export function randomUUID(): string {
  return globalThis.crypto.randomUUID()
}
