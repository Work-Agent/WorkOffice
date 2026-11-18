import { sha256 } from '@noble/hashes/sha2'
import { bytesToHex, concatBytes } from '@noble/hashes/utils'

export function createHash(algorithm: string) {
  if (algorithm.toLowerCase() !== 'sha256') {
    throw new Error(`Unsupported hash in page: ${algorithm}`)
  }
  let data: Uint8Array<ArrayBufferLike> = new Uint8Array()
  return {
    update(input: string | Uint8Array) {
      const bytes =
        typeof input === 'string' ? new TextEncoder().encode(input) : input
      data = concatBytes(data, bytes)
      return this
    },
    digest(encoding?: string): string | Uint8Array {
      const hash = sha256(data)
      return encoding === 'hex' ? bytesToHex(hash) : (hash as Uint8Array<ArrayBuffer>)
    },
  }
}

export function randomUUID(): string {
  return crypto.randomUUID()
}
