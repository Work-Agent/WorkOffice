// Synchronous SHA-256 for xlsx-gateway's package manifest helpers.
// node:crypto's createHash API is sync; Web Crypto is async, so use noble.
import { sha256 } from '@noble/hashes/sha2'
import { bytesToHex, concatBytes } from '@noble/hashes/utils'

export function createHash(algorithm: string) {
  if (algorithm.toLowerCase() !== 'sha256') {
    throw new Error(`Unsupported hash in page: ${algorithm}`)
  }
  let data = new Uint8Array()
  return {
    update(input: string | Uint8Array): unknown {
      const bytes =
        typeof input === 'string' ? new TextEncoder().encode(input) : input
      data = concatBytes(data, bytes)
      return this
    },
    digest(encoding?: string): string | Uint8Array {
      const hash = sha256(data)
      return encoding === 'hex' ? bytesToHex(hash) : hash
    },
  }
}
