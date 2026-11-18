import { deflateSync as fflateDeflateSync } from 'fflate'

export function deflateSync(data: Uint8Array): Uint8Array {
  return fflateDeflateSync(data)
}

export function inflateSync(data: Uint8Array): Uint8Array {
  throw new Error('node:zlib inflateSync is not available in the page')
}
