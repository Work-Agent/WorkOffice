// Must be the first import in the page entry: JSZip snapshots
// `typeof Buffer !== "undefined"` at module-evaluation time.
import { Buffer } from 'buffer'

if (!(globalThis as { Buffer?: typeof Buffer }).Buffer) {
  ;(globalThis as { Buffer?: typeof Buffer }).Buffer = Buffer
}
