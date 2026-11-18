// ─── ZIP 原始条目读取（不解压） ────────────────────────────────────────
//
// 用来断言 I1：`save_archive` 之后，**未被改动的条目压缩流逐字节不变**。
// 用解压后的 hash 比对是不够的 —— JSZip 那类实现会重压缩整包，解压内容
// 一样但压缩流全变了，正是我们要禁止的行为。
//
// 只实现读取（central directory → local header → 压缩数据切片），不写包。

export interface RawZipEntry {
  readonly name: string
  /** 压缩后的原始字节（含 deflate 头，不含 local header）。 */
  readonly raw: Uint8Array
  readonly method: number
}

const EOCD_SIGNATURE = 0x06054b50
const CENTRAL_SIGNATURE = 0x02014b50
const LOCAL_SIGNATURE = 0x04034b50

export function rawZipEntries(bytes: Uint8Array): Map<string, RawZipEntry> {
  const view = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const eocd = findEocd(view)
  const count = view.readUInt16LE(eocd + 10)
  let cursor = view.readUInt32LE(eocd + 16)
  const entries = new Map<string, RawZipEntry>()
  for (let index = 0; index < count; index += 1) {
    if (view.readUInt32LE(cursor) !== CENTRAL_SIGNATURE)
      throw new Error('ZIP central directory 损坏')
    const method = view.readUInt16LE(cursor + 10)
    const compressedSize = view.readUInt32LE(cursor + 20)
    const nameLength = view.readUInt16LE(cursor + 28)
    const extraLength = view.readUInt16LE(cursor + 30)
    const commentLength = view.readUInt16LE(cursor + 32)
    const localOffset = view.readUInt32LE(cursor + 42)
    const name = view.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8')
    if (view.readUInt32LE(localOffset) !== LOCAL_SIGNATURE)
      throw new Error(`ZIP local header 损坏：${name}`)
    const localNameLength = view.readUInt16LE(localOffset + 26)
    const localExtraLength = view.readUInt16LE(localOffset + 28)
    const dataStart = localOffset + 30 + localNameLength + localExtraLength
    entries.set(name, {
      name,
      raw: new Uint8Array(bytes.buffer, bytes.byteOffset + dataStart, compressedSize),
      method,
    })
    cursor += 46 + nameLength + extraLength + commentLength
  }
  return entries
}

function findEocd(view: Buffer): number {
  const earliest = Math.max(0, view.length - 22 - 0xffff)
  for (let offset = view.length - 22; offset >= earliest; offset -= 1) {
    if (view.readUInt32LE(offset) === EOCD_SIGNATURE) return offset
  }
  throw new Error('找不到 ZIP EOCD（不是 zip 包，或是 zip64 且无兼容记录）')
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false
  for (let index = 0; index < left.byteLength; index += 1)
    if (left[index] !== right[index]) return false
  return true
}

/**
 * 断言：`touched` 之外的条目压缩流逐字节不变。
 * 返回被意外改动的条目名（空数组 = 通过），便于调用方决定报错还是告警。
 */
export function driftedEntries(
  before: Uint8Array,
  after: Uint8Array,
  touched: Iterable<string>,
): string[] {
  const keep = new Set(touched)
  const beforeEntries = rawZipEntries(before)
  const afterEntries = rawZipEntries(after)
  const drifted: string[] = []
  for (const [name, entry] of beforeEntries) {
    if (keep.has(name)) continue
    const next = afterEntries.get(name)
    if (!next || !sameBytes(entry.raw, next.raw)) drifted.push(name)
  }
  return drifted
}
