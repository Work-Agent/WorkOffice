// ─── zip 条目级重写（只替换一个 part） ─────────────────────────────────
//
// 契约要求「只重写被编辑的 OOXML 片段」：除 `word/document.xml` 之外的 zip
// 条目必须**逐字节**不变（含压缩流、data descriptor、时间戳、权限位）。
// fflate 的 `zipSync` 会把所有条目重新压缩，做不到这一点，所以这里直接操作
// zip 结构：
//
//   [local record]×n → [central directory] → [EOCD]
//
// 未改条目的 local record 与 central directory record 都原样拷贝（只改 CD
// record 里的 localOffset）；只有目标条目重新 deflate 并重算 CRC。
//
// 限制（README「已知差异」）：zip64 包直接拒绝（docx 极少见，且需要重写
// zip64 EOCD / extra 字段，收益不足以在本期实现）。

import { deflateSync } from 'fflate';

const SIG_LOCAL = 0x04034b50;
const SIG_CD = 0x02014b50;
const SIG_EOCD = 0x06054b50;

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let value = i;
    for (let bit = 0; bit < 8; bit += 1)
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    table[i] = value >>> 0;
  }
  return table;
})();

export function crc32(bytes: Uint8Array, seed = 0): number {
  let crc = (seed ^ 0xffffffff) >>> 0;
  for (const byte of bytes) crc = (CRC_TABLE[(crc ^ byte) & 0xff]! ^ (crc >>> 8)) >>> 0;
  return (crc ^ 0xffffffff) >>> 0;
}

export class ZipShapeError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ZipShapeError';
  }
}

interface CdEntry {
  name: string;
  versionMadeBy: number;
  versionNeeded: number;
  flags: number;
  method: number;
  modTime: number;
  modDate: number;
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
  internalAttributes: number;
  externalAttributes: number;
  localOffset: number;
  /** 原始 CD record 字节（未改条目原样回写）。 */
  cdRecord: Uint8Array;
  nameBytes: Uint8Array;
}

interface ZipLayout {
  entries: CdEntry[];
  cdStart: number;
  comment: Uint8Array;
}

const decoder = new TextDecoder();

function readU16(view: DataView, offset: number): number {
  return view.getUint16(offset, true);
}

function readU32(view: DataView, offset: number): number {
  return view.getUint32(offset, true);
}

function viewOf(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function layoutOf(bytes: Uint8Array): ZipLayout {
  const view = viewOf(bytes);
  const length = bytes.byteLength;
  if (length < 22) throw new ZipShapeError('不是有效的 ZIP（太短）');
  let eocd = -1;
  const stop = Math.max(0, length - 22 - 0xffff);
  for (let pos = length - 22; pos >= stop; pos -= 1) {
    if (readU32(view, pos) !== SIG_EOCD) continue;
    const commentLength = readU16(view, pos + 20);
    if (pos + 22 + commentLength === length) {
      eocd = pos;
      break;
    }
  }
  if (eocd < 0) throw new ZipShapeError('找不到 ZIP 目录尾（EOCD）');
  const disk = readU16(view, eocd + 4);
  const cdDisk = readU16(view, eocd + 6);
  const entryCount = readU16(view, eocd + 10);
  const cdSize = readU32(view, eocd + 12);
  const cdStart = readU32(view, eocd + 16);
  const commentLength = readU16(view, eocd + 20);
  if (disk !== 0 || cdDisk !== 0)
    throw new ZipShapeError('不支持分卷 ZIP');
  if (entryCount === 0xffff || cdSize === 0xffffffff || cdStart === 0xffffffff)
    throw new ZipShapeError('不支持 zip64 文档（请先用 Word 另存为普通 .docx）');
  if (cdStart + cdSize > eocd)
    throw new ZipShapeError('ZIP 目录越界');

  const entries: CdEntry[] = [];
  let pos = cdStart;
  for (let index = 0; index < entryCount; index += 1) {
    if (pos + 46 > eocd || readU32(view, pos) !== SIG_CD)
      throw new ZipShapeError(`ZIP 中央目录损坏（第 ${index} 项）`);
    const nameLength = readU16(view, pos + 28);
    const extraLength = readU16(view, pos + 30);
    const entryCommentLength = readU16(view, pos + 32);
    const next = pos + 46 + nameLength + extraLength + entryCommentLength;
    if (next > eocd) throw new ZipShapeError('ZIP 中央目录越界');
    const nameBytes = bytes.slice(pos + 46, pos + 46 + nameLength);
    entries.push({
      name: decoder.decode(nameBytes),
      versionMadeBy: readU16(view, pos + 4),
      versionNeeded: readU16(view, pos + 6),
      flags: readU16(view, pos + 8),
      method: readU16(view, pos + 10),
      modTime: readU16(view, pos + 12),
      modDate: readU16(view, pos + 14),
      crc: readU32(view, pos + 16),
      compressedSize: readU32(view, pos + 20),
      uncompressedSize: readU32(view, pos + 24),
      internalAttributes: readU16(view, pos + 36),
      externalAttributes: readU32(view, pos + 38),
      localOffset: readU32(view, pos + 42),
      cdRecord: bytes.slice(pos, next),
      nameBytes,
    });
    pos = next;
  }
  const comment = bytes.slice(eocd + 22, eocd + 22 + commentLength);
  return { entries, cdStart, comment };
}

/**
 * 用 `newData` 替换 `partName` 指定的 zip 条目，其余条目字节不变。
 *
 * `word/document.xml` 的路径大小写按 zip 规范敏感（DOS 产物可能是大写），
 * 因此按完全相等匹配；找不到即抛错（调用方回 `command.invalid`）。
 */
export function rewriteZipPart(
  bytes: Uint8Array,
  partName: string,
  newData: Uint8Array,
): Uint8Array {
  const layout = layoutOf(bytes);
  const target = layout.entries.find((entry) => entry.name === partName);
  if (!target) throw new ZipShapeError(`DOCX 缺少 ${partName}`);

  const ordered = [...layout.entries].sort(
    (a, b) => a.localOffset - b.localOffset,
  );
  for (const [index, entry] of ordered.entries()) {
    const end = index + 1 < ordered.length ? ordered[index + 1]!.localOffset : layout.cdStart;
    if (entry.localOffset < 0 || end < entry.localOffset || end > layout.cdStart)
      throw new ZipShapeError(`ZIP 局部记录越界：${entry.name}`);
  }

  const chunks: Uint8Array[] = [];
  let offset = 0;
  const push = (chunk: Uint8Array): void => {
    chunks.push(chunk);
    offset += chunk.byteLength;
  };

  const newOffsets = new Map<string, number>();
  for (const [index, entry] of ordered.entries()) {
    const end = index + 1 < ordered.length ? ordered[index + 1]!.localOffset : layout.cdStart;
    newOffsets.set(entry.name, offset);
    if (entry.name !== partName) {
      push(bytes.subarray(entry.localOffset, end));
      continue;
    }
    push(localRecordOf(entry, newData));
  }

  // central directory：未改条目原样拷贝（只改 localOffset），目标条目重建。
  const cdChunks: Uint8Array[] = [];
  let cdSize = 0;
  for (const entry of ordered) {
    const built =
      entry.name === partName
        ? cdRecordOf(entry, newData, newOffsets.get(entry.name)!)
        : relocatedCdRecord(entry, newOffsets.get(entry.name)!);
    cdChunks.push(built);
    cdSize += built.byteLength;
  }
  const cdStart = offset;
  for (const chunk of cdChunks) push(chunk);

  const eocd = new Uint8Array(22 + layout.comment.byteLength);
  const view = viewOf(eocd);
  view.setUint32(0, SIG_EOCD, true);
  view.setUint16(8, ordered.length, true);
  view.setUint16(10, ordered.length, true);
  view.setUint32(12, cdSize, true);
  view.setUint32(16, cdStart, true);
  view.setUint16(20, layout.comment.byteLength, true);
  eocd.set(layout.comment, 22);
  push(eocd);

  const out = new Uint8Array(offset);
  let cursor = 0;
  for (const chunk of chunks) {
    out.set(chunk, cursor);
    cursor += chunk.byteLength;
  }
  return out;
}

function localRecordOf(entry: CdEntry, data: Uint8Array): Uint8Array {
  const compressed = compress(data);
  const flags = (entry.flags & ~0x0008) | (hasNonAscii(entry.nameBytes) ? 0x0800 : 0);
  const header = new Uint8Array(30 + entry.nameBytes.byteLength);
  const view = viewOf(header);
  view.setUint32(0, SIG_LOCAL, true);
  view.setUint16(4, Math.max(20, entry.versionNeeded), true);
  view.setUint16(6, flags, true);
  view.setUint16(8, compressed.method, true);
  view.setUint16(10, entry.modTime, true);
  view.setUint16(12, entry.modDate, true);
  view.setUint32(14, crc32(data), true);
  view.setUint32(18, compressed.bytes.byteLength, true);
  view.setUint32(22, data.byteLength, true);
  view.setUint16(26, entry.nameBytes.byteLength, true);
  view.setUint16(28, 0, true);
  header.set(entry.nameBytes, 30);
  return concat([header, compressed.bytes]);
}

function cdRecordOf(
  entry: CdEntry,
  data: Uint8Array,
  localOffset: number,
): Uint8Array {
  const compressed = compress(data);
  const flags = (entry.flags & ~0x0008) | (hasNonAscii(entry.nameBytes) ? 0x0800 : 0);
  const record = new Uint8Array(46 + entry.nameBytes.byteLength);
  const view = viewOf(record);
  view.setUint32(0, SIG_CD, true);
  view.setUint16(4, entry.versionMadeBy, true);
  view.setUint16(6, Math.max(20, entry.versionNeeded), true);
  view.setUint16(8, flags, true);
  view.setUint16(10, compressed.method, true);
  view.setUint16(12, entry.modTime, true);
  view.setUint16(14, entry.modDate, true);
  view.setUint32(16, crc32(data), true);
  view.setUint32(20, compressed.bytes.byteLength, true);
  view.setUint32(24, data.byteLength, true);
  view.setUint16(28, entry.nameBytes.byteLength, true);
  view.setUint16(30, 0, true);
  view.setUint16(32, 0, true);
  view.setUint16(34, 0, true);
  view.setUint16(36, entry.internalAttributes, true);
  view.setUint32(38, entry.externalAttributes, true);
  view.setUint32(42, localOffset, true);
  record.set(entry.nameBytes, 46);
  return record;
}

function relocatedCdRecord(entry: CdEntry, localOffset: number): Uint8Array {
  const record = entry.cdRecord.slice();
  viewOf(record).setUint32(42, localOffset, true);
  return record;
}

/** 压缩策略与 fflate 默认一致（raw deflate）；不可压缩时降级 store。 */
function compress(data: Uint8Array): { bytes: Uint8Array; method: number } {
  const deflated = deflateSync(data, { level: 6 });
  if (deflated.byteLength >= data.byteLength)
    return { bytes: data, method: 0 };
  return { bytes: deflated, method: 8 };
}

function hasNonAscii(name: Uint8Array): boolean {
  for (const byte of name) if (byte > 0x7f) return true;
  return false;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, part) => sum + part.byteLength, 0);
  const out = new Uint8Array(total);
  let cursor = 0;
  for (const part of parts) {
    out.set(part, cursor);
    cursor += part.byteLength;
  }
  return out;
}

/** 诊断用：读 zip 条目名（测试 / `check`）。 */
export function zipEntryNames(bytes: Uint8Array): string[] {
  return layoutOf(bytes).entries.map((entry) => entry.name);
}

/** 诊断 / 测试用：每个条目的原始 local record（未改条目必须逐字节相等）。 */
export function zipEntryRecords(bytes: Uint8Array): Map<string, Uint8Array> {
  const layout = layoutOf(bytes);
  const ordered = [...layout.entries].sort((a, b) => a.localOffset - b.localOffset);
  const map = new Map<string, Uint8Array>();
  for (const [index, entry] of ordered.entries()) {
    const end = index + 1 < ordered.length ? ordered[index + 1]!.localOffset : layout.cdStart;
    map.set(entry.name, bytes.slice(entry.localOffset, end));
  }
  return map;
}
