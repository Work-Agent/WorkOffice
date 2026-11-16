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

/** 一个 part 的改写意图：给出新字节即替换 / 新增，`delete` 即移除。 */
export interface ZipPartPatch {
  path: string;
  /** 新内容（无则必须带 `delete`）。 */
  data?: Uint8Array;
  delete?: boolean;
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
  // 单 part 语义 = 替换：缺失即报错（不能静默变成「新增一个 part」）。
  if (!zipEntryNames(bytes).includes(partName))
    throw new ZipShapeError(`DOCX 缺少 ${partName}`);
  return rewriteZipParts(bytes, [{ path: partName, data: newData }]);
}

/**
 * 一次替换 / 新增 / 删除多个 part（P5 的整包通道）。
 *
 * 哲学与单 part 版本一致：**未涉及的条目逐字节拷贝**（local record 与 CD record
 * 都只改 `localOffset`）。替换 / 新增的条目重新 raw-deflate + 重算 CRC；新增条目
 * 追加在已有 local record 之后，CD 顺序 = local 顺序。
 */
export function rewriteZipParts(
  bytes: Uint8Array,
  patches: readonly ZipPartPatch[],
): Uint8Array {
  const layout = layoutOf(bytes);
  const byPath = new Map<string, ZipPartPatch>();
  for (const patch of patches) {
    if (!patch.path) throw new ZipShapeError('part 补丁缺少 path');
    if (byPath.has(patch.path)) throw new ZipShapeError(`part 补丁重复：${patch.path}`);
    if (!patch.delete && !patch.data)
      throw new ZipShapeError(`part 补丁缺少内容：${patch.path}`);
    byPath.set(patch.path, patch);
  }
  // 目标必须存在（替换 / 删除）；不存在的路径按「新增」处理
  for (const [path, patch] of byPath) {
    const exists = layout.entries.some((entry) => entry.name === path);
    if (!exists && patch.delete)
      throw new ZipShapeError(`要删除的 part 不存在：${path}`);
  }

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

  // 追加的新条目（不在原 CD 里的路径），按 path 排序保证输出稳定
  const added: ZipPartPatch[] = [...byPath.values()]
    .filter((patch) => !ordered.some((entry) => entry.name === patch.path))
    .sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));

  const newOffsets = new Map<string, number>();
  const replaced = new Map<string, Uint8Array>();
  for (const [index, entry] of ordered.entries()) {
    const patch = byPath.get(entry.name);
    if (patch?.delete) continue;
    const end = index + 1 < ordered.length ? ordered[index + 1]!.localOffset : layout.cdStart;
    newOffsets.set(entry.name, offset);
    if (!patch?.data) {
      push(bytes.subarray(entry.localOffset, end));
      continue;
    }
    replaced.set(entry.name, patch.data);
    push(localRecordOf(entry, patch.data));
  }
  for (const patch of added) {
    newOffsets.set(patch.path, offset);
    replaced.set(patch.path, patch.data!);
    push(localRecordOf(newEntryOf(patch.path), patch.data!));
  }

  // central directory：未改条目原样拷贝（只改 localOffset），改动的条目重建。
  const cdChunks: Uint8Array[] = [];
  let cdSize = 0;
  for (const entry of ordered) {
    if (byPath.get(entry.name)?.delete) continue;
    const data = replaced.get(entry.name);
    const built = data
      ? cdRecordOf(entry, data, newOffsets.get(entry.name)!)
      : relocatedCdRecord(entry, newOffsets.get(entry.name)!);
    cdChunks.push(built);
    cdSize += built.byteLength;
  }
  for (const patch of added) {
    const built = cdRecordOf(newEntryOf(patch.path), patch.data!, newOffsets.get(patch.path)!);
    cdChunks.push(built);
    cdSize += built.byteLength;
  }
  const cdStart = offset;
  for (const chunk of cdChunks) push(chunk);

  const eocd = new Uint8Array(22 + layout.comment.byteLength);
  const view = viewOf(eocd);
  view.setUint32(0, SIG_EOCD, true);
  const finalCount =
    ordered.length - [...byPath.values()].filter((patch) => patch.delete).length + added.length;
  view.setUint16(8, finalCount, true);
  view.setUint16(10, finalCount, true);
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

/** 新增 part 的合成 CD 条目：只有名字 / 时间 / 属性，其余在写入时重算。 */
function newEntryOf(path: string): CdEntry {
  const nameBytes = new TextEncoder().encode(path);
  const now = new Date();
  return {
    name: path,
    versionMadeBy: 20,
    versionNeeded: 20,
    flags: 0,
    method: 8,
    modTime: (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1),
    modDate:
      ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate(),
    crc: 0,
    compressedSize: 0,
    uncompressedSize: 0,
    internalAttributes: 0,
    externalAttributes: 0,
    localOffset: 0,
    cdRecord: new Uint8Array(0),
    nameBytes,
  };
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
