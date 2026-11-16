// ─── zip 部件指纹（页面保存前的能力闸门） ─────────────────────────────
//
// 执行点只重写 `word/document.xml`；其余条目原样拷贝。所以页面在提交计划前
// 必须确认「这次保存没有动到别的部件」——否则图片 / 批注 / 分节属性这类改动
// 会静默丢失（甚至产出引用不存在 media 的坏文件）。
//
// 这里只读中央目录（不解压、不建 zip 对象），拿到每个条目的 `name + crc + size`
// 即可判断部件是否变化。zip64 / 分卷返回 null（与 exec 的拒绝口径一致）。

import { unzipSync } from 'fflate';

const SIG_EOCD = 0x06054b50;
const SIG_CD = 0x02014b50;

function readU16(view: DataView, offset: number): number {
  return view.getUint16(offset, true);
}

function readU32(view: DataView, offset: number): number {
  return view.getUint32(offset, true);
}

/** `name → "crc:size"`；无法安全解析（zip64 / 分卷 / 损坏）返回 null。 */
export function zipPartFingerprints(bytes: Uint8Array): Map<string, string> | null {
  const length = bytes.byteLength;
  if (length < 22) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  const stop = Math.max(0, length - 22 - 0xffff);
  for (let pos = length - 22; pos >= stop; pos -= 1) {
    if (readU32(view, pos) !== SIG_EOCD) continue;
    if (pos + 22 + readU16(view, pos + 20) === length) {
      eocd = pos;
      break;
    }
  }
  if (eocd < 0) return null;
  if (readU16(view, eocd + 4) !== 0 || readU16(view, eocd + 6) !== 0) return null;
  const entryCount = readU16(view, eocd + 10);
  const cdSize = readU32(view, eocd + 12);
  const cdStart = readU32(view, eocd + 16);
  if (entryCount === 0xffff || cdSize === 0xffffffff || cdStart === 0xffffffff) return null;
  if (cdStart + cdSize > eocd) return null;

  const decoder = new TextDecoder();
  const parts = new Map<string, string>();
  let pos = cdStart;
  for (let index = 0; index < entryCount; index += 1) {
    if (pos + 46 > eocd || readU32(view, pos) !== SIG_CD) return null;
    const nameLength = readU16(view, pos + 28);
    const extraLength = readU16(view, pos + 30);
    const commentLength = readU16(view, pos + 32);
    const next = pos + 46 + nameLength + extraLength + commentLength;
    if (next > eocd) return null;
    const name = decoder.decode(bytes.subarray(pos + 46, pos + 46 + nameLength));
    const crc = readU32(view, pos + 16);
    const size = readU32(view, pos + 24);
    parts.set(name, `${crc}:${size}`);
    pos = next;
  }
  return parts;
}

export interface ZipPartChange {
  path: string;
  kind: 'replace' | 'add' | 'remove';
}

/** 基线与产出的部件差异（正文 `word/document.xml` 也在里面，调用方自行过滤）。 */
export function diffZipParts(
  baseline: Uint8Array,
  produced: Uint8Array,
): ZipPartChange[] | 'unreadable' {
  const before = zipPartFingerprints(baseline);
  const after = zipPartFingerprints(produced);
  if (!before || !after) return 'unreadable';
  const changes: ZipPartChange[] = [];
  for (const [path, fingerprint] of after) {
    const previous = before.get(path);
    if (previous === fingerprint) continue;
    changes.push({ path, kind: previous === undefined ? 'add' : 'replace' });
  }
  for (const path of before.keys()) if (!after.has(path)) changes.push({ path, kind: 'remove' });
  return changes.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0));
}

/** 从 zip 里读出指定条目的原始内容（只解压这些条目）。 */
export function readZipEntries(
  bytes: Uint8Array,
  paths: readonly string[],
): Map<string, Uint8Array> {
  const wanted = new Set(paths);
  const out = new Map<string, Uint8Array>();
  if (wanted.size === 0) return out;
  const archive = unzipSync(bytes, { filter: (file) => wanted.has(file.name) });
  for (const [path, data] of Object.entries(archive)) out.set(path, data);
  return out;
}

/** XML / rels 走文本补丁（可读、可断言）；二进制走 base64。 */
export function isTextPart(path: string): boolean {
  return /\.(xml|rels)$/i.test(path);
}

/**
 * 部件 → 用户能听懂的功能名（超限提示 / 诊断用）。
 * 未命中的部件按目录猜（`word/media/` → 图片…），最后兜底为原路径。
 */
export function describePart(path: string): string {
  if (path.startsWith('word/media/')) return '图片 / 媒体';
  if (path.startsWith('word/embeddings/')) return '嵌入对象';
  if (/^word\/header\d*\.xml$/.test(path)) return '页眉';
  if (/^word\/footer\d*\.xml$/.test(path)) return '页脚';
  if (/^word\/footnotes\.xml$/.test(path)) return '脚注';
  if (/^word\/endnotes\.xml$/.test(path)) return '尾注';
  if (/^word\/comments/.test(path)) return '批注';
  if (/^word\/numbering\.xml$/.test(path)) return '编号 / 列表定义';
  if (/^word\/styles\.xml$/.test(path)) return '样式表';
  if (/^word\/theme\//.test(path) || /^word\/theme\d*\.xml$/.test(path)) return '主题';
  if (/^word\/settings\.xml$/.test(path)) return '文档设置（页面颜色 / 兼容性等）';
  if (/^word\/_rels\//.test(path) || /^\[Content_Types\]\.xml$/.test(path)) return '关系 / 内容类型';
  if (/^word\/charts\//.test(path)) return '图表';
  if (/^word\/diagrams\//.test(path)) return 'SmartArt 图形';
  if (/^word\/drawings\//.test(path)) return '绘图';
  if (/^docProps\//.test(path)) return '文档属性';
  if (/^word\/(people|fontTable|webSettings)\.xml$/.test(path)) return '文档设置';
  return path;
}

/** `["word/media/a.png"]` → `图片 / 媒体（word/media/a.png）`，去重后逗号连接。 */
export function describeForeignParts(paths: readonly string[]): string {
  const seen = new Map<string, string[]>();
  for (const path of paths) {
    const label = describePart(path);
    const list = seen.get(label);
    if (list) list.push(path);
    else seen.set(label, [path]);
  }
  return [...seen.entries()]
    .map(([label, list]) => (label === list[0] ? label : `${label}（${list.join('、')}）`))
    .join('、');
}
