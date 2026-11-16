// ─── 页面侧计划生成（P2） ────────────────────────────────────────────
//
// 渲染器的保存流水线（`genoffice/editor/file-actions.ts` 的 `buildDocBytes`）
// 已经能产出「新 document.xml」的完整字节（`@genoffice/docx-engine` 的 `saveDocx`）。
// 而本包执行点只接受 `{kind:'original'|'xml'}` 的块序列 —— 引擎 `SaveBlock` 里的
// `generated`（结构化块）需要引擎内部上下文才能序列化，页面拿不到。
//
// 所以这里用**块级 diff** 把新字节还原成我们的 `SavePlan`：
//   新 document.xml 的顶层切片  ↔  基线切片（按哈希、保持顺序匹配）
//     命中 → `{kind:'original', docxIndex}`（exec 用基线切片原样回写，字节保真）
//     未命中 → `{kind:'xml', xml}`（exec 原样写入）
// 结果：exec 产出的 document.xml 与引擎产出一致，只差块间空白（既有归一化）。
//
// 代价：页面仍跑一次 `saveDocx`（只为拿到块的最终 XML）。P5 的 part 通道与
// `generated → xml` 直出可以去掉这趟重复序列化；届时本文件只剩「块级 diff」的兜底。

import { parseDocxBytes } from './document';
import { blockHash } from './sha256';
import { diffZipParts, describePart, isTextPart, readZipEntries } from './zip-parts';
import { MAX_PART_BYTES, MAX_PARTS_BYTES } from './plan';
import type { ModelBlock, SaveBlock, SavePartPatch, SavePlan } from './types';

/** 我们的 `SaveBlock`（`original` + `xml` 两种，执行点只认这两种）。 */
function original(docxIndex: number): SaveBlock {
  return { kind: 'original', docxIndex };
}

export interface BuildPlanInput {
  /** 打开 / 上次提交后的基线字节。 */
  baseline: Uint8Array;
  /** 本次 `saveDocx` 产出的新字节。 */
  produced: Uint8Array;
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = '';
  const chunk = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunk)
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunk));
  return btoa(binary);
}

export interface BuildPlanStats {
  total: number;
  kept: number;
  changed: number;
  deleted: number;
  /** 正文以外的 part 补丁数（P5）。 */
  parts: number;
}

export type BuildPlanResult =
  | { ok: true; plan: SavePlan; stats: BuildPlanStats }
  | { ok: false; error: string };

/** 哈希 → 基线里出现的位置（升序），用于按顺序匹配、同时识别删除。 */
function indexByHash(blocks: readonly ModelBlock[]): Map<string, number[]> {
  const index = new Map<string, number[]>();
  for (const [position, block] of blocks.entries()) {
    const hash = blockHash(block.xml);
    const list = index.get(hash);
    if (list) list.push(position);
    else index.set(hash, [position]);
  }
  return index;
}

/** 在 `position >= from` 的候选里取最小的那个（线性扫，候选通常 1 个）。 */
function nextPosition(candidates: readonly number[] | undefined, from: number): number | null {
  if (!candidates) return null;
  for (const position of candidates) if (position >= from) return position;
  return null;
}

/** 把「改写」的块补上它继承的基线 `docxIndex`（仅诊断用，exec 只读 xml）。 */
function replacesAt(blocks: SaveBlock[], position: number, docxIndex: number): void {
  const block = blocks[position];
  if (block?.kind === 'xml' && block.docxIndex === undefined) block.docxIndex = docxIndex;
}

/** 诊断用：`4836B:504b0304`（长度 + 前 4 字节魔数；PK\x03\x04 = 合法 zip）。 */
function describeBytes(bytes: Uint8Array): string {
  const head = Array.from(bytes.subarray(0, 4))
    .map((byte) => byte.toString(16).padStart(2, '0'))
    .join('');
  return `${bytes.byteLength}B:${head || 'empty'}`;
}

export function planFromBytes(input: BuildPlanInput): BuildPlanResult {
  let base;
  try {
    base = parseDocxBytes(input.baseline);
  } catch (error) {
    const detail = `基线不可解析（${describeBytes(input.baseline)}）：${error instanceof Error ? error.message : String(error)}`;
    console.error('[office-docx] plan from baseline failed', detail, input.baseline);
    return { ok: false, error: detail };
  }
  let next;
  try {
    next = parseDocxBytes(input.produced);
  } catch (error) {
    const detail = `序列化产物不可解析（${describeBytes(input.produced)}）：${error instanceof Error ? error.message : String(error)}`;
    console.error('[office-docx] plan from produced failed', detail, input.produced);
    return { ok: false, error: detail };
  }

  // body 外层（`<w:document>` 开标签 / `</w:body>` 之后）的变化无法用「块序列」表达：
  // 分节属性、页面设置、命名空间声明都必须走 part 通道（P5）。
  if (base.beforeBody !== next.beforeBody || base.afterBody !== next.afterBody) {
    return {
      ok: false,
      error: '正文外层结构发生变化（分节 / 页面设置 / 命名空间）：本版暂不支持保存，请在 Word 中修改',
    };
  }
  if (next.blocks.length === 0) return { ok: false, error: '序列化结果为空，已取消保存' };

  // 正文以外的 part（页眉页脚 / 脚注尾注 / 批注 / 编号 / 主题 / 图片与 rels）
  // 走整包补丁（P5）：把产出里**变化过的条目**原样搬进计划，执行点负责写入。
  const changes = diffZipParts(input.baseline, input.produced);
  if (changes === 'unreadable')
    return { ok: false, error: '无法读取 zip 目录（zip64 / 分卷 / 损坏）：本版不支持保存该文档' };
  const nonBody = changes.filter((change) => change.path !== 'word/document.xml');
  const contents = readZipEntries(
    input.produced,
    nonBody.filter((change) => change.kind !== 'remove').map((change) => change.path),
  );
  const parts: SavePartPatch[] = [];
  let partBytes = 0;
  for (const change of nonBody) {
    if (change.kind === 'remove') {
      parts.push({ path: change.path, delete: true });
      continue;
    }
    const data = contents.get(change.path);
    if (!data)
      return { ok: false, error: `无法从产出里读出 part：${change.path}` };
    if (data.byteLength > MAX_PART_BYTES)
      return {
        ok: false,
        error: `${describePart(change.path)} 太大（${Math.round(data.byteLength / 1024 / 1024)} MB）：单次保存上限 ${Math.round(MAX_PART_BYTES / 1024 / 1024)} MB`,
      };
    partBytes += data.byteLength;
    if (partBytes > MAX_PARTS_BYTES)
      return {
        ok: false,
        error: `本次保存涉及的其他内容超过 ${Math.round(MAX_PARTS_BYTES / 1024 / 1024)} MB 上限（图片 / 附件过多）`,
      };
    parts.push(
      isTextPart(change.path)
        ? { path: change.path, xml: new TextDecoder().decode(data) }
        : { path: change.path, base64: bytesToBase64(data) },
    );
  }

  const positions = indexByHash(base.blocks);
  const saveBlocks: SaveBlock[] = [];
  const saveBlockIndexByDocx: Record<number, number> = {};
  const matches: Array<number | null> = [];
  let cursor = 0;
  let kept = 0;
  let changed = 0;

  for (const block of next.blocks) {
    const matched = nextPosition(positions.get(blockHash(block.xml)), cursor);
    if (matched !== null) {
      const docxIndex = base.blocks[matched]!.docxIndex;
      if (typeof docxIndex === 'number') {
        saveBlockIndexByDocx[docxIndex] = saveBlocks.length;
        saveBlocks.push(original(docxIndex));
        matches.push(matched);
        cursor = matched + 1;
        kept += 1;
        continue;
      }
    }
    // 未命中的块（新增 / 改写）：整段 XML 交给 exec，字节原样写入。
    matches.push(null);
    saveBlocks.push({ kind: 'xml', xml: block.xml });
    changed += 1;
  }

  // 区分「改写」与「删除」：一个未命中的块如果身前还有没被命中消耗的基线块，
  // 就地接管一个（视为改写，docxIndex 换了内容）；否则是纯新增。剩下的基线块
  // 才是真正被删的（sectPr 除外：它总在末尾兜底，不参与计数）。
  let baselineCursor = 0;
  let replaced = 0;
  for (const [position, matched] of matches.entries()) {
    if (matched !== null) {
      baselineCursor = Math.max(baselineCursor, matched + 1);
      continue;
    }
    let nextMatched = base.blocks.length;
    for (let ahead = position + 1; ahead < matches.length; ahead += 1) {
      const value = matches[ahead];
      if (value !== null) {
        nextMatched = value;
        break;
      }
    }
    if (baselineCursor < nextMatched) {
      replacesAt(saveBlocks, position, baselineCursor);
      baselineCursor += 1;
      replaced += 1;
    }
  }

  const deleted = Math.max(0, base.blocks.length - kept - replaced);
  return {
    ok: true,
    stats: { total: next.blocks.length, kept, changed, deleted, parts: parts.length },
    plan: {
      kind: 'save',
      saveBlocks,
      saveBlockIndexByDocx,
      changedCount: changed,
      deletedCount: deleted,
      totalOriginal: base.blocks.length,
      trailingSectPrChange: false,
      ...(parts.length > 0 ? { parts } : {}),
    },

  };
}

/**
 * 两份字节的**块切片**是否一致（不看块间空白 / 其它部件）。
 * 用于提交后确认「磁盘上的已经是这次提交的内容」——写入是异步落盘的，
 * 立刻回读可能拿到旧内容或半截文件，那会让渲染器重解析回旧文档。
 */
export function sameBlockSlices(left: Uint8Array, right: Uint8Array): boolean {
  let a;
  let b;
  try {
    a = parseDocxBytes(left);
    b = parseDocxBytes(right);
  } catch {
    return false;
  }
  if (a.blocks.length !== b.blocks.length) return false;
  for (const [index, block] of a.blocks.entries()) {
    const other = b.blocks[index]!;
    if (block.docxIndex !== other.docxIndex) return false;
    if (blockHash(block.xml) !== blockHash(other.xml)) return false;
  }
  return true;
}
