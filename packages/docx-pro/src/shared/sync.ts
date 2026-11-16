// ─── baseline sync（docx 独有，契约 I15 / §8.3） ───────────────────────
//
// 一次成功提交会重写 `word/document.xml`，顶层下标整体重排。页面手里的
// `docxIndex` 若不对齐，下一份计划会**静默写错块**——这不是 `command.stale`。
//
// exec 在产出字节的同一趟里产出 manifest（它已经 parse 过，边际成本≈0）；
// 页面按 `hash` 把已有切片对回新下标，只对「页面没有这段切片」的块接收 xml。
// 任一 `hash` 对不上 → 页面走全量重解析（`reload`），不猜。

import type { ParsedDocx } from './document';
import { blockFromXml } from './document';
import type { ModelBlock, SaveBlock, SavePlan, SyncManifest } from './types';
import { blockHash } from './sha256';
import { elementKindOf } from './xml-ops';

/** 计划 + 基线 → 提交后清单（由 exec 调用，`nextBlocks` 与 saveBlocks 等长）。 */
export function buildManifest(
  parsed: ParsedDocx,
  plan: SavePlan,
): SyncManifest {
  const baseline = new Map<number, ModelBlock>();
  for (const block of parsed.blocks) baseline.set(block.docxIndex!, block);
  const blocks: SyncManifest['blocks'] = [];
  const hidden: NonNullable<SyncManifest['hidden']> = [];

  plan.saveBlocks?.forEach((entry: SaveBlock, position: number) => {
    if (entry.kind === 'original') {
      const source = baseline.get(entry.docxIndex);
      if (!source) throw new Error(`manifest: 基线块不存在 ${entry.docxIndex}`);
      blocks.push({ docxIndex: position, hash: blockHash(source.xml) });
      if (source.kind === 'sectPr')
        hidden.push({ docxIndex: position, kind: 'sectPr' });
      return;
    }
    const hash = blockHash(entry.xml);
    blocks.push({ docxIndex: position, hash, xml: entry.xml });
    if (elementKindOf(entry.xml) === 'sectPr')
      hidden.push({ docxIndex: position, kind: 'sectPr' });
  });

  return hidden.length > 0 ? { blocks, hidden } : { blocks };
}

export interface ManifestApplyResult {
  model: ModelBlock[];
  /** true = 有块对不上，调用方应回退到全量重解析。 */
  mismatched: boolean;
}

/**
 * 页面应用 manifest：只更新基线映射与切片，不重建模型（会话内 id 保留）。
 *
 * 匹配顺序：
 * 1. 同下标且 hash 相同 → 复用（常见情况：整批都在原位）
 * 2. 任意未使用且 hash 相同 → 复用（块被移动 / 前后插入了别的块）
 * 3. manifest 带 xml → 新建块（新插入 / 被重写的块）
 * 4. 都没有 → 对不上，返回 mismatched
 */
export function applyManifest(
  model: readonly ModelBlock[],
  manifest: SyncManifest,
): ManifestApplyResult {
  const unused = new Set(model.map((_block, index) => index));
  const used = new Map<number, number>(); // manifest 内置：块序号 → 模型下标
  const hashes = model.map((block) => blockHash(block.xml));
  const next: ModelBlock[] = [];
  let mismatched = false;

  for (const [position, entry] of manifest.blocks.entries()) {
    const sameIndex = model[position];
    if (
      sameIndex &&
      unused.has(position) &&
      sameIndex.docxIndex === entry.docxIndex &&
      hashes[position] === entry.hash
    ) {
      unused.delete(position);
      used.set(position, position);
      next.push({ ...sameIndex, docxIndex: entry.docxIndex });
      continue;
    }
    const found = [...unused].find((index) => hashes[index] === entry.hash);
    if (found !== undefined) {
      unused.delete(found);
      const source = model[found]!;
      next.push({ ...source, docxIndex: entry.docxIndex });
      continue;
    }
    if (typeof entry.xml === 'string') {
      next.push(blockFromXml(entry.xml, entry.docxIndex, `m${entry.docxIndex}`));
      continue;
    }
    mismatched = true;
  }

  return { model: next, mismatched };
}

/** 计划 → 新顶层序列（exec 与测试共用；不做 zip 处理）。 */
export function spliceBody(parsed: ParsedDocx, plan: SavePlan): string {
  const baseline = new Map<number, ModelBlock>();
  for (const block of parsed.blocks) baseline.set(block.docxIndex!, block);
  let inner = '';
  for (const entry of plan.saveBlocks ?? []) {
    if (entry.kind === 'original') {
      const source = baseline.get(entry.docxIndex);
      if (!source) throw new Error(`计划引用了不存在的基线块：${entry.docxIndex}`);
      inner += source.xml;
      continue;
    }
    inner += entry.xml;
  }
  return `${parsed.beforeBody}${inner}${parsed.afterBody}`;
}
