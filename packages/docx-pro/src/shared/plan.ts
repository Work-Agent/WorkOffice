// ─── 读侧模型 → SavePlan ───────────────────────────────────────────────
//
// docx 没有常驻写侧模型（契约 §4.1）：计划就是「最终顶层块序列」。
// 页面（PM 或本页的块模型）与执行点各自把模型映射成同一个形状：
//   - 与基线切片一致 → `{kind:'original', docxIndex}`（原字节回写）
//   - 改动 / 新增     → `{kind:'xml', xml}`（只重写这一块）
//
// 提交后 `Receipt.sync` 必然是 `manifest`（契约 §4.2.1 / §8.3）：一次成功
// 提交会重排 `docxIndex`，页面必须 rebase，否则下一份计划静默写错块。

import type { ParsedDocx } from './document';
import { tierFor } from './document';
import type { ModelBlock, Rejection, SaveBlock, SavePlan } from './types';
import { blockHash } from './sha256';
import { isWellFormedElement, paragraphTextOf } from './xml-ops';

export interface BaselineSlices {
  hashes: Map<number, string>;
  total: number;
  /** 基线尾部 sectPr 的 `docxIndex`（没有则 null）。 */
  trailingSectPrIndex: number | null;
  trailingSectPrHash: string | null;
}

export function baselineSlices(parsed: ParsedDocx): BaselineSlices {
  const hashes = new Map<number, string>();
  let trailingIndex: number | null = null;
  let trailingHash: string | null = null;
  for (const block of parsed.blocks) {
    const index = block.docxIndex!;
    const hash = blockHash(block.xml);
    hashes.set(index, hash);
    if (block.kind === 'sectPr' && block.hidden) {
      trailingIndex = index;
      trailingHash = hash;
    }
  }
  return {
    hashes,
    total: parsed.blocks.length,
    trailingSectPrIndex: trailingIndex,
    trailingSectPrHash: trailingHash,
  };
}

/** 模型 → SavePlan（页面与执行点共用）。 */
export function planFromModel(
  model: readonly ModelBlock[],
  baseline: BaselineSlices,
  applied = 0,
): SavePlan {
  const saveBlocks: SaveBlock[] = [];
  const saveBlockIndexByDocx: Record<number, number> = {};
  const present = new Set<number>();
  let changedCount = 0;

  for (const block of model) {
    const docxIndex = block.docxIndex;
    if (
      docxIndex !== null &&
      baseline.hashes.get(docxIndex) === blockHash(block.xml)
    ) {
      saveBlockIndexByDocx[docxIndex] = saveBlocks.length;
      saveBlocks.push({ kind: 'original', docxIndex });
      present.add(docxIndex);
      continue;
    }
    if (docxIndex !== null) saveBlockIndexByDocx[docxIndex] = saveBlocks.length;
    saveBlocks.push({
      kind: 'xml',
      xml: block.xml,
      ...(docxIndex !== null ? { docxIndex } : {}),
    });
    if (docxIndex !== null) present.add(docxIndex);
    changedCount += 1;
  }

  // 尾部 sectPr 必须留在最后：命令词汇不允许删 / 移动它；这里再兜一次底，
  // 因为它一旦丢失，整份文档的分节（页面尺寸 / 页眉页脚）都会失效。
  let trailingSectPrChange = false;
  const trailingIndex = baseline.trailingSectPrIndex;
  if (trailingIndex !== null) {
    const tail = saveBlocks[saveBlocks.length - 1];
    const intact =
      tail?.kind === 'original' && tail.docxIndex === trailingIndex;
    if (!intact) {
      // 若模型仍带着这个块但被改写过，用原始切片替换最后一项之外的同位项。
      for (let i = saveBlocks.length - 1; i >= 0; i -= 1) {
        const entry = saveBlocks[i]!;
        if (entry.kind === 'xml' && entry.docxIndex === trailingIndex)
          saveBlocks.splice(i, 1);
      }
      saveBlockIndexByDocx[trailingIndex] = saveBlocks.length;
      saveBlocks.push({ kind: 'original', docxIndex: trailingIndex });
      present.add(trailingIndex);
      trailingSectPrChange = true;
    }
  }

  let deletedCount = 0;
  for (const index of baseline.hashes.keys())
    if (!present.has(index)) deletedCount += 1;

  return {
    kind: 'save',
    saveBlocks,
    saveBlockIndexByDocx,
    changedCount,
    deletedCount,
    totalOriginal: baseline.total,
    trailingSectPrChange,
    applied,
  };
}

/** 页面侧撤销 = 模型历史，不需要逆计划；历史版本面板用这份「全 original」。 */
export function inversePlan(parsed: ParsedDocx): SavePlan {
  const saveBlocks: SaveBlock[] = parsed.blocks.map((block) => ({
    kind: 'original',
    docxIndex: block.docxIndex!,
  }));
  const saveBlockIndexByDocx: Record<number, number> = {};
  parsed.blocks.forEach((block, index) => {
    saveBlockIndexByDocx[block.docxIndex!] = index;
  });
  return {
    kind: 'save',
    saveBlocks,
    saveBlockIndexByDocx,
    changedCount: 0,
    deletedCount: 0,
    totalOriginal: parsed.blocks.length,
    trailingSectPrChange: false,
    applied: 0,
  };
}

/** 只读诊断（`check`）：不产字节、不落盘。 */
export function checkPlan(parsed: ParsedDocx): SavePlan {
  const paragraphs = parsed.blocks.filter((block) => block.kind === 'p');
  const tables = parsed.blocks.filter((block) => block.kind === 'tbl');
  const headings = paragraphs
    .filter((block) => block.heading > 0)
    .slice(0, 200)
    .map((block) => ({ level: block.heading, text: block.text.slice(0, 160) }));
  const emptyParagraphs = paragraphs.filter(
    (block) => paragraphTextOf(block.xml).trim().length === 0,
  ).length;
  return {
    kind: 'check',
    changedCount: 0,
    deletedCount: 0,
    totalOriginal: parsed.blocks.length,
    trailingSectPrChange: false,
    applied: 0,
    audit: {
      blocks: parsed.blocks.length,
      paragraphs: paragraphs.length,
      tables: tables.length,
      headings,
      emptyParagraphs,
      textLength: parsed.textLength,
      weight: parsed.weight,
      tier: tierFor(parsed.weight),
    },
  };
}

/** 被拒的批：apply 不落盘，回执精确到命令下标（契约 §9.1）。 */
export function rejectedPlan(
  rejection: Rejection,
  parsed: ParsedDocx | null,
): SavePlan {
  return {
    kind: 'reject',
    changedCount: 0,
    deletedCount: 0,
    totalOriginal: parsed?.blocks.length ?? 0,
    trailingSectPrChange: false,
    applied: 0,
    rejected: [
      {
        index: rejection.index,
        code: rejection.code,
        message: rejection.message,
      },
    ],
  };
}

export class PlanShapeError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PlanShapeError';
  }
}

const MAX_PLAN_BLOCKS = 250_000;

/** 计划自足性校验（页面提交的计划不可信，exec 落盘前必须过一遍）。 */
export function assertPlanShape(
  value: unknown,
  baseline: BaselineSlices,
): SavePlan {
  if (!value || typeof value !== 'object')
    throw new PlanShapeError('计划不是对象');
  const plan = value as Partial<SavePlan>;
  if (plan.kind !== 'save' && plan.kind !== 'check' && plan.kind !== 'reject')
    throw new PlanShapeError(`未知计划类型：${String(plan.kind)}`);
  if (plan.kind !== 'save') return plan as SavePlan;
  const saveBlocks = plan.saveBlocks;
  if (!Array.isArray(saveBlocks) || saveBlocks.length === 0)
    throw new PlanShapeError('saveBlocks 为空');
  if (saveBlocks.length > MAX_PLAN_BLOCKS)
    throw new PlanShapeError(`saveBlocks 超过上限：${saveBlocks.length}`);
  for (const [position, block] of saveBlocks.entries()) {
    if (!block || typeof block !== 'object')
      throw new PlanShapeError(`saveBlocks[${position}] 不是对象`);
    if (block.kind === 'original') {
      if (typeof block.docxIndex !== 'number')
        throw new PlanShapeError(`saveBlocks[${position}] 缺少 docxIndex`);
      if (!baseline.hashes.has(block.docxIndex))
        throw new PlanShapeError(
          `saveBlocks[${position}] 引用了不存在的基线块：${block.docxIndex}`,
        );
      continue;
    }
    if (block.kind === 'xml') {
      if (typeof block.xml !== 'string')
        throw new PlanShapeError(`saveBlocks[${position}] 缺少 xml`);
      if (!isWellFormedElement(block.xml))
        throw new PlanShapeError(`saveBlocks[${position}] 不是良构的顶层元素`);
      continue;
    }
    throw new PlanShapeError(
      `saveBlocks[${position}] 的 kind 非法：${String((block as { kind?: unknown }).kind)}`,
    );
  }
  return plan as SavePlan;
}
