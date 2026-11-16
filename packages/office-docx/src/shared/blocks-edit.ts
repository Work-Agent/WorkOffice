// ─── 跨块编辑 → 命令批（纯函数，页面与单测共用） ─────────────────────────
//
// 页面是一个块一个 contenteditable 宿主，浏览器的编辑流在块边界处被切断，
// 所以「段首退格合并 / 跨段选区被覆盖 / 多行粘贴」这几类原生意图由页面翻译成
// 命令批（`editor.ts` 只负责取 DOM 意图，这里只负责模型翻译）。
//
// 约束：产出的仍是同一套 `DocxCommand`，因此写盘路径完全不变——`planFromModel`
// 照样产出「最终顶层块序列」的 `SavePlan`，未改块仍是 `{kind:'original'}`。

import type { BlockRange, DocxCommand, ModelBlock } from './types';

/** 一次跨块编辑的翻译结果。 */
export type EditBatch =
  | {
      kind: 'apply';
      commands: DocxCommand[];
      /** 应用后光标应落在哪（已知目标块坐标时使用）。 */
      focus?: { index: number; offset: number };
      /** 未知坐标的插入：从 `anchor` 之后数第 `count` 段的 `offset` 处（粘贴用）。 */
      tail?: { anchor: number; count: number; offset: number };
    }
  /** 边界情况，无事可做（首段退格、末段 Delete…）。 */
  | { kind: 'skip' }
  /** 意图本身不支持（跨表格等），调用方负责提示。 */
  | { kind: 'refuse'; reason: string };

/** 一次粘贴最多展开成多少段：多余的行并入末段（软换行），避免一次批过大。 */
export const MAX_PASTE_BLOCKS = 200;

/** 段落合并：段首退格（back）把本段并入上一段，段尾 Delete（forward）并入下一段。 */
export function mergeCommands(
  model: readonly ModelBlock[],
  index: number,
  direction: 'back' | 'forward',
): EditBatch {
  const visible = model.filter((block) => !block.hidden);
  const position = visible.findIndex((block) => block.docxIndex === index);
  if (position < 0) return { kind: 'skip' };
  const current = visible[position]!;
  const neighbor = direction === 'back' ? visible[position - 1] : visible[position + 1];
  if (!neighbor) return { kind: 'skip' };
  if (current.kind !== 'p' || neighbor.kind !== 'p')
    return { kind: 'refuse', reason: '表格与其他结构块暂不支持段落合并。' };
  const back = direction === 'back';
  const target = back ? neighbor : current;
  const absorbed = back ? current : neighbor;
  if (typeof target.docxIndex !== 'number' || typeof absorbed.docxIndex !== 'number')
    return { kind: 'skip' };
  const head = back ? neighbor.text : current.text;
  const tail = back ? current.text : neighbor.text;
  return {
    kind: 'apply',
    commands: [
      { op: 'setText', index: target.docxIndex, text: head + tail },
      { op: 'deleteBlock', index: absorbed.docxIndex },
    ],
    // 光标停在接缝处：本段原文本与邻段文本之间。
    focus: { index: target.docxIndex, offset: head.length },
  };
}

/**
 * 跨段选区被输入 / 删除覆盖：首段留头部 + 新文本 + 末段尾部，中间段整体删除。
 * `text` 为空即纯删除（Backspace / Delete / 剪切）。
 */
export function rangeCommands(
  model: readonly ModelBlock[],
  range: BlockRange,
  text: string,
): EditBatch {
  const visible = model.filter((block) => !block.hidden);
  const fromPosition = visible.findIndex((block) => block.docxIndex === range.from);
  const toPosition = visible.findIndex((block) => block.docxIndex === range.to);
  if (fromPosition < 0 || toPosition < 0) return { kind: 'skip' };
  const start = Math.min(fromPosition, toPosition);
  const end = Math.max(fromPosition, toPosition);
  const span = visible.slice(start, end + 1);
  if (span.some((block) => block.kind !== 'p'))
    return { kind: 'refuse', reason: '选区跨越表格或其他结构块，暂不支持批量删除。' };
  const first = span[0]!;
  const last = span[span.length - 1]!;
  if (typeof first.docxIndex !== 'number' || typeof last.docxIndex !== 'number')
    return { kind: 'skip' };
  const forward = fromPosition <= toPosition;
  const head = first.text.slice(0, forward ? range.fromOffset : range.toOffset);
  const tail = last.text.slice(forward ? range.toOffset : range.fromOffset);
  const commands: DocxCommand[] = [
    { op: 'setText', index: first.docxIndex, text: head + text + tail },
  ];
  for (const block of span.slice(1)) {
    if (typeof block.docxIndex !== 'number') continue;
    commands.push({ op: 'deleteBlock', index: block.docxIndex });
  }
  return {
    kind: 'apply',
    commands,
    focus: { index: first.docxIndex, offset: head.length + text.length },
  };
}

/**
 * 多行粘贴：每行一段（Word「只保留文本」）。首行接光标前的内容（`currentText`
 * 的 `offset` 之前），末行接光标后的内容；新段沿用锚点段的大纲级别。
 */
export function pasteCommands(
  model: readonly ModelBlock[],
  index: number,
  currentText: string,
  offset: number,
  lines: readonly string[],
): EditBatch {
  if (lines.length === 0) return { kind: 'skip' };
  const anchor = model.find((block) => block.docxIndex === index);
  if (!anchor) return { kind: 'skip' };
  if (anchor.kind !== 'p')
    return { kind: 'refuse', reason: '只能在段落内粘贴多行文本。' };
  const kept = lines.slice(0, MAX_PASTE_BLOCKS);
  const rest = lines.slice(MAX_PASTE_BLOCKS);
  if (rest.length > 0)
    kept[kept.length - 1] = `${kept[kept.length - 1]}\n${rest.join('\n')}`;
  const clamp = Math.max(0, Math.min(currentText.length, offset));
  const blocks = [...kept];
  blocks[0] = currentText.slice(0, clamp) + blocks[0]!;
  blocks[blocks.length - 1] = `${blocks[blocks.length - 1]}${currentText.slice(clamp)}`;
  const commands: DocxCommand[] = [{ op: 'setText', index, text: blocks[0]! }];
  // 逆序插入：每条都插在锚点块之后，最终顺序与 blocks 一致，且锚点坐标始终有效。
  for (let line = blocks.length - 1; line >= 1; line -= 1)
    commands.push({ op: 'insertParagraph', after: index, text: blocks[line]!, heading: anchor.heading });
  return {
    kind: 'apply',
    commands,
    tail: {
      anchor: index,
      count: blocks.length,
      offset: (blocks[blocks.length - 1] ?? '').length,
    },
  };
}
