// ─── 命令批 → 读侧模型 ─────────────────────────────────────────────────
//
// 人类（页面工具栏 / 键盘）与 agent 跑**同一套命令**（契约 G3）。页面在自己的
// 模型上应用，执行点在 headless 的模型上应用；随后都走同一个
// `planFromModel` → `SavePlan`，因此「页面路径与执行点路径产同一计划」可以
// 用单测钉住（契约 §14.1 验收 2）。
//
// 批级原子（`isolation: 'batch'`）：任一条非法 → 抛 `CommandRejected`，
// 调用方（页面 / engine.plan）把它变成 `rejected[]`，整批不落盘。

import type { DocxCommand, ModelBlock } from './types';
import { CommandRejected as Rejection } from './types';
import { blockFromXml as modelBlockFromXml } from './document';
import {
  applyRunStyle,
  patchParagraphProps,
  patchRunProps,
  paragraphHeadingOf,
  paragraphTextOf,
  renderParagraph,
  replaceRunText,
  rewriteParagraphText,
} from './xml-ops';

export interface ApplyResult {
  model: ModelBlock[];
  /** 成功应用的命令条数（批级原子时要么全成、要么抛错）。 */
  applied: number;
}

export function applyCommands(
  baseline: readonly ModelBlock[],
  commands: readonly DocxCommand[],
): ApplyResult {
  const model = baseline.map((block) => ({ ...block }));
  let applied = 0;
  commands.forEach((command, batchIndex) => {
    try {
      applyCommand(model, command);
      applied += 1;
    } catch (error) {
      if (error instanceof Rejection) {
        throw new Rejection(batchIndex, error.code, error.message);
      }
      throw error;
    }
  });
  return { model, applied };
}

function applyCommand(model: ModelBlock[], command: DocxCommand): void {
  switch (command.op) {
    case 'insertParagraph': {
      const heading = normalizeHeading(command.heading);
      const xml = renderParagraph(String(command.text ?? ''), heading);
      const block = modelBlockFromXml(xml, null);
      const at =
        command.after === null || command.after === undefined
          ? 0
          : indexOf(model, command.after) + 1;
      block.docxIndex = nextLocalIndex(model);
      model.splice(at, 0, block);
      return;
    }
    case 'setText': {
      const target = requireBlock(model, command.index);
      requireParagraph(target, 'setText');
      const text = String(command.text ?? '');
      const { start, end } = command;
      if (start === undefined && end === undefined) {
        target.xml = rewriteParagraphText(target.xml, text);
      } else {
        const from = Math.trunc(Number(start));
        const to = Math.trunc(Number(end));
        if (!Number.isFinite(from) || !Number.isFinite(to))
          throw reject('command.invalid', 'setText 的 start / end 必须是数字且成对出现');
        if (from < 0 || from > to || to > target.text.length)
          throw reject(
            'command.invalid',
            `setText 区间非法：${from}..${to}（段长 ${target.text.length}）`,
          );
        target.xml = replaceRunText(target.xml, from, to, text);
      }
      retargetText(target);
      return;
    }
    case 'setHeadingLevel': {
      const target = requireBlock(model, command.index);
      requireParagraph(target, 'setHeadingLevel');
      const level = normalizeHeading(command.level);
      target.xml = patchParagraphProps(target.xml, { heading: level });
      target.heading = paragraphHeadingOf(target.xml);
      return;
    }
    case 'setTextStyle': {
      const target = requireBlock(model, command.index);
      requireParagraph(target, 'setTextStyle');
      const { align, start, end, ...style } = command;
      let xml = target.xml;
      // `start` / `end` 是段内字符区间：两者都缺省 → 整段（旧行为）。
      if (
        style.bold !== undefined ||
        style.italic !== undefined ||
        style.underline !== undefined ||
        style.size !== undefined ||
        style.color !== undefined
      ) {
        if (start === undefined && end === undefined) {
          xml = patchRunProps(xml, style);
        } else {
          const from = Math.trunc(Number(start));
          const to = Math.trunc(Number(end));
          if (!Number.isFinite(from) || !Number.isFinite(to))
            throw reject('command.invalid', 'setTextStyle 的 start / end 必须是数字且成对出现');
          if (from < 0 || from >= to || to > target.text.length)
            throw reject(
              'command.invalid',
              `setTextStyle 区间非法：${from}..${to}（段长 ${target.text.length}）`,
            );
          xml = applyRunStyle(xml, from, to, style);
        }
      }
      if (align !== undefined) xml = patchParagraphProps(xml, { align });
      target.xml = xml;
      return;
    }
    case 'deleteBlock': {
      const target = requireBlock(model, command.index);
      if (target.hidden) throw reject('block.hidden', '尾部不可见节点不能删除');
      model.splice(indexOf(model, command.index), 1);
      return;
    }
    case 'moveBlock': {
      const target = requireBlock(model, command.index);
      if (target.hidden) throw reject('block.hidden', '尾部不可见节点不能移动');
      const from = indexOf(model, command.index);
      if (command.to === command.index) return;
      const [moved] = model.splice(from, 1);
      const at =
        command.to === null || command.to === undefined
          ? 0
          : model.findIndex((block) => block.docxIndex === command.to) + 1;
      if (at <= 0) throw reject('block.not-found', `目标位置不存在：${String(command.to)}`);
      model.splice(at > model.length ? model.length : at, 0, moved!);
      return;
    }
    case 'findReplace': {
      const find = String(command.find ?? '');
      if (!find) throw reject('command.invalid', 'findReplace 的 find 不能为空');
      const replace = String(command.replace ?? '');
      let replaced = 0;
      for (const block of model) {
        if (block.hidden || block.kind !== 'p') continue;
        if (!block.text.includes(find)) continue;
        const next = command.all === false
          ? block.text.replace(find, replace)
          : block.text.split(find).join(replace);
        if (next === block.text) continue;
        block.xml = rewriteParagraphText(block.xml, next);
        retargetText(block);
        replaced += 1;
        if (command.all === false) break;
      }
      if (replaced === 0)
        throw reject('command.not-found', `没有找到：${find}`);
      return;
    }
    case 'check':
      return;
    default: {
      const unknown = command as { op?: string };
      throw reject('command.invalid', `未知命令：${String(unknown.op)}`);
    }
  }
}

function retargetText(block: ModelBlock): void {
  block.text = paragraphTextOf(block.xml);
  block.heading = paragraphHeadingOf(block.xml);
}

/** 页面内新块的临时坐标：取模型里最小的负下标再 -1（不会与基线冲突）。 */
function nextLocalIndex(model: readonly ModelBlock[]): number {
  let min = 0;
  for (const block of model)
    if (typeof block.docxIndex === 'number' && block.docxIndex < min)
      min = block.docxIndex;
  return min - 1;
}

function normalizeHeading(value: unknown): number {
  const level = Math.trunc(Number(value ?? 0));
  if (!Number.isFinite(level)) return 0;
  return Math.max(0, Math.min(6, level));
}

function requireBlock(model: readonly ModelBlock[], index: unknown): ModelBlock {
  const target = model.find((block) => block.docxIndex === index);
  if (!target)
    throw reject('block.not-found', `块不存在：docxIndex=${String(index)}`);
  return target;
}

function requireParagraph(block: ModelBlock, op: string): void {
  if (block.hidden) throw reject('block.hidden', `不可见块不能 ${op}`);
  if (block.kind !== 'p')
    throw reject('block.not-paragraph', `${op} 只支持段落（当前 ${block.kind}）`);
}

function indexOf(model: readonly ModelBlock[], docxIndex: unknown): number {
  const index = model.findIndex((block) => block.docxIndex === docxIndex);
  if (index < 0)
    throw reject('block.not-found', `块不存在：docxIndex=${String(docxIndex)}`);
  return index;
}

function reject(code: string, message: string): Rejection {
  return new Rejection(-1, code, message);
}
