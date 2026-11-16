// ─── agent 页面代理（P2b） ───────────────────────────────────────────
//
// 宿主（WorkStudio）用 bus 的 `agent-request` 让页面扮演「该 ref 的唯一编辑模型」：
//   kind:'context' → 回 live 投影（与 `shared/project.ts` 同一份格式）
//   kind:'apply'   → 在实时 PM 文档上应用命令，flush 成 plan 提交，再回执
//
// 命令词汇是本包的 `DocxCommand`（与 engine 侧同一套）。落地分两条路：
//   - 文本 / 结构（setText / insertParagraph / deleteBlock / moveBlock / check）：
//     直接走 ProseMirror 事务（渲染器的 schema：docParagraph / docHeading / docTable…）。
//   - 格式（setTextStyle / setHeadingLevel / findReplace）：
//     复用渲染器自带的 ops 引擎（`ai/ops.ts` 的 executeOps），它已处理 marks、
//     段落属性、表格单元格与撤销集成，比手写稳。
//
// 坐标：对外只用 `docxIndex`（PM 顶层节点的 attrs.docxIndex），与 `[n]` 投影一致；
// 落到 PM 前再映射成顶层序号 / 位置。批次先整体校验，任何一条非法就整批不执行。

import type { Editor } from '@tiptap/core';
import type { Node as PmNode } from '@tiptap/pm/model';
import { executeOps } from './genoffice/editor/ai/ops';
import { unwrapAgentCommands } from '../shared';
import { sessionState } from './host';
import type { DocxCommand } from '../shared/types';
import type { BusMessage, Receipt } from '../page/env';

const MAX_PROJECT_BLOCKS = 400;
const MAX_PROJECT_TEXT = 160;

interface AiDocsBridge {
  editor?: Editor | null;
  save?: () => Promise<boolean>;
  getStatus?: () => string;
}

function bridge(): AiDocsBridge | undefined {
  return (window as unknown as { __aidocs?: AiDocsBridge }).__aidocs;
}

function editorNow(): Editor | null {
  const current = bridge()?.editor;
  if (!current || current.isDestroyed) return null;
  return current;
}

/** 等渲染器把编辑器挂出来（页面刚 loaded 时 agent 可能立刻来要 context）。 */
export async function waitForEditor(timeoutMs = 15_000): Promise<Editor | null> {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    const current = editorNow();
    if (current) return current;
    if (Date.now() > deadline) return null;
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
}

function coded(code: string, message: string): Error {
  const error = new Error(message) as Error & { code: string };
  error.code = code;
  return error;
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function errorCodeOf(error: unknown): string | undefined {
  const code = (error as { code?: unknown })?.code;
  return typeof code === 'string' ? code : undefined;
}

// ─── 坐标 ─────────────────────────────────────────────────────────────

interface BlockRef {
  /** PM 顶层序号（0 基，含表格等所有顶层节点）。 */
  pmIndex: number;
  /** 该节点在文档中的起始位置（等于 `node.nodeSize` 的前缀和）。 */
  pos: number;
  node: PmNode;
  docxIndex: number | null;
}

function blocksOf(editor: Editor): BlockRef[] {
  const refs: BlockRef[] = [];
  let pos = 0;
  editor.state.doc.forEach((node) => {
    const raw = node.attrs['docxIndex'];
    refs.push({
      pmIndex: refs.length,
      pos,
      node,
      docxIndex: typeof raw === 'number' ? raw : null,
    });
    pos += node.nodeSize;
  });
  return refs;
}

function refOf(refs: readonly BlockRef[], docxIndex: number): BlockRef | undefined {
  return refs.find((ref) => ref.docxIndex === docxIndex);
}

/** 段内字符偏移 → PM 文档位置（只统计文本，与 `paragraphTextOf` 同口径）。 */
function charToPos(block: BlockRef, offset: number): number {
  let remaining = Math.max(0, offset);
  let result: number | null = null;
  block.node.descendants((child, childPos) => {
    if (result !== null) return false;
    if (!child.isText) return true;
    const length = child.text?.length ?? 0;
    if (remaining <= length) {
      result = block.pos + 1 + childPos + remaining;
      return false;
    }
    remaining -= length;
    return true;
  });
  return result ?? block.pos + 1 + block.node.content.size;
}

// ─── 投影（与 shared/project.ts 同格式，只是数据来自 live PM 文档） ────

function oneLine(text: string): string {
  const flat = text.replace(/\s+/g, ' ').trim();
  return flat.length > MAX_PROJECT_TEXT ? `${flat.slice(0, MAX_PROJECT_TEXT)}…` : flat;
}

export interface ProjectMeta {
  name: string;
  version: string;
  epoch: number;
  tier?: string;
}

export function projectLive(editor: Editor, meta: ProjectMeta): string {
  const refs = blocksOf(editor);
  const tables = refs.filter((ref) => ref.node.type.name === 'docTable').length;
  const header = [
    `# ${meta.name}`,
    `blocks=${refs.length}`,
    tables > 0 ? `tables=${tables}` : '',
    meta.version ? `version=${meta.version.slice(0, 8)}` : '',
    `epoch=${meta.epoch}`,
    meta.tier ? `tier=${meta.tier}` : '',
    'engine=js sync=manifest',
  ]
    .filter(Boolean)
    .join(' ');
  const lines: string[] = [header];
  for (const ref of refs.slice(0, MAX_PROJECT_BLOCKS)) {
    const label = `[${ref.docxIndex ?? '?'}]`;
    const name = ref.node.type.name;
    if (name === 'docHeading') {
      const level = Number(ref.node.attrs['level']) || 1;
      lines.push(`${label} h${Math.min(6, Math.max(1, level))} ${oneLine(ref.node.textContent)}`);
      continue;
    }
    if (name === 'docTable') {
      const rows = ref.node.childCount;
      const cols = ref.node.firstChild?.childCount ?? 0;
      lines.push(`${label} table ${rows}x${cols}`);
      continue;
    }
    if (name === 'docParagraph' || name === 'docListItem') {
      lines.push(`${label} p ${oneLine(ref.node.textContent)}`);
      continue;
    }
    lines.push(`${label} ${name}`);
  }
  if (refs.length > MAX_PROJECT_BLOCKS)
    lines.push(
      `… 还有 ${refs.length - MAX_PROJECT_BLOCKS} 个块未列出（用 findReplace 或分段读取）`,
    );
  lines.push(
    '# 用 docx_apply(ref,{commands:[…]}) 修改；[n] 即命令的 index；先 guide 再写；写前 dryRun',
  );
  return lines.join('\n');
}

// ─── 命令校验（整批原子：先全验，再执行） ─────────────────────────────

function isTextBlock(node: PmNode): boolean {
  return node.type.name === 'docParagraph' || node.type.name === 'docHeading' || node.type.name === 'docListItem';
}

function validateCommand(refs: readonly BlockRef[], command: DocxCommand): string | null {
  switch (command.op) {
    case 'check':
      return null;
    case 'insertParagraph':
      if (command.after !== null && command.after !== undefined) {
        const anchor = refOf(refs, command.after);
        if (!anchor) return `insertParagraph：after=${command.after} 不存在`;
      }
      return null;
    case 'findReplace':
      if (!String(command.find ?? '')) return 'findReplace：find 不能为空';
      return null;
    case 'setText':
    case 'setTextStyle':
    case 'setHeadingLevel':
    case 'deleteBlock':
    case 'moveBlock': {
      const block = refOf(refs, command.index);
      if (!block) return `${command.op}：index=${command.index} 不存在`;
      if (command.op === 'moveBlock') {
        if (command.to !== null && command.to !== undefined && !refOf(refs, command.to))
          return `moveBlock：to=${command.to} 不存在`;
        if (!isTextBlock(block.node) && block.node.type.name !== 'docTable')
          return `moveBlock：index=${command.index} 不是可移动的块`;
        return null;
      }
      if (!isTextBlock(block.node))
        return `${command.op}：index=${command.index} 不是段落（${block.node.type.name}）`;
      if (command.op === 'setText' || command.op === 'setTextStyle') {
        const length = block.node.textContent.length;
        const hasRange = command.start !== undefined || command.end !== undefined;
        if (!hasRange) return null;
        const from = Math.trunc(Number(command.start));
        const to = Math.trunc(Number(command.end));
        if (!Number.isFinite(from) || !Number.isFinite(to))
          return `${command.op}：start / end 必须是数字且成对出现`;
        if (from < 0 || from > length || to > length || (command.op === 'setTextStyle' && from >= to))
          return `${command.op}：区间 ${from}..${to} 越界（段长 ${length}）`;
      }
      return null;
    }
    default: {
      const unknown = command as { op?: string };
      return `未知命令：${String(unknown.op)}`;
    }
  }
}

// ─── 执行 ─────────────────────────────────────────────────────────────

export interface ApplyOptions {
  dryRun?: boolean;
}

export interface ApplyResult {
  ok: boolean;
  applied: number;
  error?: string;
}

/** 走渲染器 ops 引擎的格式 / 选择器命令。 */
function engineOp(command: DocxCommand, editor: Editor): { op: Record<string, unknown>; ctx: Record<string, unknown> } | null {
  const refs = blocksOf(editor);
  switch (command.op) {
    case 'setHeadingLevel': {
      const block = refOf(refs, command.index);
      if (!block) return null;
      return {
        op: { op: 'setHeadingLevel', target: { blockIndexes: [block.pmIndex] }, level: command.level },
        ctx: {},
      };
    }
    case 'findReplace':
      return {
        op: {
          op: 'findReplace',
          find: command.find,
          replace: command.replace,
          matchCase: true,
        },
        ctx: {},
      };
    case 'setTextStyle': {
      const block = refOf(refs, command.index);
      if (!block) return null;
      const font: Record<string, unknown> = {};
      if (command.bold !== undefined) font['bold'] = command.bold;
      if (command.italic !== undefined) font['italic'] = command.italic;
      if (command.underline !== undefined) font['underline'] = command.underline;
      if (command.size !== undefined) font['fontSize'] = command.size;
      if (command.color !== undefined) font['color'] = `#${command.color.replace(/^#/, '')}`;
      const hasRange = command.start !== undefined && command.end !== undefined;
      if (hasRange) {
        const from = charToPos(block, Number(command.start));
        const to = charToPos(block, Number(command.end));
        return {
          op: { op: 'setFont', target: { scope: 'selection' }, ...font },
          ctx: {
            selection: {
              startIndex: block.pmIndex,
              endIndex: block.pmIndex,
              from,
              to,
            },
          },
        };
      }
      return {
        op: { op: 'setFont', target: { blockIndexes: [block.pmIndex] }, ...font },
        ctx: {},
      };
    }
    default:
      return null;
  }
}

/** 直接走 PM 事务的文本 / 结构命令。 */
function applyCustom(command: DocxCommand, editor: Editor): boolean {
  const refs = blocksOf(editor);
  const schema = editor.schema;
  const tr = editor.state.tr;
  switch (command.op) {
    case 'setText': {
      const block = refOf(refs, command.index);
      if (!block) return false;
      const text = String(command.text ?? '');
      const hasRange = command.start !== undefined || command.end !== undefined;
      if (!hasRange) {
        const marks = block.node.firstChild?.marks ?? [];
        const content = text.length > 0 ? schema.text(text, marks) : null;
        tr.replaceWith(block.pos + 1, block.pos + block.node.nodeSize - 1, content ?? []);
        break;
      }
      const from = charToPos(block, Number(command.start));
      const to = charToPos(block, Number(command.end));
      const marks = editor.state.doc.resolve(from).marks();
      tr.replaceWith(from, to, text.length > 0 ? schema.text(text, marks) : []);
      break;
    }
    case 'insertParagraph': {
      const heading = Math.max(0, Math.min(6, Math.trunc(Number(command.heading ?? 0))));
      const text = String(command.text ?? '');
      const content = text.length > 0 ? schema.text(text) : null;
      const node =
        heading > 0
          ? schema.nodes['docHeading']!.create({ level: heading }, content)
          : schema.nodes['docParagraph']!.create(null, content);
      const at =
        command.after === null || command.after === undefined
          ? 0
          : (() => {
              const anchor = refOf(refs, command.after);
              return anchor ? anchor.pos + anchor.node.nodeSize : null;
            })();
      if (at === null) return false;
      tr.insert(at, node);
      break;
    }
    case 'deleteBlock': {
      const block = refOf(refs, command.index);
      if (!block) return false;
      tr.delete(block.pos, block.pos + block.node.nodeSize);
      break;
    }
    case 'moveBlock': {
      const block = refOf(refs, command.index);
      if (!block) return false;
      if (command.to === command.index) return true;
      const target =
        command.to === null || command.to === undefined ? null : (refOf(refs, command.to) ?? null);
      if (command.to !== null && command.to !== undefined && target === null) return false;
      const node = block.node;
      tr.delete(block.pos, block.pos + node.nodeSize);
      const insertPos = target ? tr.mapping.map(target.pos) + target.node.nodeSize : 0;
      tr.insert(insertPos, node);
      break;
    }
    default:
      return false;
  }
  editor.view.dispatch(tr);
  return true;
}

export function applyCommands(commands: readonly DocxCommand[], options: ApplyOptions = {}): ApplyResult {
  const editor = editorNow();
  if (!editor) return { ok: false, applied: 0, error: '页面尚未就绪' };

  // 整批先校验：任何一条非法都不执行（与 engine 侧的批级原子一致）
  const refs = blocksOf(editor);
  for (const [index, command] of commands.entries()) {
    const problem = validateCommand(refs, command);
    if (problem) return { ok: false, applied: 0, error: `[${index}] ${problem}` };
  }
  if (options.dryRun) return { ok: true, applied: commands.length };

  let applied = 0;
  for (const [index, command] of commands.entries()) {
    if (command.op === 'check') {
      applied += 1;
      continue;
    }
    const engine = engineOp(command, editor);
    if (engine) {
      const outcome = executeOps(editor, [engine.op], {
        ...engine.ctx,
        source: 'ai',
        markAi: false,
      });
      if (!outcome.ok) return { ok: false, applied, error: `[${index}] ${outcome.error ?? outcome.summary}` };
      applied += 1;
      continue;
    }
    if (!applyCustom(command, editor))
      return { ok: false, applied, error: `[${index}] ${command.op} 执行失败` };
    applied += 1;
  }
  return { ok: true, applied };
}

/** check 运维用：只读统计（不落盘）。 */
export function auditOf(editor: Editor): Record<string, unknown> {
  const refs = blocksOf(editor);
  const paragraphs = refs.filter((ref) => isTextBlock(ref.node)).length;
  const tables = refs.filter((ref) => ref.node.type.name === 'docTable').length;
  return { blocks: refs.length, paragraphs, tables, tier: 'normal' };
}

// ─── agent-request 处理 ───────────────────────────────────────────────

export interface AgentSessionInfo {
  ref: string;
  name: string;
  tier?: string;
}

interface AgentReplyPayload {
  type: 'agent-result';
  id: string;
  ok: boolean;
  context?: unknown;
  receipt?: unknown;
  error?: { code: string; message: string };
}

/** 触发渲染器的保存（走 plan 通道），返回是否已落盘。 */
export async function flushDocument(): Promise<boolean> {
  return flush()
}

/** 触发渲染器的保存（走 plan 通道），返回是否已落盘。 */
async function flush(): Promise<boolean> {
  const save = bridge()?.save;
  if (!save) return false;
  try {
    return await save();
  } catch {
    return false;
  }
}

export async function handleAgentRequest(
  message: BusMessage,
  session: AgentSessionInfo,
  sessionId: string,
): Promise<void> {
  const id = String(message.id ?? '');
  const reply = (payload: AgentReplyPayload): void => {
    void window.extension.ura.page
      .agentResult(sessionId, payload)
      .then(() => undefined)
      .catch(() => undefined);
  };

  try {
    const editor = await waitForEditor();
    if (!editor)
      throw coded('command.page-unavailable', '页面尚未就绪（编辑器未挂载）');
    const state = sessionState();

    if (message.kind === 'context') {
      reply({
        type: 'agent-result',
        id,
        ok: true,
        context: {
          kind: 'docx',
          ref: session.ref,
          version: state.version,
          epoch: state.epoch,
          docRev: String(state.docRev),
          text: projectLive(editor, {
            name: session.name,
            version: state.version,
            epoch: state.epoch,
            ...(session.tier ? { tier: session.tier } : {}),
          }),
        },
      });
      return;
    }

    const commands = unwrapAgentCommands(message.commands);
    if (commands.length === 0) throw coded('command.invalid', 'agent 批次为空');
    if (typeof message.epoch === 'number' && message.epoch !== state.epoch)
      throw coded(
        'command.stale-index',
        `epoch 已变化（${message.epoch} → ${state.epoch}），请重新 guide/read`,
      );

    // agent 的坐标按磁盘基线给：先把本地未保存改动落盘
    const synced = await flush();
    if (!synced) throw coded('command.stale', '存在未保存的人类编辑，agent 批次已拒绝');

    // `check` 是只读诊断：在实时文档上直接算（不落盘、不动 version）
    const wantsCheck = commands.some((command) => command.op === 'check');
    const audit = wantsCheck ? auditOf(editor) : undefined;

    const result = applyCommands(commands, { dryRun: message.dryRun === true });
    if (!result.ok) throw coded('command.invalid', result.error ?? '命令执行失败');

    if (message.dryRun) {
      reply({
        type: 'agent-result',
        id,
        ok: true,
        receipt: {
          applied: result.applied,
          version: sessionState().version,
          epoch: sessionState().epoch,
          dryRun: true,
        },
      });
      return;
    }

    if (!(await flush())) throw coded('command.invalid', '命令已应用，但保存失败（见状态栏）');
    const after = sessionState();
    const receipt: Partial<Receipt> & Record<string, unknown> = {
      ref: session.ref,
      applied: result.applied,
      version: after.version,
      epoch: after.epoch,
      ...(audit ? { audit } : {}),
    };
    reply({ type: 'agent-result', id, ok: true, receipt });
  } catch (error) {
    reply({
      type: 'agent-result',
      id,
      ok: false,
      error: {
        code: errorCodeOf(error) ?? 'command.invalid',
        message: messageOf(error),
      },
    });
  }
}
