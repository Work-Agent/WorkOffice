// ─── office-docx 页面会话 ──────────────────────────────────────────────
//
// 页面 = 该 ref 的唯一编辑模型（契约 I10）：
// - 内容自取（`content: 'pulled'`）：`get` 元数据 + `bytes` 分块读 + 本地解析
// - 人类编辑 → 命令 → 模型（`applyCommands`，与 agent 同一套词汇）
// - flush：模型 → `SavePlan` → `/commit`（宿主锁内 exec → 写盘）
// - 回执：`version` / `epoch` 镜像 + `sync.manifest` rebase（I15）
// - Ctrl+Z = 模型历史（I16）；无逆计划栈
// - `agent-request`：context 回 live 投影；apply 在页面模型应用后提交（§7.5）

import {
  parseDocxBytes,
  tierFor,
  type ParsedDocx,
} from '../shared/document';
import { applyCommands } from '../shared/ops';
import {
  mergeCommands,
  pasteCommands,
  rangeCommands,
  type EditBatch,
} from '../shared/blocks-edit';
import { baselineSlices, planFromModel, type BaselineSlices } from '../shared/plan';
import { applyManifest } from '../shared/sync';
import { projectModel } from '../shared/project';
import {
  CommandRejected,
  type BlockRange,
  type DocxCommand,
  type ModelBlock,
  type SyncManifest,
  type SyncManifestBlock,
} from '../shared/types';
import { runStyleAt } from '../shared/xml-ops';
import { BlockListEditor } from './editor';
import type { BusMessage, PageSessionInfo, Receipt } from './env';

const SESSION_ID = new URLSearchParams(location.search).get('session');

// ─── 状态 ─────────────────────────────────────────────────────────────

interface SessionState {
  info: PageSessionInfo | null;
  ref: string;
  name: string;
  version: string;
  epoch: number;
  tier: 'normal' | 'lite' | 'readOnly' | 'refused';
  parsed: ParsedDocx | null;
  model: ModelBlock[];
  committed: ModelBlock[];
  baseline: BaselineSlices;
  readOnly: boolean;
  conflict: boolean;
  syncing: boolean;
  loading: boolean;
  initialized: boolean;
  modelRev: number;
  pendingApplied: number;
}

const state: SessionState = {
  info: null,
  ref: '',
  name: 'document.docx',
  version: '',
  epoch: 0,
  tier: 'normal',
  parsed: null,
  model: [],
  committed: [],
  baseline: { hashes: new Map(), total: 0, trailingSectPrIndex: null, trailingSectPrHash: null },
  readOnly: false,
  conflict: false,
  syncing: false,
  loading: false,
  initialized: false,
  modelRev: 0,
  pendingApplied: 0,
};

const history: ModelBlock[][] = [];
let historyIndex = -1;
let lastGroup: { key: string; at: number } | null = null;
let flushing = false;
let queued = false;
let flushWait: Promise<void> | null = null;
let flushDone: (() => void) | null = null;
let retries = 0;
let lastDirtySent: boolean | null = null;
let idleTimer: ReturnType<typeof setTimeout> | null = null;
let maxTimer: ReturnType<typeof setTimeout> | null = null;

// ─── DOM ──────────────────────────────────────────────────────────────

const blocksRoot = requireElement<HTMLElement>('blocks');
const moreButton = requireElement<HTMLButtonElement>('more');
const noticeBar = requireElement<HTMLElement>('notice');
const statusName = requireElement<HTMLElement>('status-name');
const statusState = requireElement<HTMLElement>('status-state');
const statusVersion = requireElement<HTMLElement>('status-version');
const statusEpoch = requireElement<HTMLElement>('status-epoch');
const statusBlocks = requireElement<HTMLElement>('status-blocks');
const statusTier = requireElement<HTMLElement>('status-tier');

const extension = window.extension;
const editor = new BlockListEditor(blocksRoot, moreButton, {
  onInput: (index, text) => {
    const block = state.model.find((item) => item.docxIndex === index);
    if (!block || block.kind !== 'p') return;
    // 只替换改动的字符区间：整段重写会把区间格式（半段加粗）冲掉。
    mutate([{ op: 'setText', index, ...textDiff(block.text, text) }], {
      group: `text:${index}`,
      render: false,
    });
  },
  onSplit: (index, before, after) => {
    const ok = mutate(
      [
        { op: 'setText', index, text: before },
        { op: 'insertParagraph', after: index, text: after },
      ],
      {},
      );
    if (ok) {
      const newest = newestLocalIndex();
      if (newest !== null) editor.focusIndex(newest);
    }
  },
  onDelete: (index) => {
    mutate([{ op: 'deleteBlock', index }], {});
  },
  onMerge: (index, direction) => mergeParagraph(index, direction),
  onReplaceRange: (range, text) => replaceRange(range, text),
  onPasteLines: (index, currentText, offset, lines) =>
    pasteLines(index, currentText, offset, lines),
  onUndo: () => void undo(),
  onRedo: () => void redo(),
  onSave: () => void flush(true),
  onShortcut: (action) => void runCommand(action),
  onActiveChange: () => updateStatus(),
});

// ─── 启动 ─────────────────────────────────────────────────────────────

void boot();

async function boot(): Promise<void> {
  if (!SESSION_ID) {
    fatal('缺少页面会话参数（session）');
    return;
  }
  extension.bus.on(onBusMessage);
  bindToolbar();
  extension.bus.loaded();
  try {
    await loadFromDisk();
    state.initialized = true;
    updateStatus();
  } catch (error) {
    fatal(`打开失败：${messageOf(error)}`);
  }
}

function onBusMessage(message: BusMessage): void {
  switch (message.type) {
    case 'theme':
      document.documentElement.dataset['theme'] =
        message.theme === 'dark' ? 'dark' : message.theme === 'light' ? 'light' : 'system';
      return;
    case 'content':
      void onExternalContent();
      return;
    case 'action':
      if (message.id === 'save') void saveThenReply();
      return;
    case 'agent-request':
      void onAgentRequest(message);
      return;
    default:
      return;
  }
}

// ─── 载入 / 重载 ───────────────────────────────────────────────────────

async function loadFromDisk(): Promise<void> {
  if (state.loading) return;
  state.loading = true;
  try {
    const info = await extension.ura.page.get<PageSessionInfo>(SESSION_ID!);
    const bytes = await fetchAllBytes(info.size);
    const parsed = parseDocxBytes(bytes);
    state.info = info;
    state.ref = info.ref;
    state.name = nameOf(info.ref);
    state.version = info.version;
    state.epoch = info.epoch;
    // 体积分级按**本文档**算（宿主 info.tier 是插件级默认值，不能替文档判级）。
    state.tier = tierFor(parsed.weight);
    state.parsed = parsed;
    state.baseline = baselineSlices(parsed);
    state.model = parsed.blocks.map((block) => ({ ...block }));
    state.committed = cloneModel(state.model);
    state.readOnly =
      info.write !== 'plan' ||
      state.tier === 'readOnly' ||
      state.tier === 'refused';
    state.conflict = false;
    state.pendingApplied = 0;
    state.modelRev += 1;
    resetHistory(state.model);
    editor.setReadOnly(state.readOnly);
    editor.render(state.model);
    clearNotice();
    if (state.readOnly)
      showNotice(
        state.tier === 'refused'
          ? '文档过大（>230k 权重），本版本只解析不渲染；请在 Word 中打开。'
          : '文档已进入只读分级：可查看，自动保存已停止。',
      );
    updateRuntimeNotice();
    notifyDirty();
    updateStatus();
  } finally {
    state.loading = false;
  }
}

async function onExternalContent(): Promise<void> {
  if (!state.initialized) return;
  if (isDirty()) {
    state.conflict = true;
    showConflictNotice();
    updateStatus();
    return;
  }
  await loadFromDisk().catch((error) => showNotice(`重新载入失败：${messageOf(error)}`));
}

async function fetchAllBytes(size?: number): Promise<Uint8Array> {
  const chunkSize = 1024 * 1024;
  const limit = size ?? 0;
  const chunks: Uint8Array[] = [];
  let offset = 0;
  for (;;) {
    const end = limit > 0 ? Math.min(limit, offset + chunkSize) : offset + chunkSize;
    const chunk = await extension.ura.page.bytes(SESSION_ID!, { start: offset, end });
    if (!chunk || chunk.byteLength === 0) break;
    chunks.push(chunk);
    offset += chunk.byteLength;
    if (limit > 0 && offset >= limit) break;
    if (chunk.byteLength < chunkSize) break;
    if (offset > 512 * 1024 * 1024) throw new Error('文档超过 512 MB 上限');
  }
  const out = new Uint8Array(offset);
  let cursor = 0;
  for (const chunk of chunks) {
    out.set(chunk, cursor);
    cursor += chunk.byteLength;
  }
  return out;
}

// ─── 模型变更 ─────────────────────────────────────────────────────────

interface MutateOptions {
  /** 历史分组键：同一组内的连续编辑只产生一个撤销步骤。 */
  group?: string;
  render?: boolean;
}

function mutate(
  commands: readonly DocxCommand[],
  options: MutateOptions = {},
): boolean {
  if (state.readOnly) return false;
  try {
    const result = applyCommands(state.model, commands);
    state.model = result.model;
    state.modelRev += 1;
    state.pendingApplied += result.applied;
    pushHistory(options.group);
    if (options.render !== false) editor.render(state.model, true);
    notifyDirty();
    updateStatus();
    scheduleFlush();
    return true;
  } catch (error) {
    if (error instanceof CommandRejected) {
      showNotice(`命令被拒 [${error.index}] ${error.code}：${error.message}`);
      return false;
    }
    showNotice(`命令失败：${messageOf(error)}`);
    return false;
  }
}

/** 新旧文本的最小差异（公共前缀 / 后缀）→ 只替换改动的字符区间。 */
function textDiff(
  before: string,
  after: string,
): { start: number; end: number; text: string } {
  const max = Math.min(before.length, after.length);
  let start = 0;
  while (start < max && before[start] === after[start]) start += 1;
  let tail = 0;
  while (tail < max - start && before[before.length - 1 - tail] === after[after.length - 1 - tail])
    tail += 1;
  return {
    start,
    end: before.length - tail,
    text: after.slice(start, after.length - tail),
  };
}

function isDirty(): boolean {
  if (state.model.length !== state.committed.length) return true;
  for (const [index, block] of state.model.entries()) {
    const base = state.committed[index]!;
    if (block.xml !== base.xml || block.docxIndex !== base.docxIndex) return true;
  }
  return false;
}

// ─── 跨块编辑：块级宿主 → 跨块命令的翻译层 ─────────────────────────────
//
// 一段一个 contenteditable 拿不到浏览器的跨段编辑结果，所以「段首退格 /
// 段尾 Delete / 跨段选区被覆盖 / 多行粘贴」由 editor.ts 取 DOM 意图，
// shared/blocks-edit.ts 把意图翻成命令批。写盘路径不变：计划仍是「最终顶层
// 块序列」，只是块数变多了。

function visibleBlocks(): ModelBlock[] {
  return state.model.filter((block) => !block.hidden);
}

function visiblePosition(index: number | null | undefined): number {
  if (typeof index !== 'number') return -1;
  return visibleBlocks().findIndex((block) => block.docxIndex === index);
}

/** 段首退格 / 段尾 Delete：并入相邻段，光标停在接缝处。 */
function mergeParagraph(index: number, direction: 'back' | 'forward'): void {
  if (state.readOnly) return;
  applyBatch(mergeCommands(state.model, index, direction));
}

/** 跨段选区被输入 / 删除覆盖（`text` 为空 = 纯删除）。 */
function replaceRange(range: BlockRange, text: string): void {
  if (state.readOnly) return;
  applyBatch(rangeCommands(state.model, range, text));
}

/** 粘贴多行：每行一段。 */
function pasteLines(
  index: number,
  currentText: string,
  offset: number,
  lines: string[],
): void {
  if (state.readOnly) return;
  applyBatch(pasteCommands(state.model, index, currentText, offset, lines));
}

/** 统一应用一份跨块编辑批：拒绝给提示，成功后把光标落回接缝。 */
function applyBatch(batch: EditBatch): void {
  if (batch.kind === 'refuse') {
    showNotice(batch.reason);
    return;
  }
  if (batch.kind === 'skip') return;
  if (!mutate(batch.commands, {})) return;
  if (batch.focus) {
    editor.focusIndex(batch.focus.index, batch.focus.offset);
    return;
  }
  if (batch.tail) {
    const position = visiblePosition(batch.tail.anchor);
    if (position >= 0)
      editor.focusPosition(position + batch.tail.count - 1, batch.tail.offset);
  }
}

function notifyDirty(): void {
  const dirty = isDirty();
  if (dirty === lastDirtySent) return;
  lastDirtySent = dirty;
  extension.bus.send({ type: 'dirty', dirty });
}

// ─── 历史（Ctrl+Z = 模型历史，I16） ───────────────────────────────────

function resetHistory(model: readonly ModelBlock[]): void {
  history.length = 0;
  history.push(cloneModel(model));
  historyIndex = 0;
  lastGroup = null;
}

function pushHistory(group?: string): void {
  const now = Date.now();
  const snapshot = cloneModel(state.model);
  if (group && lastGroup && lastGroup.key === group && now - lastGroup.at < 1500) {
    if (historyIndex >= 0) history[historyIndex] = snapshot;
    lastGroup.at = now;
    return;
  }
  history.splice(historyIndex + 1);
  history.push(snapshot);
  if (history.length > 200) history.shift();
  historyIndex = history.length - 1;
  lastGroup = group ? { key: group, at: now } : null;
}

async function undo(): Promise<void> {
  if (state.readOnly) return;
  await flush();
  if (historyIndex <= 0) return;
  historyIndex -= 1;
  state.model = cloneModel(history[historyIndex]!);
  state.modelRev += 1;
  state.pendingApplied += 1;
  lastGroup = null;
  editor.render(state.model, true);
  notifyDirty();
  updateStatus();
  scheduleFlush(0);
}

async function redo(): Promise<void> {
  if (state.readOnly) return;
  await flush();
  if (historyIndex >= history.length - 1) return;
  historyIndex += 1;
  state.model = cloneModel(history[historyIndex]!);
  state.modelRev += 1;
  state.pendingApplied += 1;
  lastGroup = null;
  editor.render(state.model, true);
  notifyDirty();
  updateStatus();
  scheduleFlush(0);
}

// ─── 提交 ─────────────────────────────────────────────────────────────

function scheduleFlush(delay = 400): void {
  if (state.readOnly || state.conflict) return;
  if (idleTimer) clearTimeout(idleTimer);
  idleTimer = setTimeout(() => {
    idleTimer = null;
    void flush();
  }, delay);
  if (!maxTimer)
    maxTimer = setTimeout(() => {
      maxTimer = null;
      void flush();
    }, 2000);
}

async function flush(explicit = false): Promise<void> {
  if (idleTimer) {
    clearTimeout(idleTimer);
    idleTimer = null;
  }
  if (state.readOnly || state.conflict || !state.initialized) return;
  if (flushing) {
    queued = true;
    // 等正在进行的提交落地，否则 agent / undo 会在「半提交」状态上做判断。
    if (flushWait) await flushWait;
    return;
  }
  if (!isDirty()) {
    if (explicit) showSaved();
    return;
  }
  flushing = true;
  flushWait = new Promise<void>((resolve) => {
    flushDone = resolve;
  });
  setStatus('保存中…');
  try {
    let rounds = 0;
    while (isDirty() && !state.conflict && rounds < 3) {
      queued = false;
      const snapshot = cloneModel(state.model);
      const applied = state.pendingApplied;
      state.pendingApplied = 0;
      const plan = planFromModel(snapshot, state.baseline, applied);
      const receipt = await commitPlan(plan, 'page');
      if (receipt.rejected?.length) {
        const [first] = receipt.rejected;
        state.conflict = true;
        showNotice(
          `提交被拒（${first?.code ?? 'command.invalid'}）：${first?.message ?? '批次被拒'}。磁盘未改动，请重新载入后再编辑。`,
          { label: '重新载入', run: () => void loadFromDisk() },
        );
        updateStatus();
        return;
      }
      await applyReceipt(receipt, snapshot);
      rounds += 1;
    }
    retries = 0;
    if (maxTimer) {
      clearTimeout(maxTimer);
      maxTimer = null;
    }
    if (!isDirty()) showSaved();
    else if (!state.conflict)
      showNotice('保存未收敛（已重试 3 轮）：仍有未提交修改，稍后会再次自动保存。');
  } catch (error) {
    // 失败时 pendingApplied 不清零：下一次回执仍能反映真实工作量。
    onCommitError(error);
  } finally {
    flushing = false;
    queued = false;
    const done = flushDone;
    flushDone = null;
    flushWait = null;
    done?.();
    updateStatus();
  }
}

function commitPlan(plan: unknown, origin: 'page' | 'agent'): Promise<Receipt> {
  return extension.ura.page.commit<Receipt>(SESSION_ID!, {
    txnId: newTxnId(),
    baseVersion: state.version,
    expect: { docRev: String(state.modelRev), epoch: state.epoch },
    payload: { type: 'plan', plan },
    origin,
  });
}

async function applyReceipt(receipt: Receipt, snapshot: readonly ModelBlock[]): Promise<void> {
  state.version = receipt.version;
  state.epoch = receipt.epoch;
  const sync = receipt.sync;
  if (sync?.kind === 'manifest' && Array.isArray(sync.blocks) && sync.blocks.length > 0) {
    const hidden = (sync as unknown as { hidden?: SyncManifest['hidden'] }).hidden;
    const manifest: SyncManifest = {
      blocks: sync.blocks as SyncManifestBlock[],
      ...(Array.isArray(hidden) ? { hidden } : {}),
    };
    const applied = applyManifest(snapshot, manifest);
    if (applied.mismatched) {
      showNotice('提交成功但 baseline sync 对不上，正在重新解析…');
      await loadFromDisk();
      return;
    }
    state.committed = applied.model;
    // 提交在途期间的本地编辑：按 id 保留，只换 docxIndex（I15 的 rebase）。
    const byId = new Map(applied.model.map((block) => [block.id, block]));
    state.model = state.model.map((block) => {
      const base = byId.get(block.id);
      return base ? { ...block, docxIndex: base.docxIndex } : block;
    });
    state.baseline = baselineFromManifest(manifest);
    // 历史快照有同一批 docxIndex 坐标：一起 rebase，否则撤销后会把旧坐标
    // 套到新基线（I15 的「静默写错块」）。对不上新基线的块改负坐标 → 必写成 xml。
    rebaseHistory(applied.model);
    editor.render(state.model, true);
    notifyDirty();
    return;
  }
  // `bytes` / `version` / `none`：页面无法 rebase 位置坐标 → 全量重读，不猜。
  await loadFromDisk();
}

function rebaseHistory(appliedModel: readonly ModelBlock[]): void {
  const seen = new Map<string, number>();
  for (const block of appliedModel) seen.set(block.id, block.docxIndex ?? -1);
  for (const [index, snapshot] of history.entries()) {
    let fresh = 0;
    for (const block of snapshot)
      if (typeof block.docxIndex === 'number' && block.docxIndex < fresh)
        fresh = block.docxIndex;
    fresh -= 1;
    history[index] = snapshot.map((block) => {
      const next = seen.get(block.id);
      if (next !== undefined && next >= 0) return { ...block, docxIndex: next };
      return { ...block, docxIndex: fresh-- };
    });
  }
}

function baselineFromManifest(manifest: SyncManifest): BaselineSlices {
  const hashes = new Map<number, string>();
  for (const block of manifest.blocks) hashes.set(block.docxIndex, block.hash);
  const sectPr = (manifest.hidden ?? []).filter((item) => item.kind === 'sectPr').pop();
  return {
    hashes,
    total: manifest.blocks.length,
    trailingSectPrIndex: sectPr?.docxIndex ?? null,
    trailingSectPrHash: sectPr ? (hashes.get(sectPr.docxIndex) ?? null) : null,
  };
}

function onCommitError(error: unknown): void {
  const code = errorCodeOf(error);
  const message = messageOf(error);
  if (code === 'command.stale') {
    state.conflict = true;
    showConflictNotice();
    updateStatus();
    return;
  }
  if (code === 'command.stale-index') {
    state.conflict = true;
    showNotice(`坐标已失效（${message}）。请重新载入后再编辑。`, {
      label: '重新载入',
      run: () => void loadFromDisk(),
    });
    updateStatus();
    return;
  }
  if (code === 'command.syncing') {
    queued = true;
    scheduleFlush(150);
    return;
  }
  if (code === 'command.page-unavailable') {
    showNotice('页面会话不可用：请在标签里重新打开文档后再保存。');
    return;
  }
  if (code === 'command.invalid' || code === 'command.forbidden') {
    showNotice(`提交被拒：${message}`);
    return;
  }
  retries += 1;
  const delay = retries === 1 ? 100 : retries === 2 ? 1000 : 3000;
  showNotice(`保存失败（第 ${retries} 次）：${message}`);
  if (retries <= 3) setTimeout(() => void flush(), delay);
  else showNotice('连续保存失败，已转为离线草稿：请检查文件是否可写，或另存副本。');
}

// ─── 工具栏 ───────────────────────────────────────────────────────────

function bindToolbar(): void {
  for (const button of document.querySelectorAll<HTMLButtonElement>('[data-cmd]'))
    button.addEventListener('click', () => void runCommand(button.dataset['cmd'] ?? ''));
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') void flush();
  });
}

/** 工具栏作用对象：多块选区优先，否则光标所在段落。 */
interface FormatTarget {
  index: number;
  /** 段内区间（仅当选区只覆盖该段的一部分时才有）。 */
  start?: number;
  end?: number;
}

/** 选区 → 格式目标：同段内取字符区间，跨段时首末段取部分、中间段整段。 */
function formatTargets(active: number | null): FormatTarget[] {
  const range = editor.selection();
  if (range) {
    if (range.from === range.to) {
      const start = Math.min(range.fromOffset, range.toOffset);
      const end = Math.max(range.fromOffset, range.toOffset);
      if (end > start) return [{ index: range.from, start, end }];
    } else {
      const visible = visibleBlocks();
      const fromPosition = visible.findIndex((block) => block.docxIndex === range.from);
      const toPosition = visible.findIndex((block) => block.docxIndex === range.to);
      if (fromPosition >= 0 && toPosition >= 0) {
        const forward = fromPosition <= toPosition;
        const start = Math.min(fromPosition, toPosition);
        const end = Math.max(fromPosition, toPosition);
        const span = visible
          .slice(start, end + 1)
          .filter((block) => block.kind === 'p' && typeof block.docxIndex === 'number');
        const headLength = forward ? range.fromOffset : range.toOffset;
        const tailLength = forward ? range.toOffset : range.fromOffset;
        const targets: FormatTarget[] = [];
        span.forEach((block, offset) => {
          const index = block.docxIndex as number;
          const length = block.text.length;
          if (offset === 0) {
            if (headLength < length) targets.push({ index, start: headLength, end: length });
            return;
          }
          if (offset === span.length - 1) {
            if (tailLength > 0) targets.push({ index, start: 0, end: Math.min(tailLength, length) });
            return;
          }
          targets.push({ index });
        });
        if (targets.length > 0) return targets;
      }
    }
  }
  return active === null ? [] : [{ index: active }];
}

/** 格式命令后把选区还回去（Word 也会保持选中，否则读数上像“取消选中”）。 */
function keepSelection(applied: boolean, range: BlockRange | null): void {
  if (applied && range) editor.select(range);
}

/** 目标 → 命令里的区间字段（整段时两个字段都不出现 = 旧行为）。 */
function styleRange(target: FormatTarget): { index: number; start?: number; end?: number } {
  if (target.start === undefined || target.end === undefined) return { index: target.index };
  return { index: target.index, start: target.start, end: target.end };
}

async function runCommand(command: string): Promise<void> {
  if (!state.initialized) return;
  const active = editor.activeIndex();
  const selection = editor.selection();
  const targets = formatTargets(active);
  const styleIndex = active ?? targets[0]?.index ?? null;
  const styleBlock =
    styleIndex === null
      ? undefined
      : state.model.find((item) => item.docxIndex === styleIndex);
  const styleOffset = targets.find((item) => item.index === styleIndex)?.start ?? 0;
  const style =
    styleBlock && styleBlock.kind === 'p'
      ? runStyleAt(styleBlock.xml, styleOffset)
      : {};
  const group = `format:${targets.map((item) => item.index).join(',')}`;

  switch (command) {
    case 'p':
    case 'h1':
    case 'h2':
    case 'h3': {
      if (targets.length === 0) return;
      const level = command === 'p' ? 0 : Number(command.slice(1));
      keepSelection(
        mutate(
          targets.map((item) => ({ op: 'setHeadingLevel', index: item.index, level })),
          { group },
        ),
        selection,
      );
      return;
    }
    case 'bold':
    case 'italic':
    case 'underline': {
      if (targets.length === 0) return;
      const next = !style[command];
      keepSelection(
        mutate(
          targets.map(
            (item): DocxCommand =>
              command === 'bold'
                ? { op: 'setTextStyle', ...styleRange(item), bold: next }
                : command === 'italic'
                  ? { op: 'setTextStyle', ...styleRange(item), italic: next }
                  : { op: 'setTextStyle', ...styleRange(item), underline: next },
          ),
          { group },
        ),
        selection,
      );
      return;
    }
    case 'sizeUp':
    case 'sizeDown': {
      if (targets.length === 0) return;
      const size = Math.max(6, Math.min(96, (style.size ?? 11) + (command === 'sizeUp' ? 2 : -2)));
      keepSelection(
        mutate(
          targets.map((item) => ({ op: 'setTextStyle', ...styleRange(item), size })),
          { group },
        ),
        selection,
      );
      return;
    }
    case 'insert': {
      const last = [...state.model].reverse().find((item) => !item.hidden);
      const after = active ?? last?.docxIndex ?? null;
      const ok = mutate([{ op: 'insertParagraph', after, text: '', heading: 0 }], {});
      if (ok) {
        const newest = newestLocalIndex();
        if (newest !== null) editor.focusIndex(newest);
      }
      return;
    }
    case 'delete': {
      if (active === null) return;
      const position = state.model.filter((item) => !item.hidden).findIndex((item) => item.docxIndex === active);
      if (mutate([{ op: 'deleteBlock', index: active }], {}))
        editor.focusPosition(Math.max(0, position - 1));
      return;
    }
    case 'undo':
      await undo();
      return;
    case 'redo':
      await redo();
      return;
    case 'save':
      await flush(true);
      return;
    case 'check':
      await runCheck();
      return;
    default:
      return;
  }
}

async function runCheck(): Promise<void> {
  try {
    // 只读诊断：exec 不产字节 → 不落盘、version 不变，回执带 audit。
    const receipt = await extension.ura.page.commit<Receipt>(SESSION_ID!, {
      txnId: newTxnId(),
      baseVersion: state.version,
      expect: { docRev: String(state.modelRev), epoch: state.epoch },
      payload: {
        type: 'plan',
        tool: 'docx.apply',
        params: { ref: state.ref, commands: [{ op: 'check' }] },
      },
      origin: 'page',
    });
    const audit = receipt.audit as Record<string, unknown> | undefined;
    showNotice(
      audit
        ? `检查：blocks=${String(audit['blocks'])} paragraphs=${String(audit['paragraphs'])} tables=${String(audit['tables'])} tier=${String(audit['tier'])}`
        : '检查完成（无诊断数据）',
      { label: '关闭', run: clearNotice },
    );
  } catch (error) {
    showNotice(`检查失败：${messageOf(error)}`);
  }
}

// ─── 写类 agent 请求（页面代理，契约 §7.5） ───────────────────────────

async function onAgentRequest(message: BusMessage): Promise<void> {
  const id = String(message.id ?? '');
  const reply = (payload: Record<string, unknown>): void => {
    void extension.ura.page
      .agentResult(SESSION_ID!, { type: 'agent-result', id, ...payload })
      .catch(() => undefined);
  };

  try {
    if (!state.initialized) throw coded('command.page-unavailable', '页面尚未就绪');
    if (message.kind === 'context') {
      reply({
        ok: true,
        context: {
          kind: 'docx',
          ref: state.ref,
          version: state.version,
          epoch: state.epoch,
          docRev: String(state.modelRev),
          text: projectModel(state.model, {
            name: state.name,
            version: state.version,
            epoch: state.epoch,
            tier: state.tier,
          }),
        },
      });
      return;
    }

    const commands = unwrapAgentCommands(message.commands);
    if (commands.length === 0) throw coded('command.invalid', 'agent 批次为空');
    if (state.readOnly) throw coded('command.forbidden', '页面处于只读分级');
    if (state.conflict) throw coded('command.stale', '文件已被外部修改，请重新载入');
    if (typeof message.epoch === 'number' && message.epoch !== state.epoch)
      throw coded(
        'command.stale-index',
        `epoch 已变化（${message.epoch} → ${state.epoch}），请重新 guide/read`,
      );

    await flush();
    if (isDirty()) throw coded('command.stale', '存在未提交的人类编辑，agent 批次已拒绝');

    if (message.dryRun) {
      const probe = applyCommands(state.model, commands);
      reply({
        ok: true,
        receipt: {
          applied: probe.applied,
          version: state.version,
          epoch: state.epoch,
          dryRun: true,
        },
      });
      return;
    }

    const before = cloneModel(state.model);
    let result;
    try {
      result = applyCommands(state.model, commands);
    } catch (error) {
      if (error instanceof CommandRejected)
        throw coded(error.code, `[${error.index}] ${error.message}`);
      throw error;
    }
    state.model = result.model;
    state.modelRev += 1;
    pushHistory();
    editor.render(state.model, true);
    notifyDirty();
    updateStatus();

    const plan = planFromModel(state.model, state.baseline, result.applied);
    const receipt = await commitPlan(plan, 'agent');
    if (receipt.rejected?.length) {
      state.model = before;
      state.modelRev += 1;
      editor.render(state.model, true);
      notifyDirty();
      const [first] = receipt.rejected;
      throw coded(first?.code ?? 'command.invalid', first?.message ?? '批次被拒');
    }
    await applyReceipt(receipt, cloneModel(state.model));
    notifyDirty();
    updateStatus();
    reply({ ok: true, receipt });
  } catch (error) {
    reply({
      ok: false,
      error: { code: errorCodeOf(error) ?? 'command.invalid', message: messageOf(error) },
    });
  }
}

function unwrapAgentCommands(raw: unknown): DocxCommand[] {
  if (!Array.isArray(raw)) return [];
  const out: DocxCommand[] = [];
  for (const item of raw) {
    if (!item || typeof item !== 'object') continue;
    const value = item as { tool?: unknown; params?: unknown; op?: unknown };
    if (typeof value.op === 'string') {
      out.push(item as DocxCommand);
      continue;
    }
    const params = value.params as { commands?: unknown } | undefined;
    if (params && Array.isArray(params.commands))
      out.push(...(params.commands as DocxCommand[]));
  }
  return out;
}

// ─── 保存握手 ─────────────────────────────────────────────────────────

async function saveThenReply(): Promise<void> {
  await flush(true);
  const ok = !isDirty() && !state.conflict;
  extension.bus.send({ type: 'save-result', ok, ...(ok ? {} : { message: '仍有未提交修改' }) });
  if (!ok)
    showNotice(
      state.conflict
        ? '无法保存：文件已被外部修改。请重新载入或另存副本。'
        : '保存未完成：请重试。',
      { label: '重新载入', run: () => void loadFromDisk() },
    );
}

// ─── 状态与提示 ───────────────────────────────────────────────────────

function updateStatus(): void {
  statusName.textContent = state.name;
  statusVersion.textContent = state.version ? `v ${state.version.slice(0, 8)}` : '';
  statusEpoch.textContent = `epoch ${state.epoch}`;
  const visible = state.model.filter((block) => !block.hidden).length;
  statusBlocks.textContent = `${visible} 块`;
  statusTier.textContent = state.readOnly ? `只读 · ${state.tier}` : state.tier;
  if (state.conflict) statusState.textContent = '冲突';
  else if (state.readOnly) statusState.textContent = '只读';
  else if (isDirty()) statusState.textContent = '未保存';
  else statusState.textContent = '已保存';
  statusState.classList.toggle('dirty', !state.conflict && !state.readOnly && isDirty());
  statusState.classList.toggle('error', state.conflict);
}

function setStatus(text: string): void {
  statusState.textContent = text;
}

function showSaved(): void {
  setStatus('已保存');
  clearNotice();
}

function updateRuntimeNotice(): void {
  const runtime = state.info?.runtime as { degraded?: string[] } | undefined;
  if (runtime?.degraded?.includes('engine-init-failed'))
    showNotice('执行点探活失败：保存能力已降级，请重启 WorkStudio 后重试。');
}

function showConflictNotice(): void {
  showNotice('文件已被外部修改，本地编辑与磁盘不一致：自动保存已暂停。', {
    label: '重新载入（丢弃本地修改）',
    run: () => void loadFromDisk(),
  });
}

function showNotice(text: string, action?: { label: string; run: () => void }): void {
  noticeBar.hidden = false;
  noticeBar.textContent = '';
  const label = document.createElement('span');
  label.textContent = text;
  noticeBar.append(label);
  if (action) {
    const group = document.createElement('span');
    group.className = 'notice-actions';
    const button = document.createElement('button');
    button.textContent = action.label;
    button.addEventListener('click', () => {
      action.run();
    });
    group.append(button);
    noticeBar.append(group);
  }
}

function clearNotice(): void {
  noticeBar.hidden = true;
  noticeBar.textContent = '';
}

function fatal(message: string): void {
  showNotice(message);
  setStatus('打开失败');
}

// ─── 工具函数 ─────────────────────────────────────────────────────────

function cloneModel(model: readonly ModelBlock[]): ModelBlock[] {
  return model.map((block) => ({ ...block }));
}

function newestLocalIndex(): number | null {
  let min: number | null = null;
  for (const block of state.model) {
    const index = block.docxIndex;
    if (typeof index === 'number' && index < 0 && (min === null || index < min)) min = index;
  }
  return min;
}

function nameOf(ref: string): string {
  try {
    const url = new URL(ref);
    const name = url.pathname.slice(url.pathname.lastIndexOf('/') + 1);
    return decodeURIComponent(name) || 'document.docx';
  } catch {
    const name = ref.slice(ref.lastIndexOf('/') + 1);
    return name || 'document.docx';
  }
}

function newTxnId(): string {
  return crypto.randomUUID();
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function errorCodeOf(error: unknown): string | undefined {
  const code = (error as { code?: unknown })?.code;
  return typeof code === 'string' ? code : undefined;
}

function coded(code: string, message: string): Error {
  const error = new Error(message) as Error & { code: string };
  error.code = code;
  return error;
}

function requireElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) throw new Error(`页面缺少元素 #${id}`);
  return element as T;
}
