// ─── 页面视图：块列表渲染 + 编辑事件 ───────────────────────────────────
//
// 这是「可替换的读侧渲染层」：本仓库没有 GenOffice 页面包（vendor 为空），
// 因此默认实现是块级 contenteditable 列表；接口（render / activeIndex /
// focus）与 GenOffice 分页渲染器一致，接入时只需替换这个文件，页面协议
// （get / bytes / commit / agent-request）不变。
//
// 渲染层不产字节、不写盘（契约 L4）：所有修改都变成命令交给会话层。
//
// 块级宿主的代价是「浏览器的编辑流被块边界切断」：跨段选区、段首退格合并、
// 多行粘贴都拿不到原生结果。这个文件负责把这三类原生意图**翻译成跨块命令**
// （onMerge / onReplaceRange / onPasteLines），会话层再用同一套命令词应用。

import type { BlockRange, ModelBlock } from '../shared/types';
import { paragraphAlignOf, runSpans, runStyleAt, tableSizeOf } from '../shared/xml-ops';

export type { BlockRange };

export interface EditorCallbacks {
  onInput(index: number, text: string): void;
  onSplit(index: number, before: string, after: string): void;
  onDelete(index: number): void;
  /** 段首退格 / 段尾 Delete：该段并入相邻段，合并后光标落在接缝处。 */
  onMerge(index: number, direction: 'back' | 'forward'): void;
  /** 跨段选区被输入 / 删除覆盖（`text` 为空 = 纯删除）。 */
  onReplaceRange(range: BlockRange, text: string): void;
  /** 粘贴多行：`lines` 是拆好的各行；光标前后合并在 blocks-edit 里做。 */
  onPasteLines(index: number, currentText: string, offset: number, lines: string[]): void;
  onUndo(): void;
  onRedo(): void;
  onSave(): void;
  /** Ctrl+B / I / U：交给会话层按命令应用，避免浏览器 execCommand 产生假格式。 */
  onShortcut(action: 'bold' | 'italic' | 'underline'): void;
  onActiveChange(index: number | null): void;
}

/** 首次同步挂载多少块（先出首屏，剩余分帧追加）。 */
const FIRST_PAINT = 64;
/** 每个动画帧追加多少块。 */
const PAINT_CHUNK = 128;

export class BlockListEditor {
  private model: readonly ModelBlock[] = [];
  private visible: ModelBlock[] = [];
  private active: number | null = null;
  private caret: { index: number; offset: number } | null = null;
  private readOnly = false;
  private composing = false;
  /** 已挂载到 DOM 的块（`visible` 的前缀快照，与 `nodes` 一一对应）。 */
  private painted: ModelBlock[] = [];
  private nodes: HTMLElement[] = [];
  private paintHandle: number | null = null;
  private renderToken = 0;

  constructor(
    private readonly root: HTMLElement,
    private readonly moreButton: HTMLButtonElement,
    private readonly callbacks: EditorCallbacks,
  ) {
    // 按钮只做“仍在载入”的进度指示；全量挂载是后台自动完成的。
    this.moreButton.disabled = true;
  }

  setReadOnly(readOnly: boolean): void {
    this.readOnly = readOnly;
  }

  activeIndex(): number | null {
    return this.active;
  }

  /** 当前选区（块坐标 + 段内偏移；未归一化方向）。 */
  selection(): BlockRange | null {
    return this.selectionRange();
  }

  currentModel(): readonly ModelBlock[] {
    return this.model;
  }

  render(model: readonly ModelBlock[], preserveCaret = false): void {
    if (preserveCaret) this.caret = this.captureCaret();
    this.model = model;
    this.visible = model.filter((block) => !block.hidden);
    this.renderToken += 1;
    if (this.paintHandle !== null) {
      cancelAnimationFrame(this.paintHandle);
      this.paintHandle = null;
    }
    if (preserveCaret) this.syncPainted();
    else this.resetPainted();
    this.schedulePaint();
    if (preserveCaret && this.caret) {
      this.restoreCaret(this.caret);
      this.caret = null;
    }
  }

  /** 新文档：清空后先同步挂载首屏，其余分帧追加。 */
  private resetPainted(): void {
    if (this.paintHandle !== null) cancelAnimationFrame(this.paintHandle);
    this.paintHandle = null;
    this.root.replaceChildren();
    this.painted = [];
    this.nodes = [];
    this.appendRange(0, Math.min(this.visible.length, FIRST_PAINT));
  }

  /**
   * 已挂载区段的增量 patch：按 `id` 求公共前缀 / 后缀，只重建中间那几块。
   * 这是“全量渲染”能成立的前提——否则每次回车都要重建整篇 DOM。
   */
  private syncPainted(): void {
    const keep = Math.min(this.painted.length, this.visible.length);
    const target = this.visible.slice(0, keep);
    const old = this.painted;
    const max = Math.min(old.length, target.length);
    let head = 0;
    while (head < max && old[head]!.id === target[head]!.id) head += 1;
    let tail = 0;
    while (tail < max - head && old[old.length - 1 - tail]!.id === target[target.length - 1 - tail]!.id)
      tail += 1;
    const oldMiddle = old.slice(head, old.length - tail);
    const newMiddle = target.slice(head, target.length - tail);
    if (oldMiddle.length === 0 && newMiddle.length === 0 && old.length === target.length) {
      // 纯内容变化（id 全部对齐）：只更新改动的块，不换节点，光标不丢。
      for (const [offset, block] of target.entries()) {
        const previous = old[offset]!;
        if (previous.xml === block.xml && previous.docxIndex === block.docxIndex) continue;
        this.updateBlock(this.nodes[offset], block, offset);
      }
      this.painted = target;
      this.applyActive();
      return;
    }
    const removed = this.nodes.slice(head, this.nodes.length - tail);
    const reference = this.nodes[this.nodes.length - tail] ?? null;
    for (const node of removed) node.remove();
    const fragment = document.createDocumentFragment();
    const created: HTMLElement[] = [];
    for (const [offset, block] of newMiddle.entries()) {
      const element = this.buildBlock(block, head + offset);
      created.push(element);
      fragment.append(element);
    }
    if (reference) this.root.insertBefore(fragment, reference);
    else this.root.append(fragment);
    this.nodes = [...this.nodes.slice(0, head), ...created, ...this.nodes.slice(this.nodes.length - tail)];
    this.painted = target;
    this.renumber();
    this.applyActive();
  }

  /** 重排 `data-position`（中间插入 / 删除后，尾部节点的位置会整体偏移）。 */
  private renumber(): void {
    this.nodes.forEach((node, position) => {
      const value = String(position);
      if (node.dataset['position'] !== value) node.dataset['position'] = value;
    });
  }

  /** 追加下一帧的块（分帧挂载，首屏不阻塞）。 */
  private schedulePaint(): void {
    this.updateMore();
    if (this.painted.length >= this.visible.length || this.paintHandle !== null) return;
    const token = this.renderToken;
    this.paintHandle = requestAnimationFrame(() => {
      this.paintHandle = null;
      if (token !== this.renderToken) return;
      this.appendRange(this.painted.length, Math.min(
        this.visible.length,
        this.painted.length + PAINT_CHUNK,
      ));
      this.schedulePaint();
    });
  }

  /** 同步补齐到至少 `position + 1` 块已挂载（光标 / 选区定位前调用）。 */
  private ensurePainted(position: number): void {
    if (position < this.painted.length) return;
    if (this.paintHandle !== null) {
      cancelAnimationFrame(this.paintHandle);
      this.paintHandle = null;
    }
    this.appendRange(this.painted.length, Math.min(this.visible.length, position + 1));
  }

  private appendRange(from: number, to: number): void {
    if (to <= from) return;
    const fragment = document.createDocumentFragment();
    const created: HTMLElement[] = [];
    for (let position = from; position < to; position += 1) {
      const block = this.visible[position]!;
      const element = this.buildBlock(block, position);
      created.push(element);
      fragment.append(element);
    }
    this.root.append(fragment);
    this.nodes = [...this.nodes, ...created];
    this.painted = [...this.painted, ...this.visible.slice(from, to)];
    this.applyActive();
    this.updateMore();
  }

  /** `#more` 现在是“后台仍在挂载”的进度条，不再需要点击。 */
  private updateMore(): void {
    const remaining = this.visible.length - this.painted.length;
    this.moreButton.hidden = remaining <= 0;
    if (remaining > 0) this.moreButton.textContent = `正在载入… 还有 ${remaining} 块`;
  }

  private applyActive(): void {
    if (this.active === null) return;
    for (const node of this.nodes) node.classList.toggle(
      'is-active',
      node.dataset['index'] === String(this.active),
    );
  }

  private buildBlock(block: ModelBlock, position: number): HTMLElement {
    const element = document.createElement('div');
    element.className = 'blk';
    element.dataset['index'] = String(block.docxIndex);
    element.dataset['position'] = String(position);
    if (block.kind === 'p') {
      this.fillParagraph(element, block);
      this.bindParagraph(element);
    } else {
      this.fillOpaque(element, block);
    }
    return element;
  }

  /** 把已有节点更新成新块（文本 / 格式变化，不换节点）。 */
  private updateBlock(element: HTMLElement | undefined, block: ModelBlock, position: number): void {
    if (!element) return;
    const wasParagraph = element.classList.contains('para');
    element.dataset['index'] = String(block.docxIndex);
    element.dataset['position'] = String(position);
    element.className = 'blk';
    if (block.kind === 'p') {
      this.fillParagraph(element, block);
      if (!wasParagraph) this.bindParagraph(element);
      return;
    }
    this.fillOpaque(element, block);
  }

  private fillParagraph(element: HTMLElement, block: ModelBlock): void {
    element.classList.add('para');
    if (block.heading > 0) element.classList.add(`h${Math.min(6, block.heading)}`);
    element.removeAttribute('style');
    // 段落本身只负责标题级别与对齐；字符格式按 run 渲染（区间加粗能看见）。
    // 字体 / 行高 / 分页的近似仍属于 GenOffice 页面层（见 README §7）。
    const align = paragraphAlignOf(block.xml);
    if (align) element.style.textAlign = align === 'both' ? 'justify' : align;
    element.contentEditable = this.readOnly ? 'false' : 'true';
    element.spellcheck = false;
    this.paintRuns(element, block);
  }

  /** 按 run 切片上色：文本内容必须与 `block.text` 完全一致。 */
  private paintRuns(element: HTMLElement, block: ModelBlock): void {
    const spans = runSpans(block.xml);
    if (spans.length === 0) {
      element.textContent = block.text;
      return;
    }
    const fragment = document.createDocumentFragment();
    let cursor = 0;
    for (const span of spans) {
      if (span.charEnd <= cursor) continue;
      const text = block.text.slice(span.charStart, span.charEnd);
      if (!text) continue;
      const style = runStyleAt(block.xml, span.charStart);
      if (style.bold || style.italic || style.underline || style.size || style.color) {
        const run = document.createElement('span');
        if (style.bold) run.classList.add('r-b');
        if (style.italic) run.classList.add('r-i');
        if (style.underline) run.classList.add('r-u');
        if (style.size) run.style.fontSize = `${style.size}pt`;
        if (style.color) run.style.color = `#${style.color}`;
        run.textContent = text;
        fragment.append(run);
      } else {
        fragment.append(document.createTextNode(text));
      }
      cursor = span.charEnd;
    }
    // 异常文档（run 计数与文本长度对不上）时用剩余文本兜底，绝不丢字。
    if (cursor < block.text.length)
      fragment.append(document.createTextNode(block.text.slice(cursor)));
    element.replaceChildren(fragment);
  }

  private fillOpaque(element: HTMLElement, block: ModelBlock): void {
    element.classList.add('ro');
    element.contentEditable = 'false';
    element.textContent =
      block.kind === 'tbl'
        ? `▦ 表格 ${tableSizeOf(block.xml).rows}×${tableSizeOf(block.xml).cols}（表格内容请在 Word 里编辑）`
        : `▣ ${block.kind}（暂不支持在页面内编辑）`;
  }

  private bindParagraph(element: HTMLElement): void {
    // 坐标以 dataset 为准：节点会被 updateBlock 复用，docxIndex 会随 baseline
    // sync 换值，闭包里捕获的旧值会把命令打到错误的块上。
    const at = (): number => Number(element.dataset['index']);
    element.addEventListener('focus', () => {
      this.active = at();
      this.markActive(element);
      this.callbacks.onActiveChange(at());
    });
    element.addEventListener('blur', () => {
      element.classList.remove('is-active');
    });
    element.addEventListener('compositionstart', () => {
      this.composing = true;
    });
    element.addEventListener('compositionend', () => {
      this.composing = false;
    });
    element.addEventListener('input', () => {
      if (this.readOnly) return;
      const text = readText(element);
      this.callbacks.onInput(at(), text);
    });
    element.addEventListener('paste', (event) => {
      if (this.readOnly) return;
      event.preventDefault();
      const text = normalizePastedText(event.clipboardData?.getData('text/plain') ?? '');
      if (!text) return;
      const lines = text.split('\n');
      if (lines.length <= 1) {
        document.execCommand('insertText', false, text);
        return;
      }
      // 多行文本 → 多段（Word「只保留文本」的粘贴语义）。
      const offset = caretOffset(element);
      const current = readText(element);
      this.callbacks.onPasteLines(at(), current, offset, lines);
    });
    element.addEventListener('copy', (event) => {
      // 跨宿主选区：浏览器默认只拷单个可编辑区域，这里显式拷选区文本。
      const range = this.selectionRange();
      if (!range || range.from === range.to) return;
      const text = window.getSelection()?.toString() ?? '';
      if (!text) return;
      event.preventDefault();
      event.clipboardData?.setData('text/plain', text);
    });
    element.addEventListener('cut', (event) => {
      if (this.readOnly) return;
      const range = this.selectionRange();
      if (!range || range.from === range.to) return;
      event.preventDefault();
      const text = window.getSelection()?.toString() ?? '';
      if (text) event.clipboardData?.setData('text/plain', text);
      this.callbacks.onReplaceRange(range, '');
    });
    // 跨段选区上的输入 / 删除：浏览器会就地改 DOM（跨宿主），模型跟不上，下一次
    // render 会把改动抹掉（表现为删掉的文字「复活」）。这里一律转成命令。
    element.addEventListener('beforeinput', (event) => {
      if (this.readOnly || this.composing || event.isComposing) return;
      const range = this.selectionRange();
      if (!range || range.from === range.to) return;
      const type = event.inputType;
      if (type === 'insertText' || type === 'insertReplacementText') {
        event.preventDefault();
        this.callbacks.onReplaceRange(range, event.data ?? '');
        return;
      }
      if (
        type === 'deleteContentBackward' ||
        type === 'deleteContentForward' ||
        type === 'deleteContent' ||
        type === 'deleteByCut'
      ) {
        event.preventDefault();
        this.callbacks.onReplaceRange(range, '');
      }
    });
    element.addEventListener('keydown', (event) => {
      if (event.isComposing || this.composing) return;
      const modifier = event.ctrlKey || event.metaKey;
      if (modifier && event.key.toLowerCase() === 's') {
        event.preventDefault();
        this.callbacks.onSave();
        return;
      }
      if (modifier && event.key.toLowerCase() === 'z') {
        event.preventDefault();
        if (event.shiftKey) this.callbacks.onRedo();
        else this.callbacks.onUndo();
        return;
      }
      if (modifier && event.key.toLowerCase() === 'y') {
        event.preventDefault();
        this.callbacks.onRedo();
        return;
      }
      if (modifier && ['b', 'i', 'u'].includes(event.key.toLowerCase())) {
        event.preventDefault();
        if (!this.readOnly)
          this.callbacks.onShortcut(event.key.toLowerCase() as 'bold' | 'italic' | 'underline');
        return;
      }
      // Ctrl+A：浏览器在 contenteditable 里只会选中当前块；Word 是选中全篇。
      if (modifier && event.key.toLowerCase() === 'a') {
        event.preventDefault();
        this.selectAll();
        return;
      }
      if (this.readOnly) return;
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        const offset = caretOffset(element);
        const text = readText(element);
        this.callbacks.onSplit(at(), text.slice(0, offset), text.slice(offset));
        return;
      }
      const selection = window.getSelection();
      const hasSelection = !!selection && !selection.isCollapsed;
      if (event.key === 'Backspace' || event.key === 'Delete') {
        // 选区非空：同块交给 input，跨块交给 beforeinput（两边都不在 keydown 里动模型）。
        if (hasSelection) return;
        const text = readText(element);
        const position = Number(element.dataset['position'] ?? 0);
        if (event.key === 'Backspace') {
          if (text === '') {
            event.preventDefault();
            this.callbacks.onDelete(at());
            if (position > 0) this.focusPosition(position - 1, 'end');
            return;
          }
          // 段首退格 = 并入上一段（Word 行为）。
          if (caretOffset(element) === 0) {
            event.preventDefault();
            this.callbacks.onMerge(at(), 'back');
          }
          return;
        }
        if (text === '') {
          event.preventDefault();
          this.callbacks.onDelete(at());
          this.focusPosition(position, 'start');
          return;
        }
        // 段尾 Delete = 并入下一段。
        if (caretOffset(element) >= text.length) {
          event.preventDefault();
          this.callbacks.onMerge(at(), 'forward');
        }
      }
    });
  }

  /** 全选当前已渲染的段落（浏览器默认只会选中当前块）。 */
  private selectAll(): void {
    this.ensurePainted(this.visible.length - 1);
    const editable = Array.from(
      this.root.querySelectorAll<HTMLElement>('.blk[contenteditable="true"]'),
    );
    const first = editable[0];
    const last = editable[editable.length - 1];
    if (!first || !last) return;
    first.focus();
    this.active = Number(first.dataset['index']);
    this.markActive(first);
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    range.setStart(first, 0);
    range.setEnd(last, last.childNodes.length);
    selection.removeAllRanges();
    selection.addRange(range);
  }

  /** 当前段落标记：左侧细条（不使用 outline，段落不应有“框”）。 */
  private markActive(element: HTMLElement): void {
    for (const other of this.root.querySelectorAll<HTMLElement>('.blk.is-active'))
      if (other !== element) other.classList.remove('is-active');
    element.classList.add('is-active');
  }

  /** DOM 选区 → 块坐标（跨宿主时 from/to 不同）。 */
  private selectionRange(): BlockRange | null {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return null;
    const range = selection.getRangeAt(0);
    const startElement = blockElementOf(range.startContainer);
    const endElement = blockElementOf(range.endContainer);
    if (!startElement || !endElement) return null;
    const from = Number(startElement.dataset['index']);
    const to = Number(endElement.dataset['index']);
    if (!Number.isFinite(from) || !Number.isFinite(to)) return null;
    return {
      from,
      fromOffset: offsetWithin(startElement, range.startContainer, range.startOffset),
      to,
      toOffset: offsetWithin(endElement, range.endContainer, range.endOffset),
    };
  }

  focusPosition(position: number, where: 'start' | 'end' | number = 'end'): void {
    if (!Number.isFinite(position) || position < 0 || position >= this.visible.length) return;
    // 目标块可能还没挂载（后台分帧挂载中）：同步补齐到它。
    this.ensurePainted(position);
    const element = this.root.querySelector<HTMLElement>(
      `.blk[data-position="${position}"]`,
    );
    if (!element || element.contentEditable !== 'true') return;
    element.focus();
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    if (typeof where === 'number') {
      const point = pointAtOffset(element, where);
      if (!point) return;
      range.setStart(point.node, point.offset);
      range.collapse(true);
    } else {
      range.selectNodeContents(element);
      range.collapse(where === 'start');
    }
    selection.removeAllRanges();
    selection.addRange(range);
  }

  focusIndex(index: number, offset?: number): void {
    const position = this.visible.findIndex((block) => block.docxIndex === index);
    if (position >= 0) this.focusPosition(position, offset ?? 'end');
  }

  /** 重新选中一个跨块区间（格式命令后保住选区，否则用户会以为“误取消”）。 */
  select(range: BlockRange): void {
    const from = this.visible.findIndex((block) => block.docxIndex === range.from);
    const to = this.visible.findIndex((block) => block.docxIndex === range.to);
    if (from < 0 || to < 0) return;
    this.ensurePainted(Math.max(from, to));
    const startElement = this.root.querySelector<HTMLElement>(
      `.blk[data-position="${from}"]`,
    );
    const endElement = this.root.querySelector<HTMLElement>(
      `.blk[data-position="${to}"]`,
    );
    if (!startElement || !endElement) return;
    const startPoint = pointAtOffset(startElement, range.fromOffset);
    const endPoint = pointAtOffset(endElement, range.toOffset);
    if (!startPoint || !endPoint) return;
    const forward =
      startPoint.node === endPoint.node
        ? startPoint.offset <= endPoint.offset
        : (startPoint.node.compareDocumentPosition(endPoint.node) &
            Node.DOCUMENT_POSITION_FOLLOWING) !==
          0;
    const [first, last] = forward
      ? [startPoint, endPoint]
      : [endPoint, startPoint];
    startElement.focus();
    const selection = window.getSelection();
    if (!selection) return;
    const dom = document.createRange();
    dom.setStart(first.node, first.offset);
    dom.setEnd(last.node, last.offset);
    selection.removeAllRanges();
    selection.addRange(dom);
  }

  private captureCaret(): { index: number; offset: number } | null {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return null;
    const anchor = selection.anchorNode;
    const host =
      anchor && anchor.nodeType === Node.ELEMENT_NODE
        ? (anchor as HTMLElement)
        : anchor?.parentElement;
    const element = host?.closest('.blk') as HTMLElement | null | undefined;
    if (!element) return null;
    const index = Number(element.dataset['index']);
    if (!Number.isFinite(index)) return null;
    return { index, offset: caretOffset(element) };
  }

  private restoreCaret(caret: { index: number; offset: number }): void {
    const element = this.root.querySelector<HTMLElement>(
      `.blk[data-index="${caret.index}"]`,
    );
    if (!element || element.contentEditable !== 'true') return;
    element.focus();
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    const point = pointAtOffset(element, caret.offset);
    if (!point) return;
    range.setStart(point.node, point.offset);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  }
}

/** 内容可编辑块的纯文本（`innerText` 会把块级换行折成 `\n`）。 */
export function readText(element: HTMLElement): string {
  return (element.innerText ?? element.textContent ?? '').replace(/\n+$/, '');
}

/** 剪贴板纯文本 → 段分割输入：统一换行符，丢弃结尾空行。 */
export function normalizePastedText(text: string): string {
  return text.replace(/\r\n?/g, '\n').replace(/\n+$/, '');
}

/** 光标在块内的字符偏移。 */
export function caretOffset(element: HTMLElement): number {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return 0;
  const range = selection.getRangeAt(0);
  return offsetWithin(element, range.startContainer, range.startOffset);
}

/** 块内某节点位置 → 字符偏移（跨块选区取端点时不能用 `caretOffset`）。 */
export function offsetWithin(element: HTMLElement, node: Node, nodeOffset: number): number {
  if (!element.contains(node)) return 0;
  const before = document.createRange();
  before.selectNodeContents(element);
  before.setEnd(node, nodeOffset);
  return before.toString().length;
}

/** 节点所属的块元素。 */
export function blockElementOf(node: Node): HTMLElement | null {
  const host =
    node.nodeType === Node.ELEMENT_NODE ? (node as HTMLElement) : node.parentElement;
  return (host?.closest('.blk') as HTMLElement | null | undefined) ?? null;
}

function pointAtOffset(
  element: HTMLElement,
  offset: number,
): { node: Node; offset: number } | null {
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  let remaining = offset;
  let node = walker.nextNode();
  while (node) {
    const length = node.textContent?.length ?? 0;
    if (remaining <= length)
      return { node, offset: Math.max(0, Math.min(length, remaining)) };
    remaining -= length;
    node = walker.nextNode();
  }
  return { node: element, offset: element.childNodes.length };
}
