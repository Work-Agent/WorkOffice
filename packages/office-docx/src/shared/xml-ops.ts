// ─── 段落 / 文字切片改写 ───────────────────────────────────────────────
//
// 只改被编辑的那一块：`pPr`（段落属性）与 `rPr`（字符属性）按 OOXML 的
// 子元素顺序插入，其余未知子元素原样保留；文本改写只重建 run 的文本节点，
// 段属性与首个 run 的字符属性都保留（未编辑的块整段不经过这里）。
//
// 已知取舍（README「已知差异」）：
// - 一个段落里的多个 run 只在改写文本时合并为一个 run（保留首个 rPr）。
// - 域（TOC / PAGE 等）与书签在改写文本时会丢失；因此文本改写只用于
//   纯文本段落，含域代码的段落建议只改格式。

import {
  escapeXml,
  findTagEnd,
  innerOf,
  localNameOfTag,
  openTagOf,
  scanElement,
  topLevelElements,
  unescapeXml,
} from './xml';
import type { BlockKind } from './types';

const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
  'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
  'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
  'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
  'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
  'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
  'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr',
  'pPrChange',
];

const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
  'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof',
  'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern',
  'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd',
  'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
  'specVanish', 'oMath', 'rPrChange',
];

/** 元素种类（body 顶层子节点）。 */
export function elementKindOf(xml: string): BlockKind {
  const trimmed = xml.trimStart();
  const name = localNameOfTag(trimmed);
  if (name === 'p') return 'p';
  if (name === 'tbl') return 'tbl';
  if (name === 'sdt') return 'sdt';
  if (name === 'sectPr') return 'sectPr';
  return 'other';
}

/** 段落内产生可见字符的 token（计数规则必须与 `paragraphTextOf` 完全一致）。 */
const TEXT_TOKEN =
  /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:t\s*\/>|<w:tab\s*\/>|<w:br\s*\/>|<w:cr\s*\/>/g;
/** 一个 `w:r` 元素。run 不嵌套，且 `w:rPr` / `w:rFonts` 等不会被误匹配。 */
const RUN_ELEMENT = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;

/** 段落纯文本：`w:t` 拼接，`w:tab` / `w:br` / `w:cr` 转成空白。 */
export function paragraphTextOf(xml: string): string {
  if (elementKindOf(xml) !== 'p') return '';
  let text = '';
  for (const match of xml.matchAll(TEXT_TOKEN)) {
    const raw = match[1];
    if (raw !== undefined) text += unescapeXml(raw);
    else if (match[0].startsWith('<w:tab')) text += '\t';
    else if (!match[0].startsWith('<w:t')) text += '\n';
  }
  return text;
}

/** 标题级别：`pStyle` 的 Heading N，或 `outlineLvl + 1`；正文为 0。 */
export function paragraphHeadingOf(xml: string): number {
  if (elementKindOf(xml) !== 'p') return 0;
  const pPr = firstChild(xml, 'pPr');
  const scope = pPr ?? xml;
  const style = /<w:pStyle\b[^>]*\bw:val="([^"]*)"/i.exec(scope)?.[1] ?? '';
  const named = /heading\s*([1-9])/i.exec(style) ?? /^([1-9])$/i.exec(style);
  if (named) return Number(named[1]);
  const outline = /<w:outlineLvl\b[^>]*\bw:val="([0-8])"/i.exec(scope)?.[1];
  if (outline !== undefined) return Number(outline) + 1;
  return 0;
}

/** 段落属性里取一个子元素的内层（找不到返回 null）。 */
function firstChild(element: string, localName: string): string | null {
  const inner = innerOf(element);
  for (const span of topLevelElements(inner)) {
    const child = inner.slice(span.start, span.end);
    if (localNameOfTag(child) === localName) return innerOf(child);
  }
  return null;
}

/** 段落属性块（含起止标签）→ 内层字符串；没有则返回 null。 */
function pPrInnerOf(paragraphXml: string): string | null {
  const inner = innerOf(paragraphXml);
  for (const span of topLevelElements(inner)) {
    const child = inner.slice(span.start, span.end);
    if (localNameOfTag(child) === 'pPr') return innerOf(child);
  }
  return null;
}

/** 首个 run 的 `rPr`（含标签）——改写文本时保留字符属性。 */
function firstRunProperties(paragraphXml: string): string {
  const run = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/.exec(paragraphXml);
  if (!run) return '';
  const rPr = /<w:rPr\b[^>]*>([\s\S]*?)<\/w:rPr>|(<w:rPr\s*\/>)/.exec(run[1]!);
  return rPr?.[0] ?? '';
}

/** 生成段落 XML。 */
export function renderParagraph(
  text: string,
  heading = 0,
  runProps = '',
): string {
  const pPr = heading > 0 ? paragraphPropsXml({ heading }) : '';
  const runs =
    text.length > 0
      ? `<w:r>${runProps}${runsOfText(text)}</w:r>`
      : '';
  return `<w:p>${pPr}${runs}</w:p>`;
}

function runsOfText(text: string): string {
  const parts = text.split('\n');
  return parts
    .map((line, index) => {
      const piece = `<w:t xml:space="preserve">${escapeXml(line)}</w:t>`;
      return index === 0 ? piece : `<w:br/>${piece}`;
    })
    .join('');
}

/** 段落属性 XML（按 OOXML 顺序插好）。 */
export function paragraphPropsXml(props: {
  heading?: number;
  align?: 'left' | 'center' | 'right' | 'both';
}): string {
  let inner = '';
  if (props.heading !== undefined) {
    const level = Math.max(0, Math.min(9, Math.trunc(props.heading)));
    if (level > 0) {
      inner = upsertProp(
        inner,
        'pStyle',
        `<w:pStyle w:val="Heading${level}"/>`,
        PPR_ORDER,
      );
      inner = upsertProp(
        inner,
        'outlineLvl',
        `<w:outlineLvl w:val="${level - 1}"/>`,
        PPR_ORDER,
      );
    } else {
      inner = removeProp(inner, 'pStyle');
      inner = removeProp(inner, 'outlineLvl');
    }
  }
  if (props.align !== undefined) {
    inner =
      props.align === 'left'
        ? removeProp(inner, 'jc')
        : upsertProp(inner, 'jc', `<w:jc w:val="${props.align}"/>`, PPR_ORDER);
  }
  return inner ? `<w:pPr>${inner}</w:pPr>` : '';
}

/** 段落开标签：自闭合（`<w:p/>`）在重写内容时改为成对标签。 */
function paragraphOpenTag(paragraphXml: string): string {
  return openTagOf(paragraphXml).replace(/\/>$/, '>');
}

/** 只替换段落文本，保留原段属性与首个 run 的字符属性。 */
export function rewriteParagraphText(paragraphXml: string, text: string): string {
  const open = paragraphOpenTag(paragraphXml);
  const inner = innerOf(paragraphXml);
  const innerStart = inner.indexOf('<');
  const head = innerStart < 0 ? inner : inner.slice(0, innerStart);
  const pPrMatch = /^\s*<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>|^\s*<w:pPr\s*\/>/.exec(
    inner,
  );
  const pPrOpen = pPrMatch?.[0] ?? '';
  const runProps = firstRunProperties(paragraphXml);
  const trailing = /\s*$/.exec(inner)?.[0] ?? '';
  const runs = text.length > 0 ? `<w:r>${runProps}${runsOfText(text)}</w:r>` : '';
  return `${open}${head}${pPrOpen}${runs}${trailing}</w:p>`;
}

/** 段落属性补丁（标题 / 对齐）：保留 runs 原样。 */
export function patchParagraphProps(
  paragraphXml: string,
  props: { heading?: number; align?: 'left' | 'center' | 'right' | 'both' },
): string {
  const open = paragraphOpenTag(paragraphXml);
  const inner = innerOf(paragraphXml);
  const headMatch = /^\s*/.exec(inner);
  const head = headMatch?.[0] ?? '';
  const body = inner.slice(head.length);
  const pPrSpan = scanTopLevel(body, 'pPr');
  const existingInner = pPrSpan ? innerOf(body.slice(pPrSpan.start, pPrSpan.end)) : '';
  let nextInner = existingInner;
  if (props.heading !== undefined) {
    const level = Math.max(0, Math.min(9, Math.trunc(props.heading)));
    if (level > 0) {
      nextInner = upsertProp(
        nextInner,
        'pStyle',
        `<w:pStyle w:val="Heading${level}"/>`,
        PPR_ORDER,
      );
      nextInner = upsertProp(
        nextInner,
        'outlineLvl',
        `<w:outlineLvl w:val="${level - 1}"/>`,
        PPR_ORDER,
      );
    } else {
      nextInner = removeProp(nextInner, 'pStyle');
      nextInner = removeProp(nextInner, 'outlineLvl');
    }
  }
  if (props.align !== undefined) {
    nextInner =
      props.align === 'left'
        ? removeProp(nextInner, 'jc')
        : upsertProp(nextInner, 'jc', `<w:jc w:val="${props.align}"/>`, PPR_ORDER);
  }
  if (!pPrSpan) {
    return nextInner
      ? `${open}${head}<w:pPr>${nextInner}</w:pPr>${body}</w:p>`
      : paragraphXml;
  }
  const before = body.slice(0, pPrSpan.start);
  const after = body.slice(pPrSpan.end);
  const pPr = nextInner ? `<w:pPr>${nextInner}</w:pPr>` : '';
  return `${open}${head}${before}${pPr}${after}</w:p>`;
}

export interface RunStylePatch {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** 磅值。 */
  size?: number;
  /** `RRGGBB`。 */
  color?: string;
}

/** 字符属性补丁：作用于段落内每个 run（不动文本与其他子元素）。 */
export function patchRunProps(paragraphXml: string, props: RunStylePatch): string {
  return paragraphXml.replace(RUN_ELEMENT, (run) => patchRunElement(run, props));
}

/** 单个 `<w:r>` 元素的字符属性补丁（区间格式与整段格式共用同一套 upsert 语义）。 */
function patchRunElement(run: string, props: RunStylePatch): string {
  const open = openTagOf(run);
  const inner = innerOf(run);
  const head = /^\s*/.exec(inner)?.[0] ?? '';
  const body = inner.slice(head.length);
  const rPrSpan = scanTopLevel(body, 'rPr');
  const existing = rPrSpan ? innerOf(body.slice(rPrSpan.start, rPrSpan.end)) : '';
  let nextInner = existing;
  if (props.bold !== undefined)
    nextInner = props.bold
      ? upsertProp(nextInner, 'b', '<w:b/>', RPR_ORDER)
      : upsertProp(nextInner, 'b', '<w:b w:val="0"/>', RPR_ORDER);
  if (props.italic !== undefined)
    nextInner = props.italic
      ? upsertProp(nextInner, 'i', '<w:i/>', RPR_ORDER)
      : upsertProp(nextInner, 'i', '<w:i w:val="0"/>', RPR_ORDER);
  if (props.underline !== undefined)
    nextInner = props.underline
      ? upsertProp(nextInner, 'u', '<w:u w:val="single"/>', RPR_ORDER)
      : upsertProp(nextInner, 'u', '<w:u w:val="none"/>', RPR_ORDER);
  if (props.size !== undefined) {
    const half = Math.max(1, Math.round(props.size * 2));
    nextInner = upsertProp(nextInner, 'sz', `<w:sz w:val="${half}"/>`, RPR_ORDER);
    nextInner = upsertProp(nextInner, 'szCs', `<w:szCs w:val="${half}"/>`, RPR_ORDER);
  }
  if (props.color !== undefined) {
    const color = /^[0-9a-f]{6}$/i.test(props.color) ? props.color.toUpperCase() : null;
    if (color)
      nextInner = upsertProp(nextInner, 'color', `<w:color w:val="${color}"/>`, RPR_ORDER);
  }
  if (!nextInner) return run;
  if (!rPrSpan) return `${open}${head}<w:rPr>${nextInner}</w:rPr>${body}</w:r>`;
  const before = body.slice(0, rPrSpan.start);
  const after = body.slice(rPrSpan.end);
  return `${open}${head}${before}<w:rPr>${nextInner}</w:rPr>${after}</w:r>`;
}

/** run 在段落里的位置：字符区间 + 元素字符串区间。 */
export interface RunSpan {
  /** 段内字符偏移（与 `paragraphTextOf` 同口径）。 */
  charStart: number;
  charEnd: number;
  /** `<w:r` 在段落 XML 中的起点，`</w:r>` 之后的第一个字符。 */
  start: number;
  end: number;
}

/** 一个 run 内可见字符数（`w:t` 文本长度；`w:tab` / `w:br` / `w:cr` 各算 1）。 */
function runTextLength(xml: string): number {
  let length = 0;
  for (const match of xml.matchAll(TEXT_TOKEN)) {
    const raw = match[1];
    length += raw !== undefined ? unescapeXml(raw).length : 1;
  }
  return length;
}

/** 段落内每个 run 的字符区间（按文档顺序，零宽 run 也会列出）。 */
export function runSpans(paragraphXml: string): RunSpan[] {
  const spans: RunSpan[] = [];
  let chars = 0;
  for (const match of paragraphXml.matchAll(RUN_ELEMENT)) {
    const start = match.index!;
    const length = runTextLength(match[0]);
    spans.push({
      charStart: chars,
      charEnd: chars + length,
      start,
      end: start + match[0].length,
    });
    chars += length;
  }
  return spans;
}

/**
 * 在段内字符偏移处切开一个 run（两半共享原 `rPr`），边界不在 run 内部时原样返回。
 * `w:t` 切开后统一写成 `xml:space="preserve"`，避免空格被 Word 吃掉。
 */
export function splitRunAt(paragraphXml: string, offset: number): string {
  const spans = runSpans(paragraphXml);
  if (spans.length === 0) return paragraphXml;
  const limit = spans[spans.length - 1]!.charEnd;
  if (offset <= 0 || offset >= limit) return paragraphXml;
  const span = spans.find((item) => item.charStart < offset && offset < item.charEnd);
  if (!span) return paragraphXml;
  const runXml = paragraphXml.slice(span.start, span.end);
  const [left, right] = splitRunElement(runXml, offset - span.charStart);
  return paragraphXml.slice(0, span.start) + left + right + paragraphXml.slice(span.end);
}

function splitRunElement(runXml: string, keep: number): [string, string] {
  const open = openTagOf(runXml);
  const inner = innerOf(runXml);
  const rPrSpan = scanTopLevel(inner, 'rPr');
  const rPr = rPrSpan ? inner.slice(rPrSpan.start, rPrSpan.end) : '';
  const bodyStart = rPrSpan ? rPrSpan.end : 0;
  const left: string[] = [];
  const right: string[] = [];
  let used = 0;
  let pos = bodyStart;
  for (const span of topLevelElements(inner)) {
    if (span.end <= bodyStart) continue;
    const element = inner.slice(span.start, span.end);
    const gap = inner.slice(pos, span.start);
    pos = span.end;
    const length = runTextLength(element);
    if (used + length <= keep) {
      left.push(gap + element);
      used += length;
      continue;
    }
    if (used >= keep) {
      right.push(gap + element);
      continue;
    }
    // 边界落在这个元素内部：只有 `w:t` 可以切（tab/br/cr 各占 1 字符）。
    const take = keep - used;
    if (localNameOfTag(element) === 't') {
      const text = unescapeXml(innerOf(element));
      left.push(`${gap}<w:t xml:space="preserve">${escapeXml(text.slice(0, take))}</w:t>`);
      right.push(`<w:t xml:space="preserve">${escapeXml(text.slice(take))}</w:t>`);
      used = keep;
      continue;
    }
    right.push(gap + element);
  }
  const tail = inner.slice(pos);
  if (tail) right.push(tail);
  return [`${open}${rPr}${left.join('')}</w:r>`, `${open}${rPr}${right.join('')}</w:r>`];
}

/** `w:pPr` 之外的 run 字符属性块（含起止标签）；没有则空串。 */
function runPropertiesXml(runXml: string): string {
  const inner = innerOf(runXml);
  const span = scanTopLevel(inner, 'rPr');
  return span ? inner.slice(span.start, span.end) : '';
}

/**
 * 按段内字符区间替换文本（不动区间外的 run）。
 *
 * 打字、退格、打字覆盖选区都走这里：插入的文字沿用「光标前一个 run」的 rPr
 * （段首则用后一个 run 的），与 Word 的行为一致。这也是区间格式能存活的前提——
 * `rewriteParagraphText` 会把整段并成一个 run，那就把半段加粗冲掉了。
 */
export function replaceRunText(
  paragraphXml: string,
  start: number,
  end: number,
  text: string,
): string {
  // 空段落（`<w:p/>`）没有 run 可以切：直接重写（`rewriteParagraphText` 会把
  // 自闭合开标签改成成对标签，并保留 pPr）。
  if (runSpans(paragraphXml).length === 0) return rewriteParagraphText(paragraphXml, text);
  let xml = paragraphXml;
  if (end > start) {
    xml = splitRunAt(xml, end);
    xml = splitRunAt(xml, start);
  } else if (start > 0) {
    // 纯插入：先把光标处切开，否则偏移会落在 run 内部而找不到落点。
    xml = splitRunAt(xml, start);
  }
  const spans = runSpans(xml);
  const covered = spans.filter(
    (span) => span.charEnd > span.charStart && span.charStart >= start && span.charEnd <= end,
  );
  let zoneStart: number;
  let zoneEnd: number;
  let style = '';
  if (covered.length > 0) {
    zoneStart = covered[0]!.start;
    zoneEnd = covered[covered.length - 1]!.end;
    style = runPropertiesXml(xml.slice(zoneStart, zoneEnd));
  } else {
    const before = spans.filter((span) => span.charEnd === start).pop();
    const after = spans.find((span) => span.charStart === start);
    if (before) {
      zoneStart = zoneEnd = before.end;
      style = runPropertiesXml(xml.slice(before.start, before.end));
    } else if (after) {
      zoneStart = zoneEnd = after.start;
      style = runPropertiesXml(xml.slice(after.start, after.end));
    } else {
      zoneStart = zoneEnd = Math.max(0, xml.lastIndexOf('</w:p>'));
    }
  }
  const run = text.length > 0 ? `<w:r>${style}${runsOfText(text)}</w:r>` : '';
  return xml.slice(0, zoneStart) + run + xml.slice(zoneEnd);
}

/**
 * 段内区间字符格式：先按端点切 run，再只给落在 `[start, end)` 内的 run 贴 `rPr`。
 * 未涉及的 run 逐字节不变（与未改块 `{kind:'original'}` 同一套保真哲学）。
 */
export function applyRunStyle(
  paragraphXml: string,
  start: number,
  end: number,
  props: RunStylePatch,
): string {
  let xml = paragraphXml;
  if (end > start) xml = splitRunAt(xml, end);
  if (start > 0) xml = splitRunAt(xml, start);
  const spans = runSpans(xml);
  let out = xml;
  for (let index = spans.length - 1; index >= 0; index -= 1) {
    const span = spans[index]!;
    if (span.charStart === span.charEnd) continue;
    if (span.charEnd <= start || span.charStart >= end) continue;
    const run = out.slice(span.start, span.end);
    out = out.slice(0, span.start) + patchRunElement(run, props) + out.slice(span.end);
  }
  return out;
}

function scanTopLevel(inner: string, localName: string) {  for (const span of topLevelElements(inner)) {
    if (localNameOfTag(inner.slice(span.start, span.end)) === localName)
      return span;
  }
  return null;
}

/** 属性容器内层：删除同名子元素。 */
function removeProp(inner: string, name: string): string {
  let out = '';
  let pos = 0;
  for (const span of topLevelElements(inner)) {
    out += inner.slice(pos, span.start);
    const element = inner.slice(span.start, span.end);
    const unknown = localNameOfTag(element) !== name;
    if (unknown) out += element;
    pos = span.end;
  }
  out += inner.slice(pos);
  return out;
}

/** 属性容器内层：按 canonical 顺序插入 / 覆盖同名子元素。 */
function upsertProp(
  inner: string,
  name: string,
  xml: string,
  order: string[],
): string {
  const cleaned = removeProp(inner, name);
  const rank = order.indexOf(name);
  let out = '';
  let pos = 0;
  let inserted = false;
  for (const span of topLevelElements(cleaned)) {
    const element = cleaned.slice(span.start, span.end);
    const childName = localNameOfTag(element);
    const childRank = order.indexOf(childName);
    if (!inserted && rank >= 0 && childRank > rank) {
      out += xml;
      inserted = true;
    }
    out += element;
    pos = span.end;
  }
  out += cleaned.slice(pos);
  if (!inserted) out += xml;
  return out;
}

/** 顶层元素内层里的 `w:tbl` 行列数（sidecar 用）。 */
export function tableSizeOf(tableXml: string): { rows: number; cols: number } {
  const inner = innerOf(tableXml);
  let rows = 0;
  let cols = 0;
  for (const span of topLevelElements(inner)) {
    const row = inner.slice(span.start, span.end);
    if (localNameOfTag(row) !== 'tr') continue;
    rows += 1;
    let cells = 0;
    for (const cellSpan of topLevelElements(innerOf(row))) {
      const cell = innerOf(row).slice(cellSpan.start, cellSpan.end);
      if (localNameOfTag(cell) === 'tc') cells += 1;
    }
    cols = Math.max(cols, cells);
  }
  return { rows, cols };
}

/** 文本是否只由空白与段落标记组成。 */
export function isBlankParagraph(paragraphXml: string): boolean {
  return paragraphTextOf(paragraphXml).trim().length === 0;
}

/** 校验顶层片段是良构单元素（计划自足性检查的一部分）。 */
export function isWellFormedElement(xml: string): boolean {
  const trimmed = xml.trim();
  if (!trimmed.startsWith('<')) return false;
  const span = scanElement(trimmed, 0);
  if (!span) return false;
  if (span.end !== trimmed.length) return false;
  // 开标签本身必须是合法的标签边界（防止 `<w:p` 这类半截片段）。
  return findTagEnd(trimmed, 0) > 0;
}

export interface RunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** 磅值 */
  size?: number;
  color?: string;
}

/** run 元素 → 字符属性（`rPr` 缺失则为 `{}`）。 */
function styleOfRun(runXml: string): RunStyle {
  const inner = innerOf(runXml);
  const rPrSpan = scanTopLevel(inner, 'rPr');
  if (!rPrSpan) return {};
  const rPr = innerOf(inner.slice(rPrSpan.start, rPrSpan.end));
  if (!rPr) return {};
  const style: RunStyle = {};
  const toggle = (name: string): boolean | undefined => {
    const match = new RegExp(`<w:${name}\\b([^>]*)/?>`).exec(rPr);
    if (!match) return undefined;
    const val = /w:val="([^"]*)"/.exec(match[1] ?? '')?.[1];
    return val === undefined || val === '1' || val === 'true';
  };
  const bold = toggle('b');
  if (bold !== undefined) style.bold = bold;
  const italic = toggle('i');
  if (italic !== undefined) style.italic = italic;
  const underline = toggle('u');
  if (underline !== undefined) style.underline = underline;
  const size = /<w:sz\b[^>]*w:val="(\d+)"/.exec(rPr)?.[1];
  if (size) style.size = Number(size) / 2;
  const color = /<w:color\b[^>]*w:val="([0-9a-f]{6})"/i.exec(rPr)?.[1];
  if (color) style.color = color.toUpperCase();
  return style;
}

/** 段落首个 run 的字符属性（工具栏 toggle 用）。 */
export function firstRunStyleOf(paragraphXml: string): RunStyle {
  const first = runSpans(paragraphXml)[0];
  return first ? styleOfRun(paragraphXml.slice(first.start, first.end)) : {};
}

/** 段内某个字符位置所在 run 的字符属性（区间格式的 toggle 基准）。 */
export function runStyleAt(paragraphXml: string, offset: number): RunStyle {
  const spans = runSpans(paragraphXml);
  if (spans.length === 0) return {};
  const position = Math.max(0, offset);
  const span =
    spans.find((item) => position >= item.charStart && position < item.charEnd) ??
    spans[spans.length - 1]!;
  return styleOfRun(paragraphXml.slice(span.start, span.end));
}

/** 段落对齐（工具栏状态用）。 */
export function paragraphAlignOf(
  paragraphXml: string,
): 'left' | 'center' | 'right' | 'both' | undefined {
  const jc = /<w:jc\b[^>]*w:val="(left|center|right|both|start|end)"/i.exec(paragraphXml)?.[1];
  if (!jc) return undefined;
  if (jc === 'start') return 'left';
  if (jc === 'end') return 'right';
  return jc as 'left' | 'center' | 'right' | 'both';
}
