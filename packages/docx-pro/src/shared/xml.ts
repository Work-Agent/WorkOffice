// ─── 最小 XML 切片工具 ─────────────────────────────────────────────────
//
// docx 的「字节保真」靠的就是切片：未改块原样回写，改动的块重新生成。
// 因此这里不做通用 XML 解析（不建 DOM、不排序属性、不展开实体），只做
// 三件事：定位顶层元素边界、取元素名、取 / 写文本。
//
// 被解析的对象是 `word/document.xml` 的 body 内层；它必须是良构 XML，但
// 允许注释 / CDATA / 处理指令穿插（Word 会写入 `w14:paraId` 等属性，我们
// 一律保留原始开标签文本）。

export const XML_ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

export function unescapeXml(text: string): string {
  return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (whole, body: string) => {
    if (body.startsWith('#x') || body.startsWith('#X'))
      return String.fromCodePoint(Number.parseInt(body.slice(2), 16));
    if (body.startsWith('#'))
      return String.fromCodePoint(Number.parseInt(body.slice(1), 10));
    return XML_ENTITIES[body.toLowerCase()] ?? whole;
  });
}

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export interface Span {
  start: number;
  end: number;
}

/**
 * 扫描一个顶层元素（从 `pos` 处的 `<` 开始）。
 *
 * 返回元素在整个字符串里的区间（含起止标签）；遇到注释 / CDATA / 处理指令
 * 视为普通文本；自闭合元素立即结束。找不到配平则返回 null。
 */
export function scanElement(xml: string, pos: number): Span | null {
  if (xml[pos] !== '<') return null;
  // 注释 / CDATA / 处理指令不是元素。
  if (xml.startsWith('<!--', pos)) {
    const end = xml.indexOf('-->', pos + 4);
    return end < 0 ? null : { start: pos, end: end + 3 };
  }
  if (xml.startsWith('<![CDATA[', pos)) {
    const end = xml.indexOf(']]>', pos + 9);
    return end < 0 ? null : { start: pos, end: end + 3 };
  }
  if (xml.startsWith('<?', pos)) {
    const end = xml.indexOf('?>', pos + 2);
    return end < 0 ? null : { start: pos, end: end + 2 };
  }
  if (xml.startsWith('<!', pos)) {
    const end = xml.indexOf('>', pos + 2);
    return end < 0 ? null : { start: pos, end: end + 1 };
  }

  const openEnd = findTagEnd(xml, pos);
  if (openEnd < 0) return null;
  const openTag = xml.slice(pos, openEnd + 1);
  const name = localNameOfTag(openTag);
  if (!name) return null;
  if (/\/>$/.test(openTag)) return { start: pos, end: openEnd + 1 };

  let depth = 1;
  let cursor = openEnd + 1;
  while (cursor < xml.length) {
    const next = xml.indexOf('<', cursor);
    if (next < 0) return null;
    if (xml.startsWith('<!--', next)) {
      const end = xml.indexOf('-->', next + 4);
      if (end < 0) return null;
      cursor = end + 3;
      continue;
    }
    if (xml.startsWith('<![CDATA[', next)) {
      const end = xml.indexOf(']]>', next + 9);
      if (end < 0) return null;
      cursor = end + 3;
      continue;
    }
    if (xml.startsWith('<?', next)) {
      const end = xml.indexOf('?>', next + 2);
      if (end < 0) return null;
      cursor = end + 2;
      continue;
    }
    if (xml.startsWith('</', next)) {
      const end = findTagEnd(xml, next);
      if (end < 0) return null;
      depth -= 1;
      cursor = end + 1;
      if (depth === 0) return { start: pos, end: cursor };
      continue;
    }
    if (xml.startsWith('<!', next)) {
      const end = xml.indexOf('>', next + 2);
      if (end < 0) return null;
      cursor = end + 1;
      continue;
    }
    const end = findTagEnd(xml, next);
    if (end < 0) return null;
    if (!/\/>$/.test(xml.slice(next, end + 1))) depth += 1;
    cursor = end + 1;
  }
  return null;
}

/** `>` 的位置（跳过引号内的 `>`）。 */
export function findTagEnd(xml: string, pos: number): number {
  let quote = '';
  for (let i = pos + 1; i < xml.length; i += 1) {
    const ch = xml[i]!;
    if (quote) {
      if (ch === quote) quote = '';
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '>') return i;
  }
  return -1;
}

/** `<w:p …>` → `p`；`</w:p>` → `p`。 */
export function localNameOfTag(tag: string): string {
  const match = /^<\/?\s*([^\s/>]+)/.exec(tag);
  if (!match) return '';
  const full = match[1]!;
  const colon = full.indexOf(':');
  return colon < 0 ? full : full.slice(colon + 1);
}

/** 开标签的名字（含前缀）：`<w:p>` → `w:p`。 */
export function qualifiedNameOfTag(tag: string): string {
  const match = /^<\/?\s*([^\s/>]+)/.exec(tag);
  return match?.[1] ?? '';
}

/** 顶层子元素迭代器：跳过文本 / 注释 / CDATA / 处理指令，只产出元素区间。 */
export function* topLevelElements(inner: string): Generator<Span> {
  let pos = 0;
  while (pos < inner.length) {
    const next = inner.indexOf('<', pos);
    if (next < 0) return;
    if (
      inner.startsWith('<!--', next) ||
      inner.startsWith('<![CDATA[', next) ||
      inner.startsWith('<?', next) ||
      inner.startsWith('<!', next)
    ) {
      const span = scanElement(inner, next);
      if (!span) return;
      pos = span.end;
      continue;
    }
    const span = scanElement(inner, next);
    if (!span) return;
    yield span;
    pos = span.end;
  }
}

/** 顶层元素名列表（调色板 / 校验用）。 */
export function topLevelElementNames(inner: string): string[] {
  const names: string[] = [];
  for (const span of topLevelElements(inner)) {
    names.push(localNameOfTag(inner.slice(span.start, span.end)));
  }
  return names;
}

/** 取元素的开标签（含属性），用于原样保留 `w14:paraId` 这类属性。 */
export function openTagOf(element: string): string {
  const end = findTagEnd(element, 0);
  return end < 0 ? element : element.slice(0, end + 1);
}

/** 取元素的紧邻内层（去掉起止标签）。 */
export function innerOf(element: string): string {
  const openEnd = findTagEnd(element, 0);
  if (openEnd < 0) return '';
  if (/\/>$/.test(element.slice(0, openEnd + 1))) return '';
  const closeStart = element.lastIndexOf('</');
  if (closeStart < 0 || closeStart < openEnd) return '';
  return element.slice(openEnd + 1, closeStart);
}

/** 顶层子元素的内层（父元素内层里第一个匹配 `localName` 的元素）。 */
export function firstChildInner(
  parentInner: string,
  localName: string,
): string | null {
  for (const span of topLevelElements(parentInner)) {
    const element = parentInner.slice(span.start, span.end);
    if (localNameOfTag(element) === localName) return innerOf(element);
  }
  return null;
}
