// ─── DOCX 包与 document.xml 顶层的切片解析 ─────────────────────────────
//
// 只解析 `word/document.xml`（必要时加 `word/styles.xml` 的存在性），
// 不建 DOM、不解析 rels：docx 的写回粒度是「body 顶层子节点」，
// 其余 zip 条目由 exec 原样拷贝（见 engine/zip.ts）。
//
// 页面与执行点共用这份解析（I17 的单侧同源：同一份实现、同一个 bundle）。

import { strFromU8, unzipSync } from 'fflate';
import { blockHash } from './sha256';
import type { ModelBlock } from './types';
import { localNameOfTag, topLevelElements } from './xml';
import {
  elementKindOf,
  paragraphHeadingOf,
  paragraphTextOf,
} from './xml-ops';

export const DOCUMENT_PART = 'word/document.xml';

export interface ParsedDocx {
  /** 完整 `word/document.xml` 文本。 */
  documentXml: string;
  /** body 开标签之前的内容（含前导空白）。 */
  beforeBody: string;
  /** body 闭标签之后的内容（含尾随空白）。 */
  afterBody: string;
  /** body 顶层子节点序列（含尾部 sectPr）。 */
  blocks: ModelBlock[];
  /** 段落文本总长（tier 判定用）。 */
  textLength: number;
  /** `blocks + textLength / 110`（与 GenOffice 的 openTierFor 同口径）。 */
  weight: number;
}

export class DocxParseError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'DocxParseError';
  }
}

/** 从 zip 字节解析；只解压 `word/document.xml`。 */
export function parseDocxBytes(bytes: Uint8Array): ParsedDocx {
  let archive: Record<string, Uint8Array>;
  try {
    archive = unzipSync(bytes, {
      filter: (file) => file.name === DOCUMENT_PART,
    });
  } catch {
    throw new DocxParseError(
      '不是有效的 DOCX 包（ZIP 解压失败；加密或旧版 .doc 请另存为 .docx）',
    );
  }
  const part = archive[DOCUMENT_PART];
  if (!part) throw new DocxParseError('DOCX 缺少 word/document.xml');
  return parseDocumentXml(strFromU8(part));
}

/** 解析 `word/document.xml` 文本。 */
export function parseDocumentXml(xml: string): ParsedDocx {
  const open = /<w:body(?:\s[^>]*)?>/i.exec(xml);
  if (!open || open.index === undefined)
    throw new DocxParseError('document.xml 缺少 w:body');
  const bodyOpenEnd = open.index + open[0].length;
  const bodyClose = xml.lastIndexOf('</w:body>');
  if (bodyClose < bodyOpenEnd)
    throw new DocxParseError('document.xml 的 w:body 未闭合');

  const rawInner = xml.slice(bodyOpenEnd, bodyClose);
  const leading = /^\s*/.exec(rawInner)?.[0] ?? '';
  const trailing = /\s*$/.exec(rawInner)?.[0] ?? '';
  const inner = rawInner.slice(leading.length, rawInner.length - trailing.length);

  const blocks: ModelBlock[] = [];
  let textLength = 0;
  const spans = [...topLevelElements(inner)];
  // 末尾残留（非空白）说明有未闭合 / 非法节点：宁可拒绝，也不要写出半份 XML。
  const tail = inner.slice(spans.length > 0 ? spans[spans.length - 1]!.end : 0);
  if (tail.trim().length > 0)
    throw new DocxParseError('document.xml 存在未闭合或非法节点');
  spans.forEach((span, position) => {
    const element = inner.slice(span.start, span.end);
    const kind = elementKindOf(element);
    const text = kind === 'p' ? paragraphTextOf(element) : '';
    textLength += text.length;
    const trailingSectPr =
      kind === 'sectPr' && position === spans.length - 1;
    blocks.push({
      id: `d${position}`,
      docxIndex: position,
      kind,
      xml: element,
      text,
      heading: kind === 'p' ? paragraphHeadingOf(element) : 0,
      ...(trailingSectPr || (kind === 'other' && text.trim().length === 0)
        ? { hidden: true }
        : {}),
    });
  });

  return {
    documentXml: xml,
    beforeBody: `${xml.slice(0, bodyOpenEnd)}${leading}`,
    afterBody: `${trailing}${xml.slice(bodyClose)}`,
    blocks,
    textLength,
    weight: blocks.length + textLength / 110,
  };
}

/** 基线切片摘要：`docxIndex → hash(originalXml)`。 */
export function baselineHashes(parsed: ParsedDocx): Map<number, string> {
  const map = new Map<number, string>();
  for (const block of parsed.blocks)
    map.set(block.docxIndex!, blockHash(block.xml));
  return map;
}

/** 读侧模型 = 基线块（页面自己维护 id；执行点用确定性 id）。 */
export function modelFromParsed(parsed: ParsedDocx): ModelBlock[] {
  return parsed.blocks.map((block) => ({ ...block }));
}

export type Tier = 'normal' | 'lite' | 'readOnly' | 'refused';

/** 与 GenOffice 同一分级口径：≤4k normal / ≤115k lite / ≤230k readOnly。 */
export function tierFor(weight: number): Tier {
  if (weight <= 4_000) return 'normal';
  if (weight <= 115_000) return 'lite';
  if (weight <= 230_000) return 'readOnly';
  return 'refused';
}

/** 从切片文本取顶层元素名（`w:p` → `p`）。 */
export function elementNameOf(xml: string): string {
  return localNameOfTag(xml.trimStart());
}

/** 用一段顶层 XML 构造模型块（新插入块 / manifest 里页面没有的块）。 */
export function blockFromXml(
  xml: string,
  docxIndex: number | null,
  id?: string,
): ModelBlock {
  const kind = elementKindOf(xml);
  const hidden = kind === 'sectPr';
  return {
    id: id ?? `n${Math.random().toString(36).slice(2, 10)}`,
    docxIndex,
    kind,
    xml,
    text: kind === 'p' ? paragraphTextOf(xml) : '',
    heading: kind === 'p' ? paragraphHeadingOf(xml) : 0,
    ...(hidden ? { hidden: true } : {}),
  };
}
