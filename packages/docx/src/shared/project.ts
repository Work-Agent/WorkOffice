// ─── agent 读取面（sidecar / store 的投影） ────────────────────────────
//
// 契约 §4.5：sidecar 的索引与命令的目标编号必须完全一致——读到什么就能改什么。
// 这里打印的 `[n]` 就是 `docxIndex`（body 顶层下标），不是「可见块序号」：
// 双坐标会让模型改错块（I13 的坐标系混用）。
//
// 有页面时宿主走页面代理拿 live 模型；无页面才由执行点解析字节重建。

import type { ParsedDocx } from './document';
import type { ModelBlock } from './types';
import { tableSizeOf } from './xml-ops';

export const MAX_PROJECT_BLOCKS = 400;
export const MAX_PROJECT_TEXT = 160;

export interface ProjectMeta {
  name?: string;
  version?: string;
  epoch?: number;
  tier?: string;
}

/** 块投影（`docx_guide` / `read` / sidecar 共用）。 */
export function projectModel(
  model: readonly ModelBlock[],
  meta: ProjectMeta = {},
): string {
  const visible = model.filter((block) => !block.hidden);
  const tables = visible.filter((block) => block.kind === 'tbl').length;
  const header = [
    `# ${meta.name ?? 'document.docx'}`,
    `blocks=${visible.length}`,
    tables > 0 ? `tables=${tables}` : '',
    meta.version ? `version=${short(meta.version)}` : '',
    meta.epoch !== undefined ? `epoch=${meta.epoch}` : '',
    meta.tier ? `tier=${meta.tier}` : '',
    `engine=js sync=manifest`,
  ]
    .filter(Boolean)
    .join(' ');

  const lines: string[] = [header];
  for (const block of visible.slice(0, MAX_PROJECT_BLOCKS)) {
    if (block.kind === 'p') {
      const label = block.heading > 0 ? `h${block.heading}` : 'p';
      lines.push(`[${block.docxIndex}] ${label} ${oneLine(block.text)}`);
      continue;
    }
    if (block.kind === 'tbl') {
      const { rows, cols } = tableSizeOf(block.xml);
      lines.push(`[${block.docxIndex}] table ${rows}x${cols}`);
      continue;
    }
    lines.push(`[${block.docxIndex}] ${block.kind}`);
  }
  if (visible.length > MAX_PROJECT_BLOCKS)
    lines.push(
      `… 还有 ${visible.length - MAX_PROJECT_BLOCKS} 个块未列出（用 docx_guide / findReplace 或分段读取）`,
    );
  lines.push(
    `# 用 docx_apply(ref,{commands:[…]}) 修改；[n] 即命令的 index；先 guide 再写；写前 dryRun`,
  );
  return lines.join('\n');
}

export function projectParsed(parsed: ParsedDocx, meta: ProjectMeta = {}): string {
  return projectModel(parsed.blocks, meta);
}

/** 结构化查询（`store` / `query` 工具）：blocks / sections / headings。 */
export interface QueryInput {
  select?: string[];
  where?: Record<string, unknown>;
  limit?: number;
}

export function queryModel(
  model: readonly ModelBlock[],
  query: QueryInput = {},
): unknown {
  const limit = Math.max(1, Math.min(2_000, Number(query.limit ?? 200)));
  const where = query.where ?? {};
  let rows = model
    .filter((block) => !block.hidden)
    .map((block) => ({
      index: block.docxIndex,
      type: block.kind,
      heading: block.heading,
      text: block.text,
      length: block.text.length,
    }));
  if (typeof where['type'] === 'string')
    rows = rows.filter((row) => row.type === where['type']);
  if (typeof where['heading'] === 'number')
    rows = rows.filter((row) => row.heading === where['heading']);
  if (typeof where['text'] === 'string')
    rows = rows.filter((row) => row.text.includes(String(where['text'])));
  const selected = Array.isArray(query.select) && query.select.length > 0
    ? rows.map((row) => {
        const out: Record<string, unknown> = {};
        for (const key of query.select!) if (key in row) out[key] = row[key as keyof typeof row];
        return out;
      })
    : rows.slice(0, limit);
  return { total: rows.length, rows: selected };
}

function oneLine(text: string): string {
  const flat = text.replace(/\s+/g, ' ').trim();
  return flat.length > MAX_PROJECT_TEXT
    ? `${flat.slice(0, MAX_PROJECT_TEXT)}…`
    : flat;
}

function short(version: string): string {
  return version.length > 12 ? `${version.slice(0, 12)}…` : version;
}
