// ─── agent 批次的解包（纯函数，页面与单测共用） ─────────────────────────
//
// 宿主把工具调用透传给页面时，commands 有两种形状：
//   - 直接是命令数组：`[{ op: 'setText', index: 1, text }]`
//   - 工具调用包装：`[{ tool: 'docx.apply', params: { commands: [...] } }]`
// 这里统一拆平成 `DocxCommand[]`（无法识别的项丢弃，交给上层报空批次）。

import type { DocxCommand } from './types';

export function unwrapAgentCommands(raw: unknown): DocxCommand[] {
  if (!Array.isArray(raw)) return [];
  const out: DocxCommand[] = [];
  for (const item of raw) {
    if (!item || typeof item !== 'object') continue;
    const value = item as { op?: unknown; params?: unknown };
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
