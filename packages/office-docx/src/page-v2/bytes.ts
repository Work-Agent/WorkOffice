// ─── 页面会话的字节读取（内容自取：`get` + `bytes` 分块） ─────────────
//
// 页面打开、提交后刷新基线、探针都走这里：一次 `get` 拿 `size`，再按 1 MB
// 窗口循环读。返回的缓冲严格等于 `size`（末尾块不足时截断）。

import type { PageSessionInfo } from '../page/env';

const CHUNK = 1024 * 1024;
const LIMIT = 512 * 1024 * 1024;

export async function readSessionBytes(sessionId: string): Promise<Uint8Array> {
  const info = await window.extension.ura.page.get<PageSessionInfo>(sessionId);
  return readBytesWindow(sessionId, info.size);
}

export async function readBytesWindow(sessionId: string, size?: number): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let offset = 0;
  for (;;) {
    const end = size && size > 0 ? Math.min(size, offset + CHUNK) : offset + CHUNK;
    const chunk = await window.extension.ura.page.bytes(sessionId, { start: offset, end });
    if (!chunk || chunk.byteLength === 0) break;
    chunks.push(chunk);
    offset += chunk.byteLength;
    if (size && size > 0 && offset >= size) break;
    if (chunk.byteLength < CHUNK) break;
    if (offset > LIMIT) throw new Error('文档超过 512 MB 上限');
  }
  const out = new Uint8Array(offset);
  let cursor = 0;
  for (const chunk of chunks) {
    out.set(chunk, cursor);
    cursor += chunk.byteLength;
  }
  return out;
}
