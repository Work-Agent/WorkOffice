// ─── office-docx 页面入口 v2（P1：挂载 docs 渲染器，只读） ─────────────
//
// 数据通路与旧的 `src/page/main.ts` 一致（内容自取：`get` + `bytes` 分块），
// 但渲染层换成 `./genoffice/editor`（GenOffice Docs 渲染器 fork，含 Ribbon
// 工具栏与分页）。写盘在 P2 接到 `/commit` 的 plan 通道。
//
// 渲染器是**复制**进来的（不从 packages/docx 跨包引用）：两个插件最终只留一个，
// 复制避免长期耦合；代价是渲染器补丁只在本包内打（P3/P5 都在这里改）。
//
// 页面只做三件事：
//   1. 取会话元数据与字节（1 MB 分块）
//   2. 安装 `window.desktop` 适配层（host.ts，从 v1 移植）
//   3. 启动渲染器并回报 loaded

import { readSessionBytes } from './bytes';
import { flushDocument, handleAgentRequest } from './agent';
import { hashOf, hostControls, installDocxHost } from './host';
import type { BusMessage, PageSessionInfo } from '../page/env';

const SESSION_ID = new URLSearchParams(location.search).get('session');

function fail(message: string): void {
  const root = document.getElementById('root');
  if (!root) return;
  const box = document.createElement('div');
  box.className = 'boot';
  box.dataset.error = '1';
  box.textContent = message;
  root.replaceChildren(box);
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

function applyTheme(theme: 'light' | 'dark'): void {
  document.documentElement.dataset.theme = theme;
  document.documentElement.style.colorScheme = theme;
}

/** 关闭守卫 / Ctrl+S 的保存握手：落盘后回 `save-result`（宿主据此决定能否关标签）。 */
async function replySave(): Promise<void> {
  const ok = await flushDocument();
  window.extension.bus.send({
    type: 'save-result',
    ok,
    ...(ok ? {} : { message: '仍有未提交的修改' }),
  });
}

/**
 * 磁盘上的文件被外部改过：
 * - 没有本地改动 → 整篇重载（走渲染器的「打开文档」入口，docxIndex / 分节一起重置）；
 * - 有本地改动 → 不动。下一次提交带 `baseVersion` 会拿到 `command.stale`，
 *   页面显示"文件已被外部修改"，绝不把本地改动盖到别人的文件上（I12）。
 */
async function onExternalContent(): Promise<void> {
  const controls = hostControls();
  if (!controls || controls.isDirty()) return;
  await controls.reloadFromDisk();
}

async function waitForEditor(): Promise<void> {
  const deadline = Date.now() + 20_000;
  while (Date.now() < deadline) {
    if (document.querySelector('.app')) return;
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  throw new Error('编辑器初始化超时');
}

async function boot(): Promise<void> {
  if (!SESSION_ID) {
    fail('缺少页面会话参数（session）');
    return;
  }
  try {
    applyTheme(
      window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light',
    );
    const info = await window.extension.ura.page.get<PageSessionInfo>(SESSION_ID);
    const bytes = await readSessionBytes(SESSION_ID);
    const name = nameOf(info.ref);
    const blobUrl = URL.createObjectURL(
      new Blob([bytes as unknown as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    );

    installDocxHost({
      sessionId: SESSION_ID,
      ref: info.ref,
      name,
      bytes,
      blobUrl,
      hash: hashOf(bytes),
      version: info.version,
      epoch: info.epoch,
    });

    // 宿主 → 页面：agent 代理（P2b）/ 关闭守卫保存握手 / 外部变更重载
    const session = { ref: info.ref, name, ...(info.tier ? { tier: info.tier } : {}) };
    window.extension.bus.on((message: BusMessage) => {
      if (message.type === 'agent-request') {
        void handleAgentRequest(message, session, SESSION_ID);
        return;
      }
      if (message.type === 'action' && message.id === 'save') {
        void replySave();
        return;
      }
      if (message.type === 'content') {
        void onExternalContent();
        return;
      }
    });

    await import('./genoffice/editor/main');
    await waitForEditor();
    window.extension.bus.loaded();
    window.extension.bus.send({ type: 'dirty', dirty: false });
    document.querySelector<HTMLElement>('.boot')?.remove();
  } catch (error) {
    fail(`打开失败：${error instanceof Error ? error.message : String(error)}`);
  }
}

void boot();
