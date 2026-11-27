// ─── 页面入口 ─────────────────────────────────────────────────────────
//
// 1. `?ref=` → pptx.load 一次拿整份字节（没有分块、没有会话 id）
// 2. 装 Slides host：保存 / 另存 / 上报都接到页面通道
// 3. 挂载编辑器 → attach 领取租约 → 订阅 events → 处理 pending
//
// 页面是唯一写入者：编辑与助手的写都在这里落盘（pptx.save）。

// pptx-engine 在页面里依赖 node:Buffer / node:crypto；shim 必须先于任何
// pptx-engine 的调用执行（openPptx 在 host 安装时就跑）。
import '../editor/shims/buffer-global'

import { channel } from './channel'
import { installWorkStudioSlidesHost, type SlidesHost } from './host'
import { installPageQueue } from './lease'

function fail(message: string): void {
  const boot = document.querySelector<HTMLElement>('.ws-boot')
  if (!boot) return
  boot.dataset.error = '1'
  boot.textContent = message
}

function applyTheme(theme: 'light' | 'dark'): void {
  document.documentElement.dataset.theme = theme
  document.documentElement.style.colorScheme = theme
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

async function boot(): Promise<void> {
  const ref = new URLSearchParams(location.search).get('ref')
  if (!ref) {
    fail('缺少文件地址：请从工作区打开演示文稿')
    return
  }

  const pageId = crypto.randomUUID()
  let theme: 'light' | 'dark' = 'light'
  try {
    theme = await channel.theme()
  } catch {
    /* 主题拿不到不影响打开 */
  }
  applyTheme(theme)

  let loaded
  try {
    loaded = await channel.load(ref)
  } catch (error) {
    fail(`演示文稿打开失败：${messageOf(error)}`)
    return
  }
  if (!loaded.bytes || loaded.bytes.byteLength === 0) {
    fail('演示文稿没有内容')
    return
  }

  window.__pptxProBoot = {
    ref: loaded.ref,
    name: loaded.name,
    pageId,
    hash: loaded.hash,
    bytes: loaded.bytes,
  }

  // ── 保存 / 另存 / 上报：全部落到页面通道 ─────────────────────────
  let pendingReport: ReturnType<typeof setTimeout> | null = null
  let lastReport: { docRev: number; dirty: boolean } = { docRev: 0, dirty: false }

  const report = (state: { hash?: string; docRev: number; dirty: boolean }): void => {
    lastReport = { docRev: state.docRev, dirty: state.dirty }
    if (pendingReport) return
    pendingReport = setTimeout(() => {
      pendingReport = null
      void channel
        .report({
          ref: loaded.ref,
          pageId,
          ...(state.hash ? { hash: state.hash } : {}),
          docRev: lastReport.docRev,
          dirty: lastReport.dirty,
        })
        .catch(() => undefined)
    }, 400)
  }

  const themeHandlers = new Set<(next: 'light' | 'dark') => void>()
  channel.onTheme((next) => {
    theme = next
    applyTheme(next)
    for (const handler of themeHandlers) handler(next)
  })

  let host: SlidesHost
  try {
    host = await installWorkStudioSlidesHost({
      ref: loaded.ref,
      name: loaded.name,
      pageId,
      bytes: loaded.bytes,
      hash: loaded.hash,
      save: async (bytes, baseHash) => {
        const receipt = await channel.save({ ref: loaded.ref, pageId, bytes, baseHash })
        return { hash: receipt.hash, size: receipt.size }
      },
      saveAs: async (bytes) => {
        const token = await channel.pickSaveFile(loaded.name)
        if (!token) return null
        const receipt = await channel.saveAs({ ref: loaded.ref, pageId, token, bytes })
        return { hash: receipt.hash, size: receipt.size }
      },
      report,
      theme: () => theme,
      onTheme: (handler) => {
        themeHandlers.add(handler)
        return () => themeHandlers.delete(handler)
      },
    })
  } catch (error) {
    fail(`演示文稿解析失败：${messageOf(error)}`)
    return
  }

  window.__pptxProHost = host as unknown as Window['__pptxProHost']

  try {
    const { mountEditor } = await import('../editor/main')
    mountEditor()
  } catch (error) {
    fail(`编辑器加载失败：${messageOf(error)}`)
    return
  }

  // ── 租约 + 请求队列 ───────────────────────────────────────────────
  let queue: ReturnType<typeof installPageQueue> | null = null
  try {
    const attached = await channel.attach({ ref: loaded.ref, pageId })
    queue = installPageQueue({ ref: loaded.ref, pageId, host })
    for (const request of attached.pending) void queue.handle(request.id)
  } catch (error) {
    fail(`页面登记失败：${messageOf(error)}`)
    return
  }

  window.addEventListener('beforeunload', () => {
    queue?.stop()
  })

  document.querySelector<HTMLElement>('.ws-boot')?.remove()
}

void boot()
