// ─── DOCX 插件页面入口 ──────────────────────────────────────────────
//
// 页面只做三件事：从页面会话取字节、安装 window.desktop 宿主适配层、
// 启动 GenOffice Docs 渲染器。文档生命周期由插件注册入口与页面会话负责。

import {
  docxInfoOf,
  hashOf,
  installDocxHost,
} from './host'
import type { PageSessionInfo } from './env'

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

async function currentTheme(): Promise<'light' | 'dark'> {
  try {
    const theme = await window.appBridge?.peek<string>('/theme')
    if (theme === 'dark' || theme === 'light') return theme
  } catch {
    /* 宿主未就绪时按系统主题兜底 */
  }
  return window.matchMedia('(prefers-color-scheme: dark)').matches
    ? 'dark'
    : 'light'
}

async function waitForEditor(): Promise<void> {
  const deadline = Date.now() + 20_000
  while (Date.now() < deadline) {
    if (document.querySelector('.app')) return
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  throw new Error('编辑器初始化超时')
}

async function boot(): Promise<void> {
  const sessionId = new URLSearchParams(location.search).get('session')
  if (!sessionId) {
    fail('缺少页面会话：请从 WorkStudio 资源面板打开 DOCX 文档')
    return
  }

  try {
    applyTheme(await currentTheme())
    const info = await window.extension.ura.page.get<PageSessionInfo>(sessionId)
    const { bytes, name } = docxInfoOf(info)
    const blobUrl = URL.createObjectURL(
      new Blob([bytes as unknown as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    )

    installDocxHost({
      sessionId,
      ref: info.ref,
      name,
      bytes,
      blobUrl,
      hash: hashOf(bytes),
    })

    // 页面层源码已 fork 到 src/editor；这里只负责会话数据与宿主边界。
    await import('../editor/main')
    await waitForEditor()
    window.extension.bus.loaded()
    window.extension.bus.send({ type: 'dirty', dirty: false })
    document.querySelector<HTMLElement>('.ws-boot')?.remove()
  } catch (error) {
    fail(
      `DOCX 打开失败：${error instanceof Error ? error.message : String(error)}`,
    )
  }
}

void boot()
