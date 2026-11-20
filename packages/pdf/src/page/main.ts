// ─── PDF 页面入口 ────────────────────────────────────────────────────
//
// 1. 宿主能力自检（page.get / page.bytes / bus）
// 2. page.get 取元数据；page.bytes 分块取字节（content: pulled）
// 3. 注入 __pdfBoot，启动查看器
// 4. 页面只有查看与打印：不调用 saveBytes / commit / save。

import { assembleBytes, MAX_PDF_BYTES } from '../shared/chunk'
import type { PdfBoot, PdfPageMeta } from '../shared/pdf-types'
import type { PageSessionInfo } from './env'

const CHUNK = 1024 * 1024

function fail(message: string): void {
  const boot = document.querySelector<HTMLElement>('.ws-boot')
  if (!boot) return
  boot.dataset.error = '1'
  boot.textContent = message
}

function assertHost(): void {
  const page = (window.extension as { ura?: { page?: Record<string, unknown> } })?.ura?.page
  const bus = (window.extension as { bus?: Record<string, unknown> })?.bus
  const missing = ['get', 'bytes'].filter((name) => typeof page?.[name] !== 'function')
  if (typeof bus?.loaded !== 'function') missing.push('bus.loaded')
  if (typeof bus?.send !== 'function') missing.push('bus.send')
  const ura = (window.extension as { ura?: Record<string, unknown> })?.ura
  if (typeof ura?.onChange !== 'function') missing.push('ura.onChange')
  if (missing.length > 0) {
    throw new Error(
      `宿主版本过低：缺少 ${missing.map((name) => `ura.${name}`).join(' / ')}。\n` +
        'PDF 只读查看器需要 apiVersion 2 的 WorkStudio（content: pulled + 分块读）。',
    )
  }
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
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

async function fetchBytes(sessionId: string, size: number): Promise<Uint8Array> {
  return assembleBytes(
    (range) => window.extension.ura.page.bytes(sessionId, range),
    { size, chunkBytes: CHUNK, maxBytes: MAX_PDF_BYTES },
  )
}

async function metaOf(sessionId: string, fallbackSize: number): Promise<{
  data: PdfPageMeta
  version: string
  size: number
  degraded: boolean
}> {
  const info = await window.extension.ura.page.get<PageSessionInfo>(sessionId)
  const size = info.size ?? info.data.size ?? fallbackSize
  return { data: info.data, version: info.version, size, degraded: info.degraded }
}

async function boot(): Promise<void> {
  const sessionId = new URLSearchParams(location.search).get('session')
  if (!sessionId) {
    fail('缺少页面会话：请从 WorkStudio 资源面板打开 PDF 文档')
    return
  }

  try {
    assertHost()
    applyTheme(await currentTheme())
    window.extension.bus.on((message) => {
      if (message.type === 'theme') applyTheme(message.theme === 'dark' ? 'dark' : 'light')
    })

    const first = await metaOf(sessionId, 0)
    const bytes = await fetchBytes(sessionId, first.size)
    const bootData: PdfBoot = {
      sessionId,
      data: first.data,
      bytes,
      version: first.version,
      degraded: first.degraded,
      refetch: async () => {
        const latest = await metaOf(sessionId, first.size)
        return fetchBytes(sessionId, latest.size)
      },
      reload: async () => {
        const latest = await metaOf(sessionId, first.size)
        const next = await fetchBytes(sessionId, latest.size)
        return { bytes: next, data: latest.data, version: latest.version }
      },
    }
    window.__pdfBoot = bootData
    await import('../viewer/main')
  } catch (error) {
    fail(`PDF 打开失败：${error instanceof Error ? error.message : String(error)}`)
  }
}

void boot()
