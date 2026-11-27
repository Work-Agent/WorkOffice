// ─── docx 页面入口（WorkStudio SDK） ───────────────────────────────
//
// 页面只做：取 ref → 取字节 → 装 window.desktop 适配层 → 启动编辑器
// → 登记页面租约 → 处理 Agent 请求 / 保存回执。
//
// 不订阅 workspace、不管理标签 / 文件列表；这些归宿主 shell。

import { installDocxHost, hashOf } from './host'
import { installPageQueue } from './lease'
import type { LoadResult, PageRequest } from './env'

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

async function watchTheme(): Promise<void> {
  try {
    const snapshot = await window.workStudio.subscribe('shell', 'appearance')
    const theme = (snapshot as { theme?: unknown } | null)?.theme
    applyTheme(theme === 'dark' ? 'dark' : 'light')
  } catch {
    applyTheme(matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light')
  }
  window.workStudio.onUpdate((update) => {
    if (update.domain !== 'shell' || update.resource !== 'appearance') return
    if (update.type !== 'snapshot' && update.type !== 'patch') return
    const theme = (update.data as { theme?: unknown }).theme
    if (theme === 'light' || theme === 'dark') applyTheme(theme)
  })
}

async function waitForEditor(): Promise<void> {
  const deadline = Date.now() + 30_000
  while (Date.now() < deadline) {
    if (document.querySelector('.app')) return
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  throw new Error('编辑器初始化超时')
}

async function loadDocument(ref: string): Promise<LoadResult> {
  const result = await window.workStudio.command<LoadResult>('docx', 'docx', 'load', {
    ref,
  })
  if (!result.ok) throw new Error(`${result.error.code}: ${result.error.message}`)
  if (result.data.bytes instanceof Uint8Array) return result.data
  // 极端情况下 invoke 没有保留 typed array：退到 base64。
  const fallback = await window.workStudio.command<LoadResult>('docx', 'docx', 'load', {
    ref,
    base64: true,
  })
  if (!fallback.ok) throw new Error(`${fallback.error.code}: ${fallback.error.message}`)
  return fallback.data
}

async function boot(): Promise<void> {
  const params = new URLSearchParams(location.search)
  const ref = params.get('ref') ?? ''
  const locator = params.get('locator') ?? ''
  if (!ref) {
    fail('缺少 ref：请从 WorkStudio 资源面板打开 DOCX 文档')
    return
  }

  try {
    await watchTheme()
    const loaded = await loadDocument(ref)
    const bytes = loaded.bytes instanceof Uint8Array
      ? loaded.bytes
      : new Uint8Array(
          atob(loaded.data ?? '')
            .split('')
            .map((char) => char.charCodeAt(0)),
        )
    const name = loaded.name || 'document.docx'
    const blobUrl = URL.createObjectURL(
      new Blob([bytes as unknown as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    )
    const pageId = crypto.randomUUID()
    const host = installDocxHost({
      ref,
      pageId,
      name,
      bytes,
      blobUrl,
      hash: loaded.hash || hashOf(bytes),
      version: loaded.hash || hashOf(bytes),
      epoch: 0,
    })
    const queue = installPageQueue({ ref, pageId, host, locator })

    const attached = await window.workStudio.command<{ pending?: PageRequest[] }>(
      'docx',
      'docx',
      'attach',
      { ref, pageId },
    )
    if (!attached.ok) throw new Error(`${attached.error.code}: ${attached.error.message}`)
    queue.touch()

    await import('../editor/main')
    await waitForEditor()
    void queue.handlePending(attached.data?.pending ?? [])
    document.querySelector<HTMLElement>('.ws-boot')?.remove()
  } catch (error) {
    fail(`DOCX 打开失败：${error instanceof Error ? error.message : String(error)}`)
  }
}

void boot()
