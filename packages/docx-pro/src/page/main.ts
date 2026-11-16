// ─── docx-pro 页面入口 ─────────────────────────────────────────────────
//
// 页面只做四件事：
//   1. 从页面会话取元数据 + 分块取字节（content: 'pulled'）
//   2. 安装 window.desktop 宿主适配层（保存 = 计划型 commit）
//   3. 安装页面代理（agent-request → 编辑器事务 → 同一条 commit）
//   4. 启动本包自己的 docs 编辑器副本（src/editor，从 packages/docx fork 而来）
//
// 字节只经 commit 写回；路径、版本、冲突与串行写归 WorkStudio 的会话层。

import { hashOf, installDocxHost } from './host'
import { docxMetaOf } from './host'
import { installAgentBridge } from './agent'
import type { PageSessionInfo } from './env'

const CHUNK = 1024 * 1024

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
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

async function fetchBytes(sessionId: string, size: number): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  let offset = 0
  for (;;) {
    const chunk = await window.extension.ura.page.bytes(sessionId, {
      start: offset,
      end: offset + CHUNK,
    })
    const bytes =
      chunk instanceof Uint8Array ? chunk : new Uint8Array(chunk as ArrayBufferLike)
    if (bytes.byteLength === 0) break
    chunks.push(bytes)
    offset += bytes.byteLength
    if (size > 0 && offset >= size) break
    if (chunks.length > 512) break
  }
  if (chunks.length === 0) throw new Error('页面会话没有返回内容')
  if (chunks.length === 1) return chunks[0]!
  const total = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0)
  const all = new Uint8Array(total)
  let at = 0
  for (const chunk of chunks) {
    all.set(chunk, at)
    at += chunk.byteLength
  }
  return all
}

/**
 * 宿主能力自检：docx-pro 是 apiVersion 2 插件（计划型写 + 分块读 + 页面代理）。
 * 老宿主（0.2.x）的 ura.page 只有 get/save/saveBytes，没有 bytes/commit/agentResult，
 * 直接跑会在取内容时抛 “undefined is not a function”，页面白屏且难诊断。
 * 这里提前判一次，把「无视图 / 白屏」变成一条明确的版本提示。
 */
function assertHost(): void {
  const page = (window.extension as { ura?: { page?: Record<string, unknown> } })
    ?.ura?.page
  const missing = ['bytes', 'commit', 'agentResult'].filter(
    (name) => typeof page?.[name] !== 'function',
  )
  if (missing.length > 0)
    throw new Error(
      `宿主版本过低：缺少 ura.page.${missing.join(' / ')}。\n` +
        'docx-pro 需要 apiVersion 2 的 WorkStudio（0.3.x+，含计划型写与页面代理）。\n' +
        '在这个宿主上请改用 apiVersion 1 的 docx 插件。',
    )
}

async function waitForEditor(): Promise<void> {
  const deadline = Date.now() + 30_000
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
    assertHost()
    applyTheme(await currentTheme())
    const info = await window.extension.ura.page.get<PageSessionInfo>(sessionId)
    const { name, size } = docxMetaOf(info)
    const bytes = await fetchBytes(sessionId, size ?? info.size ?? 0)
    const blobUrl = URL.createObjectURL(
      new Blob([bytes as unknown as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    )

    const host = installDocxHost({
      sessionId,
      ref: info.ref,
      name,
      bytes,
      blobUrl,
      hash: hashOf(bytes),
      version: info.version,
      epoch: info.epoch ?? 0,
      onReceipt: (receipt) => {
        window.extension.bus.send({
          type: 'docx-receipt',
          version: receipt.version,
          epoch: receipt.epoch,
          applied: receipt.applied,
        })
      },
    })

    installAgentBridge({ sessionId, host })

    // 编辑器源码是本包自带的副本（src/editor）；这里只接会话与宿主边界。
    await import('../editor/main')
    await waitForEditor()
    window.extension.bus.loaded()
    window.extension.bus.send({ type: 'dirty', dirty: false })
    document.querySelector<HTMLElement>('.ws-boot')?.remove()
  } catch (error) {
    fail(`DOCX 打开失败：${error instanceof Error ? error.message : String(error)}`)
  }
}

void boot()
