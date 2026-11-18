// ─── pptx-pro 页面入口 ────────────────────────────────────────────────
//
// 1. 宿主能力自检（bytes / commit / agentResult）
// 2. page.get 取元数据；page.bytes 分块取字节（content: pulled）
// 3. 注入 __pptxProBoot，启动本包 editor 副本
// 4. 安装 agent 桥（页面只做 dry-run / commit；实际 mutate 由 exec 执行）

import { installAgentBridge } from './agent'
import type { PageSessionInfo, PptxProBoot } from './env'

const CHUNK = 1024 * 1024

function fail(message: string): void {
  const boot = document.querySelector<HTMLElement>('.ws-boot')
  if (!boot) return
  boot.dataset.error = '1'
  boot.textContent = message
}

function assertHost(): void {
  const page = (window.extension as { ura?: { page?: Record<string, unknown> } })?.ura?.page
  const missing = ['bytes', 'commit', 'agentResult'].filter(
    (name) => typeof page?.[name] !== 'function',
  )
  if (missing.length > 0)
    throw new Error(
      `宿主版本过低：缺少 ura.page.${missing.join(' / ')}。\n` +
        'pptx-pro 需要 apiVersion 2 的 WorkStudio（含 pulled bytes、计划型提交与页面代理）。',
    )
}

async function fetchBytes(sessionId: string, size: number): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  let offset = 0
  for (;;) {
    const chunk = await window.extension.ura.page.bytes(sessionId, {
      start: offset,
      end: offset + CHUNK,
    })
    const bytes = chunk instanceof Uint8Array ? chunk : new Uint8Array(chunk as ArrayBufferLike)
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

async function boot(): Promise<void> {
  const sessionId = new URLSearchParams(location.search).get('session')
  if (!sessionId) {
    fail('缺少页面会话：请从 WorkStudio 资源面板打开 PPTX 演示文稿')
    return
  }

  try {
    assertHost()
    const info = await window.extension.ura.page.get<PageSessionInfo>(sessionId)
    const bytes = await fetchBytes(sessionId, info.size ?? info.data.size ?? 0)
    const bootData: PptxProBoot = {
      sessionId,
      data: info.data,
      bytes,
      version: info.version,
      epoch: info.epoch ?? 0,
    }
    window.__pptxProBoot = bootData
    await import('../editor/main')
    installAgentBridge({ sessionId })
    document.querySelector<HTMLElement>('.ws-boot')?.remove()
  } catch (error) {
    fail(`PPTX 打开失败：${error instanceof Error ? error.message : String(error)}`)
  }
}

void boot()
