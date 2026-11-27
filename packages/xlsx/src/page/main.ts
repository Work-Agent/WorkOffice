/**
 * xlsx 页面入口。
 *
 *   1. `?ref=` → `page status` 拿哈希与体积
 *   2. `page meta`：wasm 可用时是完整元数据（流式打开），否则是薄清单
 *   3. 装 `window.__xlsxPlus`（编辑器唯一 I/O 出口）
 *   4. 流式：只注入元数据；回退：取整份字节并本地解析出工作表快照
 *   5. 启动 `@sheets/main`（GenOffice Sheets 渲染层，fork 只覆盖改过的文件）
 */

import '@work-studio/sdk/global'

import { installAgent } from './agent'
import { createBridge, setHostTheme, type EditorBridge } from './bridge'
import { Channel, PageError } from './channel'
import { askDraft } from './draft'
import { parseWorkbook } from '../shared/xlsx-parse'

/** 没有 wasm 时页面要自己解析整份字节，超过这个体积就明确拒绝。 */
const MAX_INLINE_BYTES = 64 * 1024 * 1024

function rootEl(): HTMLElement {
  const root = document.getElementById('root')
  if (!root) throw new Error('缺少 #root')
  return root
}

function showError(message: string): void {
  const root = rootEl()
  root.innerHTML = ''
  root.className = 'boot'
  root.dataset.error = '1'
  root.textContent = message
}

function watchTheme(): void {
  const bridge = window.workStudio
  if (!bridge) return
  void bridge
    .subscribe('shell', 'appearance')
    .then((snapshot) => setHostTheme((snapshot as { theme?: string } | null)?.theme))
    .catch(() => undefined)
  bridge.onUpdate((update) => {
    if (update.domain !== 'shell' || update.resource !== 'appearance') return
    if (update.type !== 'snapshot' && update.type !== 'patch') return
    const data = update.data as { theme?: string } | undefined
    if (data && 'theme' in data) setHostTheme(data.theme)
  })
}

declare global {
  interface Window {
    __xlsxPlus?: EditorBridge
  }
}

async function boot(): Promise<void> {
  const ref = new URLSearchParams(location.search).get('ref') ?? ''
  if (!ref) {
    showError('缺少 ?ref=：请从工作区打开一个 .xlsx 文件。')
    return
  }

  watchTheme()

  const pageId =
    typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID()
      : `page-${Date.now()}-${Math.random().toString(36).slice(2)}`
  const channel = new Channel(ref, pageId)

  try {
    rootEl().textContent = '正在读取工作簿…'
    const status = await channel.read<{ name: string; size: number; hash: string }>('status')

    try {
      await channel.attach()
    } catch (error) {
      if (error instanceof PageError) await channel.attach(true)
      else throw error
    }

    // 崩溃恢复副本优先：有草稿就用草稿字节当启动数据（走整份字节那条路）。
    const draft = await askDraft(channel)
    if (draft) {
      const bridge = createBridge(channel, { hash: status.hash, supportsEngineRange: false })
      window.__xlsxPlus = bridge
      const sheets = parseWorkbook(draft)
      const container = rootEl()
      container.textContent = ''
      container.removeAttribute('data-error')
      container.className = ''
      window.__xlsxProBoot = {
        sessionId: ref,
        data: { name: status.name, ref, version: status.hash, size: draft.byteLength, sheets: sheets as never },
        epoch: 0,
        bytes: draft,
      }
      await import('@sheets/main')
      installAgent(channel, bridge)
      return
    }

    const meta = await channel.read<Record<string, unknown>>('meta')
    const streaming = meta['source'] === 'wasm'
    if (!streaming && status.size > MAX_INLINE_BYTES)
      throw new PageError(
        'command.too-large',
        `工作簿 ${status.size} 字节超过页面直读上限（wasm 引擎不可用）；请先安装完整引擎`,
      )

    const bridge = createBridge(channel, { hash: status.hash, supportsEngineRange: streaming })
    window.__xlsxPlus = bridge

    let bytes: Uint8Array | undefined
    let sheets: unknown[] = []
    let version = status.hash
    let size = status.size
    if (!streaming) {
      const opened = await bridge.bytes()
      bytes = opened.bytes
      sheets = parseWorkbook(bytes)
      version = opened.hash
      size = bytes.byteLength
    } else {
      // 占位快照：真正的 sheets/styles 由宿主在启动时用 `meta` 重建；
      // 万一 meta 解析失败，至少保证工作表名/id 与引擎会话一致。
      const listed = Array.isArray(meta['sheets']) ? (meta['sheets'] as Array<Record<string, unknown>>) : []
      sheets = listed.map((sheet) => ({ id: String(sheet['id'] ?? ''), name: String(sheet['name'] ?? ''), cells: {} }))
    }

    const container = rootEl()
    container.textContent = ''
    container.removeAttribute('data-error')
    container.className = ''

    window.__xlsxProBoot = {
      sessionId: ref,
      data: {
        name: status.name,
        ref,
        version,
        size,
        sheets: sheets as never,
      },
      epoch: 0,
      ...(bytes ? { bytes } : {}),
    }

    await import('@sheets/main')
    installAgent(channel, bridge)

    const beat = setInterval(() => {
      void channel.report({ hash: bridge.hash(), dirty: false, docRev: 0 }).catch(() => undefined)
    }, 15_000)
    window.addEventListener('pagehide', () => {
      clearInterval(beat)
      channel.detach()
    })
  } catch (error) {
    const message =
      error instanceof PageError
        ? `${error.code}: ${error.message}`
        : error instanceof Error
          ? error.message
          : String(error)
    showError(`打开失败：${message}`)
  }
}

void boot()
