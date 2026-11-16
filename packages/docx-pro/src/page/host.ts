// ─── GenOffice Docs 页面 ↔ WorkStudio 页面会话适配层 ─────────────────
//
// 这是两套交互模型的唯一接缝，不是 GenOffice 主进程的复制：
//   GenOffice 打开 = 文件路径 → 这里 = page.get / page.bytes 的字节
//   GenOffice 保存 = 写路径 → 这里 = ura.page.commit（计划型写：
//     payload.type='plan', plan={kind:'bytes'}，exec 是唯一字节产生者）
//   GenOffice 对话框 / 菜单 → 这里 = 不承载，交给 WorkStudio 壳层
//
// 与 packages/docx/page/host.ts 的唯一区别：写通道从 saveBytes 换成 commit，
// 因此人写与 agent 写共用一条提交函数，version / epoch 永远同源。

import type { PageSessionInfo, Receipt } from './env'

export interface DocxHostInput {
  sessionId: string
  ref: string
  name: string
  bytes: Uint8Array
  blobUrl: string
  hash: string
  version: string
  epoch: number
  /** 提交回执（version / epoch / applied / rejected）；agent 桥用它回执。 */
  onReceipt?: (receipt: Receipt) => void
}

export interface PageWriteState {
  version: string
  epoch: number
  /** 下一次提交的来源：agent 批次期间为 'agent'。 */
  origin: 'page' | 'agent'
  lastReceipt: Receipt | null
  lastError: string | null
}

type Listener<T> = (value: T) => void

async function saveText(
  defaultName: string,
  text: string,
  mime: string,
): Promise<{ ok: boolean; path?: string; error?: string }> {
  const picker = (
    window as unknown as {
      showSaveFilePicker?: (options: {
        suggestedName?: string
        types?: Array<{ description: string; accept: Record<string, string[]> }>
      }) => Promise<{
        name: string
        createWritable(): Promise<{
          write(data: string): Promise<void>
          close(): Promise<void>
        }>
      }>
    }
  ).showSaveFilePicker
  if (picker) {
    try {
      const handle = await picker({
        suggestedName: defaultName,
        types: [{ description: 'Document', accept: { [mime]: [`.${defaultName.split('.').pop() ?? 'txt'}`] } }],
      })
      const writable = await handle.createWritable()
      await writable.write(text)
      await writable.close()
      return { ok: true, path: handle.name }
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return { ok: false }
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  }
  const blob = new Blob([text], { type: mime })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = defaultName
  anchor.click()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
  return { ok: true, path: defaultName }
}

export function installDocxHost(input: DocxHostInput) {
  const { ura, bus } = window.extension
  const themeListeners = new Set<Listener<'light' | 'dark'>>()
  const path = input.ref
  let pendingConsumed = false
  let currentPath = path

  const theme = (): 'light' | 'dark' =>
    document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light'

  const setTheme = (next: 'light' | 'dark'): void => {
    document.documentElement.dataset.theme = next
    document.documentElement.style.colorScheme = next
    for (const listener of themeListeners) listener(next)
  }

  bus.on((message) => {
    if (message.type === 'theme') {
      setTheme(message.theme === 'dark' ? 'dark' : 'light')
    }
  })

  // 编辑区输入 → 壳层脏态；保存成功后再清掉。GenOffice 自己不向宿主
  // 上报 dirty，这是两边交互模型差异需要补的一小段胶水。
  document.addEventListener(
    'input',
    (event) => {
      const target = event.target
      if (!(target instanceof Element)) return
      if (target.closest('.ProseMirror, [contenteditable="true"]')) {
        bus.send({ type: 'dirty', dirty: true })
      }
    },
    true,
  )

  const write: PageWriteState = {
    version: input.version,
    epoch: input.epoch,
    origin: 'page',
    lastReceipt: null,
    lastError: null,
  }

  async function commit(
    data: ArrayBuffer,
  ): Promise<{ ok: boolean; error?: string }> {
    try {
      const bytes = new Uint8Array(data)
      if (bytes.byteLength === 0) return { ok: false, error: '序列化产出空字节' }
      const receipt = await ura.page.commit<Receipt>(input.sessionId, {
        txnId: `docx-pro-${Date.now()}-${Math.random().toString(16).slice(2)}`,
        baseVersion: write.version,
        origin: write.origin,
        payload: { type: 'plan', plan: { kind: 'bytes', bytes } },
      })
      write.version = receipt.version
      write.epoch = receipt.epoch
      write.lastReceipt = receipt
      write.lastError = null
      if (write.origin === 'page') bus.send({ type: 'dirty', dirty: false })
      input.onReceipt?.(receipt)
      return { ok: true }
    } catch (error) {
      write.lastError = error instanceof Error ? error.message : String(error)
      return { ok: false, error: write.lastError }
    }
  }

  const save = commit

  const base: Record<string, unknown> = {
    // 文档生命周期
    consumePendingOpenDocx: async () => {
      if (pendingConsumed) return null
      pendingConsumed = true
      return {
        path,
        name: input.name,
        dataUrl: input.blobUrl,
        hash: input.hash,
      }
    },
    consumeNewBlankDoc: async () => false,
    consumeAiDocContent: async () => null,
    consumeHeadlessExport: async () => null,
    headlessExportDone: () => undefined,
    openDocx: async () => null,
    openDocxPath: async () => null,
    openDocxDecrypt: async () => ({ ok: false, reason: 'unsupported' }),
    confirmDocumentReplace: async () => true,
    onOpenDocx: () => () => undefined,
    onRenamedDocx: () => () => undefined,

    // 保存：全部走 commit（计划型写）
    saveDocx: async (_path: string, data: ArrayBuffer) => commit(data),
    saveDocxAs: async (_name: string, data: ArrayBuffer) => {
      const result = await commit(data)
      return result.ok ? { ok: true, path: currentPath } : result
    },
    saveDocxNew: async (_name: string, data: ArrayBuffer) => {
      const result = await commit(data)
      return result.ok ? { ok: true, path: currentPath } : result
    },
    saveDocxTo: async (_path: string, data: ArrayBuffer) => {
      const result = await commit(data)
      return result.ok ? { ok: true, path: currentPath } : result
    },
    writeRecoveryCopy: async () => ({ ok: true }),
    setDocPassword: async () => ({ ok: true }),
    docPasswordIntentRevision: async () => 0,
    discardDocPasswordIntents: async () => ({ ok: true }),

    // 主题 / 语言 / 偏好
    getLanguage: async () => 'zh',
    onLanguageChanged: () => () => undefined,
    getTheme: async () => theme(),
    onThemeChanged: (handler: Listener<'light' | 'dark'>) => {
      themeListeners.add(handler)
      return () => themeListeners.delete(handler)
    },
    getDocumentTheme: async () => 'follow',
    onDocumentThemeChanged: () => () => undefined,
    getAiPanelPrefs: async () => ({
      side: 'right',
      fontSize: 'default',
      customFontSize: 14,
      spellcheck: false,
      openInNewDocs: false,
      showAiSidebar: false,
    }),
    onAiPanelPrefsChanged: () => () => undefined,
    getAuthorName: async () => '',
    onAuthorNameChanged: () => () => undefined,
    getAutoSaveDefault: async () => ({ on: true, updatedAt: 1 }),
    onAutoSaveDefaultChanged: () => () => undefined,

    // AI 关闭：设置面板不显示，历史记录也不需要 projectApi
    aiGskStatus: async () => ({ loggedIn: false }),
    aiGskLogin: async () => undefined,
    kbList: async () => [],
    kbRead: async (_path: string, offset: number) => ({
      text: '',
      totalChars: 0,
      offset,
    }),
    getAiSettings: async () => ({
      provider: 'custom',
      providers: {},
      customEndpoints: [],
      gskToolsEnabled: false,
    }),
    setAiSettings: async () => undefined,
    onAiSettingsChanged: () => () => undefined,

    // 最近文件 / 附件 / 图片
    getRecentFiles: async () => [],
    getPathForFile: (file: File) => file.name,
    fetchImage: async (src: string) => src,
    pickImage: () =>
      new Promise<{ base64: string; mime: string; name: string } | null>((resolve) => {
        const input = document.createElement('input')
        input.type = 'file'
        input.accept = 'image/png,image/jpeg,image/gif'
        input.addEventListener('change', () => {
          const file = input.files?.[0]
          if (!file) return resolve(null)
          const reader = new FileReader()
          reader.onload = () => {
            const result = String(reader.result ?? '')
            const comma = result.indexOf(',')
            resolve({
              base64: comma >= 0 ? result.slice(comma + 1) : '',
              mime: file.type || 'image/png',
              name: file.name,
            })
          }
          reader.onerror = () => resolve(null)
          reader.readAsDataURL(file)
        })
        input.click()
      }),
    pickAttachments: async () => [],
    addAttachmentPaths: async () => [],
    readAttachment: async () => null,
    readAttachmentImage: async () => null,
    onViewImage: () => () => undefined,
    saveImageAs: async () => undefined,
    writeExportImage: async () => ({ ok: true }),
    pickExportImagesTarget: async () => null,
    takeExportPdf: async () => null,

    // 打印 / 导出：HTML / Markdown 走浏览器文件保存；PDF 仍交给打印链路
    print: () => window.print(),
    printPdfBuffer: async () => ({ ok: true }),
    saveMergedPdf: async () => ({ ok: true }),
    exportHtml: async (defaultName: string, html: string) => {
      const name = defaultName.replace(/\.docx$/i, '.html')
      return saveText(name, html, 'text/html')
    },
    exportMarkdown: async (
      defaultName: string,
      markdown: string,
      images: unknown[],
    ) => {
      const name = defaultName.replace(/\.docx$/i, '.md')
      const result = await saveText(name, markdown, 'text/markdown')
      return { ...result, imagesDropped: images.length }
    },

    // 关闭 / 菜单 / 上下文菜单：宿主不承载这些 chrome，交给页面自身
    onCloseCheck: () => () => undefined,
    onCloseSaveRequest: () => () => undefined,
    reportCloseCheck: () => undefined,
    reportCloseSaveResult: () => undefined,
    reportViewMenuState: () => undefined,
    onMenuCommand: () => () => undefined,
    onMcpCommand: () => () => undefined,
    reportMcpResult: () => undefined,
    signalMcpReady: () => undefined,
    onContextMenuRequest: () => () => undefined,
    armContextMenu: () => undefined,
    claimContextMenu: () => undefined,
    onTeardown: () => () => undefined,

    // 其它可选能力
    convertAltChunkHtml: async () => null,
    openSourcePath: () => undefined,
    openNewTab: () => undefined,
    spellDiag: () => undefined,
    onZoteroRequest: () => () => undefined,
    respondToZotero: () => undefined,
    zoteroCommand: async () => ({ ok: true }),
    webSearch: async () => [],
    skillsList: async () => [],
    skillText: async () => '',
    runCli: async () => ({ ok: true }),
  }

  // 只暴露 WorkStudio 页面会话承载的桌面语义；未实现项属于宿主 chrome，
  // 由壳层承担，不进入 GenOffice 页面。
  window.desktop = base as unknown as typeof window.desktop

  return {
    write,
    /** agent 批次期间切换提交来源（页面 flush 用 'page'，agent 应用用 'agent'）。 */
    setOrigin(origin: 'page' | 'agent') {
      write.origin = origin
    },
    getVersion: () => write.version,
    getEpoch: () => write.epoch,
    lastReceipt: () => write.lastReceipt,
    lastError: () => write.lastError,
  }
}

export type DocxHost = ReturnType<typeof installDocxHost>

export function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

export function hashOf(bytes: Uint8Array): string {
  let hash = 2166136261
  for (let i = 0; i < bytes.length; i++) {
    hash ^= bytes[i]!
    hash = Math.imul(hash, 16777619)
  }
  return `ws-${(hash >>> 0).toString(16)}`
}

/** 页面元数据（`content: 'pulled'`：整份字节不进 `get` 的 data）。 */
export function docxMetaOf(info: PageSessionInfo): { name: string; size: number } {
  const data = info.data as { name?: string; size?: number } | null
  return {
    name: data?.name || 'document.docx',
    size: Number(data?.size ?? 0),
  }
}
