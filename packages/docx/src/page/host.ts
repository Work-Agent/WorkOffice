// ─── GenOffice Docs 页面 ↔ WorkStudio 页面会话适配层 ─────────────────
//
// 这是两套交互模型的唯一接缝，不是 GenOffice 主进程的复制：
//   GenOffice 打开 = 文件路径 → 这里 = page.get 的字节
//   GenOffice 保存 = 写路径 → 这里 = ura.page.saveBytes（严格版本写）
//   GenOffice 对话框 / 菜单 → 这里 = 不承载，交给 WorkStudio 壳层
//
// 页面只持有语义化会话；路径、版本、写通道都不进入 GenOffice 的交互代码。

import type { PageCommitResult, PageSessionInfo } from './env'

export interface DocxHostInput {
  sessionId: string
  ref: string
  name: string
  bytes: Uint8Array
  blobUrl: string
  hash: string
}

type Listener<T> = (value: T) => void

export function installDocxHost(input: DocxHostInput): void {
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

  async function save(data: ArrayBuffer): Promise<{ ok: boolean; error?: string }> {
    try {
      const bytes = new Uint8Array(data)
      const result = await ura.page.saveBytes<PageCommitResult>(input.sessionId, bytes)
      if (result?.ok) {
        bus.send({ type: 'dirty', dirty: false })
        return { ok: true }
      }
      if (result?.conflict) {
        return {
          ok: false,
          error: '文件已被外部修改；请重新打开文档后再保存',
        }
      }
      return { ok: false, error: '保存失败' }
    } catch (error) {
      return {
        ok: false,
        error: error instanceof Error ? error.message : String(error),
      }
    }
  }

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

    // 保存
    saveDocx: async (_path: string, data: ArrayBuffer) => save(data),
    saveDocxAs: async (_name: string, data: ArrayBuffer) => {
      const result = await save(data)
      return result.ok ? { ok: true, path: currentPath } : result
    },
    saveDocxNew: async (_name: string, data: ArrayBuffer) => {
      const result = await save(data)
      return result.ok ? { ok: true, path: currentPath } : result
    },
    saveDocxTo: async (_path: string, data: ArrayBuffer) => {
      const result = await save(data)
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
    pickImage: async () => null,
    pickAttachments: async () => [],
    addAttachmentPaths: async () => [],
    readAttachment: async () => null,
    readAttachmentImage: async () => null,
    onViewImage: () => () => undefined,
    saveImageAs: async () => undefined,
    writeExportImage: async () => ({ ok: true }),
    pickExportImagesTarget: async () => null,
    takeExportPdf: async () => null,

    // 打印 / 导出：先交给浏览器打印，其余保持无害
    print: () => window.print(),
    printPdfBuffer: async () => ({ ok: true }),
    saveMergedPdf: async () => ({ ok: true }),

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
}

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

export function docxInfoOf(info: PageSessionInfo): {
  bytes: Uint8Array
  name: string
} {
  const data = info.data as { bytes?: string; name?: string } | null
  if (!data?.bytes) throw new Error('DOCX 页面数据缺少字节内容')
  return {
    bytes: base64ToBytes(data.bytes),
    name: data.name || 'document.docx',
  }
}
