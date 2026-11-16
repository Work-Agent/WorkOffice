// ─── GenOffice Docs 渲染器 ↔ WorkStudio 页面会话适配层（v2） ─────────
//
// 从 packages/docx/src/page/host.ts 移植：GenOffice 的交互代码只认
// `window.desktop`，这里把它接到本插件的页面会话上。
//
//   GenOffice 打开 = 文件路径 → 这里 = page.get 的字节（blobUrl）
//   GenOffice 保存 = 写路径 → 这里 = /commit 的 plan（P2）
//   GenOffice 对话框 / 菜单 → 不承载，交给 WorkStudio 壳层
//
// P2：保存走 plan —— 渲染器仍用引擎把新 document.xml 算出来（它既有的保存流水线），
// 页面把它还原成块级 SavePlan（见 plan.ts），再 `/commit` 交给执行点写盘；
// 执行点仍是唯一字节产生者，页面自己不写盘。

import type { Receipt } from '../page/env'
import { sameBlockSlices } from '../shared'
import { readSessionBytes } from './bytes'
import { planFromBytes } from '../shared'

export interface DocxHostInput {
  sessionId: string
  ref: string
  name: string
  bytes: Uint8Array
  blobUrl: string
  hash: string
  /** 打开时的 CAS 版本与坐标令牌（提交要带回去）。 */
  version: string
  epoch: number
}

type Listener<T> = (value: T) => void

/** 页面会话快照（提交后更新；agent 代理用它回执 version / epoch / docRev）。 */
export interface SessionSnapshot {
  version: string
  epoch: number
  docRev: number
}

const session: SessionSnapshot = { version: '', epoch: 0, docRev: 0 }

/** 画面侧查询 / 重载入口（main.tsx 的 bus 桥用）。 */
export interface HostControls {
  isDirty(): boolean
  reloadFromDisk(): Promise<boolean>
}

let controls: HostControls | null = null

export function hostControls(): HostControls | null {
  return controls
}

export function sessionState(): SessionSnapshot {
  return { ...session }
}

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

export function installDocxHost(input: DocxHostInput): void {
  const { bus } = window.extension
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
  let dirty = false

  const setDirty = (next: boolean): void => {
    if (dirty === next) return
    dirty = next
    bus.send({ type: 'dirty', dirty: next })
  }

  document.addEventListener(
    'input',
    (event) => {
      const target = event.target
      if (!(target instanceof Element)) return
      if (target.closest('.ProseMirror, [contenteditable="true"]')) setDirty(true)
    },
    true,
  )

  // 渲染器的「打开文档」回调：外部变更重载走的也是这条（与从 Finder 打开同一个入口）。
  const openHandlers = new Set<(result: unknown) => void>()

  // ── 提交状态（页面 = 该 ref 的唯一编辑模型） ──────────────
  session.version = input.version
  session.epoch = input.epoch
  session.docRev = 0
  let baseline = input.bytes
  let committing = false

  function errorCodeOf(error: unknown): string | undefined {
    const code = (error as { code?: unknown })?.code
    return typeof code === 'string' ? code : undefined
  }

  function commitError(error: unknown): string {
    const code = errorCodeOf(error)
    const message = error instanceof Error ? error.message : String(error)
    if (code === 'command.stale')
      return '文件已被外部修改：本地修改未写入。请重新打开该文档后再编辑。'
    if (code === 'command.stale-index')
      return '编辑坐标已失效（epoch / docRev 变化）：请重新打开该文档后再编辑。'
    if (code === 'command.too-large') return '文档超出提交上限，已取消这次保存。'
    if (code === 'command.page-unavailable') return '页面会话不可用：请重新打开该文档。'
    return message
  }

  /** 页面保存入口：新字节 → 块级计划 → /commit（执行点写盘）。 */
  async function savePlan(
    body: ArrayBuffer,
    _options?: { auto?: boolean },
  ): Promise<{ ok: boolean; error?: string }> {
    if (committing) return { ok: false, error: '上一次保存尚未完成，请稍候' }
    committing = true
    try {
      const produced = new Uint8Array(body)
      let built = planFromBytes({ baseline, produced })
      // 基线坏了（上一轮的写入尚在刷盘、或被外部程序改过）→ 重读一次再试。
      // 只对「基线侧」自愈；产出侧坏就是引擎序列化问题，不该掩盖。
      if (!built.ok && built.error.startsWith('基线不可解析')) {
        try {
          const fresh = await readSessionBytes(input.sessionId)
          const retried = planFromBytes({ baseline: fresh, produced })
          if (retried.ok) {
            baseline = fresh
            built = retried
          }
        } catch {
          /* 重读失败：把原始错误报给用户 */
        }
      }
      if (!built.ok) return { ok: false, error: built.error }
      // 无改动（引擎走了 isUnchanged 快路径、字节原样）：不产生空提交。
      // 注意 parts 也要算进来：只改页眉 / 只换图片时正文块可以完全没变，
      // 只看 changedCount 会把这些改动当成空保存丢掉。
      const partCount = built.plan.parts?.length ?? 0;
      if (built.plan.changedCount === 0 && built.plan.deletedCount === 0 && partCount === 0) {
        setDirty(false)
        return { ok: true }
      }
      const receipt = await window.extension.ura.page.commit<Receipt>(input.sessionId, {
        txnId: crypto.randomUUID(),
        baseVersion: session.version,
        expect: { docRev: String(session.docRev), epoch: session.epoch },
        payload: { type: 'plan', plan: built.plan },
        origin: 'page',
      })
      if (receipt.rejected?.length) {
        const [first] = receipt.rejected
        return {
          ok: false,
          error: `提交被拒（${first?.code ?? 'command.invalid'}）：${first?.message ?? ''}`,
        }
      }
      session.version = receipt.version
      session.epoch = receipt.epoch
      session.docRev += 1
      setDirty(false)
      // 新基线 = 本次计划产出的字节。它与 exec 落盘的块切片逐字节一致（original 条目
      // 引用的就是这些切片，xml 条目原样写入），所以不需要再从磁盘回读——
      // 刚写完/未刷盘时的回读会把下一个计划的基准弄成半截文件。
      baseline = produced
      return { ok: true }
    } catch (error) {
      return { ok: false, error: commitError(error) }
    } finally {
      committing = false
    }
  }

  /**
   * 提交后回读：渲染器要靠它重解析（docxIndex / 分节 / 页眉页脚状态）。
   * 写入是异步的（临时文件 + rename），刚返回就读可能拿到旧内容或半截文件：
   * 读到不可解析的字节就短暂重试；全失败时退回本次已提交的字节，
   * 至少让渲染器能完成重解析（下次保存会继续用同一基准）。
   */
  async function readDocxBytes(): Promise<Uint8Array> {
    for (let attempt = 0; attempt < 5; attempt += 1) {
      try {
        const bytes = await readSessionBytes(input.sessionId)
        // 不只看「能不能解析」：还要确认磁盘上就是刚提交的内容，
        // 否则渲染器会拿旧文件重解析，把用户刚改的东西“撤”回去。
        if (sameBlockSlices(bytes, baseline)) {
          baseline = bytes
          return bytes
        }
      } catch {
        /* 读失败 / 半截文件：重试 */
      }
      await new Promise((resolve) => setTimeout(resolve, 80))
    }
    // 回读不可用：用本次已提交的字节重解析（内容与刚写的一致）
    return baseline
  }

  /** 页面是否有未提交的本地改动。 */
  function isDirty(): boolean {
    return dirty
  }

  /** 把磁盘上的当前文件包成渲染器的 `OpenDocxResult`（一次性 blob URL）。 */
  async function openResultFromDisk(): Promise<{
    path: string
    name: string
    dataUrl: string
    hash: string
  }> {
    const bytes = await readSessionBytes(input.sessionId)
    const info = await window.extension.ura.page.get<{ version: string; epoch: number }>(
      input.sessionId,
    )
    session.version = info.version
    session.epoch = info.epoch
    baseline = bytes
    return {
      path: input.ref,
      name: input.name,
      dataUrl: URL.createObjectURL(
        new Blob([bytes as unknown as BlobPart], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        }),
      ),
      hash: hashOf(bytes),
    }
  }

  /**
   * 外部变更（宿主通知磁盘被改）：没脏就整篇重载；脏就交给 CAS——
   * 下一次提交会拿到 `command.stale`，页面显示"文件已被外部修改"，
   * 不会把本地改动盖到别人的文件上。
   *
   * 重载走渲染器自己的「打开最近文件」入口（`__aidocs.openPath` → `openDocxPath`），
   * 与从 Finder 打开同一条流水线：先 confirmDocumentReplace，再整篇重解析。
   */
  async function reloadFromDisk(): Promise<boolean> {
    if (dirty) return false
    try {
      const result = await openResultFromDisk()
      const opener = (window as unknown as { __aidocs?: { openPath?: (path: string) => unknown } })
        .__aidocs?.openPath
      if (opener) {
        await opener(input.ref)
        return true
      }
      // 渲染器还没挂出 `__aidocs`：退回「推给打开回调」的方式（有登记才生效）
      for (const handler of openHandlers) handler(result)
      return openHandlers.size > 0
    } catch {
      return false
    }
  }

  /** 非 plan 通道（渲染器里少数直接写字节的路径）一律拒绝。 */
  async function save(_data: ArrayBuffer): Promise<{ ok: boolean; error?: string }> {
    return { ok: false, error: '写盘只能走 plan 通道（saveDocxPlan）' }
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
    // 渲染器的「打开最近文件 / 重载」入口：只认本会话的 ref，其余一律空。
    openDocxPath: async (path: string) => {
      if (path !== input.ref) return null
      try {
        return await openResultFromDisk()
      } catch {
        return null
      }
    },
    openDocxDecrypt: async () => ({ ok: false, reason: 'unsupported' }),
    confirmDocumentReplace: async () => true,
    onOpenDocx: (handler: (result: unknown) => void) => {
      openHandlers.add(handler)
      return () => {
        openHandlers.delete(handler)
      }
    },
    onRenamedDocx: () => () => undefined,

    // 保存：plan 通道（渲染器算出新 XML，页面还原块级计划 → /commit → exec 写盘）
    saveDocxPlan: (data: ArrayBuffer, options?: { auto?: boolean }) => savePlan(data, options),
    readDocxBytes: () => readDocxBytes(),
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
    // plan 通道已接通：恢复自动保存（节奏由渲染器控制）。
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
  controls = { isDirty, reloadFromDisk }
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

