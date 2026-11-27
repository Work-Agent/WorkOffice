// ─── WorkStudio ⇄ GenOffice Sheets host adapter ─────────────────────
//
// GenOffice Sheets renderer 通过 window.desktopApi 访问宿主。这里用页面
// 会话实现同一个接口：
//   - 页面数据 → WorkbookFile / readWorkbookRange
//   - 保存 → window.__xlsxPlus.save（页面交保存请求或整份字节，引擎产字节）
//   - 主题 / 语言 / 脏态 → 页面桥（shell/appearance + xlsx.page.report）
//
// 这是“react 页面层”和 WorkStudio 的唯一接缝；renderer 本身尽量不改。

import type { CellLike, SheetLike } from '@xlsx/shared/sheet-types'
import { planWorkbookSave, saveWorkbookEdits as saveWorkbookEditsToBytes } from './workstudio-save'
import { shouldUseSavePlan } from './save-plan-budget'
import { parseWorkbook } from '@xlsx/shared/xlsx-parse'

export interface XlsxPageData {
  /** 老宿主兼容回退：base64 整份字节。新路径走 `HostOptions.bytes`。 */
  bytes?: string
  name: string
  ref: string
  version: string
  size?: number
  sheets: SheetLike[]
}

export interface XlsxProReceipt {
  ref?: string
  txnId?: string
  version?: string
  epoch?: number
  applied?: number
  /** 保存产物摘要与体积：页面据此对齐基线，不回读整包。 */
  hash?: string
  size?: number
  at?: number
  warnings?: string[]
  notices?: string[]
  /** `sync.sha256`：保存产物摘要。页面据此重建 file 模型，不回读整包。 */
  sync?: { kind: string; sheets?: unknown[]; sha256?: string }
  audit?: { sha256?: string; touchedEntries?: string[] }
}

interface HostOptions {
  sessionId: string
  data: XlsxPageData
  /** 启动字节（同进程直传，避免 base64）；缺失时回退 data.bytes。 */
  bytes?: Uint8Array | undefined
  epoch?: number
  onDirty: (dirty: boolean) => void
  onToast: (message: string) => void
  onReceipt?: (receipt: XlsxProReceipt) => void
}

const DEFAULT_STYLE = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: false,
  diagonalUp: false,
  diagonalDown: false,
}

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

function parseAddress(address: string): { row: number; column: number } | null {
  const match = /^([A-Z]+)(\d+)$/.exec(address)
  if (!match) return null
  let column = 0
  for (const ch of match[1]!) column = column * 26 + (ch.charCodeAt(0) - 64)
  return { row: Number(match[2]) - 1, column: column - 1 }
}

function parseArea(ref: string): {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
} | null {
  const [a, b = a] = ref.split(':')
  const start = parseAddress(a ?? '')
  const end = parseAddress(b ?? '')
  if (!start || !end) return null
  return {
    startRow: Math.min(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endRow: Math.max(start.row, end.row),
    endColumn: Math.max(start.column, end.column),
  }
}

function columnLabel(index: number): string {
  let label = ''
  let value = index + 1
  while (value > 0) {
    const rem = (value - 1) % 26
    label = String.fromCharCode(65 + rem) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

function styleOf(cell: CellLike): Record<string, unknown> | null {
  const style = cell.style
  if (!style) return null
  return {
    ...DEFAULT_STYLE,
    ...(style.bold ? { bold: true } : {}),
    ...(style.italic ? { italic: true } : {}),
    ...(style.underline ? { underline: true } : {}),
    ...(style.fontFamily ? { fontFamily: style.fontFamily } : {}),
    ...(style.fontSize ? { fontSize: style.fontSize } : {}),
    ...(style.fontColor ? { fontColor: style.fontColor } : {}),
    ...(style.fillColor ? { fillColor: style.fillColor } : {}),
    ...(style.numberFormat ? { numberFormat: style.numberFormat } : {}),
    ...(style.align ? { horizontalAlignment: style.align } : {}),
    ...(style.wrap ? { wrapText: true } : {}),
  }
}

function normalizeValue(value: CellLike['value']): string | number | boolean | null {
  if (value === undefined) return null
  return value
}

/// 范围读取缓存：有界 LRU（大文件滚动会持续产生新 key，无界 Map 会吃光内存）。
class RangeCache {
  private readonly map = new Map<string, unknown>()
  constructor(private readonly limit = 200) {}
  has(key: string): boolean {
    return this.map.has(key)
  }
  get(key: string): unknown {
    const value = this.map.get(key)
    if (value !== undefined) {
      // 命中即刷新为最近使用
      this.map.delete(key)
      this.map.set(key, value)
    }
    return value
  }
  set(key: string, value: unknown): void {
    this.map.delete(key)
    this.map.set(key, value)
    while (this.map.size > this.limit) {
      const oldest = this.map.keys().next()
      if (oldest.done) break
      this.map.delete(oldest.value)
    }
  }
  clear(): void {
    this.map.clear()
  }
}

/// 用引擎的 `open` 元数据重建 file 模型 —— 这是**原始协议形状**（fork 就是照它写的），
/// 比 `buildWorkbookFile` 从字节里猜的那份更完整（freeze / hidden / tabColor /
/// columnWidths / tables / comments / pivot 都是真值）。
/// 注意：`styles` 与范围读取返回的 `styleIndex` 必须同源，所以两者要么都用引擎的，
/// 要么都不用 —— 见 `rangeSource`。
function fileFromMeta(meta: any, options: HostOptions, sha256: string): any {
  return {
    sessionId: options.sessionId,
    name: meta?.name ?? options.data.name,
    sha256,
    entryCount: Number(meta?.entryCount ?? 0),
    sheets: Array.isArray(meta?.sheets) ? meta.sheets : [],
    activeTab: Number(meta?.activeTab ?? 0),
    styles: Array.isArray(meta?.styles) && meta.styles.length > 0 ? meta.styles : [{ ...DEFAULT_STYLE }],
    dxfStyles: meta?.dxfStyles ?? [],
    visuals: meta?.visuals ?? [],
    definedNames: meta?.definedNames ?? [],
    ...(meta?.themeColors ? { themeColors: meta.themeColors } : {}),
    ...(meta?.themeFonts ? { themeFonts: meta.themeFonts } : {}),
    ...(meta?.normalFontName ? { normalFontName: meta.normalFontName } : {}),
    ...(meta?.workbookProtection ? { workbookProtection: meta.workbookProtection } : {}),
    ...(meta?.date1904 === undefined ? {} : { date1904: meta.date1904 }),
    ...(meta?.shortDateFormat ? { shortDateFormat: meta.shortDateFormat } : {}),
    readOnly: false,
  }
}

function buildWorkbookFile(options: HostOptions, sha256: string): any {
  const styleList: Array<Record<string, unknown>> = [{ ...DEFAULT_STYLE }]
  const styleIndex = new Map<string, number>()
  const cellStyleIndex = new Map<string, number>()
  let entryCount = 0

  for (const sheet of options.data.sheets) {
    for (const [address, cell] of Object.entries(sheet.cells)) {
      entryCount += 1
      const style = styleOf(cell)
      if (!style) continue
      const key = JSON.stringify(style)
      let index = styleIndex.get(key)
      if (index === undefined) {
        index = styleList.length
        styleList.push(style)
        styleIndex.set(key, index)
      }
      cellStyleIndex.set(`${sheet.id}:${address}`, index)
    }
  }

  const sheets = options.data.sheets.map((sheet) => {
    let maxRow = 0
    let maxColumn = 0
    for (const address of Object.keys(sheet.cells)) {
      const parsed = parseAddress(address)
      if (!parsed) continue
      maxRow = Math.max(maxRow, parsed.row)
      maxColumn = Math.max(maxColumn, parsed.column)
    }
    const columnWidths: Array<Record<string, unknown>> = []
    for (const [label, width] of Object.entries(sheet.colWidths ?? {})) {
      const parsed = parseAddress(`${label}1`)
      if (!parsed) continue
      columnWidths.push({
        startColumn: parsed.column,
        endColumn: parsed.column,
        width: Math.max(0.1, (width - 5) / 7),
        hidden: false,
      })
    }
    return {
      id: sheet.id,
      name: sheet.name,
      rowCount: Math.max(1000, maxRow + 101, sheet.gridRows ?? 0),
      columnCount: Math.max(52, maxColumn + 11, sheet.gridColumns ?? 0),
      columnWidths,
      defaultRowHeight: 15,
      defaultColumnWidth: null,
      freeze: null,
      hidden: false,
      tabColor: null,
      showGridLines: true,
      tables: [],
      comments: [],
      pivotRanges: [],
      pivotTables: [],
      sparklines: [],
      cellImages: [],
    }
  })

  return {
    sessionId: options.sessionId,
    name: options.data.name,
    sha256,
    entryCount,
    sheets,
    activeTab: 0,
    styles: styleList,
    dxfStyles: [],
    visuals: [],
    definedNames: [],
    readOnly: false,
  }
}

function sheetOf(data: XlsxPageData, sheetId: string): SheetLike | null {
  return data.sheets.find((sheet) => sheet.id === sheetId) ?? null
}

async function digest(bytes: Uint8Array): Promise<string> {
  try {
    const hash = await crypto.subtle.digest('SHA-256', bytes as BufferSource)
    return [...new Uint8Array(hash)]
      .map((b) => b.toString(16).padStart(2, '0'))
      .join('')
  } catch {
    return '0'.repeat(64)
  }
}

/** 读取宿主主题：`shell/appearance` 订阅由页面桥缓存，缺失时回退系统偏好。 */
export async function readHostTheme(): Promise<'light' | 'dark'> {
  try {
    const theme = await window.__xlsxPlus?.theme()
    if (theme === 'dark' || theme === 'light') return theme
  } catch {
    /* fallthrough */
  }
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

export async function installWorkStudioHost(options: HostOptions): Promise<void> {
  // P5：打开时可能没有整份字节（宿主有 invoke 时页面不读文件）。
  // 需要字节的只有两条回退路径，统一走 `ensureBytes()`。
  const initialBytes =
    options.bytes ??
    (options.data.bytes === undefined ? null : base64ToBytes(options.data.bytes))
  let currentBytes: Uint8Array | null = initialBytes
  // 没有字节时用宿主的内容哈希当 file 身份（与 `version` 同源）。
  const initialSha = initialBytes
    ? await digest(initialBytes)
    : String(options.data.version ?? '')
  let currentFile: any = buildWorkbookFile(options, initialSha)
  let version = options.data.version
  let currentHash = initialSha
  let epoch = options.epoch ?? 0
  let origin: 'page' | 'agent' = 'page'
  let lastReceipt: XlsxProReceipt | null = null

  // 引擎读通道：宿主没接线（无 invoke）或 wasm 不可用时保持 'local'，
  // 一旦 'local' 就永远不用引擎的范围结果（styles 与 styleIndex 必须同源）。
  // 流式打开：wasm 可用时范围读取走 `__xlsxPlus.range`，否则本地解析整份字节。
  let rangeSource = (window.__xlsxPlus?.supportsEngineRange ? 'engine' : 'local') as
    | 'engine'
    | 'local'
  let readFallbackWarned = false

  const rangeReads = new RangeCache(200)

  /// 调只读工具（`xlsx_read what:'cells'|'meta'`），结果可能被宿主包了信封。
  async function invokeRead(input: Record<string, unknown>): Promise<any> {
    const bridge = needBridge()
    if (input['what'] === 'meta') return bridge.meta()
    if (input['what'] === 'cells') return bridge.range(String(input['sheetId']), input['range'])
    throw new Error(`不支持的只读工具：${String(input['what'])}`)
  }

  function noteReadFallback(what: string, error: unknown): void {
    rangeSource = 'local'
    if (readFallbackWarned) return
    readFallbackWarned = true
    console.warn(
      `[xlsx] 引擎范围读取不可用（${what}），回退到本地解析：`,
      error instanceof Error ? error.message : error,
    )
  }

  /// 保存后只取元数据（几 KB），不再回读整包字节。
  /// 返回 false 表示拿不到元数据 —— 调用方需要走原来的整包刷新。
  async function refreshFileFromMeta(sha256: string): Promise<boolean> {
    if (rangeSource !== 'engine') return false
    try {
      const meta = await invokeRead({ what: 'meta' })
      if (!meta?.sheets) return false
      currentFile = fileFromMeta(meta, options, sha256)
      rangeReads.clear()
      return true
    } catch (error) {
      noteReadFallback('meta', error)
      return false
    }
  }

  /// 按需取整份字节（惰性、带版本守卫）。只在回退路径调用。
  async function ensureBytes(force = false): Promise<Uint8Array> {
    if (currentBytes && !force) return currentBytes
    const bridge = needBridge()
    const fetched = await bridge.bytes()
    currentBytes = fetched.bytes
    currentHash = fetched.hash
    options.data.size = fetched.size
    return currentBytes
  }

  function needBridge() {
    const bridge = window.__xlsxPlus
    if (!bridge) throw new Error('页面桥未就绪（window.__xlsxPlus）')
    return bridge
  }

  /** 另存为：宿主对话框签 `save:` token，引擎写新文件；通道随之指向新 ref。 */
  async function saveAsToFile(request: any): Promise<{
    canceled: boolean
    file: unknown
    touchedEntries: string[]
  }> {
    const choice = await needBridge().pickSaveFile(options.data.name || 'workbook.xlsx')
    if (!choice) return { canceled: true, file: currentFile, touchedEntries: [] }

    const saved = await saveWorkbookEditsToBytes(await ensureBytes(true), request, currentFile)
    const receipt = await needBridge().saveAs(saved.bytes, choice.token)
    currentHash = String(receipt.hash)
    version = currentHash
    currentBytes = saved.bytes
    options.data.name = choice.name ?? options.data.name
    options.data.ref = receipt.ref
    options.data.size = receipt.size
    options.data.version = currentHash
    try {
      options.data.sheets = parseWorkbook(currentBytes)
    } catch (error) {
      console.warn('[xlsx] 另存后解析失败，范围读取仍用旧快照', error)
    }
    currentFile = {
      ...saved.file,
      ...buildWorkbookFile(options, currentHash),
      name: options.data.name,
      fileBytes: currentBytes.length,
      sha256: currentHash,
    }
    rangeReads.clear()
    lastReceipt = receipt
    if (origin === 'page') options.onDirty(false)
    options.onReceipt?.(receipt)
    return { canceled: false, file: currentFile, touchedEntries: saved.touchedEntries }
  }

  const styleIndexFor = (sheetId: string, address: string): number | undefined => {
    // Recompute the style index lazily from the file's style table; the table
    // is small and the lookup is only used for visible ranges.
    const sheet = sheetOf(options.data, sheetId)
    const cell = sheet?.cells[address]
    if (!cell?.style) return undefined
    const style = JSON.stringify(styleOf(cell))
    const index = currentFile.styles.findIndex(
      (item: Record<string, unknown>) => JSON.stringify(item) === style,
    )
    return index > 0 ? index : undefined
  }

  /// 大编辑集的分块传输（> SAVE_EDITS_INLINE_MAX 时 App 不再把 edits 放进请求）。
  /// 页面宿主不跨进程，所以这里只是把分块收进内存再合并 —— 但不实现就等于
  /// 静默丢掉这些编辑（旧实现把三个方法留空，超大批量保存会丢数据）。
  const saveEditTransfers = new Map<string, { total: number; chunks: Map<number, string> }>()

  function takeStagedEdits(request: any): any {
    const transferId = request?.editsTransferId
    if (typeof transferId !== 'string') return request
    const transfer = saveEditTransfers.get(transferId)
    saveEditTransfers.delete(transferId)
    if (!transfer) return request
    const staged: unknown[] = []
    for (const [, json] of [...transfer.chunks.entries()].sort((left, right) => left[0] - right[0])) {
      const slice = JSON.parse(json)
      if (Array.isArray(slice)) staged.push(...slice)
    }
    if (staged.length !== transfer.total)
      throw new Error(
        `保存数据不完整：编辑分块收到 ${staged.length} 条，期望 ${transfer.total} 条`,
      )
    return { ...request, edits: [...(request.edits ?? []), ...staged] }
  }

  /// 回读权威字节：`kind:'save'` 的提交只上行请求，页面不再持有权威字节。
  /// 保存后的 reopen（save-actions → openLazyWorkbook）要从 `options.data.sheets`
  /// 取数，所以这里必须重新解析一次，否则重开会把刚保存的修改退回旧值。
  async function refreshAuthoritativeBytes(): Promise<void> {
    const bytes = await ensureBytes(true)
    try {
      options.data.sheets = parseWorkbook(bytes)
    } catch (error) {
      console.warn('[xlsx] 保存后解析工作簿失败，range 仍使用旧快照', error)
    }
    const sha256 = await digest(bytes)
    currentFile = {
      ...currentFile,
      ...buildWorkbookFile(options, sha256),
      fileBytes: bytes.length,
      sha256,
    }
    rangeReads.clear()
    // 这份 file 模型的 styles 是本地去重表，与引擎的 styleIndex 不同源 ——
    // 之后的范围读取必须也回本地解析。
    rangeSource = 'local'
  }

  /// `plan { kind:'save' }`：上行只有请求（几十 KB），字节由执行点产出。
  /// 提交成功后页面基线按回执对齐，再回读一次权威字节供 reopen 使用。
  async function saveViaPlan(
    plan: ReturnType<typeof planWorkbookSave>,
    decision: ReturnType<typeof shouldUseSavePlan>,
  ): Promise<{ canceled: boolean; file: unknown; touchedEntries: string[]; planKind: 'save' }> {
    let receipt: XlsxProReceipt
    try {
      // 页面只交「保存请求」；字节由引擎的 gateway 产出。请求带上基线哈希，
      // 外部改动会让引擎在锁内拒绝（stale），不会覆盖别人的写入。
      receipt = await needBridge().save({ request: plan, baseHash: currentHash })
    } catch (error) {
      console.error('[xlsx] save request failed', {
        baseHash: currentHash,
        planBytes: decision.planBytes,
        error,
      })
      throw error
    }
    currentHash = String(receipt.hash ?? currentHash)
    version = currentHash
    if (typeof receipt.size === 'number') options.data.size = receipt.size
    lastReceipt = receipt
    // 流式打开时只重建元数据（几 KB）；否则回读整份字节。
    const rebuilt = await refreshFileFromMeta(currentHash)
    if (!rebuilt) await refreshAuthoritativeBytes()
    if (origin === 'page') options.onDirty(false)
    options.onReceipt?.(receipt)
    void needBridge().clearDraft()
    return { canceled: false, file: currentFile, touchedEntries: [], planKind: 'save' }
  }

  // 启动：优先用引擎元数据重建 file（与范围读取的 styleIndex 同源）。
  // 拿不到就保留上面那份从字节猜出来的模型，范围读取也走本地解析。
  if (rangeSource === 'engine') {
    const ok = await refreshFileFromMeta(currentFile.sha256)
    if (!ok) currentFile = buildWorkbookFile(options, initialBytes ? await digest(initialBytes) : initialSha)
  }

  const api: any = {
    getLanguage: async () => 'zh',
    onLanguageChanged: () => () => {},
    getTheme: async () => readHostTheme(),
    onThemeChanged: (handler: (theme: string) => void) =>
      needBridge().onTheme((next) => handler(next)),
    getDocumentTheme: async () => 'follow',
    onDocumentThemeChanged: () => () => {},
    getAutoSaveDefault: async () => ({ on: true, updatedAt: 0 }),
    onAutoSaveDefaultChanged: () => () => {},
    getAiPanelPrefs: async () => null,
    setAiPanelPrefs: async () => null,
    onAiPanelPrefsChanged: () => () => {},
    onChromePressed: () => () => {},
    selectWorkbook: async () => currentFile,
    selectWorkbooksForMerge: async () => null,
    openWorkbooksForMerge: async () => null,
    reopenWorkbook: async () => currentFile,
    onSidecarCrashed: () => () => {},
    readWorkbookRange: async (request: any) => {
      const key = `${request.sheetId}:${JSON.stringify(request.range)}`
      const cached = rangeReads.get(key)
      if (cached !== undefined) return cached
      // 引擎优先：O(范围)，且 styles 与 styleIndex 同源。
      if (rangeSource === 'engine') {
        try {
          const data = await invokeRead({
            what: 'cells',
            sheetId: request.sheetId,
            range: request.range,
          })
          if (data && typeof data === 'object') {
            const result = {
              ...(data as Record<string, unknown>),
              sessionId: currentFile.sessionId,
              sheetId: request.sheetId,
            }
            rangeReads.set(key, result)
            return result
          }
          throw new Error('引擎返回了空的 RangeResult')
        } catch (error) {
          noteReadFallback('cells', error)
        }
      }
      const sheet = sheetOf(options.data, request.sheetId)
      if (!sheet) throw new Error(`Unknown worksheet ${request.sheetId}`)
      const { startRow, endRow, startColumn, endColumn } = request.range
      const cells: any[] = []
      for (const [address, cell] of Object.entries(sheet.cells)) {
        const parsed = parseAddress(address)
        if (!parsed) continue
        if (
          parsed.row < startRow ||
          parsed.row > endRow ||
          parsed.column < startColumn ||
          parsed.column > endColumn
        )
          continue
        cells.push({
          row: parsed.row,
          column: parsed.column,
          value: normalizeValue(cell.value),
          ...(cell.formula ? { formula: cell.formula } : {}),
          ...(styleIndexFor(sheet.id, address) !== undefined
            ? { styleIndex: styleIndexFor(sheet.id, address) }
            : {}),
        })
      }
      const rows: any[] = []
      for (let row = startRow; row <= endRow; row++) {
        const height = sheet.rowHeights?.[String(row + 1)]
        rows.push({
          row,
          ...(height ? { height: (height * 72) / 96, customHeight: true } : {}),
          hidden: false,
        })
      }
      const merges = (sheet.merges ?? [])
        .map(parseArea)
        .filter(
          (area): area is NonNullable<typeof area> =>
            area !== null &&
            area.endRow >= startRow &&
            area.startRow <= endRow &&
            area.endColumn >= startColumn &&
            area.startColumn <= endColumn,
        )
      const result = {
        cells,
        rows,
        merges,
        hyperlinks: [],
        conditionalRules: [],
        autoFilter: null,
        autoFilterColumns: [],
        dataValidations: [],
        sheetProtection: null,
        rowBreaks: [],
        colBreaks: [],
        pageSetup: null,
        protectedRanges: [],
        indexedThroughRow: sheet.gridRows ?? 1000,
        indexingComplete: true,
        sessionId: currentFile.sessionId,
        sheetId: request.sheetId,
      }
      rangeReads.set(key, result)
      return result
    },
    readWorkbookFormulas: async () => ({
      cells: [],
      sharedGroups: [],
      indexingComplete: true,
      truncated: false,
    }),
    findWorkbookCells: async () => ({
      matches: [],
      indexingComplete: true,
      truncated: false,
    }),
    readWorkbookRowOutline: async () => ({ rows: [], indexingComplete: true }),
    recalcWorkbook: async () => ({ values: [], indexingComplete: true }),
    readWorkbookMedia: async () => null,
    readPivotDefinition: async () => null,
    readLocalImage: async () => null,
    captureScreenSources: async () => ({ status: 'denied', sources: [] }),
    captureScreenSource: async () => null,
    saveWorkbookEdits: async (rawRequest: any) => {
      const request = takeStagedEdits(rawRequest)
      if (request?.mode === 'save-as') return saveAsToFile(request)
      const plan = planWorkbookSave(request, currentFile)
      const decision = shouldUseSavePlan(plan, { mode: request?.mode ?? 'save' })
      if (decision.usePlan) return saveViaPlan(plan, decision)
      const saved = await saveWorkbookEditsToBytes(await ensureBytes(), request, currentFile)
      let receipt: XlsxProReceipt
      try {
        receipt = await needBridge().save({ bytes: saved.bytes, baseHash: currentHash })
      } catch (error) {
        console.error('[xlsx] save bytes failed', {
          baseHash: currentHash,
          bytes: saved.bytes?.byteLength,
          error,
        })
        throw error
      }
      currentHash = String(receipt.hash ?? currentHash)
      version = currentHash
      lastReceipt = receipt
      currentBytes = saved.bytes
      // 保存后文件长度会变：把新体积写回页面数据，分块读取据此判"截断"。
      options.data.size = saved.bytes.byteLength
      // Ctrl+S 后页面会 reopen 同一个 session；host 的 range 读取仍从
      // `options.data.sheets` 取数，必须用刚保存的字节刷新快照，否则
      // 重开会把用户刚保存的修改“退回”成旧值。
      try {
        options.data.sheets = parseWorkbook(currentBytes)
      } catch (error) {
        console.warn('[xlsx] 保存后解析工作簿失败，range 仍使用旧快照', error)
      }
      const sha256 = await digest(currentBytes)
      currentFile = {
        ...saved.file,
        ...buildWorkbookFile(options, sha256),
        name: saved.file.name,
        fileBytes: currentBytes.length,
        sha256,
      }
      rangeReads.clear()
      if (origin === 'page') options.onDirty(false)
      options.onReceipt?.(receipt)
      return {
        canceled: false,
        file: currentFile,
        touchedEntries: saved.touchedEntries,
      }
    },
    beginSaveEditsTransfer: async (transfer: {
      transferId: string
      total: number
    }) => {
      saveEditTransfers.set(transfer.transferId, { total: transfer.total, chunks: new Map() })
    },
    sendSaveEditsChunk: async (chunk: {
      transferId: string
      seq: number
      editsJson: string
    }) => {
      saveEditTransfers.get(chunk.transferId)?.chunks.set(chunk.seq, chunk.editsJson)
    },
    abortSaveEditsTransfer: async (transfer: { transferId: string }) => {
      saveEditTransfers.delete(transfer.transferId)
    },
    writeWorkbookRecovery: async (payload: unknown) => {
      try {
        return await needBridge().draft(payload)
      } catch (error) {
        console.warn('[xlsx] 恢复副本写入失败', error)
        return { ok: false }
      }
    },
    autoRenameWorkbook: async (_sessionId: string, name: string) => {
      currentFile = { ...currentFile, name }
      return currentFile
    },
    exportPdf: async () => ({ ok: false, error: '暂不支持' }),
    printWorkbook: async () => ({ ok: false, error: '暂不支持' }),
    listPrinters: async () => [],
    exportCsv: async () => ({ ok: false, error: '暂不支持' }),
    confirmCsvSave: async () => 'cancel',
    closeWorkbook: async () => {},
    openExternal: async () => {},
    onMenuAction: () => () => {},
    onWorkbookRenamed: () => () => {},
    notifyPendingEdits: (count: number) => options.onDirty(count > 0),
    onCloseSaveRequest: () => () => {},
    reportCloseSaveResult: () => {},
    onRecoveryPrompt: () => () => {},
    replyRecoveryPrompt: () => {},
    consumeNewBlankWorkbook: async () => false,
    onMcpCommand: () => () => {},
    reportMcpResult: () => {},
    signalMcpReady: () => {},
    hasQueuedWorkbook: async () => true,
    consumeHeadlessExport: async () => null,
    headlessExportDone: () => {},
    getUserDisplayName: async () => 'WorkStudio',
    onUserDisplayNameChanged: () => () => {},
    getPathForFile: () => '',
    pickAttachments: async () => null,
    addAttachmentPaths: async () => ({ accepted: [], rejected: [] }),
    addPastedImage: async () => ({ accepted: [], rejected: [] }),
    readAttachment: async () => ({ ok: false }),
    readAttachmentImage: async () => ({ ok: false }),
    openSourcePath: async () => {},
    fetchImage: async () => null,
    onAiStream: () => () => {},
  }

  ;(window as unknown as { desktopApi: unknown }).desktopApi = api
  ;(window as unknown as { __xlsxHost?: unknown }).__xlsxHost = {
    getVersion: () => version,
    getEpoch: () => epoch,
    getReceipt: () => lastReceipt,
    getFile: () => currentFile,
    getBytes: () => currentBytes ?? new Uint8Array(0),
    setOrigin: (next: 'page' | 'agent') => {
      origin = next
    },
    commitBytes: async (bytes: Uint8Array, source: 'page' | 'agent' = 'page') => {
      void source
      const receipt = await needBridge().save({ bytes, baseHash: currentHash })
      currentHash = String(receipt.hash ?? currentHash)
      version = currentHash
      lastReceipt = receipt
      return receipt
    },
  }
}
