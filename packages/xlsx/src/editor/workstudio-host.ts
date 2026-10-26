// ─── WorkStudio ⇄ GenOffice Sheets host adapter ─────────────────────
//
// GenOffice Sheets renderer 通过 window.desktopApi 访问宿主。这里用页面
// 会话实现同一个接口：
//   - 页面数据 → WorkbookFile / readWorkbookRange
//   - 保存 → ura.page.saveBytes（后续接 xlsx-gateway）
//   - 主题 / 语言 / 脏态 → WorkStudio 扩展总线
//
// 这是“react 页面层”和 WorkStudio 的唯一接缝；renderer 本身尽量不改。

import type { CellLike, SheetLike } from '../ura-types'
import { saveWorkbookEdits as saveWorkbookEditsToBytes } from './workstudio-save'

export interface XlsxPageData {
  bytes: string
  name: string
  ref: string
  version: string
  sheets: SheetLike[]
}

interface HostOptions {
  sessionId: string
  data: XlsxPageData
  onDirty: (dirty: boolean) => void
  onToast: (message: string) => void
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

export async function installWorkStudioHost(options: HostOptions): Promise<void> {
  let currentBytes = base64ToBytes(options.data.bytes)
  let currentFile = buildWorkbookFile(options, await digest(currentBytes))

  const rangeReads = new Map<string, any>()
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

  const api: any = {
    getLanguage: async () => 'zh',
    onLanguageChanged: () => () => {},
    getTheme: async () => {
      const ui = await window.extension.invoke<{ theme?: string }>('app:ui.get')
      return ui?.theme === 'dark' ? 'dark' : 'light'
    },
    onThemeChanged: (handler: (theme: string) => void) => {
      return window.extension.bus.on((message) => {
        if (message.type === 'theme') handler(message.theme === 'dark' ? 'dark' : 'light')
      })
    },
    getDocumentTheme: async () => 'follow',
    onDocumentThemeChanged: () => () => {},
    getAutoSaveDefault: async () => ({ on: false, updatedAt: 0 }),
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
      if (rangeReads.has(key)) return rangeReads.get(key)
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
    saveWorkbookEdits: async (request: any) => {
      const saved = await saveWorkbookEditsToBytes(currentBytes, request, currentFile)
      const result = await window.extension.ura.page.saveBytes<{
        ok: boolean
        version: string
      }>(options.sessionId, saved.bytes)
      if (!result?.ok) throw new Error('保存失败')
      currentBytes = saved.bytes
      currentFile = {
        ...saved.file,
        sha256: await digest(currentBytes),
        fileBytes: currentBytes.length,
      }
      rangeReads.clear()
      options.onDirty(false)
      return {
        canceled: false,
        file: currentFile,
        touchedEntries: saved.touchedEntries,
      }
    },
    beginSaveEditsTransfer: async () => {},
    sendSaveEditsChunk: async () => {},
    abortSaveEditsTransfer: async () => {},
    writeWorkbookRecovery: async () => ({ ok: true }),
    autoRenameWorkbook: async (_sessionId: string, name: string) => {
      currentFile = { ...currentFile, name }
      return currentFile
    },
    exportPdf: async () => ({ ok: false, error: '暂不支持' }),
    printWorkbook: async () => ({ ok: false, error: '暂不支持' }),
    listPrinters: async () => [],
    exportCsv: async () => ({ ok: false, error: '暂不支持' }),
    confirmCsvSave: async () => 'cancel',
    createDocument: async () => ({ ok: false, error: '暂不支持' }),
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
    getAiSettings: async () => null,
    setAiSettings: async () => {},
    onAiSettingsChanged: () => () => {},
    openAiModelSettings: async () => {},
    aiChat: async () => ({ text: '' }),
    aiStream: async () => {},
    aiStreamCancel: async () => {},
    aiGskStatus: async () => ({ loggedIn: false }),
    aiGskLogin: async () => {},
    webSearch: async () => ({ results: [], method: 'unavailable' }),
    imageSearch: async () => ({ images: [], method: 'unavailable' }),
    generateImage: async () => ({ error: 'AI 已移除' }),
    fetchImage: async () => null,
    onAiStream: () => () => {},
  }

  ;(window as unknown as { desktopApi: unknown }).desktopApi = api
}
