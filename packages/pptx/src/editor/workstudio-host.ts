// ─── WorkStudio ⇄ GenOffice Slides host adapter ─────────────────────
//
// GenOffice Slides renderer 通过 window.slidesApi 访问主进程的 pptx-engine。
// 这里把同一套 API 落到页面内：
//   - consumePendingOpen / openPptx → openPptx + buildRenderSlide
//   - save / saveAs → savePptx → ura.page.saveBytes
//   - 主题 / 语言 / 字体 / 附件 → WorkStudio 页面桥
//
// 编辑操作（editText / transform / slide ops）随后逐步接入页面内 pptx-ops；
// 未实现的调用由 no-op 兜底，保证页面外壳可以先完整呈现。

import { Buffer } from 'buffer'
import {
  EMU_PER_PT,
  copyElementData,
  elementSpid,
  getElementLink,
  getRunLinks,
  getSections,
  getSlideAnimations,
  getSlideAdvanceTime,
  getSlideLinks,
  getSlideNotes,
  getSlideTransition,
  listMasterParts,
  listSlideLayouts,
  materializeSlide,
  openPptx,
  parseMasterPart,
  readHeaderFooter,
  savePptx,
  type OpenedPptx,
} from '@genoffice/pptx-engine'
import { buildRenderSlide, EMU_PER_PX_96, type RenderSlide } from '@genoffice/pptx-render'
import { runTxn } from '@genoffice/pptx-ops'
import { PDFDocument } from 'pdf-lib'

interface PptxPageData {
  bytes: string
  name: string
  ref: string
  version: string
  slideCount: number
}

interface HostOptions {
  sessionId: string
  data: PptxPageData
}

const FIT_WIDTH = 960
const MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/emf',
  wmf: 'image/wmf',
}

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

function mimeOf(path: string): string {
  const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
  return MIME[ext] ?? 'application/octet-stream'
}

function buildSlides(opened: OpenedPptx, fitWidthPx: number): RenderSlide[] {
  const media = (ref: string): string | undefined => {
    const path = ref.replace(/^\/+/, '')
    const file = opened.archive.readBytes(path)
    if (!file) return undefined
    return `data:${mimeOf(path)};base64,${Buffer.from(file).toString('base64')}`
  }
  return opened.deck.slides.map((slide, index) =>
    buildRenderSlide(slide, opened.deck.size, {
      fitWidthPx,
      media,
      slideNo: index + 1,
    }),
  )
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

export async function installWorkStudioSlidesHost(options: HostOptions): Promise<void> {
  let currentBytes = base64ToBytes(options.data.bytes)
  let opened = await openPptx(currentBytes)
  let slides = buildSlides(opened, FIT_WIDTH)
  let dirty = false
  let closed = false
  const undoStack: Uint8Array[] = []
  const redoStack: Uint8Array[] = []
  let transformPreview = false
  let elementClipboard: any[] | null = null
  let masterPath: string | null = null
  let masterSlide: any = null
  let masterItems: any[] = []
  const masterParsed = new Map<string, any>()
  let pipWindow: any = null
  let pipDeck: string[] = []
  let pipState: any = null
  let pipStrokes: Array<{ color: string; points: Array<{ x: number; y: number }> }> = []
  let pipLaser: { x: number; y: number } | null = null
  const audienceNavHandlers = new Set<(action: 'next' | 'prev' | 'exit') => void>()
  const openedHandlers = new Set<(result: any) => void>()
  const historyHandlers = new Set<(state: { canUndo: boolean; canRedo: boolean }) => void>()
  const deckHandlers = new Set<(state: any) => void>()
  const renamedHandlers = new Set<(path: string) => void>()
  const closeSaveHandlers = new Set<() => void>()

  const result = () => ({
    path: options.data.ref || options.data.name,
    slides,
    size: { cx: opened.deck.size.cx, cy: opened.deck.size.cy },
  })

  const notifyHistory = (): void => {
    const state = {
      canUndo: undoStack.length > 0,
      canRedo: redoStack.length > 0,
    }
    for (const handler of historyHandlers) handler(state)
  }

  const notifyDeck = (): void => {
    for (const handler of deckHandlers) handler({ slides, size: result().size })
  }

  const rebuildSlide = (index: number): RenderSlide | null => {
    const slide = opened.deck.slides[index]
    if (!slide) return null
    const rendered = buildRenderSlide(slide, opened.deck.size, {
      fitWidthPx: FIT_WIDTH,
      slideNo: index + 1,
      media: (ref: string) => {
        const path = ref.replace(/^\/+/, '')
        const file = opened.archive.readBytes(path)
        if (!file) return undefined
        return `data:${mimeOf(path)};base64,${Buffer.from(file).toString('base64')}`
      },
    })
    slides = slides.map((item, i) => (i === index ? rendered : item))
    return rendered
  }

  const applyOps = (
    ops: any[],
    slideIndex?: number,
  ): RenderSlide | null => {
    const result = runTxn(opened, { ops })
    if (!result.applied) {
      console.warn('[pptx] transaction rejected', result.failures)
      return null
    }
    dirty = true
    window.extension.bus.send({ type: 'dirty', dirty: true })
    notifyHistory()
    if (slideIndex !== undefined) {
      const rendered = rebuildSlide(slideIndex)
      notifyDeck()
      return rendered
    }
    slides = buildSlides(opened, FIT_WIDTH)
    notifyDeck()
    return null
  }

  const runDeckEdit = async (ops: any[]): Promise<RenderSlide[] | null> => {
    undoStack.push(await savePptx(opened))
    redoStack.length = 0
    const result = runTxn(opened, { ops })
    if (!result.applied) {
      undoStack.pop()
      return null
    }
    dirty = true
    window.extension.bus.send({ type: 'dirty', dirty: true })
    notifyHistory()
    slides = buildSlides(opened, FIT_WIDTH)
    notifyDeck()
    return slides
  }

  const runEdit = async (ops: any[], slideIndex?: number): Promise<RenderSlide | null> => {
    undoStack.push(await savePptx(opened))
    redoStack.length = 0
    const rendered = applyOps(ops, slideIndex)
    if (!rendered) {
      undoStack.pop()
      return null
    }
    return rendered
  }

  const runStructuralEdit = async (
    ops: any[],
    slideIndex: number,
  ): Promise<RenderSlide | null> => {
    undoStack.push(await savePptx(opened))
    redoStack.length = 0
    const result = runTxn(opened, { ops })
    if (!result.applied) {
      undoStack.pop()
      return null
    }
    dirty = true
    window.extension.bus.send({ type: 'dirty', dirty: true })
    notifyHistory()
    if (!materializeSlide(opened, slideIndex)) {
      undoStack.pop()
      return null
    }
    return rebuildSlide(slideIndex)
  }

  const elementIndex = (slideIndex: number, sourceId: string): number => {
    const slide = opened.deck.slides[slideIndex]
    return slide ? slide.elements.findIndex((el) => el.id === sourceId) : -1
  }

  const chartData = (slideIndex: number, sourceId: string) => {
    const slide: any = opened.deck.slides[slideIndex]
    const el = slide?.elements?.find((item: any) => item.id === sourceId && item.type === 'chart')
    if (!el) return null
    const chart = el.chart
    return {
      kind: chart.kind,
      title: chart.title ?? '',
      categories: [...chart.categories],
      series: chart.series.map((s: any) => ({
        name: s.name ?? '',
        values: s.values.map((v: number | null) => v ?? 0),
      })),
      seriesColors: chart.series.map((s: any) => s.color),
      pointColors: chart.series.map((s: any) => s.pointColors),
    }
  }

  const CHART_SCHEME_LIST = [
    { key: 'default', label: '默认配色', colors: ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000'] },
    { key: 'blue', label: '蓝色', colors: ['#2E75B6', '#4472C4', '#5B9BD5', '#70AD47'] },
    { key: 'warm', label: '暖色', colors: ['#ED7D31', '#FFC000', '#FF0000', '#C55A11'] },
    { key: 'cool', label: '冷色', colors: ['#0070C0', '#00B0F0', '#00B0A0', '#7030A0'] },
    { key: 'mono', label: '单色', colors: ['#404040', '#666666', '#888888', '#AAAAAA'] },
  ]

  const chartSchemeColors = (key: string): string[] | undefined => {
    if (!key || key === 'default') return undefined
    const named = CHART_SCHEME_LIST.find((s) => s.key === key)
    if (named) return named.colors
    if (key === 'colorful') return ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
    if (key === 'colorful2') return ['#264478', '#9E480E', '#636363', '#997300', '#255E91', '#43682B']
    const mono = /^mono-accent(\d)$/.exec(key)
    if (mono) {
      const accents = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
      const base = accents[Number(mono[1]) - 1] ?? accents[0]!
      return [base, base + 'CC', base + '99', base + '66']
    }
    return undefined
  }

  const toEmu = (px: number, fitWidthPx: number): number => {
    const baseWidthPx = opened.deck.size.cx / EMU_PER_PX_96
    const scale = (fitWidthPx || FIT_WIDTH) / baseWidthPx
    return Math.round((px / scale) * EMU_PER_PX_96)
  }
  const toEmuWith = (fitWidthPx: number, px: number): number => toEmu(px, fitWidthPx)

  const transformOp = (op: any, item: any) => ({
    op: 'setTransform',
    target: { slide: op.slideIndex, el: item.sourceId },
    box: {
      x: toEmu(item.xPx, op.fitWidthPx),
      y: toEmu(item.yPx, op.fitWidthPx),
      cx: toEmu(item.wPx, op.fitWidthPx),
      cy: toEmu(item.hPx, op.fitWidthPx),
    },
    rotDeg: item.rotationDeg,
  })

  const save = async (): Promise<{
    ok: boolean
    path?: string
    error?: string
    slides?: RenderSlide[]
  }> => {
    try {
      const bytes = await savePptx(opened)
      const saved = await window.extension.ura.page.saveBytes<{
        ok: boolean
        version: string
      }>(options.sessionId, bytes)
      if (!saved?.ok) return { ok: false, error: '保存失败' }
      currentBytes = bytes
      opened = await openPptx(currentBytes)
      slides = buildSlides(opened, FIT_WIDTH)
      dirty = false
      window.extension.bus.send({ type: 'dirty', dirty: false })
      notifyDeck()
      return { ok: true, path: options.data.ref, slides }
    } catch (error) {
      return {
        ok: false,
        error: error instanceof Error ? error.message : String(error),
      }
    }
  }

  const openResult = async (fitWidthPx = FIT_WIDTH) => {
    opened = await openPptx(currentBytes)
    slides = buildSlides(opened, fitWidthPx)
    return {
      path: options.data.ref || options.data.name,
      slides,
      size: { cx: opened.deck.size.cx, cy: opened.deck.size.cy },
    }
  }

  const api: any = {
    // 生命周期 / 主题 / 语言
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
    getAutoSaveDefault: async () => ({ on: false, updatedAt: 0 }),
    onAutoSaveDefaultChanged: () => () => {},
    setAutoSavePref: () => {},
    onChromePressed: () => () => {},

    // 打开 / 渲染
    consumePendingOpen: (fitWidthPx: number) => openResult(fitWidthPx),
    openPptx: (fitWidthPx: number) => openResult(fitWidthPx),
    openPptxPath: (fitWidthPx: number) => openResult(fitWidthPx),
    newBlank: (fitWidthPx: number) => openResult(fitWidthPx),
    consumeHeadlessExport: async () => null,
    headlessExportDone: () => {},
    onOpened: (handler: (result: any) => void) => {
      openedHandlers.add(handler)
      return () => openedHandlers.delete(handler)
    },
    getRenderSlides: async () => slides,
    getSlideSize: async () => ({ cx: opened.deck.size.cx, cy: opened.deck.size.cy }),
    getLayouts: async () => ({
      layouts: listSlideLayouts(opened.archive),
      size: { cx: opened.deck.size.cx, cy: opened.deck.size.cy },
    }),
    getRecentFiles: async () => [],
    onRenamed: (handler: (path: string) => void) => {
      renamedHandlers.add(handler)
      return () => renamedHandlers.delete(handler)
    },
    onOpenedDeck: () => () => {},

    // 保存 / 历史
    save,
    saveAs: () => save(),
    isDirty: async () => dirty,
    onHistoryChanged: (handler: (state: { canUndo: boolean; canRedo: boolean }) => void) => {
      historyHandlers.add(handler)
      handler({ canUndo: false, canRedo: false })
      return () => historyHandlers.delete(handler)
    },
    onDeckChanged: (handler: (state: any) => void) => {
      deckHandlers.add(handler)
      return () => deckHandlers.delete(handler)
    },
    onCloseSaveRequest: (handler: () => void) => {
      closeSaveHandlers.add(handler)
      return () => closeSaveHandlers.delete(handler)
    },
    reportCloseSaveResult: () => {},
    undo: async () => {
      const previous = undoStack.pop()
      if (!previous) return null
      redoStack.push(await savePptx(opened))
      opened = await openPptx(previous)
      slides = buildSlides(opened, FIT_WIDTH)
      dirty = true
      window.extension.bus.send({ type: 'dirty', dirty: true })
      notifyHistory()
      return slides
    },
    redo: async () => {
      const next = redoStack.pop()
      if (!next) return null
      undoStack.push(await savePptx(opened))
      opened = await openPptx(next)
      slides = buildSlides(opened, FIT_WIDTH)
      dirty = true
      window.extension.bus.send({ type: 'dirty', dirty: true })
      notifyHistory()
      return slides
    },
    beginHistoryBatch: () => {},
    endHistoryBatch: () => {},
    cancelHistoryBatch: () => {},

    // Notes / comments
    getNotes: async (slideIndex: number) => {
      const slide = opened.deck.slides[slideIndex]
      return slide ? getSlideNotes(opened.archive, slide.path) : ''
    },
    setNotes: async (op: any) => {
      const next = await runDeckEdit([{ op: 'setNotes', target: { slide: op.slideIndex }, text: op.text }])
      return next !== null
    },
    getTransition: async (slideIndex: number) => {
      const slide = opened.deck.slides[slideIndex]
      return slide ? getSlideTransition(slide) : 'none'
    },
    getShowTiming: async (slideIndex: number) => {
      const slide = opened.deck.slides[slideIndex]
      return slide
        ? { advanceMs: getSlideAdvanceTime(slide), narrationIds: [] }
        : { advanceMs: null, narrationIds: [] }
    },
    getHeaderFooter: async (slideIndex: number) => {
      const slide = opened.deck.slides[slideIndex]
      return slide ? readHeaderFooter(slide) : { footer: null, slideNum: false, date: null }
    },
    getAnimations: async () => [],
    getSlideLinks: async (slideIndex: number) =>
      getSlideLinks(opened, slideIndex).map(({ elementId, target }) => ({
        sourceId: elementId,
        target,
      })),
    getRunLinks: async (slideIndex: number) =>
      getRunLinks(opened, slideIndex).map(({ elementId, ...rest }) => ({
        sourceId: elementId,
        ...rest,
      })),
    getLink: async (slideIndex: number, sourceId: string) =>
      getElementLink(opened, slideIndex, sourceId),
    getChartData: async (slideIndex: number, sourceId: string) => chartData(slideIndex, sourceId),
    getChartColorSchemes: async () => CHART_SCHEME_LIST,
    addComment: async () => null,
    deleteComment: async () => null,
    getComments: async () => [],
    getSlideComments: async () => [],
    setThreadResolved: async () => false,

    // 字体 / 媒体
    fontCatalog: async () => [],
    privateFontFaces: async () => [],
    privateFontData: async () => null,
    fontMissing: async () => [],
    fontDownload: async () => ({ ok: false, error: '字体下载不可用' }),
    fontInstallLocal: async () => ({ families: [] }),
    onFontsChanged: () => () => {},
    readLocalImage: async () => null,
    pickPictureFile: async () => null,
    getMediaData: async () => null,
    getModel3dData: async () => null,

    // 剪贴板
    hasSlideClipboard: async () => false,
    copyElements: async () => 0,
    copyElementsImage: async () => false,
    copySlides: async () => false,
    pasteSlide: async () => null,
    repasteSlide: async () => null,
    pasteElements: async () => null,
    duplicateElements: async () => null,

    // 导出 / 打印（页面内暂不支持，交给后续 WorkStudio 命令）
    pickExportDir: async () => 'download',
    exportImages: async (op: any) => {
      const paths: string[] = []
      op.pngsBase64?.forEach((png: string, index: number) => {
        const name = `${op.baseName || 'slide'}-${String(index + 1).padStart(2, '0')}.png`
        const anchor = document.createElement('a')
        anchor.href = `data:image/png;base64,${png}`
        anchor.download = name
        anchor.click()
        paths.push(name)
      })
      return { ok: true, paths }
    },

    savePicture: async (op: any) => {
      const name = op.defaultName ? `${op.defaultName}.png` : (op.fileName ?? 'selection.png')
      const anchor = document.createElement('a')
      anchor.href = `data:image/png;base64,${op.pngBase64 ?? op.base64 ?? ''}`
      anchor.download = name
      anchor.click()
      return { ok: true, path: name }
    },
    setShowFullScreen: async (on: boolean) => {
      try {
        if (on && !document.fullscreenElement) await document.documentElement.requestFullscreen()
        else if (!on && document.fullscreenElement) await document.exitFullscreen()
      } catch {
        /* ignore */
      }
    },
    openExternal: async () => {},
    onMenuCommand: () => () => {},

    // 图片 URL 插入
    fetchImage: async () => null,

    // 附件
    pickAttachments: async () => null,
    addAttachmentPaths: async () => ({ accepted: [], rejected: [] }),
    addPastedImage: async () => ({ accepted: [], rejected: [] }),
    readAttachment: async () => ({ ok: false }),
    readAttachmentImage: async () => ({ ok: false }),
    getPathForFile: () => '',
    openSourcePath: async () => {},
    getUserDisplayName: async () => 'WorkStudio',
    onUserDisplayNameChanged: () => () => {},
  }

  // 编辑 API：先用 no-op 兜底；后续替换为页面内 pptx-ops 实现。
  api.addSlide = async (op: any) => {
    const next = await runDeckEdit([
      {
        op: 'duplicateSlide',
        target: { slide: op.sourceIndex },
        ...(op.clearText ? { clearText: true } : {}),
      },
    ])
    return next ? { slides: next, index: op.sourceIndex + 1 } : null
  }

  api.addBlankSlide = async (op: any) => {
    const next = await runDeckEdit([
      { op: 'addBlankSlide', target: { slide: op.sourceIndex } },
      ...(op.before
        ? [{ op: 'moveSlide', target: { slide: op.sourceIndex }, to: op.sourceIndex + 1 }]
        : []),
    ])
    return next
      ? { slides: next, index: op.before ? op.sourceIndex : op.sourceIndex + 1 }
      : null
  }

  api.addSlideWithLayout = async (op: any) => {
    const ops: any[] = [
      {
        op: 'addSlideWithLayout',
        target: { slide: op.sourceIndex },
        layoutPath: op.layoutPath,
      },
    ]
    if (op.before) {
      ops.push({ op: 'moveSlide', target: { slide: op.sourceIndex + 1 }, to: op.sourceIndex })
    }
    const next = await runDeckEdit(ops)
    const index = op.before ? op.sourceIndex : op.sourceIndex + 1
    return next ? { slides: next, index } : null
  }

  api.moveSlide = async (op: any) => {
    const next = await runDeckEdit([
      { op: 'moveSlide', target: { slide: op.fromIndex }, to: op.toIndex },
    ])
    return next ? { slides: next, index: op.toIndex } : null
  }

  api.moveSlides = async (op: any) => {
    const indexes = [...new Set(op.slideIndexes as number[])].sort((a, b) => a - b)
    if (!indexes.length) return null
    const ops: any[] = []
    let insertAt = op.insertAt
    for (const index of indexes) {
      ops.push({ op: 'moveSlide', target: { slide: index }, to: insertAt })
      insertAt += 1
    }
    const next = await runDeckEdit(ops)
    return next ? { slides: next, sections: getSections(opened) } : null
  }

  api.moveSection = async (op: any) => {
    const next = await runDeckEdit([{ op: 'moveSection', id: op.id, dir: op.dir }])
    return next ? { slides: next, sections: getSections(opened) } : null
  }

  api.removeSectionSlides = async (op: any) => {
    const total = opened.deck.slides.length
    const sections = getSections(opened)
    let removed: number[] = []
    if (op.id == null) {
      const covered = new Set(sections.flatMap((s) => s.slideIndices))
      const first = sections.length
        ? Math.min(...sections.flatMap((s) => s.slideIndices))
        : total
      removed = Array.from({ length: first }, (_, i) => i).filter((i) => !covered.has(i))
    } else {
      const target = sections.find((s) => s.id === op.id)
      if (!target) return null
      removed = [...target.slideIndices]
    }
    if (!removed.length || removed.length >= total) return null
    const removedSet = new Set(removed)
    const shifted = (i: number) => i - removed.filter((r) => r < i).length
    const rebuilt = sections
      .filter((s) => s.id !== op.id)
      .map((s) => ({
        ...s,
        slideIndices: s.slideIndices
          .filter((i) => !removedSet.has(i))
          .map((i) => shifted(i)),
      }))
      .filter((s) => s.slideIndices.length > 0)
    const ops: any[] = removed
      .sort((a, b) => b - a)
      .map((i) => ({ op: 'deleteSlide', target: { slide: i } }))
    ops.push({ op: 'setSections', sections: rebuilt })
    const next = await runDeckEdit(ops)
    return next ? { slides: next, sections: getSections(opened) } : null
  }

  api.deleteSlide = async (slideIndex: number) =>
    runDeckEdit([{ op: 'deleteSlide', target: { slide: slideIndex } }])

  api.deleteSlides = async (op: any) => {
    const indexes = [...new Set(op.slideIndexes as number[])].sort((a, b) => b - a)
    if (!indexes.length || indexes.length >= opened.deck.slides.length) return null
    return runDeckEdit(indexes.map((i) => ({ op: 'deleteSlide', target: { slide: i } })))
  }

  api.duplicateSlides = async (op: any) => {
    const indexes = [...new Set(op.slideIndexes as number[])].sort((a, b) => a - b)
    if (!indexes.length) return null
    const ops: any[] = []
    let insertAt = Math.max(...indexes) + 1
    for (const index of indexes) {
      ops.push({ op: 'duplicateSlide', target: { slide: index } })
      ops.push({ op: 'moveSlide', target: { slide: insertAt }, to: insertAt })
      insertAt += 1
    }
    // The duplicate lands directly after its source; move it into the copied block.
    const next = await runDeckEdit(ops)
    return next ? { slides: next, index: Math.max(...indexes) + 1 } : null
  }

  api.reorderSlides = async (op: any) => {
    const next = await runDeckEdit(
      [{ op: 'moveSlide', target: { slide: op.slideIndex }, to: op.toIndex }],
    )
    return next
  }

  api.moveSlidesTo = async (op: any) => {
    const indexes = [...new Set(op.slideIndexes as number[])].sort((a, b) => a - b)
    if (!indexes.length) return null
    const ops: any[] = []
    let insertAt = op.insertAt
    for (const index of indexes) {
      ops.push({ op: 'moveSlide', target: { slide: index }, to: insertAt })
      insertAt += 1
    }
    return runDeckEdit(ops)
  }

  api.setSlidesHidden = async (op: any) =>
    runDeckEdit(
      op.slideIndexes.map((i: number) => ({
        op: 'setHidden',
        target: { slide: i },
        hidden: op.hidden,
      })),
    )

  api.addElement = async (op: any) => {
    const baseWidthPx = opened.deck.size.cx / EMU_PER_PX_96
    const scale = (op.fitWidthPx || FIT_WIDTH) / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const paragraphs = op.paragraphs?.length
      ? op.paragraphs
      : op.text
        ? op.text.split('\n').map((line: string) => ({ runs: [{ text: line }] }))
        : undefined
    const beforeIds = new Set(
      (slides[op.slideIndex]?.nodes ?? []).map((node) => node.sourceId),
    )
    const next = await runDeckEdit([
      {
        op: 'addElement',
        target: { slide: op.slideIndex },
        kind: op.kind,
        offset: {
          x: toEmu(op.xPx),
          y: toEmu(op.yPx),
          cx: toEmu(op.wPx),
          cy: toEmu(op.hPx),
        },
        ...(paragraphs ? { paragraphs } : {}),
        ...(op.fillColor ? { fill: op.fillColor } : {}),
        ...(op.stroke
          ? {
              stroke: {
                color: op.stroke.color,
                widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT),
              },
            }
          : {}),
        ...(op.bodyPr ? { bodyPr: op.bodyPr } : {}),
      },
    ])
    if (!next) return null
    const slide = next[op.slideIndex]
    const sourceId = slide?.nodes.find((node) => !beforeIds.has(node.sourceId))?.sourceId
    return slide && sourceId ? { slide, sourceId } : null
  }

  api.editStroke = async (op: any) => {
    const stroke = op.stroke
      ? {
          color: op.stroke.color,
          widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT),
          ...(op.stroke.dash ? { dash: op.stroke.dash } : {}),
          ...(op.stroke.cap ? { cap: op.stroke.cap } : {}),
          ...(op.stroke.join ? { join: op.stroke.join } : {}),
          ...(op.stroke.compound ? { compound: op.stroke.compound } : {}),
          ...(op.stroke.gradient
            ? {
                gradient: {
                  stops: op.stroke.gradient.stops,
                  angle: Math.round(op.stroke.gradient.angleDeg * 60000),
                },
              }
            : {}),
        }
      : null
    return runEdit(
      [
        {
          op: 'setStroke',
          target: { slide: op.slideIndex, el: op.sourceId },
          stroke,
          ...(op.groupId ? { group: op.groupId } : {}),
        },
      ],
      op.slideIndex,
    )
  }

  api.editFill = async (op: any) => {
    const fill =
      typeof op.fill === 'string'
        ? op.fill
        : {
            stops: op.fill.gradient.stops,
            ...(op.fill.gradient.path
              ? { path: op.fill.gradient.path }
              : { angle: Math.round((op.fill.gradient.angleDeg ?? 0) * 60000) }),
          }
    return runEdit(
      [
        {
          op: 'setFill',
          target: { slide: op.slideIndex, el: op.sourceId },
          fill,
          ...(op.groupId ? { group: op.groupId } : {}),
        },
      ],
      op.slideIndex,
    )
  }

  api.flipElements = async (op: any) =>
    runEdit(
      [
        {
          op: 'flipElements',
          target: { slide: op.slideIndex },
          els: op.sourceIds,
          axis: op.axis,
          ...(op.groupId ? { group: op.groupId } : {}),
        },
      ],
      op.slideIndex,
    )

  api.reorderElement = async (op: any) =>
    runEdit(
      [{ op: 'reorderElement', target: { slide: op.slideIndex, el: op.sourceId }, dir: op.dir }],
      op.slideIndex,
    )

  api.groupElements = async (op: any) => {
    const before = new Set((slides[op.slideIndex]?.nodes ?? []).map((n) => n.sourceId))
    const slide = await runEdit(
      [{ op: 'groupElements', target: { slide: op.slideIndex }, els: op.sourceIds }],
      op.slideIndex,
    )
    if (!slide) return null
    const groupId = slide.nodes.find((n) => !before.has(n.sourceId))?.sourceId
    return groupId ? { slide, groupId } : null
  }

  api.ungroupElement = async (op: any) =>
    runEdit(
      [{ op: 'ungroupElement', target: { slide: op.slideIndex, el: op.sourceId } }],
      op.slideIndex,
    )

  api.changeShape = async (op: any) =>
    runEdit(
      [
        {
          op: 'setShapeGeometry',
          target: { slide: op.slideIndex, el: op.sourceId },
          prst: op.prst,
          ...(op.groupId ? { group: op.groupId } : {}),
        },
      ],
      op.slideIndex,
    )

  api.setShapeAdjust = async (op: any) =>
    runEdit(
      [
        {
          op: 'setShapeAdjust',
          target: { slide: op.slideIndex, el: op.sourceId },
          adjust: op.adjust,
          ...(op.groupId ? { group: op.groupId } : {}),
        },
      ],
      op.slideIndex,
    )

  api.setShapeGeometry = async (op: any) => {
    const toEmu = (px: number) => toEmuWith(op.fitWidthPx, px)
    return runEdit(
      [
        {
          op: 'setShapeCustomGeometry',
          target: { slide: op.slideIndex, el: op.sourceId },
          path: {
            w: toEmu(op.pathPx.w),
            h: toEmu(op.pathPx.h),
            cmds: op.pathPx.cmds.map((c: any) => ({
              op: c.op,
              pts: c.pts.map(toEmu),
            })),
          },
          ...(op.groupId ? { group: op.groupId } : {}),
        },
      ],
      op.slideIndex,
    )
  }

  api.setTextAnchor = async (op: any) =>
    runEdit(
      [
        {
          op: 'setTextAnchor',
          target: { slide: op.slideIndex, el: op.sourceId },
          anchor: op.anchor,
        },
      ],
      op.slideIndex,
    )

  api.setTextBodyProps = async (op: any) =>
    runEdit(
      [
        {
          op: 'setTextBodyProps',
          target: { slide: op.slideIndex, el: op.sourceId },
          props: op.props,
        },
      ],
      op.slideIndex,
    )

  api.setEffects = async (op: any) =>
    runEdit(
      [
        {
          op: 'setEffects',
          target: { slide: op.slideIndex, el: op.sourceId },
          effects: op.effects,
        },
      ],
      op.slideIndex,
    )

  api.editTableCell = async (op: any) =>
    runEdit(
      [
        {
          op: 'setTableCell',
          target: { slide: op.slideIndex, el: op.sourceId },
          row: op.row,
          col: op.col,
          paragraphs: op.paragraphs,
        },
      ],
      op.slideIndex,
    )

  api.tableMerge = async (op: any) =>
    runEdit(
      [
        {
          op: 'tableMerge',
          target: { slide: op.slideIndex, el: op.sourceId },
          kind: op.kind,
          row: op.row,
          col: op.col,
        },
      ],
      op.slideIndex,
    )

  api.tableStructure = async (op: any) =>
    runEdit(
      [
        {
          op: 'tableStructure',
          target: { slide: op.slideIndex, el: op.sourceId },
          kind: op.kind,
          index: op.index,
          ...(op.before ? { before: true } : {}),
        },
      ],
      op.slideIndex,
    )

  api.setTableRowHeight = async (op: any) =>
    runEdit(
      [
        {
          op: 'setTableRowHeight',
          target: { slide: op.slideIndex, el: op.sourceId },
          row: op.row,
          hEmu: toEmuWith(op.fitWidthPx, op.hPx),
        },
      ],
      op.slideIndex,
    )

  api.setTableCellAnchor = async (op: any) =>
    runEdit(
      [
        {
          op: 'setTableCellAnchor',
          target: { slide: op.slideIndex, el: op.sourceId },
          row: op.row,
          col: op.col,
          anchor: op.anchor,
        },
      ],
      op.slideIndex,
    )

  api.setTableColWidth = async (op: any) =>
    runEdit(
      [
        {
          op: 'setTableColWidth',
          target: { slide: op.slideIndex, el: op.sourceId },
          col: op.col,
          wEmu: toEmuWith(op.fitWidthPx, op.wPx),
        },
      ],
      op.slideIndex,
    )

  api.editBackground = async (op: any) => {
    if (op.kind === 'image') return null
    const common =
      op.kind === 'solid'
        ? { kind: 'solid', color: op.color }
        : op.kind === 'gradient'
          ? {
              kind: 'gradient',
              from: op.from,
              to: op.to,
              ...(op.angleDeg !== undefined ? { angleDeg: op.angleDeg } : {}),
              ...(op.radial ? { radial: true } : {}),
            }
          : op.kind === 'reset'
            ? { kind: 'reset' }
            : { kind: 'graphics', hidden: op.hidden }
    const targets =
      op.slideIndex === -1 ? opened.deck.slides.map((_, i) => i) : [op.slideIndex]
    const next = await runDeckEdit(
      targets.map((i) => ({ op: 'setBackground', target: { slide: i }, ...common })),
    )
    return next
  }

  const newSourceId = (
    slideIndex: number,
    before: Set<string>,
    slide?: RenderSlide | null,
    type?: string,
  ) =>
    (slide ?? slides[slideIndex])?.nodes.find(
      (n) => !before.has(n.sourceId) && (!type || n.type === type),
    )?.sourceId

  const pickFile = (accept: string) =>
    new Promise<{ name: string; base64: string; ext: string; bytes: Uint8Array } | null>(
      (resolve) => {
        const input = document.createElement('input')
        input.type = 'file'
        input.accept = accept
        input.onchange = () => {
          const file = input.files?.[0]
          if (!file) return resolve(null)
          const ext = file.name.split('.').pop()?.toLowerCase() || 'bin'
          const reader = new FileReader()
          reader.onload = () => {
            const data = String(reader.result ?? '')
            const base64 = data.slice(data.indexOf(',') + 1)
            const binary = atob(base64)
            const bytes = new Uint8Array(binary.length)
            for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
            resolve({ name: file.name, base64, ext, bytes })
          }
          reader.onerror = () => resolve(null)
          reader.readAsDataURL(file)
        }
        input.click()
      },
    )

  api.pickPictureFile = () =>
    new Promise<{ base64: string; ext: string } | null>((resolve) => {
      const input = document.createElement('input')
      input.type = 'file'
      input.accept = 'image/*'
      input.onchange = () => {
        const file = input.files?.[0]
        if (!file) return resolve(null)
        const ext = file.name.split('.').pop()?.toLowerCase() || 'png'
        const reader = new FileReader()
        reader.onload = () => {
          const data = String(reader.result ?? '')
          resolve({ base64: data.slice(data.indexOf(',') + 1), ext })
        }
        reader.onerror = () => resolve(null)
        reader.readAsDataURL(file)
      }
      input.click()
    })

  api.insertImage = async (slideIndex: number, fitWidthPx: number) => {
    const picked = await api.pickPictureFile()
    if (!picked) return null
    return api.addImageBytes({
      slideIndex,
      fitWidthPx,
      base64: picked.base64,
      ext: picked.ext,
      xPx: 120,
      yPx: 120,
      wPx: 420,
      hPx: 280,
    })
  }

  api.addTable = async (op: any) => {
    const before = new Set((slides[op.slideIndex]?.nodes ?? []).map((n) => n.sourceId))
    const rowH = op.rowHeightEmu && op.rowHeightEmu > 0 ? Math.round(op.rowHeightEmu) : null
    const slide = await runEdit(
      [
        {
          op: 'addTable',
          target: { slide: op.slideIndex },
          rows: op.rows,
          cols: op.cols,
          offset: {
            x: toEmuWith(op.fitWidthPx, op.xPx),
            y: toEmuWith(op.fitWidthPx, op.yPx),
            cx: toEmuWith(op.fitWidthPx, op.wPx),
            cy: rowH ? rowH * op.rows : toEmuWith(op.fitWidthPx, op.hPx),
          },
          ...(rowH
            ? { rowHeightsEmu: Array.from({ length: op.rows }, () => rowH) }
            : {}),
        },
      ],
      op.slideIndex,
    )
    const sourceId = newSourceId(op.slideIndex, before, slide, 'table')
    return slide && sourceId ? { slide, sourceId } : null
  }

  api.addChart = async (op: any) => {
    const before = new Set((slides[op.slideIndex]?.nodes ?? []).map((n) => n.sourceId))
    const slide = await runEdit(
      [
        {
          op: 'addChart',
          target: { slide: op.slideIndex },
          kind: op.kind === 'barH' ? 'bar' : op.kind,
          ...(op.kind === 'barH' ? { barDir: 'bar' } : {}),
          ...(op.title ? { title: op.title } : {}),
          categories: op.categories,
          series: op.series,
          offset: {
            x: toEmuWith(op.fitWidthPx, op.xPx),
            y: toEmuWith(op.fitWidthPx, op.yPx),
            cx: toEmuWith(op.fitWidthPx, op.wPx),
            cy: toEmuWith(op.fitWidthPx, op.hPx),
          },
        },
      ],
      op.slideIndex,
    )
    const sourceId = newSourceId(op.slideIndex, before, slide, 'chart')
    return slide && sourceId ? { slide, sourceId } : null
  }

  api.addSmartArt = async (op: any) => {
    const before = new Set((slides[op.slideIndex]?.nodes ?? []).map((n) => n.sourceId))
    const slide = await runEdit(
      [
        {
          op: 'addSmartArt',
          target: { slide: op.slideIndex },
          layout: op.layout,
          items: op.items,
          offset: {
            x: toEmuWith(op.fitWidthPx, op.xPx),
            y: toEmuWith(op.fitWidthPx, op.yPx),
            cx: toEmuWith(op.fitWidthPx, op.wPx),
            cy: toEmuWith(op.fitWidthPx, op.hPx),
          },
        },
      ],
      op.slideIndex,
    )
    const sourceId = newSourceId(op.slideIndex, before, slide)
    return slide && sourceId ? { slide, sourceId } : null
  }

  api.addImageBytes = async (op: any) => {
    const before = new Set((slides[op.slideIndex]?.nodes ?? []).map((n) => n.sourceId))
    const bytes = new Uint8Array(Buffer.from(op.base64, 'base64'))
    const slide = await runEdit(
      [
        {
          op: 'addPicture',
          target: { slide: op.slideIndex },
          bytes,
          ext: op.ext,
          offset: {
            x: toEmuWith(op.fitWidthPx, op.xPx ?? 0),
            y: toEmuWith(op.fitWidthPx, op.yPx ?? 0),
            cx: Math.max(1, toEmuWith(op.fitWidthPx, op.wPx ?? 320)),
            cy: Math.max(1, toEmuWith(op.fitWidthPx, op.hPx ?? 180)),
          },
          ...(op.name ? { name: op.name } : {}),
        },
      ],
      op.slideIndex,
    )
    const sourceId = newSourceId(op.slideIndex, before, slide, 'picture')
    return slide && sourceId ? { slide, sourceId } : { error: 'unsupported', ext: op.ext }
  }

  api.addInk = async (op: any) => {
    const before = new Set((slides[op.slideIndex]?.nodes ?? []).map((n) => n.sourceId))
    const slide = await runEdit(
      [
        {
          op: 'addPicture',
          target: { slide: op.slideIndex },
          bytes: new Uint8Array(Buffer.from(op.base64, 'base64')),
          ext: 'png',
          offset: {
            x: toEmuWith(op.fitWidthPx, op.xPx),
            y: toEmuWith(op.fitWidthPx, op.yPx),
            cx: Math.max(1, toEmuWith(op.fitWidthPx, op.wPx)),
            cy: Math.max(1, toEmuWith(op.fitWidthPx, op.hPx)),
          },
          name: `aislides-ink ${Date.now().toString(36)}`,
        },
      ],
      op.slideIndex,
    )
    const sourceId = newSourceId(op.slideIndex, before, slide)
    return slide && sourceId ? { slide, sourceId } : null
  }

  api.setSlideSize = async (op: any) =>
    runDeckEdit([{ op: 'setSlideSize', cx: op.cx, cy: op.cy }])

  api.setSlideLayout = async (op: any) =>
    runEdit(
      [
        {
          op: 'setSlideLayout',
          target: { slide: op.slideIndex },
          ...(op.layoutPath ? { layoutPath: op.layoutPath } : {}),
        },
      ],
      op.slideIndex,
    )

  api.findReplace = async (op: any) => {
    const before = await savePptx(opened)
    const result = runTxn(opened, {
      ops: [
        {
          op: 'findReplace',
          find: op.find,
          replace: op.replace,
          matchCase: op.matchCase,
          firstOnly: op.firstOnly,
          slideIndex: op.slideIndex,
          elementId: op.elementId,
        },
      ],
    })
    if (!result.applied) return { count: 0, slides: null }
    undoStack.push(before)
    redoStack.length = 0
    dirty = true
    notifyHistory()
    slides = buildSlides(opened, FIT_WIDTH)
    notifyDeck()
    const count = (result.records?.[0]?.after as { count?: number } | undefined)?.count ?? 0
    return { count, slides }
  }

  api.setLink = async (op: any) =>
    runEdit(
      [
        {
          op: 'setLink',
          target: { slide: op.slideIndex, el: op.sourceId },
          link: op.target,
        },
      ],
      op.slideIndex,
    )

  api.applyHeaderFooter = async (op: any) => {
    const next = await runDeckEdit([
      {
        op: 'applyHeaderFooter',
        settings: {
          footer: op.footer ?? null,
          slideNum: !!op.slideNum,
          date: op.date ?? null,
          ...(op.dateAuto ? { dateAuto: true } : {}),
        },
      },
    ])
    return next
  }

  api.setTransition = async (op: any) => {
    const indexes =
      op.slideIndex === -1
        ? opened.deck.slides.map((_, i) => i)
        : opened.deck.slides[op.slideIndex]
          ? [op.slideIndex]
          : []
    if (!indexes.length) return false
    const next = await runDeckEdit(
      indexes.map((i) => ({ op: 'setTransition', target: { slide: i }, kind: op.kind })),
    )
    return next !== null
  }

  api.setAdvanceTimes = async (op: any) => {
    const targets = op.times.filter((t: any) => opened.deck.slides[t.slideIndex])
    if (!targets.length) return false
    const next = await runDeckEdit(
      targets.map((t: any) => ({ op: 'setAdvanceTime', target: { slide: t.slideIndex }, ms: t.ms })),
    )
    return next !== null
  }

  api.applyTheme = async (op: any) => {
    const next = await runDeckEdit([
      {
        op: 'applyTheme',
        name: op.name,
        colors: op.colors,
        ...(op.majorFont ? { majorFont: op.majorFont } : {}),
        ...(op.minorFont ? { minorFont: op.minorFont } : {}),
      },
    ])
    return next ?? { error: '主题应用失败' }
  }

  api.insertMedia = async (slideIndex: number, kind: 'video' | 'audio', fitWidthPx: number) => {
    const picked = await pickFile(kind === 'video' ? 'video/*' : 'audio/*')
    if (!picked) return null
    return api.addMediaBytes({
      slideIndex,
      kind,
      base64: picked.base64,
      ext: picked.ext,
      name: picked.name,
      fitWidthPx,
    })
  }

  api.addMediaBytes = async (op: any) => {
    const bytes = op.base64
      ? new Uint8Array(Buffer.from(op.base64, 'base64'))
      : op.bytes
    if (!bytes) return null
    const before = new Set((slides[op.slideIndex]?.nodes ?? []).map((n) => n.sourceId))
    const cx = Math.round(opened.deck.size.cx * (op.kind === 'video' ? 0.6 : 0.12))
    const cy = Math.round(opened.deck.size.cy * (op.kind === 'video' ? 0.4 : 0.12))
    const slide = await runEdit(
      [
        {
          op: 'addMedia',
          target: { slide: op.slideIndex },
          kind: op.kind,
          bytes,
          ext: op.ext,
          offset: {
            x: Math.round((opened.deck.size.cx - cx) / 2),
            y: Math.round((opened.deck.size.cy - cy) / 2),
            cx,
            cy,
          },
          ...(op.name ? { name: op.name } : {}),
        },
      ],
      op.slideIndex,
    )
    const sourceId = newSourceId(op.slideIndex, before, slide)
    return slide && sourceId ? { slide, sourceId } : null
  }

  api.getMediaData = async (slideIndex: number, sourceId: string) => {
    const slide = opened.deck.slides[slideIndex]
    const el: any = slide?.elements.find((item: any) => item.id === sourceId)
    const media = el?.media
    if (!media?.target) return null
    if (media.external) return { kind: media.kind, dataUrl: media.target }
    const bytes = opened.archive.readBytes(media.target)
    if (!bytes) return null
    const ext = media.target.split('.').pop()?.toLowerCase() ?? ''
    const mime =
      ext === 'mp4' || ext === 'm4v' || ext === 'mov'
        ? 'video/mp4'
        : ext === 'webm'
          ? 'video/webm'
          : ext === 'wav'
            ? 'audio/wav'
            : ext === 'ogg'
              ? 'audio/ogg'
              : media.kind === 'video'
                ? 'video/mp4'
                : 'audio/mpeg'
    return { kind: media.kind, dataUrl: `data:${mime};base64,${Buffer.from(bytes).toString('base64')}` }
  }

  api.pickModel3d = async () => {
    const picked = await pickFile('.glb,model/gltf-binary')
    return picked ? { name: picked.name, bytes: picked.bytes } : null
  }

  api.addModel3d = async (op: any) => {
    const before = new Set((slides[op.slideIndex]?.nodes ?? []).map((n) => n.sourceId))
    const cy = Math.round(opened.deck.size.cy * 0.5)
    const cx = Math.round(cy * op.aspect)
    const slide = await runEdit(
      [
        {
          op: 'addModel3d',
          target: { slide: op.slideIndex },
          bytes: op.bytes,
          ext: 'glb',
          ...(op.posterBase64
            ? { poster: { bytes: new Uint8Array(Buffer.from(op.posterBase64, 'base64')), ext: 'png' } }
            : {}),
          offset: {
            x: Math.round((opened.deck.size.cx - cx) / 2),
            y: Math.round((opened.deck.size.cy - cy) / 2),
            cx,
            cy,
          },
          name: op.name,
        },
      ],
      op.slideIndex,
    )
    const sourceId = newSourceId(op.slideIndex, before, slide)
    return slide && sourceId ? { slide, sourceId } : null
  }

  api.getModel3dData = async (slideIndex: number, sourceId: string) => {
    const el: any = opened.deck.slides[slideIndex]?.elements.find((item: any) => item.id === sourceId)
    const target = el?.model3d?.target
    const bytes = target ? opened.archive.readBytes(target) : undefined
    return bytes ? new Uint8Array(bytes) : null
  }

  api.setModel3dView = async (op: any) =>
    runEdit(
      [
        {
          op: 'setModel3dView',
          target: { slide: op.slideIndex, el: op.sourceId },
          ...op.rot,
          ...(op.posterBase64 ? { poster: op.posterBase64, posterExt: 'png' } : {}),
        },
      ],
      op.slideIndex,
    )

  api.replacePictureBytes = async (op: any) =>
    runEdit(
      [
        {
          op: 'replacePicture',
          target: { slide: op.slideIndex, el: op.sourceId },
          bytes: new Uint8Array(Buffer.from(op.base64, 'base64')),
          ext: op.ext,
          ...(op.keepSrcRect ? { keepSrcRect: true } : {}),
        },
      ],
      op.slideIndex,
    )

  api.getAnimations = async (slideIndex: number) => {
    const slide: any = opened.deck.slides[slideIndex]
    if (!slide) return []
    const bySpid = new Map<number, any>()
    for (const el of slide.elements) {
      const spid = elementSpid(el)
      if (spid != null && !bySpid.has(spid)) bySpid.set(spid, el)
    }
    return getSlideAnimations(slide)
      .map((a: any) => {
        const el = bySpid.get(a.spid)
        if (!el) return null
        return {
          sourceId: el.id,
          targetName: el.name || el.type || '对象',
          effect: a.effect,
          trigger: a.trigger,
          durationMs: a.durationMs,
          delayMs: a.delayMs,
          ...(a.direction != null ? { direction: a.direction } : {}),
          ...(a.motionPath != null ? { motionPath: a.motionPath } : {}),
          ...(a.paragraph != null ? { paragraph: a.paragraph } : {}),
        }
      })
      .filter(Boolean)
  }

  api.setAnimations = async (op: any) => {
    const next = await runDeckEdit([
      { op: 'setAnimations', target: { slide: op.slideIndex }, items: op.items },
    ])
    return next !== null
  }

  api.getShapeKeys = async (slideIndex: number) => {
    const slide: any = opened.deck.slides[slideIndex]
    return (slide?.elements ?? []).map((el: any) => ({
      sourceId: el.id,
      spid: elementSpid(el),
      name: el.name ?? '',
    }))
  }

  api.getSections = async () => getSections(opened)

  api.setSections = async (sections: any) => {
    const next = await runDeckEdit([{ op: 'setSections', sections }])
    return next ? getSections(opened) : null
  }

  api.addSection = async (op: any) => {
    const next = await runDeckEdit([
      { op: 'addSection', atSlideIndex: op.atSlideIndex, name: op.name },
    ])
    return next ? getSections(opened) : null
  }

  api.renameSection = async (op: any) => {
    const next = await runDeckEdit([{ op: 'renameSection', id: op.id, name: op.name }])
    return next ? getSections(opened) : null
  }

  api.removeSection = async (op: any) => {
    const next = await runDeckEdit([{ op: 'removeSection', id: op.id }])
    return next ? getSections(opened) : null
  }

  const download = (bytes: Uint8Array, name: string, mime: string) => {
    const blob = new Blob([bytes as unknown as BlobPart], { type: mime })
    const url = URL.createObjectURL(blob)
    const anchor = document.createElement('a')
    anchor.href = url
    anchor.download = name
    anchor.click()
    setTimeout(() => URL.revokeObjectURL(url), 1000)
  }

  const rasterizeSvg = (svg: string, width: number, height: number): Promise<Uint8Array> =>
    new Promise((resolve, reject) => {
      const image = new Image()
      image.onload = () => {
        const canvas = document.createElement('canvas')
        canvas.width = Math.max(1, Math.round(width))
        canvas.height = Math.max(1, Math.round(height))
        const ctx = canvas.getContext('2d')
        if (!ctx) return reject(new Error('canvas unavailable'))
        ctx.drawImage(image, 0, 0, canvas.width, canvas.height)
        const data = canvas.toDataURL('image/png').split(',')[1] ?? ''
        const binary = atob(data)
        const bytes = new Uint8Array(binary.length)
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
        resolve(bytes)
      }
      image.onerror = () => reject(new Error('SVG rasterize failed'))
      image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`
    })

  api.pickExportPdfPath = async (defaultName: string) => defaultName || 'slides.pdf'

  api.exportPdf = async (op: any) => {
    try {
      const pdf = await PDFDocument.create()
      const pageWidth = Math.max(1, op.widthPx) * 0.75
      const pageHeight = Math.max(1, op.heightPx) * 0.75
      for (const page of op.pages ?? []) {
        const png = 'png' in page
          ? (() => {
              const binary = atob(page.png)
              const bytes = new Uint8Array(binary.length)
              for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
              return bytes
            })()
          : await rasterizeSvg(page.svg, op.widthPx, op.heightPx)
        const image = await pdf.embedPng(png)
        const target = pdf.addPage([pageWidth, pageHeight])
        target.drawImage(image, { x: 0, y: 0, width: pageWidth, height: pageHeight })
      }
      const bytes = await pdf.save()
      const name = (op.filePath ?? 'slides.pdf').split('/').pop() || 'slides.pdf'
      download(bytes, name, 'application/pdf')
      return { ok: true, path: name }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  }

  api.printSlides = async (op: any) => {
    try {
      const html = (op.pngsBase64 ?? [])
        .map(
          (png: string) =>
            `<img src="data:image/png;base64,${png}" style="width:100%;page-break-after:always;display:block" />`,
        )
        .join('')
      const iframe = document.createElement('iframe')
      iframe.style.cssText = 'position:fixed;left:-9999px;top:0;width:1px;height:1px'
      document.body.append(iframe)
      const doc = iframe.contentWindow?.document
      if (!doc) return { ok: false, error: 'print unavailable' }
      doc.open()
      doc.write(`<html><body style="margin:0">${html}</body></html>`)
      doc.close()
      await new Promise((resolve) => setTimeout(resolve, 300))
      iframe.contentWindow?.focus()
      iframe.contentWindow?.print()
      setTimeout(() => iframe.remove(), 1000)
      return { ok: true }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  }

  api.clipboardProbe = async () => {
    if (elementClipboard) return true
    try {
      return (await navigator.clipboard.readText()).trim().length > 0
    } catch {
      return false
    }
  }

  api.clipboardExternal = async () => {
    try {
      const items = await navigator.clipboard.read()
      for (const item of items) {
        if (item.types.includes('image/png')) {
          const blob = await item.getType('image/png')
          const binary = new Uint8Array(await blob.arrayBuffer())
          return { kind: 'image', base64: Buffer.from(binary).toString('base64'), ext: 'png' }
        }
      }
      for (const item of items) {
        if (item.types.includes('text/plain')) {
          const text = await (await item.getType('text/plain')).text()
          if (text.trim()) return { kind: 'text', text }
        }
      }
    } catch {
      try {
        const text = await navigator.clipboard.readText()
        if (text.trim()) return { kind: 'text', text }
      } catch {
        /* ignore */
      }
    }
    return { kind: 'none' }
  }

  api.copyElements = async (op: any) => {
    const slide: any = opened.deck.slides[op.slideIndex]
    if (!slide) return 0
    const items = (op.sourceIds as string[])
      .map((id) => slide.elements.find((el: any) => el.id === id))
      .filter(Boolean)
      .map((el: any) => copyElementData(opened, slide, el))
    if (items.length) elementClipboard = items
    return items.length
  }

  api.copyElementsImage = async (_token: string, pngBase64: string) => {
    try {
      const blob = await (await fetch(`data:image/png;base64,${pngBase64}`)).blob()
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })])
      return true
    } catch {
      return false
    }
  }

  api.pasteElements = async (op: any) => {
    if (!elementClipboard?.length) {
      const external = await api.clipboardExternal()
      if (external?.kind === 'image') {
        const added = await api.addImageBytes({
          slideIndex: op.slideIndex,
          fitWidthPx: op.fitWidthPx,
          base64: external.base64,
          ext: external.ext,
          xPx: 120,
          yPx: 120,
          wPx: 420,
          hPx: 280,
        })
        return added?.slide ? { slide: added.slide, sourceIds: [added.sourceId] } : null
      }
      if (external?.kind === 'text') {
        const added = await api.addElement({
          slideIndex: op.slideIndex,
          fitWidthPx: op.fitWidthPx,
          kind: 'textbox',
          xPx: 120,
          yPx: 120,
          wPx: 500,
          hPx: 120,
          text: external.text,
        })
        return added?.slide ? { slide: added.slide, sourceIds: [added.sourceId] } : null
      }
      return null
    }
    undoStack.push(await savePptx(opened))
    redoStack.length = 0
    const result = runTxn(opened, {
      ops: [
        {
          op: 'pasteElements',
          target: { slide: op.slideIndex },
          items: elementClipboard,
          dx: 0,
          dy: 0,
        },
      ],
    })
    if (!result.applied) {
      undoStack.pop()
      return null
    }
    dirty = true
    window.extension.bus.send({ type: 'dirty', dirty: true })
    notifyHistory()
    slides = buildSlides(opened, FIT_WIDTH)
    notifyDeck()
    const slide = slides[op.slideIndex]
    return slide
      ? { slide, sourceIds: result.records?.[0]?.created ?? [] }
      : null
  }

  api.duplicateElements = async (op: any) => {
    const slide: any = opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const items = (op.sourceIds as string[])
      .map((id) => slide.elements.find((el: any) => el.id === id))
      .filter(Boolean)
      .map((el: any) => copyElementData(opened, slide, el))
    if (!items.length) return null
    undoStack.push(await savePptx(opened))
    redoStack.length = 0
    const result = runTxn(opened, {
      ops: [
        {
          op: 'pasteElements',
          target: { slide: op.slideIndex },
          items,
          dx: toEmuWith(op.fitWidthPx, op.dxPx),
          dy: toEmuWith(op.fitWidthPx, op.dyPx),
        },
      ],
    })
    if (!result.applied) {
      undoStack.pop()
      return null
    }
    dirty = true
    window.extension.bus.send({ type: 'dirty', dirty: true })
    notifyHistory()
    slides = buildSlides(opened, FIT_WIDTH)
    notifyDeck()
    const next = slides[op.slideIndex]
    return next
      ? { slide: next, sourceIds: result.records?.[0]?.created ?? [] }
      : null
  }

  api.editTransform = async (op: any) => {
    const payload = transformOp(op, op)
    if (op.preview) {
      if (!transformPreview) {
        undoStack.push(await savePptx(opened))
        redoStack.length = 0
        transformPreview = true
      }
      return applyOps([payload], op.slideIndex)
    }
    if (transformPreview) {
      transformPreview = false
      return applyOps([payload], op.slideIndex)
    }
    return runEdit([payload], op.slideIndex)
  }

  api.editTransformMulti = async (op: any) => {
    const ops = op.items.map((item: any) => transformOp(op, item))
    if (!ops.length) return null
    return runEdit(ops, op.slideIndex)
  }

  api.batchEditTransform = async (op: any) => {
    const ops = op.items.map((item: any) => ({
      op: 'setTransform',
      target: { slide: op.slideIndex, el: item.sourceId },
      box: {
        x: toEmu(item.xPx, op.fitWidthPx),
        y: toEmu(item.yPx, op.fitWidthPx),
        cx: toEmu(item.wPx, op.fitWidthPx),
        cy: toEmu(item.hPx, op.fitWidthPx),
      },
      rotDeg: item.rotationDeg,
    }))
    return runEdit(ops, op.slideIndex)
  }

  api.cancelTransformPreview = async () => {
    transformPreview = false
    return null
  }

  api.deleteElement = async (op: any) =>
    runEdit(
      [{ op: 'deleteElement', target: { slide: op.slideIndex, el: op.sourceId } }],
      op.slideIndex,
    )

  api.deleteElements = async (op: any) => {
    if (!op.sourceIds.length) return null
    return runEdit(
      op.sourceIds.map((id: string) => ({
        op: 'deleteElement',
        target: { slide: op.slideIndex, el: id },
      })),
      op.slideIndex,
    )
  }

  api.editText = async (op: any) =>
    runEdit(
      [
        {
          op: 'setText',
          target: { slide: op.slideIndex, el: op.sourceId },
          paragraphs: op.paragraphs,
          ...(op.groupId ? { group: op.groupId } : {}),
        },
      ],
      op.slideIndex,
    )

  api.setElementFont = async (op: any) =>
    runEdit(
      op.sourceIds.map((id: string) => ({
        op: 'setFont',
        target: { slide: op.slideIndex, el: id },
        font: {
          fontFamily: op.fontFamily,
          fontSizePt: op.fontSizePt,
          fontSizeStep: op.fontSizeStep,
          strike: op.strike,
          bold: op.bold,
          italic: op.italic,
          underline: op.underline,
          color: op.color,
        },
        ...(op.groupId ? { group: op.groupId } : {}),
      })),
      op.slideIndex,
    )

  api.setElementParagraphFormat = async (op: any) =>
    runEdit(
      op.sourceIds.map((id: string) => ({
        op: 'setParagraphFormat',
        target: { slide: op.slideIndex, el: id },
        format: {
          bullet: op.bullet,
          bulletChar: op.bulletChar,
          level: op.level,
          align: op.align,
          lineSpacing: op.lineSpacing,
          spaceBefore: op.spaceBefore,
          spaceAfter: op.spaceAfter,
          indent: op.indent,
          rtl: op.rtl,
        },
        ...(op.groupId ? { group: op.groupId } : {}),
      })),
      op.slideIndex,
    )

  api.editConnectorEndpoints = async (op: any) =>
    runEdit(
      [
        {
          op: 'setConnectorEndpoints',
          target: { slide: op.slideIndex, el: op.sourceId },
          p1: {
            x: toEmuWith(op.fitWidthPx, op.x1Px),
            y: toEmuWith(op.fitWidthPx, op.y1Px),
          },
          p2: {
            x: toEmuWith(op.fitWidthPx, op.x2Px),
            y: toEmuWith(op.fitWidthPx, op.y2Px),
          },
          ...(op.start !== undefined ? { start: op.start } : {}),
          ...(op.end !== undefined ? { end: op.end } : {}),
        },
      ],
      op.slideIndex,
    )

  api.editPictureSrcRect = async (op: any) =>
    runEdit(
      [
        {
          op: 'setPictureSrcRect',
          target: { slide: op.slideIndex, el: op.sourceId },
          srcRect: op.srcRect ?? null,
          ...(op.boxPx
            ? {
                box: {
                  x: toEmuWith(op.fitWidthPx, op.boxPx.x),
                  y: toEmuWith(op.fitWidthPx, op.boxPx.y),
                  cx: toEmuWith(op.fitWidthPx, op.boxPx.w),
                  cy: toEmuWith(op.fitWidthPx, op.boxPx.h),
                },
              }
            : {}),
        },
      ],
      op.slideIndex,
    )

  api.editImageFill = async (op: any) => {
    const source = op.source ?? (await api.pickPictureFile())
    if (!source) return null
    return runEdit(
      op.targets.map((target: any) => ({
        op: 'setImageFill',
        target: { slide: op.slideIndex, el: target.sourceId },
        source: {
          bytes: Uint8Array.from(Buffer.from(source.base64, 'base64')),
          ext: source.ext,
        },
        tile: op.mode === 'tile',
        ...(target.groupId ? { group: target.groupId } : {}),
      })),
      op.slideIndex,
    )
  }

  api.editTableStyle = async (op: any) => {
    const elIndex = elementIndex(op.slideIndex, op.sourceId)
    if (elIndex < 0) return null
    const rendered = await runStructuralEdit(
      [
        {
          op: 'setTableStyle',
          target: { slide: op.slideIndex, el: op.sourceId },
          ...(op.styleName != null ? { styleName: op.styleName } : {}),
          ...(op.firstRow != null ? { firstRow: op.firstRow } : {}),
          ...(op.bandRow != null ? { bandRow: op.bandRow } : {}),
          ...(op.rtl != null ? { rtl: op.rtl } : {}),
          ...(op.shadingColor != null ? { shadingColor: op.shadingColor } : {}),
          ...(op.borderColor != null ? { borderColor: op.borderColor } : {}),
          ...(op.borderWidthPt != null ? { borderWidthPt: op.borderWidthPt } : {}),
          ...(op.borderPreset != null ? { borderPreset: op.borderPreset } : {}),
          ...(op.cells != null ? { cells: op.cells } : {}),
        },
      ],
      op.slideIndex,
    )
    if (!rendered) return null
    const sourceId = opened.deck.slides[op.slideIndex]?.elements[elIndex]?.id ?? null
    return { slide: rendered, sourceId }
  }

  api.editChart = async (op: any) => {
    const elIndex = elementIndex(op.slideIndex, op.sourceId)
    if (elIndex < 0) return null
    const patch: any = {}
    for (const key of [
      'kind',
      'categories',
      'series',
      'title',
      'legendPos',
      'dataLabels',
      'gridlines',
      'catAxisTitle',
      'valAxisTitle',
      'gapWidthPct',
      'switchRowCol',
      'pointColors',
    ]) {
      if (op[key] !== undefined) patch[key] = op[key]
    }
    if (patch.kind === 'barH') {
      patch.kind = 'bar'
      patch.barDir = 'bar'
    }
    if (op.colorScheme !== undefined) {
      const colors = chartSchemeColors(op.colorScheme)
      if (colors) patch.colorScheme = colors
    }
    const rendered = await runStructuralEdit(
      [{ op: 'setChart', target: { slide: op.slideIndex, el: op.sourceId }, patch }],
      op.slideIndex,
    )
    if (!rendered) return null
    const sourceId = opened.deck.slides[op.slideIndex]?.elements[elIndex]?.id ?? null
    return { slide: rendered, sourceId }
  }

  const emitAudienceNav = (action: 'next' | 'prev' | 'exit'): void => {
    for (const handler of audienceNavHandlers) handler(action)
  }

  const drawPipInk = (): void => {
    const doc = pipWindow?.document
    if (!doc) return
    const canvas = doc.getElementById('audience-ink') as HTMLCanvasElement | null
    if (!canvas) return
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    const w = canvas.width
    const h = canvas.height
    ctx.clearRect(0, 0, w, h)
    ctx.lineCap = 'round'
    ctx.lineJoin = 'round'
    for (const stroke of pipStrokes) {
      if (stroke.points.length < 2) continue
      ctx.strokeStyle = stroke.color
      ctx.lineWidth = 4
      ctx.beginPath()
      stroke.points.forEach((point, index) => {
        const x = point.x * w
        const y = point.y * h
        if (index === 0) ctx.moveTo(x, y)
        else ctx.lineTo(x, y)
      })
      ctx.stroke()
    }
    if (pipLaser && pipLaser.x >= 0) {
      ctx.fillStyle = '#e53935'
      ctx.beginPath()
      ctx.arc(pipLaser.x * w, pipLaser.y * h, 9, 0, Math.PI * 2)
      ctx.fill()
    }
  }

  const renderPip = (): void => {
    const doc = pipWindow?.document
    if (!doc) return
    const img = doc.getElementById('audience-slide') as HTMLImageElement | null
    if (!img) return
    doc.body.style.background = pipState?.white ? '#fff' : '#000'
    if (pipState?.black || pipState?.white) {
      img.style.visibility = 'hidden'
    } else {
      img.style.visibility = 'visible'
      const png = pipDeck[pipState?.idx ?? 0]
      if (png) img.src = `data:image/png;base64,${png}`
    }
    drawPipInk()
  }

  api.presenterStart = async () => {
    if (pipWindow) return { audience: true }
    const pip = (window as any).documentPictureInPicture
    if (!pip?.requestWindow) return { audience: false }
    try {
      const win = await pip.requestWindow({
        width: Math.min(1280, window.screen.availWidth),
        height: Math.min(720, window.screen.availHeight),
      })
      pipWindow = win
      const doc = win.document
      doc.title = '观众视图'
      doc.body.style.cssText = 'margin:0;overflow:hidden;background:#000;display:flex;align-items:center;justify-content:center;'
      const img = doc.createElement('img')
      img.id = 'audience-slide'
      img.style.cssText = 'width:100%;height:100%;object-fit:contain;display:block;'
      const canvas = doc.createElement('canvas')
      canvas.id = 'audience-ink'
      canvas.style.cssText = 'position:fixed;inset:0;width:100%;height:100%;pointer-events:none;'
      doc.body.append(img, canvas)
      const resize = (): void => {
        canvas.width = win.innerWidth
        canvas.height = win.innerHeight
        drawPipInk()
      }
      resize()
      win.addEventListener('resize', resize)
      win.addEventListener('pagehide', () => {
        pipWindow = null
        pipDeck = []
        pipState = null
        pipStrokes = []
        pipLaser = null
      })
      doc.addEventListener('keydown', (event: KeyboardEvent) => {
        if (event.key === 'ArrowRight' || event.key === 'PageDown' || event.key === ' ') {
          event.preventDefault()
          emitAudienceNav('next')
        } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
          event.preventDefault()
          emitAudienceNav('prev')
        } else if (event.key === 'Escape') {
          event.preventDefault()
          emitAudienceNav('exit')
        }
      })
      doc.addEventListener('click', (event: MouseEvent) => {
        if ((event.target as HTMLElement | null)?.id === 'audience-slide') emitAudienceNav('next')
      })
      renderPip()
      return { audience: true }
    } catch {
      pipWindow = null
      return { audience: false }
    }
  }

  api.presenterDeck = (pngs: string[]) => {
    pipDeck = pngs
    renderPip()
  }

  api.presenterEnd = async () => {
    try {
      pipWindow?.close()
    } catch {
      /* already closed */
    }
    pipWindow = null
    pipDeck = []
    pipState = null
    pipStrokes = []
    pipLaser = null
  }

  api.presenterSync = (state: any) => {
    pipState = state
    renderPip()
  }

  api.presenterInk = (event: any) => {
    if (event.type === 'clear') {
      pipStrokes = []
      pipLaser = null
    } else if (event.type === 'stroke-start') {
      pipStrokes.push({ color: event.color ?? '#e53935', points: [{ x: event.x, y: event.y }] })
      pipLaser = null
    } else if (event.type === 'stroke-move') {
      pipStrokes.at(-1)?.points.push({ x: event.x, y: event.y })
    } else if (event.type === 'laser') {
      pipLaser = event.x < 0 ? null : { x: event.x, y: event.y }
    }
    drawPipInk()
  }

  api.presenterSwap = async () => {
    if (!pipWindow) return false
    try {
      const screens = await (window as any).getScreenDetails?.()
      if (!screens?.screens?.length) return false
      const current = pipWindow.screen
      const other = screens.screens.find((screen: any) => screen !== current)
      if (!other) return false
      pipWindow.moveTo(other.availLeft, other.availTop)
      pipWindow.resizeTo(other.availWidth, other.availHeight)
      return true
    } catch {
      return false
    }
  }

  api.audienceNav = () => {}
  api.audienceReady = async () => pipState
  api.onShowSync = () => () => {}
  api.onShowInk = () => () => {}
  api.onAudienceNav = (handler: (action: 'next' | 'prev' | 'exit') => void) => {
    audienceNavHandlers.add(handler)
    return () => audienceNavHandlers.delete(handler)
  }
  api.localGeneratePage = async () => null
  api.landGeneratedPages = async () => null

  const renderMasterPart = (slide: any): RenderSlide =>
    buildRenderSlide(slide, opened.deck.size, {
      fitWidthPx: FIT_WIDTH,
      media: (ref: string) => {
        const path = ref.replace(/^\/+/, '')
        const file = opened.archive.readBytes(path)
        if (!file) return undefined
        return `data:${mimeOf(path)};base64,${Buffer.from(file).toString('base64')}`
      },
    })

  const runMasterEdit = async (ops: any[]): Promise<RenderSlide | null> => {
    if (!masterPath) return null
    if (!masterSlide) masterSlide = masterParsed.get(masterPath) ?? parseMasterPart(opened.archive, masterPath)
    if (!masterSlide) return null
    const result = runTxn(opened, { ops, parts: new Map([[masterPath, masterSlide]]) })
    if (!result.applied) {
      console.warn('[pptx] master transaction rejected', JSON.stringify(result.failures))
      return null
    }
    dirty = true
    window.extension.bus.send({ type: 'dirty', dirty: true })
    const rendered = renderMasterPart(masterSlide)
    masterItems = masterItems.map((item) =>
      item.partPath === masterPath ? { ...item, slide: rendered } : item,
    )
    return rendered
  }

  api.masterEnter = async () => {
    const parts = listMasterParts(opened.archive)
    masterParsed.clear()
    masterItems = parts.flatMap((info) => {
      const parsed = parseMasterPart(opened.archive, info.partPath)
      if (!parsed) return []
      masterParsed.set(info.partPath, parsed)
      return [{ ...info, slide: renderMasterPart(parsed) }]
    })
    const first = masterItems[0]
    masterPath = first?.partPath ?? null
    masterSlide = masterPath ? masterParsed.get(masterPath) : null
    return { items: masterItems }
  }

  api.masterOpen = async (partPath: string) => {
    const parsed = parseMasterPart(opened.archive, partPath)
    if (!parsed) return null
    masterPath = partPath
    masterSlide = parsed
    masterParsed.set(partPath, parsed)
    return renderMasterPart(parsed)
  }

  api.masterClose = async () => {
    masterPath = null
    masterSlide = null
    masterItems = []
    masterParsed.clear()
    slides = buildSlides(opened, FIT_WIDTH)
    notifyDeck()
    return slides
  }

  api.masterEditText = async (op: any) =>
    runMasterEdit([
      {
        op: 'setText',
        target: { part: masterPath, el: op.sourceId },
        paragraphs: op.paragraphs,
      },
    ])

  api.masterEditTransform = async (op: any) =>
    runMasterEdit([
      {
        op: 'setTransform',
        target: { part: masterPath, el: op.sourceId },
        box: {
          x: toEmu(op.xPx, op.fitWidthPx),
          y: toEmu(op.yPx, op.fitWidthPx),
          cx: toEmu(op.wPx, op.fitWidthPx),
          cy: toEmu(op.hPx, op.fitWidthPx),
        },
        rotDeg: op.rotationDeg,
      },
    ])

  api.masterEditFill = async (op: any) => {
    const fill =
      typeof op.fill === 'string'
        ? op.fill
        : {
            stops: op.fill.gradient.stops,
            ...(op.fill.gradient.path
              ? { path: op.fill.gradient.path }
              : { angle: Math.round((op.fill.gradient.angleDeg ?? 0) * 60000) }),
          }
    return runMasterEdit([
      { op: 'setFill', target: { part: masterPath, el: op.sourceId }, fill },
    ])
  }

  api.masterEditStroke = async (op: any) =>
    runMasterEdit([
      {
        op: 'setStroke',
        target: { part: masterPath, el: op.sourceId },
        stroke: op.stroke
          ? {
              color: op.stroke.color,
              widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT),
            }
          : null,
      },
    ])

  api.masterDeleteElement = async (op: any) =>
    runMasterEdit([{ op: 'deleteElement', target: { part: masterPath, el: op.sourceId } }])

  api.applyTxn = async (op: any) => {
    const result = runTxn(opened, {
      ops: op.ops,
      ...(op.isolation ? { isolation: op.isolation } : {}),
      ...(op.dryRun ? { dryRun: true } : {}),
    })
    if (op.dryRun) {
      return {
        applied: false,
        dryRun: true,
        ...(result.plan ? { plan: result.plan } : {}),
        ...(result.failures
          ? { failures: result.failures.map((f) => ({ index: f.index, error: f.error })) }
          : {}),
      }
    }
    if (!result.applied) {
      return {
        applied: false,
        ...(result.failures
          ? { failures: result.failures.map((f) => ({ index: f.index, error: f.error })) }
          : {}),
      }
    }
    undoStack.push(await savePptx(opened))
    redoStack.length = 0
    dirty = true
    window.extension.bus.send({ type: 'dirty', dirty: true })
    notifyHistory()
    slides = buildSlides(opened, FIT_WIDTH)
    notifyDeck()
    return {
      applied: true,
      records: (result.records ?? []).map((record) => ({
        op: record.op.op,
        ...(record.op.target?.slide !== undefined
          ? { target: `slide:${record.op.target.slide}` }
          : {}),
        ...(record.created ? { created: record.created } : {}),
      })),
      slides,
    }
  }

  api.nativeClipboard = async (kind: 'cut' | 'copy' | 'paste') => {
    try {
      if (kind === 'paste') {
        const text = await navigator.clipboard.readText().catch(() => '')
        if (text) {
          document.execCommand('insertText', false, text)
          return
        }
      }
      document.execCommand(kind)
    } catch {
      /* clipboard permission or unsupported */
    }
  }

  api.microphoneAccess = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true })
      stream.getTracks().forEach((track) => track.stop())
      return true
    } catch {
      return false
    }
  }

  api.saveRecording = async (op: any) => {
    const ops: any[] = []
    for (const slide of op.slides ?? []) {
      ops.push({
        op: 'addNarration',
        target: { slide: slide.slideIndex },
        bytes: Uint8Array.from(Buffer.from(slide.wavBase64, 'base64')),
        ext: 'wav',
        durationMs: slide.durationMs,
      })
      ops.push({
        op: 'setAdvanceTime',
        target: { slide: slide.slideIndex },
        ms: slide.advanceMs,
      })
    }
    return ops.length > 0 ? runDeckEdit(ops) : null
  }

  api.clearNarration = async (indexes: number[]) =>
    indexes.length > 0
      ? runDeckEdit(
          indexes.map((index) => ({ op: 'clearNarration', target: { slide: index } })),
        )
      : null

  const slidesApi = new Proxy(api, {
    get(target, prop) {
      if (prop in target) return (target as Record<PropertyKey, unknown>)[prop]
      if (typeof prop === 'string' && prop.startsWith('on')) {
        return () => () => {}
      }
      return async () => null
    },
  })
  ;(window as unknown as { slidesApi: unknown }).slidesApi = slidesApi
  ;(window as unknown as { desktop: unknown }).desktop = new Proxy({
    pickAttachments: async () => null,
    addAttachmentPaths: async () => ({ accepted: [], rejected: [] }),
    addPastedImage: async () => ({ accepted: [], rejected: [] }),
    readAttachment: async () => ({ ok: false }),
    readAttachmentImage: async () => ({ ok: false }),
    getPathForFile: () => '',
    openExternal: async () => {},
    safeExternalUrl: (url: string) => url,
  }, {
    get(target, prop) {
      if (prop in target) return (target as Record<PropertyKey, unknown>)[prop]
      if (typeof prop === 'string' && prop.startsWith('on')) return () => () => {}
      return async () => null
    },
  })

  // 初始打开结果交给 renderer 的 consumePendingOpen；这里只保留一份通知能力。
  void digest(currentBytes)
  void closed
  void notifyHistory
  void notifyDeck
  void openResult
  void options.sessionId
}
