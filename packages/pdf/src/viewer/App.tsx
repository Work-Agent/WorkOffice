import type { PDFDocumentLoadingTask, PDFDocumentProxy } from 'pdfjs-dist'
import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { PageGeom } from '../shared/pdf-types'
import { EmptyState } from './components/EmptyState'
import { PasswordDialog } from './components/PasswordDialog'
import { SearchBar } from './components/SearchBar'
import { Sidebar } from './components/Sidebar'
import { Toolbar } from './components/Toolbar'
import {
  geomDispSize,
  geomOfPage,
  sameGeom,
  withViewRotation,
} from './geometry'
import { createI18n } from './i18n'
import { loadOutline, pageOfDest } from './outline'
import type { OutlineNode } from './outline'
import { PdfPage } from './PdfPage'
import { useVisibleSet } from './useVisibleSet'
import { loadPdf } from './pdf-document'
import { printPdf } from './print'
import { buildSearchIndex, matchContext, searchInIndex } from './search'
import type { SearchIndex, SearchMatch } from './search'
import {
  MAX_SCALE,
  MIN_SCALE,
  PAGE_GAP,
  SCROLL_PAD,
  ZOOM_STEPS,
} from './view-config'
import type { FitMode } from './view-config'

const SAVE_COPY_LIMIT = 64 * 1024 * 1024
const TOAST_MS = 6000

type Status = 'loading' | 'password' | 'ready' | 'error'

function clampScale(value: number): number {
  if (!Number.isFinite(value) || value <= 0) return 1
  return Math.min(MAX_SCALE, Math.max(MIN_SCALE, value))
}

/** 资源面板 / agent open 带的定位：pdf:page:12、#p=12、12。 */
function pageOfLocator(locator: string): number {
  if (!locator) return 0
  for (const pattern of [/^pdf:page:(\d+)$/i, /[#?&]p=(\d+)/i, /^p(?:age)?=(\d+)$/i]) {
    const match = pattern.exec(locator)
    if (match) return Number(match[1])
  }
  const value = Number(locator)
  return Number.isInteger(value) && value > 0 ? value : 0
}

async function probeInitialGeom(doc: PDFDocumentProxy): Promise<PageGeom[]> {
  const out: PageGeom[] = []
  let base: PageGeom
  try {
    base = geomOfPage(await doc.getPage(1))
  } catch {
    base = { pw: 595, ph: 842, rot: 0, userUnit: 1 }
  }
  for (let i = 0; i < doc.numPages; i++) out.push(base)
  // 首屏附近的页面先探测；其余滚动到再补（页数上限内成本可控）。
  const probeCount = Math.min(doc.numPages, 64)
  for (let n = 1; n <= probeCount; n++) {
    try {
      out[n - 1] = geomOfPage(await doc.getPage(n))
    } catch {
      /* 单页失败沿用 base */
    }
  }
  return out
}

/**
 * 首屏之外的后台几何探测：每批 24 页，让混合页面尺寸的文档滚动高度
 * 逐步收敛；只填补仍是占位几何的页（用户已看到的页优先）。
 */
async function probeRemaining(
  doc: PDFDocumentProxy,
  placeholder: PageGeom,
  cancelled: () => boolean,
  apply: (updates: Array<[number, PageGeom]>, placeholder: PageGeom) => void,
): Promise<void> {
  const limit = Math.min(doc.numPages, 2000)
  let pageNo = 65
  while (pageNo <= limit) {
    if (cancelled()) return
    const updates: Array<[number, PageGeom]> = []
    for (let i = 0; i < 24 && pageNo <= limit; i++, pageNo++) {
      try {
        updates.push([pageNo, geomOfPage(await doc.getPage(pageNo))])
      } catch {
        /* 损坏页保持占位几何 */
      }
    }
    if (cancelled()) return
    if (updates.length > 0) apply(updates, placeholder)
    await new Promise((resolve) => setTimeout(resolve, 16))
  }
}

export function App() {
  const boot = window.__pdfBoot!
  const i18n = useMemo(() => createI18n(), [])
  const t = i18n.t

  const [status, setStatus] = useState<Status>('loading')
  const [loadProgress, setLoadProgress] = useState(0)
  const [error, setError] = useState<string | null>(null)
  const [doc, setDoc] = useState<PDFDocumentProxy | null>(null)
  const [pages, setPages] = useState<PageGeom[]>([])
  const [current, setCurrent] = useState(1)
  const [scale, setScale] = useState(1)
  const [fit, setFit] = useState<FitMode>(null)
  const [viewRotation, setViewRotation] = useState(0)
  const [sidebar, setSidebar] = useState<'thumbs' | 'outline' | null>(null)
  const [searchOpen, setSearchOpen] = useState(false)
  const [query, setQuery] = useState('')
  const [matchCase, setMatchCase] = useState(false)
  const [wholeWord, setWholeWord] = useState(false)
  const [indexing, setIndexing] = useState(false)
  const [indexProgress, setIndexProgress] = useState(0)
  const [matches, setMatches] = useState<SearchMatch[]>([])
  const [matchIndex, setMatchIndex] = useState(0)
  const [activeMatch, setActiveMatch] = useState<SearchMatch | null>(null)
  const [wrongPassword, setWrongPassword] = useState(false)
  const [warnings, setWarnings] = useState<string[]>(() =>
    boot.data.metadataWarning ? ['metadata'] : [],
  )
  const [externalVersion, setExternalVersion] = useState<string | null>(null)
  const [reloading, setReloading] = useState(false)
  const [busy, setBusy] = useState(false)
  const [toast, setToast] = useState<string | null>(null)
  const [printProgress, setPrintProgress] = useState<{ done: number; total: number } | null>(null)
  const [outline, setOutline] = useState<OutlineNode[]>([])
  const [outlineLoading, setOutlineLoading] = useState(false)
  const [outlineError, setOutlineError] = useState<string | null>(null)
  const [viewportSize, setViewportSize] = useState({ width: 0, height: 0 })

  const taskRef = useRef<PDFDocumentLoadingTask | null>(null)
  const passwordUpdateRef = useRef<((password: string) => void) | null>(null)
  const passwordCancelledRef = useRef(false)
  const loadGenRef = useRef(0)
  const containerRef = useRef<HTMLDivElement>(null)
  const pagesRef = useRef<HTMLDivElement>(null)
  const pageElsRef = useRef(new Map<number, HTMLElement>())
  const searchIndexRef = useRef<SearchIndex | null>(null)
  const searchTokenRef = useRef(0)
  const currentRef = useRef(1)
  const scrollRafRef = useRef(0)
  const startedRef = useRef(false)
  /** 适合宽度/页面以「点按钮时的页」为锚，滚动翻页时不改缩放。 */
  const fitAnchorRef = useRef(1)

  currentRef.current = current

  const showToast = useCallback((message: string) => setToast(message), [])

  useEffect(() => {
    if (!toast) return
    const timer = setTimeout(() => setToast(null), TOAST_MS)
    return () => clearTimeout(timer)
  }, [toast])

  const openDocument = useCallback(
    async (bytes: Uint8Array, password?: string) => {
      const generation = ++loadGenRef.current
      setStatus('loading')
      setError(null)
      setLoadProgress(0)
      setWrongPassword(false)
      passwordCancelledRef.current = false
      const previous = taskRef.current
      taskRef.current = null
      if (previous) {
        try {
          await previous.destroy()
        } catch {
          /* 旧任务可能已经结束 */
        }
      }

      const task = loadPdf(bytes, {
        ...(password ? { password } : {}),
        onPassword: (updatePassword, wrong) => {
          passwordUpdateRef.current = updatePassword
          setWrongPassword(wrong)
          setStatus('password')
        },
        onProgress: (loaded, total) => {
          if (total > 0) setLoadProgress(Math.min(99, Math.round((loaded / total) * 100)))
        },
      })
      taskRef.current = task

      try {
        const next = await task.promise
        if (generation !== loadGenRef.current) {
          await task.destroy()
          return
        }
        const geoms = await probeInitialGeom(next)
        if (generation !== loadGenRef.current) {
          await task.destroy()
          return
        }
        // 新文档 = 新搜索索引；旧命中全部失效。
        searchTokenRef.current++
        searchIndexRef.current = null
        setMatches([])
        setMatchIndex(0)
        setActiveMatch(null)
        setDoc(next)
        setPages(geoms)
        if (next.numPages > 64) {
          const placeholder = geoms[64] ?? geoms[0]!
          void probeRemaining(
            next,
            placeholder,
            () => generation !== loadGenRef.current,
            (updates, base) => {
              setPages((prev) => {
                let changed = false
                const merged = prev.slice()
                for (const [pageNo, geom] of updates) {
                  if (pageNo <= merged.length && merged[pageNo - 1] === base) {
                    merged[pageNo - 1] = geom
                    changed = true
                  }
                }
                return changed ? merged : prev
              })
            },
          )
        }
        setStatus('ready')
        setWrongPassword(false)
        setLoadProgress(100)
        if (!startedRef.current) {
          startedRef.current = true
          setFit('width')
        }
      } catch (cause) {
        if (generation !== loadGenRef.current) return
        if (passwordCancelledRef.current) return
        const message = cause instanceof Error ? cause.message : String(cause)
        setError(/encrypt/i.test(message) ? t('unsupportedEncryption') : message)
        setStatus('error')
      }
    },
    [t],
  )

  // 查看器已挂载：移除入口占位（boot 失败时 main.ts 仍会写入错误文本）。
  useEffect(() => {
    document.querySelector<HTMLElement>('.ws-boot')?.remove()
  }, [])

  useEffect(() => {
    void openDocument(boot.bytes)
    return () => {
      loadGenRef.current++
      void taskRef.current?.destroy().catch(() => {})
      taskRef.current = null
    }
  }, [boot, openDocument])

  // 外部变更：只提示，不产生写冲突。
  useEffect(() => {
    return boot.watchExternal(() => setExternalVersion('changed'))
  }, [boot])

  const reloadDocument = useCallback(async () => {
    if (reloading) return
    setReloading(true)
    setBusy(true)
    try {
      const fresh = await boot.reload()
      boot.data = fresh.data
      boot.version = fresh.version
      setExternalVersion(null)
      await openDocument(fresh.bytes)
    } catch (cause) {
      showToast(t('openFailed', { message: cause instanceof Error ? cause.message : String(cause) }))
    } finally {
      setReloading(false)
      setBusy(false)
    }
  }, [boot, openDocument, reloading, showToast, t])

  // 页面几何：PdfPage 可见时自己 probe，这里只吸收变化。
  const handleGeom = useCallback((pageNo: number, geom: PageGeom) => {
    setPages((prev) => {
      if (pageNo < 1 || pageNo > prev.length) return prev
      if (sameGeom(prev[pageNo - 1], geom)) return prev
      const next = prev.slice()
      next[pageNo - 1] = geom
      return next
    })
  }, [])

  const registerEl = useCallback((pageNo: number, el: HTMLElement | null) => {
    if (el) pageElsRef.current.set(pageNo, el)
    else pageElsRef.current.delete(pageNo)
  }, [])

  const { visible, setItemRef } = useVisibleSet(
    containerRef,
    pages.length,
    '100% 0px',
    status === 'ready',
  )

  const retry = useCallback(async () => {
    try {
      const fresh = await boot.refetch()
      await openDocument(fresh)
    } catch (cause) {
      showToast(t('openFailed', { message: cause instanceof Error ? cause.message : String(cause) }))
    }
  }, [boot, openDocument, showToast, t])

  const pageBoxes = useMemo(
    () => pages.map((geom) => geomDispSize(withViewRotation(geom, viewRotation))),
    [pages, viewRotation],
  )

  useEffect(() => {
    const el = containerRef.current
    if (!el) return
    const sync = (): void => setViewportSize({ width: el.clientWidth, height: el.clientHeight })
    sync()
    const observer = new ResizeObserver(sync)
    observer.observe(el)
    return () => observer.disconnect()
  }, [status])

  // 适合宽度 / 适合页面：以点击时的页为锚，窗口 resize / 旋转时重算。
  useEffect(() => {
    if (!fit || pageBoxes.length === 0) return
    if (viewportSize.width < 50 || viewportSize.height < 50) return
    const width = Math.max(1, viewportSize.width - SCROLL_PAD * 2)
    const height = Math.max(1, viewportSize.height - SCROLL_PAD * 2)
    const box =
      pageBoxes[Math.min(fitAnchorRef.current, pageBoxes.length) - 1] ?? pageBoxes[0]!
    const next = clampScale(
      fit === 'width'
        ? width / Math.max(1, box.width)
        : Math.min(width / Math.max(1, box.width), height / Math.max(1, box.height)),
    )
    setScale(next)
  }, [fit, pageBoxes, viewportSize])

  const zoomIn = useCallback(() => {
    setFit(null)
    setScale((prev) => ZOOM_STEPS.find((step) => step > prev + 1e-6) ?? MAX_SCALE)
  }, [])

  const zoomOut = useCallback(() => {
    setFit(null)
    setScale((prev) => [...ZOOM_STEPS].reverse().find((step) => step < prev - 1e-6) ?? MIN_SCALE)
  }, [])

  const goToPage = useCallback((page: number) => {
    const target = Math.min(Math.max(1, page), pages.length || 1)
    setCurrent(target)
    const el = pageElsRef.current.get(target)
    const container = containerRef.current
    if (el && container) {
      const elRect = el.getBoundingClientRect()
      const containerRect = container.getBoundingClientRect()
      const top = Math.max(
        0,
        container.scrollTop + (elRect.top - containerRect.top) - SCROLL_PAD,
      )
      container.scrollTo({ top, behavior: 'auto' })
    }
  }, [pages.length])

  const stepPage = useCallback(
    (delta: number) => {
      goToPage(currentRef.current + delta)
    },
    [goToPage],
  )

  const onScroll = useCallback(() => {
    if (scrollRafRef.current) return
    scrollRafRef.current = requestAnimationFrame(() => {
      scrollRafRef.current = 0
      const container = containerRef.current
      if (!container) return
      const innerTop = pagesRef.current?.offsetTop ?? 0
      const target = container.scrollTop + 80 - innerTop
      let best = 1
      for (const [pageNo, el] of pageElsRef.current) {
        if (el.offsetTop <= target) best = Math.max(best, pageNo)
      }
      setCurrent((prev) => (prev === best ? prev : best))
    })
  }, [])

  useEffect(() => {
    return () => {
      if (scrollRafRef.current) cancelAnimationFrame(scrollRafRef.current)
    }
  }, [])

  // Ctrl/Cmd + 滚轮缩放
  useEffect(() => {
    const el = containerRef.current
    if (!el) return
    const onWheel = (event: WheelEvent): void => {
      if (!event.ctrlKey && !event.metaKey) return
      event.preventDefault()
      if (event.deltaY < 0) zoomIn()
      else zoomOut()
    }
    el.addEventListener('wheel', onWheel, { passive: false })
    return () => el.removeEventListener('wheel', onWheel)
  }, [zoomIn, zoomOut, status])

  // ── 搜索 ───────────────────────────────────────────────────────────
  const previewMatch = useCallback(
    (match: SearchMatch) => {
      setActiveMatch(match)
      goToPage(match.pageIndex + 1)
    },
    [goToPage],
  )

  const runSearch = useCallback(
    async (text: string, caseSensitive: boolean, whole: boolean) => {
      if (!doc) return
      // 每次都推进 token：清空查询/关闭搜索也能让正在跑的索引停下来。
      const token = ++searchTokenRef.current
      if (!text.trim()) {
        setIndexing(false)
        setMatches([])
        setMatchIndex(0)
        setActiveMatch(null)
        return
      }
      let index = searchIndexRef.current
      if (!index) {
        setIndexing(true)
        setIndexProgress(0)
        try {
          index = await buildSearchIndex(doc, {
            onProgress: (pageNumber, total, chars) => {
              if (token !== searchTokenRef.current) return false
              setIndexProgress(Math.round((pageNumber / Math.max(1, total)) * 100))
              return true
            },
          })
          if (token !== searchTokenRef.current) return
          searchIndexRef.current = index
        } catch {
          if (token === searchTokenRef.current) setIndexing(false)
          return
        } finally {
          if (token === searchTokenRef.current) setIndexing(false)
        }
      }
      const found = searchInIndex(index, text, { matchCase: caseSensitive, wholeWord: whole })
      setMatches(found)
      setMatchIndex(0)
      if (found[0]) previewMatch(found[0])
      else setActiveMatch(null)
    },
    [doc, previewMatch],
  )

  useEffect(() => {
    if (!searchOpen) return
    const timer = setTimeout(() => {
      void runSearch(query, matchCase, wholeWord)
    }, 250)
    return () => clearTimeout(timer)
  }, [query, matchCase, wholeWord, searchOpen, runSearch])

  const gotoMatch = useCallback(
    (delta: number) => {
      if (matches.length === 0) return
      const next = (matchIndex + delta + matches.length) % matches.length
      setMatchIndex(next)
      previewMatch(matches[next]!)
    },
    [matchIndex, matches, previewMatch],
  )

  const contextOf = useCallback(
    (match: SearchMatch) => matchContext(searchIndexRef.current?.[match.pageIndex], query),
    [query],
  )

  const closeSearch = useCallback(() => {
    searchTokenRef.current++
    setIndexing(false)
    setSearchOpen(false)
    setActiveMatch(null)
  }, [])

  // ── 大纲 ───────────────────────────────────────────────────────────
  useEffect(() => {
    if (!doc) return
    let cancelled = false
    setOutlineLoading(true)
    setOutlineError(null)
    void loadOutline(doc)
      .then((nodes) => {
        if (!cancelled) setOutline(nodes)
      })
      .catch((cause) => {
        if (!cancelled) setOutlineError(cause instanceof Error ? cause.message : String(cause))
      })
      .finally(() => {
        if (!cancelled) setOutlineLoading(false)
      })
    return () => {
      cancelled = true
    }
  }, [doc])

  const onGoToDest = useCallback(
    async (dest: unknown) => {
      if (!doc) return
      const page = await pageOfDest(doc, dest)
      if (page) goToPage(page)
    },
    [doc, goToPage],
  )

  // ── 打印 / 另存副本 ────────────────────────────────────────────────
  const doPrint = useCallback(async () => {
    if (!doc || busy) return
    setBusy(true)
    setPrintProgress({ done: 0, total: doc.numPages })
    try {
      await printPdf(doc, {
        onProgress: (done, total) => setPrintProgress({ done, total }),
      })
    } catch (cause) {
      showToast(t('printFailed', { message: cause instanceof Error ? cause.message : String(cause) }))
    } finally {
      setBusy(false)
      setPrintProgress(null)
    }
  }, [doc, busy, showToast, t])

  const doSaveCopy = useCallback(async () => {
    if (!doc || busy) return
    if (boot.data.size > SAVE_COPY_LIMIT) {
      showToast(t('saveCopyLarge'))
      return
    }
    setBusy(true)
    try {
      const bytes = await boot.refetch()
      await boot.saveCopy(bytes)
    } catch (cause) {
      showToast(t('saveCopyFailed', { message: cause instanceof Error ? cause.message : String(cause) }))
    } finally {
      setBusy(false)
    }
  }, [boot, busy, doc, showToast, t])

  // 键盘快捷键：用 ref 拿最新回调，监听只装一次。
  const handlersRef = useRef({ doPrint, zoomIn, zoomOut, stepPage, closeSearch })
  handlersRef.current = { doPrint, zoomIn, zoomOut, stepPage, closeSearch }
  const uiRef = useRef({ searchOpen, sidebar })
  uiRef.current = { searchOpen, sidebar }

  useEffect(() => {
    const onKey = (event: KeyboardEvent): void => {
      const target = event.target as HTMLElement | null
      const typing =
        !!target &&
        (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable)
      if (event.key === 'Escape') {
        if (uiRef.current.searchOpen) handlersRef.current.closeSearch()
        else if (uiRef.current.sidebar) setSidebar(null)
        return
      }
      if (typing) return
      const mod = event.ctrlKey || event.metaKey
      if (mod && (event.key === 'f' || event.key === 'F')) {
        event.preventDefault()
        setSearchOpen(true)
        return
      }
      if (mod && (event.key === 'p' || event.key === 'P')) {
        event.preventDefault()
        void handlersRef.current.doPrint()
        return
      }
      if (mod && (event.key === '=' || event.key === '+')) {
        event.preventDefault()
        handlersRef.current.zoomIn()
        return
      }
      if (mod && event.key === '-') {
        event.preventDefault()
        handlersRef.current.zoomOut()
        return
      }
      if (event.key === 'PageDown') {
        event.preventDefault()
        handlersRef.current.stepPage(1)
        return
      }
      if (event.key === 'PageUp') {
        event.preventDefault()
        handlersRef.current.stepPage(-1)
        return
      }
      if (event.key === 'Home') {
        event.preventDefault()
        handlersRef.current.stepPage(-1_000_000)
        return
      }
      if (event.key === 'End') {
        event.preventDefault()
        handlersRef.current.stepPage(1_000_000)
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  // 密码对话框提交：updatePassword 由 PDF.js 持有，成功后 promise 继续。
  const submitPassword = (value: string): void => {
    const update = passwordUpdateRef.current
    if (!update) return
    setStatus('loading')
    setWrongPassword(false)
    update(value)
  }

  const cancelPassword = (): void => {
    passwordCancelledRef.current = true
    passwordUpdateRef.current = null
    void taskRef.current?.destroy().catch(() => {})
    taskRef.current = null
    setError(t('passwordCancelled'))
    setStatus('error')
  }

  const toggleSidebar = (mode: 'thumbs' | 'outline'): void => {
    setSidebar((prev) => (prev === mode ? null : mode))
  }

  const failedLabel = useCallback(
    (pageNo: number) => t('pageRenderFailed', { page: pageNo }),
    [t],
  )

  const ready = status === 'ready' && doc !== null && pages.length > 0

  // 从资源面板 / agent open 带过来的定位：就绪后跳一次。
  const locatorDoneRef = useRef(false)
  useEffect(() => {
    if (locatorDoneRef.current || status !== 'ready' || pages.length === 0) return
    locatorDoneRef.current = true
    const target = pageOfLocator(boot.locator)
    if (target > 0) goToPage(target)
  }, [status, pages.length, boot.locator, goToPage])

  const statusText = printProgress
    ? t('printProgress', { done: printProgress.done, total: printProgress.total })
    : reloading
      ? t('reloading')
      : toast

  return (
    <div className="pdf-app">
      {ready && doc && (
        <Toolbar
          name={boot.data.name}
          current={current}
          numPages={pages.length}
          scale={scale}
          fit={fit}
          sidebar={sidebar}
          searching={searchOpen}
          busy={busy}
          t={t}
          onGoTo={goToPage}
          onStep={stepPage}
          onZoomIn={zoomIn}
          onZoomOut={zoomOut}
          onFit={(mode) => {
            fitAnchorRef.current = currentRef.current
            setFit(mode)
          }}
          onRotate={() => setViewRotation((prev) => (prev + 90) % 360)}
          onSidebar={toggleSidebar}
          onSearch={() => setSearchOpen((prev) => !prev)}
          onPrint={() => void doPrint()}
          onSaveCopy={() => void doSaveCopy()}
          canSaveCopy={boot.data.size <= SAVE_COPY_LIMIT}
        />
      )}

      {warnings.includes('metadata') && (
        <div className="pdf-banner" role="status">
          <span>{t('largeFileMetadata')}</span>
          <button
            type="button"
            className="pdf-banner-close"
            onClick={() => setWarnings((prev) => prev.filter((item) => item !== 'metadata'))}
            aria-label={t('dismiss')}
          >
            ×
          </button>
        </div>
      )}

      {externalVersion !== null && (
        <div className="pdf-banner pdf-banner-warn" role="status">
          <span>{t('fileUpdated')}</span>
          <button
            type="button"
            className="pdf-btn pdf-btn-text"
            onClick={() => void reloadDocument()}
            disabled={reloading}
          >
            {reloading ? t('reloading') : t('reload')}
          </button>
          <button
            type="button"
            className="pdf-banner-close"
            onClick={() => setExternalVersion(null)}
            aria-label={t('dismiss')}
          >
            ×
          </button>
        </div>
      )}

      {searchOpen && ready && (
        <SearchBar
          query={query}
          onQuery={setQuery}
          matchCase={matchCase}
          wholeWord={wholeWord}
          onToggleCase={() => setMatchCase((prev) => !prev)}
          onToggleWord={() => setWholeWord((prev) => !prev)}
          indexing={indexing}
          progress={indexProgress}
          results={matches}
          index={matchIndex}
          total={matches.length}
          contextOf={contextOf}
          onNext={() => gotoMatch(1)}
          onPrev={() => gotoMatch(-1)}
          onSelect={(match, index) => {
            setMatchIndex(index)
            previewMatch(match)
          }}
          onClose={closeSearch}
          t={t}
        />
      )}

      <div className="pdf-body">
        {ready && doc && sidebar && (
          <Sidebar
            mode={sidebar}
            doc={doc}
            numPages={pages.length}
            current={current}
            outline={outline}
            outlineLoading={outlineLoading}
            outlineError={outlineError}
            onSelectPage={goToPage}
            onSelectOutline={goToPage}
            t={t}
          />
        )}

        <div
          className="pdf-scroll"
          ref={containerRef}
          onScroll={onScroll}
          tabIndex={0}
          aria-label="PDF"
        >
          {ready && doc && (
            <div className="pdf-pages" ref={pagesRef}>
              {pages.map((geom, index) => {
                const pageNo = index + 1
                const displayGeom = withViewRotation(geom, viewRotation)
                const highlights =
                  activeMatch && activeMatch.pageIndex === index ? activeMatch.rects : undefined
                return (
                  <div
                    className="pdf-page-slot"
                    key={pageNo}
                    style={{ marginBottom: PAGE_GAP }}
                    ref={setItemRef(index)}
                    data-idx={index}
                  >
                    <PdfPage
                      doc={doc}
                      pageNo={pageNo}
                      geom={displayGeom}
                      scale={scale}
                      viewRotation={viewRotation}
                      visible={visible.has(index)}
                      {...(highlights ? { highlights } : {})}
                      onGeom={handleGeom}
                      registerEl={registerEl}
                      onGoToDest={(dest) => void onGoToDest(dest)}
                      failedLabel={failedLabel}
                    />
                  </div>
                )
              })}
            </div>
          )}
          {status === 'loading' && (
            <div className="pdf-loading" role="status">
              <div className="pdf-spinner" aria-hidden="true" />
              <p>{loadProgress > 0 ? t('loadingProgress', { percent: loadProgress }) : t('loading')}</p>
            </div>
          )}
        </div>
      </div>

      {status === 'error' && (
        <EmptyState
          title={t('emptyTitle')}
          hint={t('emptyHint')}
          {...(error ? { detail: error } : {})}
          retryLabel={t('retry')}
          onRetry={() => void retry()}
        />
      )}

      {status === 'password' && (
        <PasswordDialog
          wrong={wrongPassword}
          hint={t('passwordHint')}
          onSubmit={submitPassword}
          onCancel={cancelPassword}
          t={t}
        />
      )}

      {statusText && (
        <div className="pdf-status" role="status">
          {statusText}
        </div>
      )}
    </div>
  )
}
