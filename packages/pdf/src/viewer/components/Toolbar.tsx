import { useEffect, useState } from 'react'
import type { FitMode } from '../view-config'
import {
  ChevronLeft,
  ChevronRight,
  DownloadIcon,
  EyeIcon,
  FitPageIcon,
  FitWidthIcon,
  OutlineIcon,
  PrinterIcon,
  RotateIcon,
  SearchIcon,
  ThumbsIcon,
  ZoomInIcon,
  ZoomOutIcon,
} from './icons'

export interface ToolbarProps {
  name: string
  current: number
  numPages: number
  scale: number
  fit: FitMode
  sidebar: 'thumbs' | 'outline' | null
  searching: boolean
  busy: boolean
  t: (key: string, vars?: Record<string, string | number>) => string
  onGoTo: (page: number) => void
  onStep: (delta: number) => void
  onZoomIn: () => void
  onZoomOut: () => void
  onFit: (mode: Exclude<FitMode, null>) => void
  onRotate: () => void
  onSidebar: (mode: 'thumbs' | 'outline') => void
  onSearch: () => void
  onPrint: () => void
  onSaveCopy: () => void
  canSaveCopy: boolean
}

export function Toolbar(props: ToolbarProps) {
  const { t, current, numPages } = props
  const [pageText, setPageText] = useState(String(current))
  useEffect(() => setPageText(String(current)), [current])

  const commitPage = (): void => {
    const value = Number.parseInt(pageText, 10)
    if (Number.isFinite(value)) props.onGoTo(Math.min(numPages, Math.max(1, value)))
    else setPageText(String(current))
  }

  return (
    <header className="pdf-toolbar" role="toolbar" aria-label="PDF">
      <div className="pdf-brand" title={props.name}>
        <span className="pdf-brand-name">{props.name}</span>
      </div>

      <div className="pdf-tool-group">
        <button
          type="button"
          className="pdf-btn"
          onClick={() => props.onStep(-1)}
          disabled={current <= 1}
          aria-label={t('previousPage')}
          title={t('previousPage')}
        >
          <ChevronLeft />
        </button>
        <form
          className="pdf-page-form"
          onSubmit={(event) => {
            event.preventDefault()
            commitPage()
          }}
        >
          <input
            className="pdf-page-input"
            value={pageText}
            onChange={(event) => setPageText(event.target.value)}
            onBlur={commitPage}
            aria-label={t('pageInput')}
            inputMode="numeric"
          />
        </form>
        <span className="pdf-page-total">/ {numPages}</span>
        <button
          type="button"
          className="pdf-btn"
          onClick={() => props.onStep(1)}
          disabled={current >= numPages}
          aria-label={t('nextPage')}
          title={t('nextPage')}
        >
          <ChevronRight />
        </button>
      </div>

      <div className="pdf-tool-group">
        <button
          type="button"
          className="pdf-btn"
          onClick={props.onZoomOut}
          aria-label={t('zoomOut')}
          title={t('zoomOut')}
        >
          <ZoomOutIcon />
        </button>
        <span className="pdf-zoom-label" aria-label={t('zoomLevel')}>
          {Math.round(props.scale * 100)}%
        </span>
        <button
          type="button"
          className="pdf-btn"
          onClick={props.onZoomIn}
          aria-label={t('zoomIn')}
          title={t('zoomIn')}
        >
          <ZoomInIcon />
        </button>
        <button
          type="button"
          className={`pdf-btn${props.fit === 'width' ? ' is-active' : ''}`}
          onClick={() => props.onFit('width')}
          aria-label={t('fitWidth')}
          title={t('fitWidth')}
        >
          <FitWidthIcon />
        </button>
        <button
          type="button"
          className={`pdf-btn${props.fit === 'page' ? ' is-active' : ''}`}
          onClick={() => props.onFit('page')}
          aria-label={t('fitPage')}
          title={t('fitPage')}
        >
          <FitPageIcon />
        </button>
        <button
          type="button"
          className="pdf-btn"
          onClick={props.onRotate}
          aria-label={t('rotate')}
          title={t('rotate')}
        >
          <RotateIcon />
        </button>
      </div>

      <div className="pdf-tool-spacer" />

      <div className="pdf-tool-group">
        <button
          type="button"
          className={`pdf-btn${props.sidebar === 'thumbs' ? ' is-active' : ''}`}
          onClick={() => props.onSidebar('thumbs')}
          aria-label={t('thumbnails')}
          title={t('thumbnails')}
        >
          <ThumbsIcon />
        </button>
        <button
          type="button"
          className={`pdf-btn${props.sidebar === 'outline' ? ' is-active' : ''}`}
          onClick={() => props.onSidebar('outline')}
          aria-label={t('outline')}
          title={t('outline')}
        >
          <OutlineIcon />
        </button>
        <button
          type="button"
          className={`pdf-btn${props.searching ? ' is-active' : ''}`}
          onClick={props.onSearch}
          aria-label={t('search')}
          title={t('search')}
        >
          <SearchIcon />
        </button>
        <button
          type="button"
          className="pdf-btn"
          onClick={props.onPrint}
          disabled={props.busy}
          aria-label={t('print')}
          title={t('print')}
        >
          <PrinterIcon />
        </button>
        <button
          type="button"
          className="pdf-btn"
          onClick={props.onSaveCopy}
          disabled={props.busy || !props.canSaveCopy}
          aria-label={t('saveCopy')}
          title={props.canSaveCopy ? t('saveCopy') : t('saveCopyLarge')}
        >
          <DownloadIcon />
        </button>
        <span className="pdf-readonly" title={t('readOnly')}>
          <EyeIcon />
          {t('readOnly')}
        </span>
      </div>
    </header>
  )
}
