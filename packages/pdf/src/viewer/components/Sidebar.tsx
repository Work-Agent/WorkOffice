import type { PDFDocumentProxy } from 'pdfjs-dist'
import { useRef } from 'react'
import type { RefObject } from 'react'
import type { OutlineNode } from '../outline'
import { useVisibleSet } from '../useVisibleSet'
import { OutlinePanel } from './OutlinePanel'
import { PdfThumb } from './PdfThumb'

export function Sidebar({
  mode,
  doc,
  numPages,
  current,
  outline,
  outlineLoading,
  outlineError,
  onSelectPage,
  onSelectOutline,
  t,
}: {
  mode: 'thumbs' | 'outline'
  doc: PDFDocumentProxy
  numPages: number
  current: number
  outline: OutlineNode[]
  outlineLoading: boolean
  outlineError: string | null
  onSelectPage: (page: number) => void
  onSelectOutline: (page: number) => void
  t: (key: string, vars?: Record<string, string | number>) => string
}) {
  const scrollRef = useRef<HTMLDivElement>(null)
  const thumbsMode = mode === 'thumbs'
  const { visible, setItemRef } = useVisibleSet(
    scrollRef as RefObject<HTMLElement | null>,
    numPages,
    '600px 0px',
    thumbsMode,
  )

  return (
    <aside className={`pdf-sidebar pdf-sidebar-${mode}`} aria-label={t('sidebar')}>
      <div className="pdf-sidebar-head">{mode === 'thumbs' ? t('thumbnails') : t('outline')}</div>
      <div className="pdf-sidebar-scroll" ref={scrollRef}>
        {mode === 'thumbs' ? (
          <div className="pdf-thumbs">
            {Array.from({ length: numPages }, (_value, index) => {
              const pageNo = index + 1
              return (
                <div key={pageNo} ref={setItemRef(index)} data-idx={index}>
                  <PdfThumb
                    doc={doc}
                    pageNo={pageNo}
                    visible={visible.has(index)}
                    current={pageNo === current}
                    onSelect={onSelectPage}
                    label={t('pageInput')}
                  />
                </div>
              )
            })}
          </div>
        ) : (
          <OutlinePanel
            nodes={outline}
            loading={outlineLoading}
            error={outlineError}
            onSelect={onSelectOutline}
            t={t}
          />
        )}
      </div>
    </aside>
  )
}
