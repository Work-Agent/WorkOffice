import type { PDFDocumentProxy } from 'pdfjs-dist'
import { useEffect, useRef } from 'react'
import { THUMB_WIDTH } from '../view-config'

interface RenderedThumb {
  doc: PDFDocumentProxy
  pageNo: number
  width: number
}

/** 缩略图只画 canvas：可见时渲染，离屏释放像素。 */
export function PdfThumb({
  doc,
  pageNo,
  visible,
  current,
  onSelect,
  label,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  visible: boolean
  current: boolean
  onSelect: (pageNo: number) => void
  label: string
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const renderedRef = useRef<RenderedThumb | null>(null)

  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    if (!visible) {
      renderedRef.current = null
      canvas.width = 0
      canvas.height = 0
      return
    }
    const rendered = renderedRef.current
    if (rendered?.doc === doc && rendered.pageNo === pageNo && rendered.width === THUMB_WIDTH) return

    let cancelled = false
    let task: { cancel(): void } | null = null
    void (async () => {
      try {
        const page = await doc.getPage(pageNo)
        if (cancelled) return
        const rotation = page.rotate ?? 0
        const unit = page.getViewport({ scale: 1, rotation })
        const scale = THUMB_WIDTH / Math.max(1, unit.width)
        const viewport = page.getViewport({ scale, rotation })
        const dpr = Math.min(window.devicePixelRatio || 1, 2)
        canvas.width = Math.max(1, Math.floor(viewport.width * dpr))
        canvas.height = Math.max(1, Math.floor(viewport.height * dpr))
        canvas.style.width = `${Math.floor(viewport.width)}px`
        canvas.style.height = `${Math.floor(viewport.height)}px`
        const renderTask = page.render({
          canvas,
          viewport,
          ...(dpr !== 1
            ? { transform: [dpr, 0, 0, dpr, 0, 0] as [number, number, number, number, number, number] }
            : {}),
        })
        task = { cancel: () => renderTask.cancel() }
        await renderTask.promise
        if (!cancelled) renderedRef.current = { doc, pageNo, width: THUMB_WIDTH }
      } catch {
        renderedRef.current = null
      }
    })()
    return () => {
      cancelled = true
      task?.cancel()
    }
  }, [doc, pageNo, visible])

  useEffect(() => {
    const canvas = canvasRef.current
    return () => {
      if (canvas) {
        canvas.width = 0
        canvas.height = 0
      }
    }
  }, [])

  return (
    <button
      type="button"
      className={`pdf-thumb${current ? ' is-current' : ''}`}
      onClick={() => onSelect(pageNo)}
      aria-label={`${label} ${pageNo}`}
      aria-current={current ? 'page' : undefined}
    >
      <span className="pdf-thumb-canvas">
        <canvas ref={canvasRef} />
      </span>
      <span className="pdf-thumb-no">{pageNo}</span>
    </button>
  )
}
