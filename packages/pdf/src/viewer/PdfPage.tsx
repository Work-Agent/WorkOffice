import type { PDFDocumentProxy } from 'pdfjs-dist'
import { AnnotationMode, TextLayer } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { useEffect, useRef, useState } from 'react'
import type { CSSProperties } from 'react'
import type { PageGeom } from '../shared/pdf-types'
import { geomDispSize, geomOfPage, pdfRectToCss } from './geometry'
import { LinkLayer } from './components/LinkLayer'
import { MAX_PAGE_RENDER_PIXELS } from './view-config'

interface RenderHandle {
  cancel(): void
}

export function PdfPage({
  doc,
  pageNo,
  geom,
  scale,
  viewRotation,
  visible,
  highlights,
  onGeom,
  registerEl,
  onGoToDest,
  failedLabel,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  geom: PageGeom
  scale: number
  viewRotation: number
  visible: boolean
  highlights?: [number, number, number, number][]
  onGeom: (pageNo: number, geom: PageGeom) => void
  registerEl: (pageNo: number, el: HTMLElement | null) => void
  onGoToDest: (dest: unknown) => void
  failedLabel: (pageNo: number) => string
}) {
  const contentRef = useRef<HTMLDivElement>(null)
  const [failed, setFailed] = useState(false)

  const size = geomDispSize(geom)
  const width = Math.max(1, Math.round(size.width * scale))
  const height = Math.max(1, Math.round(size.height * scale))

  useEffect(() => {
    const content = contentRef.current
    if (!content) return
    if (!visible) {
      content.replaceChildren()
      return
    }
    let cancelled = false
    let task: RenderHandle | null = null
    let textLayer: TextLayer | null = null
    setFailed(false)
    void (async () => {
      try {
        const page = await doc.getPage(pageNo)
        if (cancelled) return
        // 页面自己 probe 几何：占位尺寸与真实尺寸不一致时由 App 修正。
        const probe = geomOfPage(page)
        if (cancelled) return
        onGeom(pageNo, probe)

        const rotation = ((page.rotate + viewRotation) % 360 + 360) % 360
        const viewport = page.getViewport({ scale, rotation })
        const dpr = Math.min(
          window.devicePixelRatio || 1,
          2,
          Math.sqrt(MAX_PAGE_RENDER_PIXELS / Math.max(1, viewport.width * viewport.height)),
        )
        const canvas = document.createElement('canvas')
        canvas.width = Math.floor(viewport.width * dpr)
        canvas.height = Math.floor(viewport.height * dpr)
        canvas.style.width = `${Math.floor(viewport.width)}px`
        canvas.style.height = `${Math.floor(viewport.height)}px`
        canvas.className = 'pdf-canvas'

        const renderTask = page.render({
          canvas,
          viewport,
          // 只读：画 /AP 外观（含已填写的表单值），不挂可交互 widget。
          annotationMode: AnnotationMode.ENABLE,
          ...(dpr !== 1 ? { transform: [dpr, 0, 0, dpr, 0, 0] as [number, number, number, number, number, number] } : {}),
        })
        task = { cancel: () => renderTask.cancel() }
        await renderTask.promise
        if (cancelled) return

        const textDiv = document.createElement('div')
        textDiv.className = 'textLayer'
        content.replaceChildren(canvas, textDiv)
        textLayer = new TextLayer({
          textContentSource: page.streamTextContent(),
          container: textDiv,
          viewport,
        })
        try {
          await textLayer.render()
        } catch {
          /* 文字层失败不影响 canvas；选择/复制降级 */
        }
      } catch (error) {
        if (cancelled) return
        if (error instanceof Error && error.name === 'RenderingCancelledException') return
        setFailed(true)
        content.replaceChildren()
      }
    })()
    return () => {
      cancelled = true
      task?.cancel()
      textLayer?.cancel()
    }
  }, [doc, pageNo, scale, viewRotation, visible, onGeom])

  return (
    <div
      ref={(el) => {
        registerEl(pageNo, el)
      }}
      className="pdf-page"
      data-page={pageNo}
      style={
        {
          width,
          height,
          '--scale-factor': String(scale),
          '--user-unit': String(geom.userUnit ?? 1),
          '--total-scale-factor': String(scale * (geom.userUnit ?? 1)),
        } as CSSProperties
      }
    >
      <div ref={contentRef} className="pdf-page-content" />
      {visible && (
        <LinkLayer
          doc={doc}
          pageNo={pageNo}
          geom={geom}
          scale={scale}
          onGoToDest={onGoToDest}
        />
      )}
      {highlights?.map((rect, index) => (
        <div
          key={index}
          className="pdf-search-hit"
          style={pdfRectToCss(geom, rect, scale) as CSSProperties}
        />
      ))}
      {failed && <div className="pdf-page-error">{failedLabel(pageNo)}</div>}
    </div>
  )
}
