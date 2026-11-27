import type { PDFDocumentProxy } from 'pdfjs-dist'
import { useEffect, useState } from 'react'
import type { CSSProperties } from 'react'
import type { PageGeom } from '../../shared/pdf-types'
import { pdfRectToCss } from '../geometry'

interface LinkItem {
  rect: [number, number, number, number]
  url?: string
  dest?: unknown
}

interface RawLinkAnnotation {
  subtype?: string
  rect?: number[]
  url?: string
  dest?: unknown
}

export const MAX_PAGE_LINKS = 500

/** 外链优先走宿主 openExternal；dev / 非 WorkStudio 环境退回 window.open。 */
function openExternal(url: string): void {
  const bridge = window.workStudio
  if (!bridge) {
    window.open(url, '_blank', 'noopener,noreferrer')
    return
  }
  void bridge
    .command('shell', 'native', 'openExternal', { url })
    .then((result) => {
      if (result.ok !== true) window.open(url, '_blank', 'noopener,noreferrer')
    })
    .catch(() => window.open(url, '_blank', 'noopener,noreferrer'))
}

function isFiniteRect(rect: unknown): rect is [number, number, number, number] {
  return (
    Array.isArray(rect) &&
    rect.length === 4 &&
    rect.every((n) => typeof n === 'number' && Number.isFinite(n))
  )
}

/** 外链白名单：只允许 http / https / mailto。 */
export function safeExternalUrl(url: string | undefined): string | null {
  if (!url) return null
  try {
    const parsed = new URL(url)
    return parsed.protocol === 'http:' || parsed.protocol === 'https:' || parsed.protocol === 'mailto:'
      ? parsed.href
      : null
  } catch {
    return null
  }
}

export function collectPageLinks(annots: RawLinkAnnotation[]): LinkItem[] {
  const out: LinkItem[] = []
  for (const annot of annots) {
    if (out.length >= MAX_PAGE_LINKS) break
    if (annot.subtype !== 'Link' || (!annot.url && !annot.dest)) continue
    if (!isFiniteRect(annot.rect)) continue
    if (annot.url && !safeExternalUrl(annot.url)) continue
    out.push({ rect: annot.rect, ...(annot.url ? { url: annot.url } : {}), ...(annot.dest ? { dest: annot.dest } : {}) })
  }
  return out
}

export function LinkLayer({
  doc,
  pageNo,
  geom,
  scale,
  onGoToDest,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  geom: PageGeom
  scale: number
  onGoToDest: (dest: unknown) => void
}) {
  const [links, setLinks] = useState<LinkItem[]>([])

  useEffect(() => {
    let cancelled = false
    void (async () => {
      try {
        const page = await doc.getPage(pageNo)
        const annots = (await page.getAnnotations({ intent: 'display' })) as RawLinkAnnotation[]
        if (!cancelled) setLinks(collectPageLinks(annots))
      } catch {
        if (!cancelled) setLinks([])
      }
    })()
    return () => {
      cancelled = true
    }
  }, [doc, pageNo])

  if (links.length === 0) return null

  return (
    <div className="pdf-link-layer">
      {links.map((link, index) => {
        const style = pdfRectToCss(geom, link.rect, scale) as CSSProperties
        const external = safeExternalUrl(link.url)
        return (
          <a
            key={index}
            className="pdf-link"
            style={style}
            href={external ?? '#'}
            rel="noreferrer noopener"
            target={external ? '_blank' : undefined}
            title={external ?? undefined}
            onClick={(event) => {
              if (external) {
                event.preventDefault()
                openExternal(external)
                return
              }
              event.preventDefault()
              onGoToDest(link.dest)
            }}
          />
        )
      })}
    </div>
  )
}
