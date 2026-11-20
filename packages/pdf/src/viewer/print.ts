// ─── 栅格打印（一期） ────────────────────────────────────────────────
//
// 逐页渲染到离屏 canvas → JPEG → 打印根；canvas 即用即弃。
// 文字不可选、矢量细节会损失；原文件向量打印需要宿主提供 print-to-pdf。

import type { PDFDocumentProxy } from 'pdfjs-dist'
import {
  MAX_PRINT_PIXELS,
  PRINT_DIALOG_TIMEOUT_MS,
  normalizePrintPages,
  printScaleForAreas,
} from '../shared/print'

export interface PrintOptions {
  pages?: number[]
  onProgress?: (done: number, total: number) => void
}

function waitForPrintDialog(): Promise<void> {
  return new Promise((resolve) => {
    let timer: ReturnType<typeof setTimeout> | null = null
    const detach = (): void => {
      if (timer !== null) clearTimeout(timer)
      timer = null
      window.removeEventListener('afterprint', onAfterPrint)
    }
    function onAfterPrint(): void {
      detach()
      resolve()
    }
    timer = setTimeout(() => {
      detach()
      resolve()
    }, PRINT_DIALOG_TIMEOUT_MS)
    window.addEventListener('afterprint', onAfterPrint)
    try {
      window.print()
    } catch {
      detach()
      resolve()
    }
  })
}

export async function printPdf(doc: PDFDocumentProxy, options: PrintOptions = {}): Promise<void> {
  const targets = normalizePrintPages(options.pages, doc.numPages)
  const total = targets.length
  let done = 0

  // 第一遍：量面积，预算共享 scale。逐页释放 PDFPageProxy。
  const areas: number[] = []
  for (const n of targets) {
    const page = await doc.getPage(n)
    try {
      const viewport = page.getViewport({ scale: 1 })
      const area = viewport.width * viewport.height
      areas.push(Number.isFinite(area) && area > 0 ? area : 0)
    } finally {
      page.cleanup()
    }
  }
  const scale = printScaleForAreas(areas)

  const root = document.createElement('div')
  root.className = 'pdf-print-root'
  const canvas = document.createElement('canvas')
  try {
    for (const n of targets) {
      const page = await doc.getPage(n)
      try {
        let viewport = page.getViewport({ scale })
        let w = Math.floor(viewport.width)
        let h = Math.floor(viewport.height)
        if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) continue
        if (w * h > MAX_PRINT_PIXELS) {
          const pageScale = scale * Math.sqrt(MAX_PRINT_PIXELS / (w * h))
          viewport = page.getViewport({ scale: pageScale })
          w = Math.floor(viewport.width)
          h = Math.floor(viewport.height)
          if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) continue
        }
        canvas.width = w
        canvas.height = h
        await page.render({ canvas, viewport }).promise
        const img = document.createElement('img')
        img.src = canvas.toDataURL('image/jpeg', 0.92)
        img.alt = `page ${n}`
        root.appendChild(img)
      } finally {
        page.cleanup()
        done++
        options.onProgress?.(done, total)
      }
    }
    canvas.width = 0
    canvas.height = 0
    document.body.appendChild(root)
    await Promise.all([...root.querySelectorAll('img')].map((img) => img.decode().catch(() => {})))
    await waitForPrintDialog()
  } finally {
    canvas.width = 0
    canvas.height = 0
    root.remove()
  }
}
