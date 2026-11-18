// ─── pptx 投影辅助：标题、元素文本、盒子、备注、大纲 ─────────────────
//
// 这里是只读投影，不参与保存；页面与执行点都可复用。

import {
  elementDurableId,
  slideDurableId,
  type OpenedPptx,
  type Slide,
  type SlideElement,
} from '@genoffice/pptx-engine'

export interface SlideSummary {
  index: number
  id: string
  path: string
  title: string
  elementCount: number
  hidden?: boolean
}

export interface ElementSummary {
  id: string
  type: string
  name?: string
  box?: { x: number; y: number; cx: number; cy: number }
  rotDeg?: number
  text?: string
  table?: { rows: number; cols: number }
  chart?: { kind?: string; title?: string }
  link?: unknown
}

export function slideIdOf(slide: Slide): string {
  return slideDurableId(slide)
}

export function elementIdOf(el: SlideElement): string {
  return elementDurableId(el) ?? el.id
}

type BodyLike = {
  paragraphs?: Array<{ runs?: Array<{ text?: string }> }>
  lines?: Array<{ runs?: Array<{ text?: string }> }>
}

function bodyLines(text?: BodyLike): string[] {
  const lines: string[] = []
  for (const paragraph of text?.paragraphs ?? []) {
    const value = (paragraph.runs ?? []).map((run) => run.text ?? '').join('')
    if (value.trim()) lines.push(value)
  }
  for (const line of text?.lines ?? []) {
    const value = (line.runs ?? []).map((run) => run.text ?? '').join('')
    if (value.trim()) lines.push(value)
  }
  return lines
}

export function textOfElement(el: SlideElement): string {
  const value = el as unknown as {
    text?: BodyLike
    cells?: Array<{ text?: BodyLike }>
  }
  const lines = bodyLines(value.text)
  for (const cell of value.cells ?? []) lines.push(...bodyLines(cell.text))
  return lines.join('\n').trim()
}

export function titleOfSlide(slide: Slide): string {
  let best: { score: number; text: string } | null = null
  for (const el of slide.elements) {
    const text = textOfElement(el)
    if (!text) continue
    const first = text.split('\n')[0]?.trim() ?? ''
    if (!first) continue
    const body = (el as unknown as { text?: { paragraphs?: Array<{ runs?: Array<{ fontSize?: number }> }> } }).text
    const fontSize = body?.paragraphs?.[0]?.runs?.[0]?.fontSize ?? 0
    const placeholder = String((el as { placeholder?: unknown }).placeholder ?? '')
    const score = fontSize + (placeholder.includes('title') ? 10_000 : 0)
    if (!best || score > best.score) best = { score, text: first }
  }
  return best?.text ?? ''
}

export function boxOfElement(el: SlideElement): ElementSummary['box'] {
  const t = (el as unknown as { transform?: { x?: number; y?: number; cx?: number; cy?: number } })
    .transform
  if (!t) return undefined
  return {
    x: Number(t.x ?? 0),
    y: Number(t.y ?? 0),
    cx: Number(t.cx ?? 0),
    cy: Number(t.cy ?? 0),
  }
}

export function summarizeElement(el: SlideElement): ElementSummary {
  const anyEl = el as unknown as {
    id: string
    type: string
    name?: string
    transform?: { rotDeg?: number }
    cells?: unknown[]
    chart?: { kind?: string; title?: string }
    link?: unknown
  }
  const table = Array.isArray(anyEl.cells)
    ? {
        rows: Number((el as unknown as { rows?: number }).rows ?? anyEl.cells.length),
        cols: Number((el as unknown as { cols?: number }).cols ?? 0),
      }
    : undefined
  return {
    id: elementIdOf(el),
    type: anyEl.type,
    ...(anyEl.name ? { name: anyEl.name } : {}),
    ...(boxOfElement(el) ? { box: boxOfElement(el) } : {}),
    ...(anyEl.transform?.rotDeg !== undefined ? { rotDeg: anyEl.transform.rotDeg } : {}),
    ...(textOfElement(el) ? { text: textOfElement(el) } : {}),
    ...(table ? { table } : {}),
    ...(anyEl.chart
      ? {
          chart: {
            ...(anyEl.chart.kind ? { kind: anyEl.chart.kind } : {}),
            ...(anyEl.chart.title ? { title: anyEl.chart.title } : {}),
          },
        }
      : {}),
    ...(anyEl.link ? { link: anyEl.link } : {}),
  }
}

export function summarizeSlide(slide: Slide, index: number): SlideSummary {
  return {
    index,
    id: slideIdOf(slide),
    path: slide.path,
    title: titleOfSlide(slide),
    elementCount: slide.elements.length,
  }
}

export function deckSummary(opened: OpenedPptx): {
  slideCount: number
  size: { cx: number; cy: number }
  slides: SlideSummary[]
} {
  return {
    slideCount: opened.deck.slides.length,
    size: { ...opened.deck.size },
    slides: opened.deck.slides.map((slide, index) => summarizeSlide(slide, index)),
  }
}

export function emuToInches(value: number): number {
  return Math.round((value / 914400) * 100) / 100
}

export function emuToPt(value: number): number {
  return Math.round((value / 12700) * 100) / 100
}
