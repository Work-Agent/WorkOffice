// ─── engine.read：字节侧只读投影（执行点） ───────────────────────────

import {
  getSections,
  getSlideAdvanceTime,
  getSlideAnimations,
  getSlideComments,
  getSlideLinks,
  getSlideNotes,
  getSlideTransition,
  listSlideLayouts,
  openPptx,
  parseTheme,
  slideDurableId,
  type OpenedPptx,
  type Slide,
  type SlideElement,
} from '@genoffice/pptx-engine'
import {
  deckSummary,
  elementIdOf,
  emuToInches,
  emuToPt,
  summarizeElement,
  summarizeSlide,
  textOfElement,
  titleOfSlide,
} from '../shared/deck'
import { coded, messageOf, type EngineCtx, type PlanRequest } from './abi'
import { getOrOpen, sha256Hex } from './state'

const MAX_SLIDES = 200
const MAX_ELEMENTS = 300
const MAX_TEXT = 1200

function clip(value: string, max = MAX_TEXT): string {
  const text = value.replace(/\s+/g, ' ').trim()
  return text.length > max ? `${text.slice(0, max)}…` : text
}

function slideOf(opened: OpenedPptx, value: unknown): { slide: Slide; index: number } | null {
  if (typeof value === 'number' && Number.isInteger(value)) {
    const slide = opened.deck.slides[value]
    return slide ? { slide, index: value } : null
  }
  const id = String(value ?? '')
  const index = opened.deck.slides.findIndex((slide) => slideDurableId(slide) === id)
  return index >= 0 ? { slide: opened.deck.slides[index]!, index } : null
}

export async function loadOpened(
  req: Pick<PlanRequest, 'ref' | 'baseVersion' | 'bytes'>,
  ctx: EngineCtx,
): Promise<OpenedPptx> {
  const bytes = req.bytes ?? (await ctx.readBytes()).bytes
  if (!bytes || bytes.byteLength === 0) throw coded('command.invalid', 'PPTX 内容为空')
  const version = req.baseVersion ?? sha256Hex(bytes)
  const entry = await getOrOpen(req.ref, version, bytes)
  return entry.opened
}

export function contextOf(
  opened: OpenedPptx,
  params: {
    name?: string
    version?: string
    epoch?: number
    docRev?: string
  } = {},
): string {
  const summary = deckSummary(opened)
  const w = emuToInches(summary.size.cx)
  const h = emuToInches(summary.size.cy)
  const lines = [
    `# ${params.name ?? 'presentation.pptx'} slides=${summary.slideCount} size=${summary.size.cx}x${summary.size.cy}EMU (${w}x${h}in)`,
    ...(params.version ? [`version=${params.version}`] : []),
    ...(params.docRev ? [`docRev=${params.docRev}`] : []),
    ...(params.epoch !== undefined ? [`epoch=${params.epoch}`] : []),
    '用 pptx_apply 修改；target.slide 用 s_* 或 0-based 索引，target.el 用 e_*。',
    '',
    ...summary.slides.slice(0, MAX_SLIDES).map((slide) => {
      const found = opened.deck.slides[slide.index]!
      const notes = getSlideNotes(opened.archive, found.path)
      return `[${slide.index}] ${slide.id} ${JSON.stringify(slide.title || '（无标题）')} elements=${slide.elementCount}${notes ? ` notes=${JSON.stringify(clip(notes, 120))}` : ''}`
    }),
    ...(summary.slideCount > MAX_SLIDES ? [`…（还有 ${summary.slideCount - MAX_SLIDES} 页）`] : []),
  ]
  return lines.join('\n')
}

export function readSlides(opened: OpenedPptx): string {
  return deckSummary(opened)
    .slides.slice(0, MAX_SLIDES)
    .map((slide) => {
      const found = opened.deck.slides[slide.index]!
      return `[${slide.index}] ${slide.id} ${JSON.stringify(slide.title || '')} elements=${slide.elementCount} hidden=${Boolean((found as unknown as { hidden?: boolean }).hidden)}`
    })
    .join('\n')
}

export function readSlide(opened: OpenedPptx, indexOrId: unknown): string {
  const found = slideOf(opened, indexOrId)
  if (!found) throw coded('command.not-found', `找不到幻灯片：${String(indexOrId)}`)
  const { slide, index } = found
  const notes = getSlideNotes(opened.archive, slide.path)
  const links = getSlideLinks(opened, index)
  const animations = getSlideAnimations(slide)
  const lines = [
    `slide ${index} ${slideDurableId(slide)} path=${slide.path}`,
    `size=${opened.deck.size.cx}x${opened.deck.size.cy}EMU`,
    `transition=${getSlideTransition(slide)} advance=${getSlideAdvanceTime(slide) ?? '-'}ms`,
    `notes=${JSON.stringify(clip(notes, 500))}`,
    `animations=${animations.length} links=${links.length}`,
    'elements:',
  ]
  for (let i = 0; i < Math.min(slide.elements.length, MAX_ELEMENTS); i++) {
    const el = slide.elements[i]!
    const summary = summarizeElement(el)
    const box = summary.box
    const geom = box
      ? ` box=${emuToInches(box.x)},${emuToInches(box.y)},${emuToInches(box.cx)},${emuToInches(box.cy)}in`
      : ''
    const text = summary.text ? ` text=${JSON.stringify(clip(summary.text, 500))}` : ''
    const table = summary.table
      ? ` table=${summary.table.rows}x${summary.table.cols}`
      : ''
    const chart = summary.chart
      ? ` chart=${summary.chart.kind ?? ''}${summary.chart.title ? `:${clip(summary.chart.title, 80)}` : ''}`
      : ''
    lines.push(
      `  [${i}] ${summary.id} ${summary.type}${summary.name ? ` ${JSON.stringify(summary.name)}` : ''}${geom}${table}${chart}${text}`,
    )
  }
  if (slide.elements.length > MAX_ELEMENTS)
    lines.push(`  …（还有 ${slide.elements.length - MAX_ELEMENTS} 个元素）`)
  return lines.join('\n')
}

export function readNotes(opened: OpenedPptx): string {
  return opened.deck.slides
    .map((slide, index) => {
      const notes = getSlideNotes(opened.archive, slide.path)
      return notes ? `[${index}] ${slideDurableId(slide)}\n${clip(notes, 2000)}` : null
    })
    .filter(Boolean)
    .join('\n\n') || '(没有备注)'
}

export function readComments(opened: OpenedPptx): string {
  const lines: string[] = []
  for (const [index, slide] of opened.deck.slides.entries()) {
    const comments = getSlideComments(opened.archive, slide.path)
    for (const comment of comments) {
      lines.push(`[${index}] ${comment.author ?? ''}: ${clip(comment.text ?? '', 500)}`)
    }
  }
  return lines.join('\n') || '(没有批注)'
}

export function readCharts(opened: OpenedPptx): string {
  const lines: string[] = []
  for (const [index, slide] of opened.deck.slides.entries()) {
    for (const el of slide.elements) {
      if (el.type !== 'chart') continue
      const chart = (el as unknown as { chart?: unknown }).chart
      lines.push(`[${index}] ${elementIdOf(el)} ${clip(JSON.stringify(chart), 3000)}`)
    }
  }
  return lines.join('\n') || '(没有图表)'
}

export function readTables(opened: OpenedPptx): string {
  const lines: string[] = []
  for (const [index, slide] of opened.deck.slides.entries()) {
    for (const el of slide.elements) {
      if (el.type !== 'table') continue
      const rows = (el as unknown as { rows?: Array<Array<{ text?: unknown }>> }).rows ?? []
      const text = rows
        .slice(0, 20)
        .map((row) => row.map((cell) => textOfElement(cell as unknown as SlideElement)).join(' | '))
        .join('\n')
      lines.push(`[${index}] ${elementIdOf(el)} ${rows.length}x${rows[0]?.length ?? 0}\n${clip(text, 3000)}`)
    }
  }
  return lines.join('\n') || '(没有表格)'
}

export function readTheme(opened: OpenedPptx): string {
  const path = [...opened.archive.entries.keys()].find((name) => /^ppt\/theme\/theme\d+\.xml$/.test(name))
  const xml = path ? opened.archive.readText(path) : null
  if (!xml) return '(没有 theme part)'
  try {
    const theme = parseTheme(xml)
    return [
      `theme=${path}`,
      `colors=${JSON.stringify(theme.colors)}`,
      `majorFont=${theme.majorFont ?? ''} minorFont=${theme.minorFont ?? ''}`,
      `majorEaFont=${theme.majorEaFont ?? ''} minorEaFont=${theme.minorEaFont ?? ''}`,
    ].join('\n')
  } catch (error) {
    return `theme 解析失败：${messageOf(error)}`
  }
}

export function audit(opened: OpenedPptx): string {
  const summary = deckSummary(opened)
  const problems: string[] = []
  const seen = new Set<string>()
  for (const [index, slide] of opened.deck.slides.entries()) {
    if (!slide.originalXml) problems.push(`slide ${index} 缺少 originalXml`)
    for (const el of slide.elements) {
      const id = elementIdOf(el)
      if (seen.has(`${index}:${id}`)) problems.push(`slide ${index} 元素 id 重复：${id}`)
      seen.add(`${index}:${id}`)
      const box = summarizeElement(el).box
      if (
        box &&
        (box.x < 0 ||
          box.y < 0 ||
          box.x + box.cx > opened.deck.size.cx * 1.5 ||
          box.y + box.cy > opened.deck.size.cy * 1.5)
      )
        problems.push(`slide ${index} 元素越界：${id} (${emuToInches(box.x)},${emuToInches(box.y)})`)
    }
  }
  const slides = opened.deck.slides.length
  const parts = opened.archive.entries.size
  return [
    `audit: slides=${slides} parts=${parts} size=${summary.size.cx}x${summary.size.cy}`,
    `layouts=${listSlideLayouts(opened.archive).length} sections=${getSections(opened).length}`,
    ...(problems.length ? problems : ['ok']),
  ].join('\n')
}

export function find(opened: OpenedPptx, query: string, limit = 50): string {
  const needle = query.toLowerCase()
  const hits: string[] = []
  for (const [index, slide] of opened.deck.slides.entries()) {
    for (const el of slide.elements) {
      const text = textOfElement(el)
      if (!text || !text.toLowerCase().includes(needle)) continue
      hits.push(`[${index}] ${elementIdOf(el)} ${JSON.stringify(clip(text, 500))}`)
      if (hits.length >= limit) break
    }
    if (hits.length >= limit) break
  }
  return hits.join('\n') || `(没有匹配：${query})`
}

export function readMany(opened: OpenedPptx, params: Record<string, unknown>): string {
  const what = String(params['what'] ?? 'context')
  if (what === 'context') {
    return contextOf(opened, {
      ...(typeof params['name'] === 'string' ? { name: params['name'] } : {}),
      ...(typeof params['version'] === 'string' ? { version: params['version'] } : {}),
      ...(typeof params['docRev'] === 'string' ? { docRev: params['docRev'] } : {}),
      ...(typeof params['epoch'] === 'number' ? { epoch: params['epoch'] } : {}),
    })
  }
  if (what === 'outline' || what === 'slides') return readSlides(opened)
  if (what === 'slide') return readSlide(opened, params['slide'] ?? params['slideIndex'] ?? 0)
  if (what === 'notes') return readNotes(opened)
  if (what === 'comments') return readComments(opened)
  if (what === 'charts') return readCharts(opened)
  if (what === 'tables') return readTables(opened)
  if (what === 'theme') return readTheme(opened)
  if (what === 'audit') return audit(opened)
  if (what === 'find') return find(opened, String(params['query'] ?? ''), Number(params['limit'] ?? 50))
  throw coded(
    'command.invalid',
    `what 非法：${what}（可用：context / outline / slides / slide / notes / comments / charts / tables / theme / audit / find）`,
  )
}

export async function project(req: PlanRequest, ctx: EngineCtx): Promise<string> {
  const opened = await loadOpened(req, ctx)
  const params = (req.params ?? {}) as Record<string, unknown>
  return contextOf(opened, {
    ...(typeof params['name'] === 'string' ? { name: params['name'] } : {}),
    ...(req.baseVersion ? { version: req.baseVersion } : {}),
  })
}

export async function query(req: PlanRequest, ctx: EngineCtx): Promise<unknown> {
  const opened = await loadOpened(req, ctx)
  const summary = deckSummary(opened)
  return {
    total: summary.slides.length,
    rows: summary.slides,
  }
}

export { summarizeSlide, titleOfSlide }
