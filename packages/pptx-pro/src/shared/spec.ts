// ─── 新建规格与轻量 Markdown 解析 ────────────────────────────────────

export interface DeckSpec {
  size?: '16:9' | '4:3' | { cx: number; cy: number }
  theme?: {
    name: string
    colors: Record<string, string>
    majorFont?: string
    minorFont?: string
  }
  slides: SlideSpec[]
}

export interface SlideSpec {
  title?: string
  body?: string | string[]
  bullets?: string[]
  notes?: string
  background?: string
}

const MAX_SLIDES = 200
const MAX_TEXT = 200_000

export class DeckSpecError extends Error {}

function asText(value: unknown, field: string): string | undefined {
  if (value === undefined || value === null) return undefined
  if (typeof value !== 'string') throw new DeckSpecError(`${field} 必须是字符串`)
  if (value.length > MAX_TEXT) throw new DeckSpecError(`${field} 超过 ${MAX_TEXT} 字符`)
  return value
}

function asStringArray(value: unknown, field: string): string[] | undefined {
  if (value === undefined || value === null) return undefined
  if (!Array.isArray(value)) throw new DeckSpecError(`${field} 必须是字符串数组`)
  return value.map((item, i) => asText(item, `${field}[${i}]`) ?? '')
}

export function normalizeSpec(params: Record<string, unknown>): DeckSpec {
  const rawSlides = params['slides']
  const content = asText(params['content'], 'content')
  const slides: SlideSpec[] = []

  if (Array.isArray(rawSlides)) {
    for (const [i, raw] of rawSlides.entries()) {
      if (!raw || typeof raw !== 'object') throw new DeckSpecError(`slides[${i}] 不是对象`)
      const item = raw as Record<string, unknown>
      const body = item['body']
      slides.push({
        ...(asText(item['title'], `slides[${i}].title`) !== undefined
          ? { title: asText(item['title'], `slides[${i}].title`) }
          : {}),
        ...(typeof body === 'string'
          ? { body }
          : Array.isArray(body)
            ? { body: asStringArray(body, `slides[${i}].body`) }
            : {}),
        ...(asStringArray(item['bullets'], `slides[${i}].bullets`) !== undefined
          ? { bullets: asStringArray(item['bullets'], `slides[${i}].bullets`) }
          : {}),
        ...(asText(item['notes'], `slides[${i}].notes`) !== undefined
          ? { notes: asText(item['notes'], `slides[${i}].notes`) }
          : {}),
        ...(asText(item['background'], `slides[${i}].background`) !== undefined
          ? { background: asText(item['background'], `slides[${i}].background`) }
          : {}),
      })
    }
  } else if (content) {
    slides.push(...markdownSlides(content))
  }

  if (slides.length > MAX_SLIDES) throw new DeckSpecError(`页数超过 ${MAX_SLIDES}`)

  const size = params['size']
  const normalizedSize =
    size === '16:9' || size === '4:3'
      ? size
      : size && typeof size === 'object' &&
          Number.isFinite((size as { cx?: unknown }).cx) &&
          Number.isFinite((size as { cy?: unknown }).cy)
        ? { cx: Number((size as { cx: number }).cx), cy: Number((size as { cy: number }).cy) }
        : undefined

  const rawTheme = params['theme']
  const theme =
    rawTheme && typeof rawTheme === 'object'
      ? (rawTheme as DeckSpec['theme'])
      : undefined

  return {
    ...(normalizedSize ? { size: normalizedSize } : {}),
    ...(theme ? { theme } : {}),
    slides,
  }
}

/**
 * 轻量 Markdown：`#` 页标题，`---` 分页，`-`/`*` 项目符号，`notes:` 备注。
 * 不做图片、表格、嵌套列表；需要精确结构时用 structured slides。
 */
export function markdownSlides(markdown: string): SlideSpec[] {
  const chunks: string[][] = []
  let current: string[] = []
  for (const line of markdown.replace(/\r\n?/g, '\n').split('\n')) {
    if (/^\s*---\s*$/.test(line)) {
      chunks.push(current)
      current = []
      continue
    }
    current.push(line)
  }
  chunks.push(current)

  const slides: SlideSpec[] = []
  for (const lines of chunks) {
    if (lines.every((line) => !line.trim())) continue
    let title: string | undefined
    const body: string[] = []
    const bullets: string[] = []
    let notes: string | undefined
    for (const raw of lines) {
      const line = raw.trimEnd()
      if (!line.trim()) continue
      if (!title && /^#\s+/.test(line)) {
        title = line.replace(/^#\s+/, '').trim()
        continue
      }
      const note = /^notes:\s*(.*)$/i.exec(line)
      if (note) {
        notes = note[1] ?? ''
        continue
      }
      const bullet = /^\s*[-*]\s+(.*)$/.exec(line)
      if (bullet) {
        bullets.push(bullet[1] ?? '')
        continue
      }
      body.push(line.trim())
    }
    slides.push({
      ...(title ? { title } : {}),
      ...(body.length ? { body } : {}),
      ...(bullets.length ? { bullets } : {}),
      ...(notes !== undefined ? { notes } : {}),
    })
  }
  return slides
}
