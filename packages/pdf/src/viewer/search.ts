// ─── 全文搜索（与 GenOffice search.ts 同构的只读子集） ───────────────

import type { PDFDocumentProxy } from 'pdfjs-dist'

export interface SearchMatch {
  pageIndex: number
  rects: [number, number, number, number][]
}

interface IndexedItem {
  start: number
  end: number
  x: number
  y: number
  w: number
  h: number
  rot?: boolean
  font?: string
}

export interface PageEntry {
  text: string
  lower: string
  items: IndexedItem[]
}

export type SearchIndex = PageEntry[]

export const MAX_MATCHES = 1000
export const MAX_INDEX_PAGES = 2000
export const MAX_INDEX_CHARS = 4_000_000

export interface SearchOptions {
  matchCase?: boolean
  wholeWord?: boolean
}

const CJK_RE =
  /[\u1100-\u11FF\u2E80-\u318F\u31A0-\u31BF\u3400-\u4DBF\u4E00-\u9FFF\uA960-\uA97F\uAC00-\uD7AF\uF900-\uFAFF\uFE30-\uFE4F\uFF00-\uFFEF\u{20000}-\u{2FA1F}]/u

/** 与 GenOffice / ui 的 foldCase 对齐：搜索大小写不敏感时统一小写。 */
export function foldCase(text: string): string {
  return text.toLowerCase()
}

export function isWordChar(ch: string): boolean {
  return /[\p{L}\p{N}_]/u.test(ch)
}

function wordBoundaryOk(text: string, start: number, len: number): boolean {
  const wordChar = (ch: string | undefined): boolean =>
    !!ch && !CJK_RE.test(ch) && isWordChar(ch)
  const left = !wordChar(text[start - 1]) || !wordChar(text[start])
  const right = !wordChar(text[start + len]) || !wordChar(text[start + len - 1])
  return left && right
}

interface RawTextItem {
  str?: string
  transform?: number[]
  width?: number
  height?: number
  hasEOL?: boolean
  fontName?: string
}

export interface BuildIndexOptions {
  /** 每页构建后回调（进度）；返回 false 可取消。 */
  onProgress?: (pageNumber: number, total: number, chars: number) => boolean | void
}

export async function buildSearchIndex(
  doc: PDFDocumentProxy,
  options: BuildIndexOptions = {},
): Promise<SearchIndex> {
  const total = Math.min(doc.numPages, MAX_INDEX_PAGES)
  const entries: PageEntry[] = []
  let chars = 0
  for (let n = 1; n <= total; n++) {
    const page = await doc.getPage(n)
    const content = await page.getTextContent()
    let text = ''
    const items: IndexedItem[] = []
    for (const raw of content.items as RawTextItem[]) {
      if (typeof raw.str !== 'string') continue
      if (raw.str.length > 0) {
        if (raw.transform) {
          const h = raw.height || Math.hypot(raw.transform[2] ?? 0, raw.transform[3] ?? 0)
          items.push({
            start: text.length,
            end: text.length + raw.str.length,
            x: raw.transform[4] ?? 0,
            y: raw.transform[5] ?? 0,
            w: raw.width ?? 0,
            h,
            ...(Math.abs(raw.transform[1] ?? 0) > h * 1e-3 ? { rot: true } : {}),
            ...(typeof raw.fontName === 'string' ? { font: raw.fontName } : {}),
          })
        }
        text += raw.str
      }
      if (raw.hasEOL) text += '\n'
    }
    chars += text.length
    entries.push({ text, lower: foldCase(text), items })
    if (options.onProgress?.(n, total, chars) === false) break
    if (chars > MAX_INDEX_CHARS) break
  }
  return entries
}

export function searchInIndex(
  index: SearchIndex,
  query: string,
  options: SearchOptions = {},
): SearchMatch[] {
  const matchCase = options.matchCase ?? false
  const wholeWord = options.wholeWord ?? false
  const needle = matchCase ? query : foldCase(query)
  if (!needle) return []
  const matches: SearchMatch[] = []
  for (let pageIndex = 0; pageIndex < index.length; pageIndex++) {
    const { text, lower, items } = index[pageIndex]!
    const hay = matchCase ? text : lower
    let from = 0
    while (matches.length < MAX_MATCHES) {
      const start = hay.indexOf(needle, from)
      if (start < 0) break
      if (wholeWord && !wordBoundaryOk(text, start, needle.length)) {
        from = start + 1
        continue
      }
      const end = start + needle.length
      from = end
      const rects: [number, number, number, number][] = []
      for (const item of items) {
        if (item.end <= start || item.start >= end) continue
        const len = item.end - item.start
        const lo = (Math.max(start, item.start) - item.start) / len
        const hi = (Math.min(end, item.end) - item.start) / len
        const x1 = item.x + item.w * lo
        const x2 = item.x + item.w * hi
        if (x2 - x1 < 0.01) continue
        rects.push([x1, item.y, x2, item.y + item.h])
      }
      if (rects.length > 0) matches.push({ pageIndex, rects })
    }
    if (matches.length >= MAX_MATCHES) break
  }
  return matches
}

/** 结果上下文：命中位置前后各 20 字。 */
export function matchContext(entry: PageEntry | undefined, needle: string): string {
  if (!entry) return ''
  const at = entry.lower.indexOf(foldCase(needle))
  if (at < 0) return entry.text.slice(0, 60)
  const from = Math.max(0, at - 20)
  const to = Math.min(entry.text.length, at + needle.length + 20)
  return `${from > 0 ? '…' : ''}${entry.text.slice(from, to)}${to < entry.text.length ? '…' : ''}`
}
