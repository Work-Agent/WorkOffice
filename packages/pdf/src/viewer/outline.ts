// ─── 大纲（嵌入式书签） ──────────────────────────────────────────────

import type { PDFDocumentProxy } from 'pdfjs-dist'

export interface OutlineNode {
  id: string
  title: string
  page: number | null
  depth: number
  children: OutlineNode[]
}

const MAX_DEPTH = 32
const MAX_NODES = 1500

interface RawOutlineItem {
  title?: string
  dest?: unknown
  items?: RawOutlineItem[]
}

/** dest 可能是字符串（命名目标）或数组；解析失败只丢链接，不丢标题。 */
export async function pageOfDest(doc: PDFDocumentProxy, dest: unknown): Promise<number | null> {
  try {
    const target = typeof dest === 'string' ? await doc.getDestination(dest) : dest
    if (!Array.isArray(target) || target.length === 0) return null
    const first = target[0] as unknown
    // 目标数组里的整数页码是 0-based（PDF 32000-1 Table 151）。
    if (typeof first === 'number' && Number.isFinite(first)) return Math.max(1, Math.round(first) + 1)
    if (first && typeof first === 'object') {
      const index = await doc.getPageIndex(first as Parameters<PDFDocumentProxy['getPageIndex']>[0])
      return index + 1
    }
    return null
  } catch {
    return null
  }
}

export async function loadOutline(doc: PDFDocumentProxy): Promise<OutlineNode[]> {
  const raw = (await doc.getOutline()) as RawOutlineItem[] | null
  if (!raw || raw.length === 0) return []
  let count = 0

  const walk = async (items: RawOutlineItem[], depth: number): Promise<OutlineNode[]> => {
    const out: OutlineNode[] = []
    for (const item of items) {
      if (count >= MAX_NODES) break
      count++
      const children =
        depth < MAX_DEPTH && item.items?.length ? await walk(item.items, depth + 1) : []
      out.push({
        id: `outline-${count}`,
        title: (item.title ?? '').replace(/\s+/g, ' ').trim() || '—',
        page: await pageOfDest(doc, item.dest),
        depth,
        children,
      })
    }
    return out
  }

  return walk(raw, 0)
}

export function flattenOutline(nodes: OutlineNode[]): OutlineNode[] {
  const out: OutlineNode[] = []
  const walk = (items: OutlineNode[]): void => {
    for (const node of items) {
      out.push(node)
      walk(node.children)
    }
  }
  walk(nodes)
  return out
}
