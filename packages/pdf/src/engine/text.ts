// ─── PDF 文本抽取（worker 入口） ──────────────────────────────────────
//
// 主进程不直接 import 本文件；engine/text-client.ts 用 worker_threads
// 把整个抽取脚本放进 worker，避免 PDF.js 阻塞主进程。
//
// 产物：dist/text.js + dist/pdf.worker.mjs（构建脚本负责复制后者）。
// PDF.js 的 fake worker 会相对 import.meta.url 找 pdf.worker.mjs。

import { readFile } from 'node:fs/promises'
import { isMainThread, parentPort, workerData } from 'node:worker_threads'
import { getDocument, type PDFDocumentProxy } from 'pdfjs-dist/legacy/build/pdf.mjs'

export interface TextPage {
  page: number
  text: string
}

export interface TextResult {
  pageCount: number
  pages: TextPage[]
  truncated?: boolean
  warning?: string
}

export interface OutlineItem {
  title: string
  page: number | null
  depth: number
  children: OutlineItem[]
}

export interface PageInfo {
  page: number
  width: number
  height: number
  rotate: number
}

export interface PagesResult {
  pageCount: number
  pages: PageInfo[]
}

export type TextTask =
  | {
      kind: 'text'
      path: string
      pages: number[] | 'all'
      maxPages: number
      maxChars: number
      perPageChars: number
    }
  | { kind: 'outline'; path: string }
  | { kind: 'pages'; path: string; limit: number }

const MAX_OUTLINE_NODES = 1500
const MAX_OUTLINE_DEPTH = 16

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

async function withDocument<T>(
  path: string,
  run: (doc: PDFDocumentProxy) => Promise<T>,
): Promise<T> {
  const data = new Uint8Array(await readFile(path))
  const loading = getDocument({
    data,
    useSystemFonts: false,
    disableFontFace: true,
    enableXfa: false,
    verbosity: 0,
  })
  try {
    const doc = await loading.promise
    return await run(doc)
  } finally {
    await loading.destroy().catch(() => undefined)
  }
}

async function pageText(
  doc: PDFDocumentProxy,
  pageNo: number,
  limit: number,
): Promise<{ text: string; truncated: boolean }> {
  const page = await doc.getPage(pageNo)
  try {
    const content = await page.getTextContent()
    let text = ''
    for (const item of content.items) {
      if (!('str' in item)) continue
      text += item.str
      if ((item as { hasEOL?: boolean }).hasEOL) text += '\n'
      if (text.length >= limit) {
        return { text: text.slice(0, limit), truncated: true }
      }
    }
    return { text, truncated: false }
  } finally {
    page.cleanup()
  }
}

async function extractText(
  path: string,
  options: {
    pages: number[] | 'all'
    maxPages: number
    maxChars: number
    perPageChars: number
  },
): Promise<TextResult> {
  return withDocument(path, async (doc) => {
    const pageCount = doc.numPages
    const targets =
      options.pages === 'all'
        ? Array.from({ length: Math.min(pageCount, options.maxPages) }, (_v, i) => i + 1)
        : [...new Set(options.pages)]
            .filter((n) => Number.isInteger(n) && n >= 1 && n <= pageCount)
            .sort((a, b) => a - b)

    const pages: TextPage[] = []
    let chars = 0
    let truncated = options.pages === 'all' && pageCount > targets.length
    for (const pageNo of targets) {
      if (chars >= options.maxChars) {
        truncated = true
        break
      }
      const remaining = Math.max(1, Math.min(options.perPageChars, options.maxChars - chars))
      const result = await pageText(doc, pageNo, remaining)
      pages.push({ page: pageNo, text: result.text })
      chars += result.text.length
      if (result.truncated) truncated = true
    }
    return { pageCount, pages, ...(truncated ? { truncated: true } : {}) }
  })
}

async function pageOfDest(doc: PDFDocumentProxy, dest: unknown): Promise<number | null> {
  try {
    const target = typeof dest === 'string' ? await doc.getDestination(dest) : dest
    if (!Array.isArray(target) || target.length === 0) return null
    const first = target[0] as unknown
    if (typeof first === 'number' && Number.isFinite(first)) {
      return Math.max(1, Math.round(first) + 1)
    }
    if (first && typeof first === 'object') {
      const index = await doc.getPageIndex(
        first as Parameters<PDFDocumentProxy['getPageIndex']>[0],
      )
      return index + 1
    }
    return null
  } catch {
    return null
  }
}

async function extractOutline(path: string): Promise<OutlineItem[]> {
  return withDocument(path, async (doc) => {
    const raw = (await doc.getOutline()) as
      | Array<{ title?: string; dest?: unknown; items?: unknown[] }>
      | null
    if (!raw || raw.length === 0) return []
    let count = 0

    const walk = async (
      items: Array<{ title?: string; dest?: unknown; items?: unknown[] }>,
      depth: number,
    ): Promise<OutlineItem[]> => {
      const out: OutlineItem[] = []
      for (const item of items) {
        if (count >= MAX_OUTLINE_NODES) break
        count += 1
        const children =
          depth < MAX_OUTLINE_DEPTH && Array.isArray(item.items)
            ? await walk(item.items as typeof items, depth + 1)
            : []
        out.push({
          title: (item.title ?? '').replace(/\s+/g, ' ').trim() || '—',
          page: await pageOfDest(doc, item.dest),
          depth,
          children,
        })
      }
      return out
    }

    return walk(raw, 0)
  })
}

async function extractPages(path: string, limit: number): Promise<PagesResult> {
  return withDocument(path, async (doc) => {
    const pageCount = doc.numPages
    const pages: PageInfo[] = []
    for (let pageNo = 1; pageNo <= Math.min(pageCount, limit); pageNo += 1) {
      const page = await doc.getPage(pageNo)
      try {
        const viewport = page.getViewport({ scale: 1 })
        pages.push({
          page: pageNo,
          width: Math.round(viewport.width * 100) / 100,
          height: Math.round(viewport.height * 100) / 100,
          rotate: page.rotate,
        })
      } finally {
        page.cleanup()
      }
    }
    return { pageCount, pages }
  })
}

export async function handleTask(task: TextTask): Promise<unknown> {
  switch (task.kind) {
    case 'text':
      return extractText(task.path, {
        pages: task.pages,
        maxPages: task.maxPages,
        maxChars: task.maxChars,
        perPageChars: task.perPageChars,
      })
    case 'outline':
      return extractOutline(task.path)
    case 'pages':
      return extractPages(task.path, task.limit)
  }
}

// ─── worker 入口 ────────────────────────────────────────────────────

if (!isMainThread && parentPort) {
  const port = parentPort
  const task = (workerData as { task?: TextTask } | undefined)?.task
  if (task) {
    handleTask(task).then(
      (result) => port.postMessage({ ok: true, result }),
      (error) => port.postMessage({ ok: false, error: messageOf(error) }),
    )
  }
}
