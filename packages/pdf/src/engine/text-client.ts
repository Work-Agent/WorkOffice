// ─── 文本抽取客户端：worker + 缓存 ──────────────────────────────────
//
// 每次请求起一个短命 worker（dist/text.js），完成即退出；超时 terminate。
// 文本按 ref + size:mtime 缓存，避免 read → find 连续调用重复解析。

import { Worker } from 'node:worker_threads'
import { messageOf, type PdfFile } from './ref'
import type { OutlineItem, PageInfo, PagesResult, TextResult, TextTask } from './text'

const TIMEOUT_MS = 20_000
const MAX_TEXT_PAGES = 500
const MAX_TEXT_CHARS = 200_000
const PER_PAGE_CHARS = 20_000

interface CacheEntry {
  key: string
  result: TextResult
}

const textCache = new Map<string, CacheEntry>()

function keyOf(file: PdfFile): string {
  return `${file.size}:${file.mtimeMs}`
}

function remember(ref: string, key: string, result: TextResult): void {
  textCache.delete(ref)
  textCache.set(ref, { key, result })
  while (textCache.size > 2) {
    const first = textCache.keys().next().value
    if (first === undefined) break
    textCache.delete(first)
  }
}

function runWorker<T>(task: TextTask): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    let worker: Worker
    try {
      worker = new Worker(new URL('./text.js', import.meta.url), {
        workerData: { task },
      })
    } catch (error) {
      reject(new Error(`command.failed: 文本 worker 启动失败：${messageOf(error)}`))
      return
    }

    let settled = false
    const finish = (run: () => void): void => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      void worker.terminate().catch(() => undefined)
      run()
    }
    const timer = setTimeout(() => {
      finish(() => reject(new Error('command.timeout: PDF 文本抽取超时')))
    }, TIMEOUT_MS)

    worker.once('message', (message: { ok?: boolean; result?: unknown; error?: unknown }) => {
      finish(() => {
        if (message?.ok === true) resolve(message.result as T)
        else reject(new Error(`command.failed: ${String(message?.error ?? 'PDF 文本抽取失败')}`))
      })
    })
    worker.once('error', (error) => {
      finish(() => reject(new Error(`command.failed: PDF 文本 worker 失败：${messageOf(error)}`)))
    })
    worker.once('exit', (code) => {
      if (code !== 0) {
        finish(() => reject(new Error(`command.failed: PDF 文本 worker 退出：${code}`)))
      }
    })
  })
}

/** 全文（grep / find 用）：按页数/字符预算截断。 */
export async function textOf(file: PdfFile): Promise<TextResult> {
  const key = keyOf(file)
  const cached = textCache.get(file.ref)
  if (cached?.key === key) return cached.result
  const result = await runWorker<TextResult>({
    kind: 'text',
    path: file.path,
    pages: 'all',
    maxPages: MAX_TEXT_PAGES,
    maxChars: MAX_TEXT_CHARS,
    perPageChars: PER_PAGE_CHARS,
  })
  remember(file.ref, key, result)
  return result
}

/** 指定页文本（pdf.read text 用），不写缓存。 */
export function pagesText(file: PdfFile, pages: number[], chars = PER_PAGE_CHARS): Promise<TextResult> {
  return runWorker<TextResult>({
    kind: 'text',
    path: file.path,
    pages,
    maxPages: Math.max(1, pages.length),
    maxChars: chars,
    perPageChars: chars,
  })
}

export function outlineOf(file: PdfFile): Promise<OutlineItem[]> {
  return runWorker<OutlineItem[]>({ kind: 'outline', path: file.path })
}

export function pageInfoOf(file: PdfFile, limit: number): Promise<PagesResult> {
  return runWorker<PagesResult>({ kind: 'pages', path: file.path, limit })
}

export function invalidateText(ref?: string): void {
  if (ref) textCache.delete(ref)
  else textCache.clear()
}

export type { OutlineItem, PageInfo, PagesResult, TextResult }
