// ─── resources.pdf：单个 PDF 文件的资源适配器 ────────────────────────
//
// 只回答“这一个文件是什么”。ls / find 恒为空：没有目录、没有集合。
// read 只返回元数据（页面打开时会订阅 workspace/<ref>，snapshot 就是 read，
// 不能在这里抽全文）；文本检索走 grep 或 pdf.read。

import type { Item, ResourceAdapter } from '@work-studio/sdk'
import { formatSize } from '../shared/format'
import { loadDoc } from './load'
import { summaryOf } from './meta'
import {
  INLINE_BYTES,
  assertPdfFile,
  readBytes,
  readHead,
  sha256Hex,
  type PdfFile,
} from './ref'
import { textOf } from './text-client'

const MAX_GREP_HITS = 50
const SNIPPET_BEFORE = 60
const SNIPPET_AFTER = 120

function headerVersionOf(head: Uint8Array): string | undefined {
  const text = String.fromCharCode(...head.subarray(0, Math.min(head.length, 1024)))
  return /%PDF-(\d+\.\d+)/.exec(text)?.[1]
}

async function docItem(ref: string, file: PdfFile): Promise<Item> {
  if (file.size > INLINE_BYTES) {
    const head = await readHead(file.path, 1024)
    const headerVersion = headerVersionOf(head)
    return {
      locator: 'pdf:doc',
      kind: 'pdf',
      ref,
      name: file.name,
      size: file.size,
      hash: null,
      version: `${file.size}:${file.mtimeMs}`,
      pageCount: null,
      encrypted: null,
      title: null,
      author: null,
      ...(headerVersion ? { headerVersion } : {}),
      metadataWarning: 'large-file-metadata-skipped',
      text: [
        `# ${file.name}`,
        `pages=? size=${formatSize(file.size)}`,
        ...(headerVersion ? [`pdf=${headerVersion}`] : []),
        'metadata: large-file-metadata-skipped',
      ].join('\n'),
    }
  }

  const bytes = await readBytes(file.path)
  const hash = sha256Hex(bytes)
  const doc = await loadDoc(ref, file.name, bytes, hash)
  return {
    locator: 'pdf:doc',
    kind: 'pdf',
    ref,
    name: file.name,
    size: file.size,
    hash,
    version: hash,
    pageCount: doc.pageCount ?? null,
    encrypted: doc.encrypted === true,
    title: doc.title ?? null,
    author: doc.author ?? null,
    ...(doc.headerVersion ? { headerVersion: doc.headerVersion } : {}),
    ...(doc.metadataWarning ? { metadataWarning: doc.metadataWarning } : {}),
    text: summaryOf(doc),
  }
}

function snippetAround(text: string, at: number, length: number): string {
  return text
    .slice(Math.max(0, at - SNIPPET_BEFORE), Math.min(text.length, at + length + SNIPPET_AFTER))
    .replace(/\s+/g, ' ')
    .trim()
}

async function grepPages(ref: string, query: string): Promise<Item[]> {
  const needle = query.trim()
  if (!needle) return []
  const file = await assertPdfFile(ref)

  // 加密 / 损坏 / 扫描件都走同一条降级路径：抽不出文本就回空。
  let text
  try {
    text = await textOf(file)
  } catch {
    return []
  }

  const lower = needle.toLowerCase()
  const hits: Item[] = []
  for (const page of text.pages) {
    const at = page.text.toLowerCase().indexOf(lower)
    if (at < 0) continue
    hits.push({
      ref,
      name: file.name,
      locator: `pdf:page:${page.page}`,
      kind: 'page',
      page: page.page,
      title: `${file.name} · 第 ${page.page} 页`,
      snippet: snippetAround(page.text, at, needle.length),
    })
    if (hits.length >= MAX_GREP_HITS) break
  }
  return hits
}

export const pdfResource: ResourceAdapter = {
  ls: async () => [],
  find: async () => [],
  read: async (ref) => [await docItem(ref, await assertPdfFile(ref))],
  grep: (ref, query) => grepPages(ref, query),
  open: async (ref, locator) => ({ ref, ...(locator ? { locator } : {}) }),
  describe: () => ({
    root: 'file://',
    label: 'PDF 文档',
    summary:
      '单个 .pdf 文件：read 取元数据，grep 查页文本，pdf.read 读大纲/单页/全文，open 交给查看器。没有目录能力。',
    examples: [
      'read file:///w/报告.pdf',
      'grep file:///w/报告.pdf 收入',
      'command pdf/read {ref, what:"text", page:3}',
    ],
  }),
}
