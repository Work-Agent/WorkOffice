// ─── pdf engine：WorkStudio SDK 入口 ──────────────────────────────────
//
// 只读文档：页面只 load / chunk / saveCopy；助手读走 resources.pdf + pdf.read。
// 没有页面租约、没有页面队列：PDF 的字节永远来自磁盘，页面不是权威模型。
//
// 该文件在宿主主进程求值：允许 Node 内置模块；不得 import viewer / React。

import { writeFile } from 'node:fs/promises'
import {
  definePlugin,
  type PluginCommand,
  type PluginContext,
  type PluginScope,
} from '@work-studio/sdk'
import type { PdfDoc, PdfPageMeta } from '../shared/pdf-types'
import { bytesCache } from './cache'
import { loadDoc } from './load'
import { summaryOf } from './meta'
import { pdfResource } from './resource'
import {
  CHUNK_BYTES,
  INLINE_BYTES,
  MAX_CHUNK_BYTES,
  MAX_PDF_BYTES,
  MAX_SAVE_BYTES,
  assertPdfFile,
  clampInt,
  fail,
  messageOf,
  refOf,
  sha256Hex,
  textOf as strOf,
  toBytes,
} from './ref'
import {
  invalidateText,
  outlineOf,
  pageInfoOf,
  pagesText,
  textOf,
  type OutlineItem,
  type PagesResult,
} from './text-client'

// ─── 状态 ───────────────────────────────────────────────────────────

let scope: PluginScope | null = null

function pageMetaOf(ref: string, name: string, doc: PdfDoc): PdfPageMeta {
  return {
    kind: 'pdf',
    name,
    ref,
    size: doc.size,
    version: doc.version,
    pageCount: doc.pageCount ?? null,
    encrypted: doc.encrypted ?? null,
    title: doc.title ?? null,
    author: doc.author ?? null,
    ...(doc.headerVersion ? { headerVersion: doc.headerVersion } : {}),
    ...(doc.metadataWarning ? { metadataWarning: doc.metadataWarning } : {}),
  }
}

function snippet(text: string, at: number, length: number): string {
  return text
    .slice(Math.max(0, at - 60), Math.min(text.length, at + length + 120))
    .replace(/\s+/g, ' ')
    .trim()
}

function formatOutline(items: OutlineItem[], depth = 0): string {
  if (items.length === 0) return '(no outline)'
  const lines: string[] = []
  for (const item of items) {
    lines.push(`${'  '.repeat(depth)}${item.title}${item.page ? `  → p.${item.page}` : ''}`)
    if (item.children.length > 0) lines.push(formatOutline(item.children, depth + 1))
  }
  return lines.join('\n')
}

function formatPages(result: PagesResult): string {
  const lines = [`pages=${result.pageCount}`]
  for (const page of result.pages) {
    lines.push(`p.${page.page}: ${page.width}×${page.height} pt rot=${page.rotate}`)
  }
  if (result.pages.length < result.pageCount) {
    lines.push(`(showing first ${result.pages.length}; use limit to read more)`)
  }
  return lines.join('\n')
}

async function safely<T>(run: () => Promise<T>): Promise<T> {
  try {
    return await run()
  } catch (error) {
    const message = messageOf(error)
    if (/password|encrypt|PasswordException/i.test(message)) {
      fail('command.page-unavailable', 'PDF 受密码保护，文本抽取需要页面解锁')
    }
    fail('command.failed', message)
  }
}

// ─── 助手读 ─────────────────────────────────────────────────────────

async function readPdf(input: Record<string, unknown>, ctx: PluginContext): Promise<string> {
  const ref = refOf(input)
  const what = strOf(input['what']) || 'context'
  const file = await assertPdfFile(ref)

  switch (what) {
    case 'context': {
      const bytes = await ctx.fs.read(ref)
      const hash = sha256Hex(bytes)
      const doc = await loadDoc(ref, file.name, bytes, hash)
      return `${summaryOf(doc)}\nhash=${hash}`
    }

    case 'outline': {
      const items = await safely(() => outlineOf(file))
      return formatOutline(items)
    }

    case 'pages': {
      const limit = clampInt(input['limit'], 100, 1, 2000)
      const result = await safely(() => pageInfoOf(file, limit))
      return formatPages(result)
    }

    case 'text': {
      const page = clampInt(input['page'], 0, 0, 1_000_000)
      if (page < 1) fail('command.invalid', 'text 需要 page（1-based）')
      const chars = clampInt(input['limit'], 20_000, 1, 200_000)
      const result = await safely(() => pagesText(file, [page], chars))
      const found = result.pages[0]
      if (!found) return `(page ${page} not found; pages=${result.pageCount})`
      return `# ${file.name} · page ${page}/${result.pageCount}\n${found.text || '(no extractable text)'}`
    }

    case 'find': {
      const query = strOf(input['query']).trim()
      if (!query) fail('command.invalid', 'find 需要 query')
      const limit = clampInt(input['limit'], 50, 1, 200)
      const result = await safely(() => textOf(file))
      const lower = query.toLowerCase()
      const hits: string[] = []
      for (const page of result.pages) {
        let from = 0
        while (hits.length < limit) {
          const at = page.text.toLowerCase().indexOf(lower, from)
          if (at < 0) break
          hits.push(`p.${page.page}: ${snippet(page.text, at, query.length)}`)
          from = at + Math.max(1, query.length)
        }
        if (hits.length >= limit) break
      }
      if (hits.length === 0) {
        const suffix = result.truncated ? ' (scan truncated)' : ''
        return `(no matches for ${query}; pages scanned=${result.pages.length}/${result.pageCount}${suffix})`
      }
      return hits.join('\n')
    }

    default:
      fail('command.invalid', `未知 what：${what}`)
  }
}

// ─── 页面命令 ───────────────────────────────────────────────────────

const commands: Record<string, PluginCommand> = {
  'pdf.read': {
    summary:
      '读 PDF：context（元数据）/ outline（书签）/ pages（页尺寸）/ text（单页文本）/ find（全文查找）。只读，不写盘。',
    args: {
      ref: { type: 'string', required: true },
      what: { enum: ['context', 'outline', 'pages', 'text', 'find'] },
      page: { type: 'number' },
      query: { type: 'string' },
      limit: { type: 'number' },
    },
    expose: { agent: true },
    run: (input, ctx) => readPdf(input as Record<string, unknown>, ctx),
  },

  'pdf.load': {
    summary: '页面加载：元数据 + hash；≤16 MB 直接带字节',
    args: {
      ref: { type: 'string', required: true },
      base64: { type: 'boolean' },
    },
    expose: { agent: false },
    run: async (input, ctx) => {
      const ref = refOf(input)
      const file = await assertPdfFile(ref)
      const bytes = await ctx.fs.read(ref)
      if (bytes.byteLength === 0) fail('command.invalid', 'PDF 文件为空')
      const hash = sha256Hex(bytes)
      const doc = await loadDoc(ref, file.name, bytes, hash)
      const inline = bytes.byteLength <= INLINE_BYTES
      if (!inline) bytesCache.put(ref, hash, bytes)
      const meta = pageMetaOf(ref, file.name, doc)
      const base = { ref, name: file.name, size: bytes.byteLength, hash, meta, inline }
      if (input['base64'] === true) {
        return { ...base, data: Buffer.from(bytes).toString('base64') }
      }
      return inline ? { ...base, bytes } : base
    },
  },

  'pdf.chunk': {
    summary: '页面分块取字节（大文件）',
    args: {
      ref: { type: 'string', required: true },
      start: { type: 'number', required: true },
      length: { type: 'number', required: true },
      hash: { type: 'string' },
      base64: { type: 'boolean' },
    },
    expose: { agent: false },
    run: async (input, ctx) => {
      const ref = refOf(input)
      const start = clampInt(input['start'], 0, 0, MAX_PDF_BYTES)
      const length = clampInt(input['length'], CHUNK_BYTES, 1, MAX_CHUNK_BYTES)
      const requestedHash = strOf(input['hash'])

      let entry = bytesCache.get(ref)
      if (!entry || (requestedHash && entry.hash !== requestedHash)) {
        await assertPdfFile(ref)
        const bytes = await ctx.fs.read(ref)
        if (bytes.byteLength > MAX_PDF_BYTES) {
          fail('command.too-large', `PDF 超过 ${MAX_PDF_BYTES / 1024 / 1024} MB 上限`)
        }
        const hash = sha256Hex(bytes)
        if (requestedHash && hash !== requestedHash) {
          fail('command.stale', '文件已被外部修改，请重新打开')
        }
        bytesCache.put(ref, hash, bytes)
        entry = bytesCache.get(ref)
      }
      if (!entry) fail('command.failed', '字节缓存不可用')

      const size = entry.bytes.byteLength
      const from = Math.min(start, size)
      const to = Math.min(from + length, size)
      const slice = entry.bytes.subarray(from, to)
      const base = { ref, start: from, end: to, size, hash: entry.hash }
      if (input['base64'] === true) {
        return { ...base, data: Buffer.from(slice).toString('base64') }
      }
      return { ...base, bytes: slice }
    },
  },

  'pdf.saveCopy': {
    summary: '另存副本（宿主 save token + 页面当前字节）',
    args: {
      ref: { type: 'string', required: true },
      token: { type: 'string', required: true },
    },
    expose: { agent: false },
    run: async (input) => {
      refOf(input)
      if (!scope) fail('command.page-unavailable', '插件未激活')
      const token = strOf(input['token'])
      if (!token) fail('command.invalid', '缺少保存 token')
      const bytes = toBytes(input as Record<string, unknown>)
      if (bytes.byteLength === 0) fail('command.invalid', '副本没有内容')
      if (bytes.byteLength > MAX_SAVE_BYTES) {
        fail('command.too-large', `副本超过 ${MAX_SAVE_BYTES / 1024 / 1024} MB`)
      }
      let target: string
      try {
        target = scope.files.resolveSave(token)
      } catch (error) {
        fail('command.invalid', `保存目标无效：${messageOf(error)}`)
      }
      await writeFile(target, bytes)
      return { ok: true, size: bytes.byteLength }
    },
  },
}

// ─── 贡献 ───────────────────────────────────────────────────────────

export const workstudio = definePlugin({
  resources: { pdf: pdfResource },
  commands,
  // 只读插件没有页面队列，不开 events 通道。
  lifecycle: {
    activate(value: PluginScope): void {
      scope = value
      value.log.info('pdf: read-only, 4 commands')
    },
    deactivate(): void {
      scope = null
      bytesCache.clear()
      invalidateText()
    },
  },
})

export default workstudio
