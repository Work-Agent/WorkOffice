// ─── L0 注册面：codec / sidecar / store / presenter ────────────────────
//
// 契约 §4.2 / §4.5 / §7.1：
// - `decode` 在主进程执行，必须便宜：只解 `word/document.xml` 并产出投影。
// - `encode` 必须拒绝：计划型插件的字节只由 exec 产生（command.forbidden）。
// - `presenter.data` 只放元数据：内容走 `bytes` 分块通道（content: 'pulled'）。
// - 不注册 list：打开入口复用宿主工作区列表。

import { DOCUMENT_PART, DocxParseError, parseDocxBytes, tierFor } from './shared/document'
import { projectModel } from './shared/project'
import { blockHash } from './shared/sha256'
import { tableSizeOf } from './shared/xml-ops'
import type { ModelBlock } from './shared/types'

const MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const CODEC_ID = 'docx-pro'
const KEY = { transport: '*', mime: MIME, codec: CODEC_ID }

interface Sdk {
  registerCodec(codec: unknown): void
  registerSidecar(provider: unknown): void
  registerStore(store: unknown): void
  registerPresenter(presenter: unknown): void
}

interface RawContent {
  bytes: Uint8Array
  version: string
  mime: string
}

interface DecodedDoc {
  doc: unknown
  anchorMap: string
  outline?: Array<{ level: number; title: string }>
}

const sdk = (globalThis as { __workstudioSdk?: Sdk }).__workstudioSdk

interface DocxDoc {
  name: string
  tier: ReturnType<typeof tierFor>
  weight: number
  blockCount: number
  tableCount: number
  projection: string
  blocks: Array<{
    index: number
    type: string
    heading: number
    text: string
    hash: string
  }>
}

const MAX_DECODE_BLOCKS = 4_000
const MAX_ANCHOR_TEXT = 1_000_000

function nameOf(ref: string): string {
  try {
    const url = new URL(ref)
    const name = url.pathname.slice(url.pathname.lastIndexOf('/') + 1)
    return decodeURIComponent(name) || 'document.docx'
  } catch {
    const name = ref.slice(ref.lastIndexOf('/') + 1)
    return name || 'document.docx'
  }
}

function decodeDoc(raw: RawContent, ref: string): DecodedDoc {
  const parsed = parseDocxBytes(raw.bytes)
  const name = nameOf(ref)
  const tier = tierFor(parsed.weight)
  const visible = parsed.blocks.filter((block) => !block.hidden)
  const doc: DocxDoc = {
    name,
    tier,
    weight: Math.round(parsed.weight),
    blockCount: visible.length,
    tableCount: visible.filter((block) => block.kind === 'tbl').length,
    projection: projectModel(parsed.blocks, { name, version: raw.version, tier }),
    blocks: visible.slice(0, MAX_DECODE_BLOCKS).map((block) => ({
      index: block.docxIndex!,
      type: block.kind,
      heading: block.heading,
      text: block.text,
      hash: blockHash(block.xml),
    })),
  }
  const text = plainText(parsed.blocks)
  return {
    doc,
    anchorMap: text,
    outline: parsed.blocks
      .filter((block) => block.heading > 0)
      .slice(0, 500)
      .map((block) => ({
        level: block.heading,
        title: block.text.replace(/\s+/g, ' ').trim().slice(0, 200),
      })),
  }
}

function plainText(blocks: readonly ModelBlock[]): string {
  const lines: string[] = []
  let length = 0
  for (const block of blocks) {
    if (block.hidden) continue
    const line =
      block.kind === 'tbl'
        ? `[表格 ${tableSizeOf(block.xml).rows}x${tableSizeOf(block.xml).cols}]`
        : block.text
    if (!line) continue
    lines.push(line)
    length += line.length + 1
    if (length > MAX_ANCHOR_TEXT) {
      lines.push('…（文本投影已截断）')
      break
    }
  }
  return lines.join('\n')
}

sdk?.registerCodec({
  id: CODEC_ID,
  keys: [{ transport: '*', mime: MIME }],
  extensions: ['docx', 'docm', 'dotx'],
  patchType: 'bytes',
  decode: (raw: RawContent, ref: string) => decodeDoc(raw, ref),
  locate: (_decoded: DecodedDoc, anchor: { start: number; end: number }) => anchor,
  // 字节只由 exec 产生（契约 I8 / I9）；整份写通道对本插件关闭。
  encode: () => {
    throw new Error('command.forbidden')
  },
})

sdk?.registerSidecar({
  key: KEY,
  sidecar: (decoded: DecodedDoc) => ({
    type: 'text',
    content: (decoded.doc as DocxDoc).projection,
  }),
})

sdk?.registerStore({
  key: KEY,
  schema: {
    blocks: {
      index: 'number',
      type: 'string',
      heading: 'number',
      text: 'string',
      hash: 'string',
    },
    doc: { name: 'string', tier: 'string', weight: 'number' },
  },
  query: (decoded: DecodedDoc, query: QueryInput) => {
    const doc = decoded.doc as DocxDoc
    const where = query.where ?? {}
    let rows = doc.blocks
    if (typeof where['type'] === 'string')
      rows = rows.filter((row) => row.type === where['type'])
    if (typeof where['heading'] === 'number')
      rows = rows.filter((row) => row.heading === where['heading'])
    if (typeof where['text'] === 'string')
      rows = rows.filter((row) => row.text.includes(String(where['text'])))
    const limit = Math.max(1, Math.min(2_000, Number(query.limit ?? 200)))
    return {
      total: rows.length,
      rows: query.select?.length
        ? rows.map((row) => pick(row, query.select!))
        : rows.slice(0, limit),
    }
  },
})

sdk?.registerPresenter({
  key: KEY,
  type: 'page',
  surface: 'docx-pro',
  requires: ['write'],
  present: (decoded: DecodedDoc, raw: RawContent, ref: string) => {
    const doc = decoded.doc as DocxDoc
    return {
      type: 'page',
      surface: 'docx-pro',
      page: '',
      // 只放元数据：字节经 `bytes` 分块通道取（content: 'pulled'）。
      data: {
        kind: 'docx',
        name: doc.name,
        ref,
        size: raw.bytes.byteLength,
        version: raw.version,
        tier: doc.tier,
        blocks: doc.blockCount,
        tables: doc.tableCount,
      },
    }
  },
})

interface QueryInput {
  select?: string[]
  where?: Record<string, unknown>
  limit?: number
}

function pick<T extends Record<string, unknown>>(row: T, keys: string[]): Partial<T> {
  const out: Partial<T> = {}
  for (const key of keys) if (key in row) out[key as keyof T] = row[key as keyof T]
  return out
}

export { DocxParseError, DOCUMENT_PART }
