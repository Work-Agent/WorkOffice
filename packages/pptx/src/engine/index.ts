// ─── pptx engine：WorkStudio SDK 入口 ──────────────────────────────────
//
// 只做演示文稿页：给定 ref 读 / 写当前文件；不做文件列表、标签页。
// Agent 的写一律经页面队列；没有页面租约时拒绝（command.page-unavailable），
// 唯一例外是 create —— 新建时还没有页面。
//
// 该文件在宿主主进程求值：允许用 Node 内置模块；不得 import 编辑器模块。

import { createHash, randomUUID } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { pathToFileURL } from 'node:url'
import {
  definePlugin,
  type PluginCommand,
  type PluginContext,
  type PluginScope,
} from '@work-studio/sdk'
import { openPptx, savePptx, type OpenedPptx } from '@genoffice/pptx-engine'
import { runTxn } from '@genoffice/pptx-ops'
import { messageOf } from './abi'
import { buildCreateOps } from './create'
import { baseDirOfRef } from './images'
import { buildDesignDeck, formatPageIssues, parseDesignDeck } from './build'
import { deckSummary, textOfElement } from '../shared/deck'
import { normalizeOps } from '../shared/ops'
import { runCheck, type CheckParams } from '../shared/check'
import { readMany } from '../shared/read'
import { MAX_FILE_BYTES } from '../shared/runtime'
import { DeckSpecError, normalizeSpec } from '../shared/spec'
import { allOps, guideText, isKnownOp } from '../tools/catalog'

// ─── 常量 ───────────────────────────────────────────────────────────

/** 页面一次 load / save 的上限；超过就拒绝，不分块。 */
const MAX_PAGE_BYTES = MAX_FILE_BYTES
const MAX_COMMANDS = 200
const LEASE_TTL_MS = 5 * 60_000
const REQUEST_TIMEOUT_MS = 30_000
const EXTS = /\.(pptx|pptm|potx)$/i

// ─── 类型 ───────────────────────────────────────────────────────────

export interface Receipt {
  ref: string
  hash: string
  size: number
  at: number
  applied?: number
  warnings?: string[]
}

type PageRequest =
  | {
      type: 'apply'
      id: string
      commands: unknown[]
      baseHash?: string
      docRev?: number
      dryRun?: boolean
    }
  | { type: 'read'; id: string; what: string; params: Record<string, unknown> }

interface Lease {
  pageId: string
  hash: string
  docRev: number
  dirty: boolean
  updatedAt: number
}

interface Pending {
  ref: string
  request: PageRequest
  resolve: (value: unknown) => void
  reject: (error: Error) => void
  timer: ReturnType<typeof setTimeout>
}

interface ResourceItem {
  locator: string
  [key: string]: unknown
}

// ─── 状态 ───────────────────────────────────────────────────────────

let scope: PluginScope | null = null
const leases = new Map<string, Lease>()
const pending = new Map<string, Pending>()
const locks = new Map<string, Promise<unknown>>()
let SEED: Uint8Array | null = null

function fail(code: string, message: string): never {
  throw new Error(`${code}: ${message}`)
}

// ─── 小工具 ─────────────────────────────────────────────────────────

function textOf(value: unknown): string {
  return typeof value === 'string' ? value : ''
}

function refOf(input: { ref?: unknown }): string {
  const ref = textOf(input.ref)
  if (!ref) fail('command.invalid', '缺少 ref')
  const path = ref.split(/[?#]/, 1)[0] ?? ref
  if (!EXTS.test(path)) fail('resource.not-found', `不是演示文稿：${ref}`)
  return ref
}

function nameOf(ref: string): string {
  const path = ref.split(/[?#]/, 1)[0] ?? ref
  const base = path.slice(path.lastIndexOf('/') + 1)
  try {
    return decodeURIComponent(base) || 'presentation.pptx'
  } catch {
    return base || 'presentation.pptx'
  }
}

function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

function toBytes(input: Record<string, unknown>): Uint8Array {
  const raw = input['bytes'] ?? input['data']
  if (raw instanceof Uint8Array) return raw
  if (raw instanceof ArrayBuffer) return new Uint8Array(raw)
  if (ArrayBuffer.isView(raw)) return new Uint8Array(raw.buffer, raw.byteOffset, raw.byteLength)
  if (typeof raw === 'string' && raw.length > 0) {
    // base64 兜底：宿主 IPC 理论上支持 typed array，留一条退路。
    return new Uint8Array(Buffer.from(raw, 'base64'))
  }
  return new Uint8Array()
}

function guardSize(bytes: Uint8Array, what: string): void {
  if (bytes.byteLength === 0) fail('command.invalid', `${what}为空`)
  if (bytes.byteLength > MAX_PAGE_BYTES)
    fail('command.too-large', `${what}超过上限：${bytes.byteLength} > ${MAX_PAGE_BYTES}`)
}

/** OOXML 包的最小校验：ZIP 魔数 + 必备 part。 */
async function guardPackage(bytes: Uint8Array): Promise<OpenedPptx> {
  if (bytes.byteLength < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b)
    fail('command.invalid', '不是有效的 PPTX 包（缺 ZIP 魔数）')
  let opened: OpenedPptx
  try {
    opened = await openPptx(bytes)
  } catch (error) {
    fail('command.invalid', `PPTX 解析失败：${messageOf(error)}`)
  }
  if (!opened.archive.entries.has('[Content_Types].xml'))
    fail('command.invalid', 'PPTX 包缺少 [Content_Types].xml')
  return opened
}

async function readRef(ctx: PluginContext, ref: string): Promise<Uint8Array> {
  const bytes = await ctx.fs.read(ref)
  const copy = new Uint8Array(bytes)
  guardSize(copy, '文件')
  return copy
}

function seedBytes(): Uint8Array {
  SEED ??= new Uint8Array(readFileSync(new URL('./seed.pptx', import.meta.url)))
  return SEED
}

/** 同一 ref 的写串行：读盘 → CAS → 写盘 之间不允许交错。 */
function withLock<T>(ref: string, run: () => Promise<T>): Promise<T> {
  const previous = locks.get(ref) ?? Promise.resolve()
  const next = previous.then(run, run)
  locks.set(
    ref,
    next.catch(() => undefined),
  )
  return next
}

// ─── 租约 ───────────────────────────────────────────────────────────

function liveLease(ref: string): Lease | null {
  const lease = leases.get(ref)
  if (!lease) return null
  if (Date.now() - lease.updatedAt > LEASE_TTL_MS) {
    leases.delete(ref)
    rejectPendingFor(ref, new Error('command.page-unavailable: 页面租约已过期'))
    return null
  }
  return lease
}

function requireLease(ref: string): Lease {
  const lease = liveLease(ref)
  if (!lease)
    fail(
      'command.page-unavailable',
      '演示文稿页面未打开；请先打开该文件再写入（编辑只在页面里发生）',
    )
  return lease
}

function rejectPendingFor(ref: string, error: Error): void {
  for (const [id, entry] of pending) {
    if (entry.ref !== ref) continue
    clearTimeout(entry.timer)
    entry.reject(error)
    pending.delete(id)
  }
}

function expireLeases(): void {
  for (const ref of [...leases.keys()]) liveLease(ref)
}

// ─── 页面队列 ───────────────────────────────────────────────────────

function requestPage<T>(ref: string, request: PageRequest): Promise<T> {
  if (!scope) fail('command.page-unavailable', '插件未激活')
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      pending.delete(request.id)
      reject(new Error('command.timeout: 页面 30 秒未回执'))
    }, REQUEST_TIMEOUT_MS)
    pending.set(request.id, {
      ref,
      request,
      resolve: resolve as (value: unknown) => void,
      reject,
      timer,
    })
    scope?.events.emit('request', { id: request.id, type: request.type })
  })
}

function ackPayload(raw: unknown): Record<string, unknown> {
  if (raw && typeof raw === 'object') return raw as Record<string, unknown>
  if (typeof raw === 'string' && raw.trim().startsWith('{')) {
    try {
      return JSON.parse(raw) as Record<string, unknown>
    } catch {
      return {}
    }
  }
  return {}
}

// ─── 只读投影 ───────────────────────────────────────────────────────

async function openFromDisk(ctx: PluginContext, ref: string): Promise<{ opened: OpenedPptx; hash: string; size: number }> {
  const bytes = await readRef(ctx, ref)
  const opened = await guardPackage(bytes)
  return { opened, hash: sha256Hex(bytes), size: bytes.byteLength }
}

/** Agent 只读：有页面读 live（含未保存修改），否则读磁盘。 */
async function readText(
  ctx: PluginContext,
  ref: string,
  what: string,
  params: Record<string, unknown>,
): Promise<string> {
  if (liveLease(ref)) {
    return requestPage<string>(ref, { type: 'read', id: randomUUID(), what, params })
  }
  const { opened, hash } = await openFromDisk(ctx, ref)
  return readMany(opened, { ...params, name: nameOf(ref), version: hash })
}

// ─── 资源适配器 ─────────────────────────────────────────────────────

function slideItems(ref: string, opened: OpenedPptx, limit: number): ResourceItem[] {
  const summary = deckSummary(opened)
  const items: ResourceItem[] = [
    {
      locator: `${ref}#deck`,
      kind: 'deck',
      title: nameOf(ref),
      slideCount: summary.slideCount,
      size: summary.size,
      text: summary.slides
        .slice(0, limit)
        .map((slide) => `${slide.index + 1}. ${slide.title || '(无标题)'}`)
        .join('\n'),
    },
  ]
  for (const slide of summary.slides.slice(0, limit)) {
    const model = opened.deck.slides[slide.index]
    const text = (model?.elements ?? [])
      .map((element) => textOfElement(element))
      .filter(Boolean)
      .join(' / ')
      .slice(0, 400)
    items.push({
      locator: `${ref}#s:${slide.index}`,
      kind: 'slide',
      index: slide.index,
      title: slide.title,
      elements: slide.elementCount,
      text,
    })
  }
  return items
}

const resources = {
  pptx: {
    // 不认领目录：结构一律为空，内容按单文件给。
    ls: async () => [],
    find: async () => [],
    read: async (ref: string, options?: { limit?: number }) => {
      const ctx = currentCtx
      if (!ctx) fail('command.page-unavailable', '插件未激活')
      const { opened } = await openFromDisk(ctx, refOf({ ref }))
      return slideItems(ref, opened, Math.max(1, Math.min(options?.limit ?? 40, 200)))
    },
    grep: async (ref: string, query: string) => {
      const ctx = currentCtx
      if (!ctx) fail('command.page-unavailable', '插件未激活')
      const needle = query.trim().toLowerCase()
      if (!needle) return []
      const { opened } = await openFromDisk(ctx, refOf({ ref }))
      const summary = deckSummary(opened)
      const hits: ResourceItem[] = []
      for (const slide of summary.slides) {
        const model = opened.deck.slides[slide.index]
        const text = (model?.elements ?? []).map((element) => textOfElement(element)).join(' / ')
        const at = text.toLowerCase().indexOf(needle)
        if (at < 0) continue
        hits.push({
          locator: `${ref}#s:${slide.index}`,
          kind: 'slide',
          index: slide.index,
          title: slide.title,
          snippet: text.slice(Math.max(0, at - 60), at + needle.length + 120),
        })
        if (hits.length >= 50) break
      }
      return hits
    },
    open: async (ref: string, locator?: string) => ({
      ref,
      ...(locator ? { locator } : {}),
    }),
    describe: () => ({
      root: 'file://',
      label: '演示文稿',
      summary: '单个 .pptx/.pptm/.potx 文件：read 取页与元素摘要，grep 查文本，open 交给页面。没有目录能力。',
      examples: ['read file:///w/a.pptx', 'grep file:///w/a.pptx 收入'],
    }),
  },
}

/** 资源适配器拿不到命令上下文，这里保留一份最近一次的引用。 */
let currentCtx: PluginContext | null = null

// ─── 命令 ───────────────────────────────────────────────────────────

const PAGES_LIMIT = 40

const commands: Record<string, PluginCommand> = {
  'pptx.guide': {
    summary: 'PPT 命令目录：无参分组，group 取组内 schema，op 取单条',
    args: {
      ref: { type: 'string', required: true },
      group: { enum: ['text', 'element', 'insert', 'table', 'slide', 'deck', 'chart'] },
      op: { type: 'string' },
    },
    expose: { agent: true },
    run: async (input: Record<string, unknown>) => {
      refOf(input)
      return guideText({
        ...(typeof input['group'] === 'string' ? { group: input['group'] } : {}),
        ...(typeof input['op'] === 'string' ? { op: input['op'] } : {}),
      })
    },
  },

  'pptx.read': {
    summary: '读演示文稿：概览 / 单页 / 备注 / 批注 / 主题 / 审计 / 查找',
    args: {
      ref: { type: 'string', required: true },
      what: {
        enum: [
          'context',
          'outline',
          'slides',
          'slide',
          'notes',
          'comments',
          'charts',
          'tables',
          'theme',
          'audit',
          'find',
        ],
      },
      slide: { type: 'number' },
      query: { type: 'string' },
      limit: { type: 'number' },
    },
    expose: { agent: true },
    run: async (input: Record<string, unknown>, ctx: PluginContext) => {
      const ref = refOf(input)
      currentCtx = ctx
      return readText(ctx, ref, textOf(input['what']) || 'context', {
        ...(input['slide'] !== undefined ? { slide: input['slide'] } : {}),
        ...(typeof input['query'] === 'string' ? { query: input['query'] } : {}),
        ...(input['limit'] !== undefined ? { limit: input['limit'] } : {}),
      })
    },
  },

  'pptx.check': {
    summary: '写前检查：outline / pages 断言或对现有文件跑几何审计',
    args: {
      ref: { type: 'string', required: true },
      outline: { type: 'string' },
      pages: { type: 'string' },
      style: { type: 'string' },
      styles: { type: 'boolean' },
      audit: { type: 'boolean' },
      skeleton: { type: 'boolean' },
      slide: { type: 'number' },
    },
    expose: { agent: true },
    run: async (input: Record<string, unknown>, ctx: PluginContext) => {
      const ref = refOf(input)
      currentCtx = ctx
      const params = checkParamsOf(input)
      if (liveLease(ref)) {
        return requestPage<string>(ref, {
          type: 'read',
          id: randomUUID(),
          what: 'check',
          params: params as Record<string, unknown>,
        })
      }
      const { opened } = await openFromDisk(ctx, ref)
      return runCheck(params, async () => opened)
    },
  },

  'pptx.create': {
    summary: '新建演示文稿（目标必须不存在）：设计稿 pages 或文字稿 slides',
    args: {
      ref: { type: 'string', required: true },
      pages: { type: 'string' },
      slides: { type: 'string' },
      style: { type: 'string' },
      content: { type: 'string' },
      theme: { type: 'string' },
      size: { type: 'string' },
    },
    expose: { agent: true },
    run: async (input: Record<string, unknown>, ctx: PluginContext) => {
      const ref = refOf(input)
      currentCtx = ctx
      const existing = await ctx.fs.read(ref).catch(() => null)
      if (existing && existing.byteLength > 0)
        fail('resource.exists', '目标已存在；修改已有演示文稿请用 pptx.apply')

      const params = jsonInput(input, ['pages', 'slides', 'theme', 'content', 'size'])
      const style = parseMaybeJson(input['style'])
      const notes: string[] = []
      let bytes: Uint8Array

      if (params['pages'] !== undefined) {
        const deck = parseDesignDeck(params['pages'], style === undefined ? undefined : style)
        notes.push(...formatPageIssues(deck.issues))
        const baseDir = baseDirOfRef(ref)
        const built = await buildDesignDeck(deck, { ...(baseDir ? { baseDir } : {}) })
        bytes = built.bytes
        notes.push(...built.imageFailures.map((f) => `page ${f.page}: 图片取不到，已跳过 — ${f.url}`))
        notes.push(...built.imageWarnings.map((w) => `page ${w.page}: ${w.message}`))
      } else {
        let spec
        try {
          spec = normalizeSpec({ ...params, ...(style === undefined ? {} : { style }) })
        } catch (error) {
          if (error instanceof DeckSpecError) fail('command.invalid', error.message)
          throw error
        }
        const opened = await openPptx(seedBytes())
        const { setup, content } = buildCreateOps(opened, spec)
        for (const [label, ops] of [
          ['setup', setup],
          ['content', content],
        ] as const) {
          if (ops.length === 0) continue
          const result = runTxn(opened, { ops })
          if (!result.applied)
            fail(
              'command.invalid',
              `新建 ${label} 事务失败：${(result.failures ?? []).map((f) => f.error).join('; ')}`,
            )
        }
        bytes = await savePptxOf(opened)
      }

      guardSize(bytes, '产出')
      await ctx.fs.write(ref, bytes)
      ctx.changed(ref, 'structure')
      return {
        ref,
        size: bytes.byteLength,
        hash: sha256Hex(bytes),
        slides: (await openPptx(bytes)).deck.slides.length,
        notes: notes.slice(0, 20),
      }
    },
  },

  'pptx.apply': {
    summary: '执行一批命令：只入队，由已打开的页面执行并保存',
    args: {
      ref: { type: 'string', required: true },
      // commands 不声明：接受 JSON 文本或数组（宿主按 typeof 校验，声明了会拒）。
      baseHash: { type: 'string' },
      docRev: { type: 'number' },
      dryRun: { type: 'boolean' },
    },
    expose: { agent: true },
    run: async (input: Record<string, unknown>, ctx: PluginContext) => {
      const ref = refOf(input)
      currentCtx = ctx
      const lease = requireLease(ref)
      const commands = commandsOf(input)
      if (commands.length === 0) fail('command.invalid', 'commands 为空')
      if (commands.length > MAX_COMMANDS)
        fail('command.too-large', `批上限 ${MAX_COMMANDS} 条（收到 ${commands.length}）`)

      const { ops } = normalizeOps(commands)
      for (const [index, op] of ops.entries()) {
        if (!isKnownOp(op.op))
          fail(
            'command.not-found',
            `commands[${index}] 未知 op：${op.op || '(空)'}（用 pptx.guide 取可用 op）`,
          )
      }

      const baseHash = textOf(input['baseHash'])
      if (baseHash && baseHash !== lease.hash) {
        const bytes = await readRef(ctx, ref).catch(() => null)
        const disk = bytes ? sha256Hex(bytes) : ''
        if (disk && disk !== baseHash) fail('command.stale', '文件已被外部修改，请重新 read')
      }
      const docRev = input['docRev']
      if (typeof docRev === 'number' && docRev !== lease.docRev)
        fail('command.stale-index', `页面模型已变（docRev=${lease.docRev}），请重新 read`)

      const dryRun = input['dryRun'] === true
      const receipt = await requestPage<Receipt | null>(ref, {
        type: 'apply',
        id: randomUUID(),
        commands,
        ...(baseHash ? { baseHash } : {}),
        docRev: lease.docRev,
        ...(dryRun ? { dryRun: true } : {}),
      })
      if (dryRun) return { ref, dryRun: true, applied: 0 }
      return receipt ?? { ref, applied: ops.length, at: Date.now() }
    },
  },

  'pptx.load': {
    summary: '页面加载：一次返回整份字节',
    args: { ref: { type: 'string', required: true } },
    expose: { agent: false },
    run: async (input: Record<string, unknown>, ctx: PluginContext) => {
      const ref = refOf(input)
      currentCtx = ctx
      const bytes = await readRef(ctx, ref)
      return {
        ref,
        name: nameOf(ref),
        size: bytes.byteLength,
        hash: sha256Hex(bytes),
        bytes,
      }
    },
  },

  'pptx.save': {
    summary: '页面保存当前 ref（页面唯一落盘点）',
    args: {
      ref: { type: 'string', required: true },
      pageId: { type: 'string', required: true },
      baseHash: { type: 'string' },
    },
    expose: { agent: false },
    run: async (input: Record<string, unknown>, ctx: PluginContext) => {
      const ref = refOf(input)
      currentCtx = ctx
      const pageId = textOf(input['pageId'])
      if (!pageId) fail('command.invalid', '缺少 pageId')
      const lease = requireLease(ref)
      if (lease.pageId !== pageId) fail('command.page-unavailable', '页面已失效')
      const bytes = toBytes(input)
      guardSize(bytes, '提交')

      return withLock(ref, async () => {
        await guardPackage(bytes)
        const current = await readRef(ctx, ref)
        const baseHash = textOf(input['baseHash'])
        if (baseHash && baseHash !== sha256Hex(current))
          fail('command.stale', '文件已被外部修改，请重新打开或另存')
        await ctx.fs.write(ref, bytes)
        const hash = sha256Hex(bytes)
        lease.hash = hash
        lease.dirty = false
        lease.updatedAt = Date.now()
        ctx.changed(ref, 'content')
        return { ref, hash, size: bytes.byteLength, at: Date.now() } satisfies Receipt
      })
    },
  },

  'pptx.saveAs': {
    summary: '另存为：token 由 shell/native.saveFile 签发',
    args: {
      ref: { type: 'string', required: true },
      pageId: { type: 'string', required: true },
      token: { type: 'string', required: true },
    },
    expose: { agent: false },
    run: async (input: Record<string, unknown>, ctx: PluginContext) => {
      refOf(input)
      currentCtx = ctx
      if (!scope) fail('command.page-unavailable', '插件未激活')
      const bytes = toBytes(input)
      guardSize(bytes, '提交')
      const token = textOf(input['token'])
      if (!token) fail('command.invalid', '缺少 save token')
      let target: string
      try {
        target = scope.files.resolveSave(token)
      } catch (error) {
        fail('command.invalid', `保存目标无效：${messageOf(error)}`)
      }
      const ref = pathToFileURL(target).href
      requireLease(refOf(input))
      return withLock(ref, async () => {
        await guardPackage(bytes)
        await ctx.fs.write(ref, bytes)
        ctx.changed(ref, 'content')
        return { ref, hash: sha256Hex(bytes), size: bytes.byteLength, at: Date.now() } satisfies Receipt
      })
    },
  },

  'pptx.attach': {
    summary: '页面登记租约',
    args: {
      ref: { type: 'string', required: true },
      pageId: { type: 'string', required: true },
    },
    expose: { agent: false },
    run: async (input: Record<string, unknown>, ctx: PluginContext) => {
      const ref = refOf(input)
      currentCtx = ctx
      const pageId = textOf(input['pageId'])
      if (!pageId) fail('command.invalid', '缺少 pageId')
      const existing = leases.get(ref)
      // 宿主销毁 WebContentsView 不保证 detach：新 pageId 视为接管。
      if (existing && existing.pageId !== pageId)
        rejectPendingFor(ref, new Error('command.page-unavailable: 页面已重新打开'))
      const current = await readRef(ctx, ref)
      const lease: Lease = {
        pageId,
        hash: sha256Hex(current),
        docRev: 0,
        dirty: false,
        updatedAt: Date.now(),
      }
      leases.set(ref, lease)
      return {
        hash: lease.hash,
        docRev: lease.docRev,
        pending: [...pending.values()]
          .filter((entry) => entry.ref === ref)
          .map((entry) => entry.request),
      }
    },
  },

  'pptx.pending': {
    summary: '页面取回待处理请求正文',
    args: { id: { type: 'string', required: true } },
    expose: { agent: false },
    run: async (input: Record<string, unknown>) => {
      const entry = pending.get(textOf(input['id']))
      if (!entry) fail('command.not-found', '请求不存在或已超时')
      return entry.request
    },
  },

  'pptx.report': {
    summary: '页面续期租约并上报状态',
    args: {
      ref: { type: 'string', required: true },
      pageId: { type: 'string', required: true },
      hash: { type: 'string' },
      docRev: { type: 'number' },
      dirty: { type: 'boolean' },
    },
    expose: { agent: false },
    run: async (input: Record<string, unknown>) => {
      const ref = refOf(input)
      const pageId = textOf(input['pageId'])
      const lease = leases.get(ref)
      if (!lease || lease.pageId !== pageId) return { ok: false }
      lease.updatedAt = Date.now()
      if (typeof input['hash'] === 'string') lease.hash = input['hash']
      if (typeof input['docRev'] === 'number') lease.docRev = input['docRev']
      if (typeof input['dirty'] === 'boolean') lease.dirty = input['dirty']
      return { ok: true, docRev: lease.docRev, hash: lease.hash }
    },
  },

  'pptx.detach': {
    summary: '页面释放租约',
    args: {
      ref: { type: 'string', required: true },
      pageId: { type: 'string', required: true },
    },
    expose: { agent: false },
    run: async (input: Record<string, unknown>) => {
      const ref = refOf(input)
      const lease = leases.get(ref)
      if (lease && lease.pageId === textOf(input['pageId'])) {
        leases.delete(ref)
        rejectPendingFor(ref, new Error('command.page-unavailable: 页面已关闭'))
      }
      return { ok: true }
    },
  },

  'pptx.ack': {
    summary: '页面回执，结算待处理请求',
    args: {
      id: { type: 'string', required: true },
      ok: { type: 'boolean', required: true },
      payload: { type: 'string' },
    },
    expose: { agent: false },
    run: async (input: Record<string, unknown>) => {
      const id = textOf(input['id'])
      const entry = pending.get(id)
      if (!entry) return { ok: false }
      pending.delete(id)
      clearTimeout(entry.timer)
      const payload = ackPayload(input['payload'])
      if (input['ok'] !== true) {
        const error = payload['error'] as { code?: unknown; message?: unknown } | undefined
        entry.reject(
          new Error(
            `${textOf(error?.code) || 'command.invalid'}: ${textOf(error?.message) || '页面执行失败'}`,
          ),
        )
        return { ok: true }
      }
      entry.resolve(payload['receipt'] ?? payload['data'] ?? null)
      return { ok: true }
    },
  },
}

// ─── 输入解析 ───────────────────────────────────────────────────────

/** 取指定键，字符串按 JSON 文本解析（pages / slides / theme…）。 */
function jsonInput(
  input: Record<string, unknown>,
  keys: readonly string[],
): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  for (const key of keys) {
    const value = parseMaybeJson(input[key])
    if (value !== undefined) out[key] = value
  }
  return out
}

function parseMaybeJson(value: unknown): unknown {
  if (typeof value !== 'string') return value
  const text = value.trim()
  if (!text) return undefined
  if (!/^[[{]/.test(text)) return value
  try {
    return JSON.parse(text) as unknown
  } catch {
    return value
  }
}

function commandsOf(input: Record<string, unknown>): unknown[] {
  const raw = input['commands']
  const parsed = parseMaybeJson(raw)
  if (Array.isArray(parsed)) return parsed
  if (Array.isArray(raw)) return raw
  fail('command.invalid', 'commands 必须是数组或 JSON 数组文本')
}

function checkParamsOf(input: Record<string, unknown>): CheckParams {
  const params: Record<string, unknown> = {}
  for (const key of ['outline', 'pages', 'style']) {
    const value = parseMaybeJson(input[key])
    if (value !== undefined) params[key] = value
  }
  for (const key of ['styles', 'audit', 'skeleton']) {
    if (typeof input[key] === 'boolean') params[key] = input[key]
  }
  if (input['slide'] !== undefined) params['slide'] = input['slide']
  return params as CheckParams
}

function savePptxOf(opened: OpenedPptx): Promise<Uint8Array> {
  return savePptx(opened)
}

// ─── 贡献 ───────────────────────────────────────────────────────────

export const workstudio = definePlugin({
  resources,
  commands,
  events: true,
  lifecycle: {
    activate(value: PluginScope): void {
      scope = value
      value.log.info(`pptx: ${Object.keys(commands).length} commands, ${allOps().length} ops`)
    },
    deactivate(): void {
      for (const entry of pending.values()) {
        clearTimeout(entry.timer)
        entry.reject(new Error('command.page-unavailable: 插件已停用'))
      }
      pending.clear()
      leases.clear()
      locks.clear()
      currentCtx = null
      scope = null
    },
  },
})

export const ready = true
export default workstudio
