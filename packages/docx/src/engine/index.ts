// ─── docx engine：WorkStudio SDK 入口 ─────────────────────────────
//
// 只做文档页：给定 ref 读 / 写当前文档；不做文件列表、新建、标签页。
// Agent 的写一律通过页面队列；没有有效页面租约时拒绝，不后台写。
//
// 该文件在宿主主进程求值：允许使用 Node 内置模块；不得 import 编辑器模块。

import { randomUUID } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'

import { parseDocxBytes, tierFor } from '../shared/document'
import { blockHash, sha256Hex } from '../shared/sha256'
import type { ModelBlock } from '../shared/types'
import { buildCommandGuide } from '../tools/catalog'
import { isDocxPackage, readMany } from './read'

// SDK 的 definePlugin 只做恒等返回；这里不引运行时依赖，保证插件自包含。
function definePlugin<T>(contribution: T): T {
  return contribution
}

interface Log {
  info(message: string): void
  warn(message: string): void
  error(message: string): void
}

interface PluginScope {
  readonly pluginId: string
  readonly events: { emit(name: string, data?: unknown): void }
  readonly log: Log
}

interface PluginContext {
  fs: {
    read(ref: string): Promise<Uint8Array>
    write(ref: string, bytes: Uint8Array): Promise<void>
  }
  changed(ref: string, type: 'content' | 'structure'): void
  log: Log
}

interface CommandInput {
  ref?: string
  [key: string]: unknown
}

type CommandRun = (input: CommandInput, ctx: PluginContext) => Promise<unknown>

interface CommandDecl {
  summary: string
  args: Record<string, unknown>
  expose?: { agent?: boolean }
  run: CommandRun
}

interface ResourceItem {
  locator: string
  [key: string]: unknown
}

interface ResourceAdapter {
  read(ref: string, options?: { limit?: number; offset?: number }): Promise<ResourceItem[]>
  grep(ref: string, query: string): Promise<ResourceItem[]>
  open?(ref: string, locator?: string): Promise<{ ref: string; locator?: string } | null>
}

interface PageLease {
  pageId: string
  docRev: number
  hash: string
  updatedAt: number
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

type PageRequestInput =
  | (Omit<Extract<PageRequest, { type: 'apply' }>, 'id'> & { id?: string })
  | (Omit<Extract<PageRequest, { type: 'read' }>, 'id'> & { id?: string })

interface PendingRequest {
  ref: string
  request: PageRequest
  resolve: (value: unknown) => void
  reject: (error: Error) => void
  timer: ReturnType<typeof setTimeout>
}

const REQUEST_TIMEOUT_MS = 30_000
const MAX_PAGE_BYTES = 16 * 1024 * 1024
const MAX_READ_BYTES = 64 * 1024 * 1024
const MAX_BLOCKS = 500

let scope: PluginScope | null = null
const leases = new Map<string, PageLease>()
const pending = new Map<string, PendingRequest>()

function fail(code: string, message: string): never {
  throw new Error(`${code}: ${message}`)
}

function textOf(value: unknown): string {
  return typeof value === 'string' ? value : ''
}

function refOf(input: CommandInput): string {
  const ref = textOf(input.ref)
  if (!ref) fail('command.invalid', '缺少 ref')
  return ref
}

function nameOf(ref: string): string {
  try {
    const url = new URL(ref)
    return decodeURIComponent(url.pathname.slice(url.pathname.lastIndexOf('/') + 1)) || 'document.docx'
  } catch {
    return ref.slice(ref.lastIndexOf('/') + 1) || 'document.docx'
  }
}

function isDocxRef(ref: string): boolean {
  const path = ref.split(/[?#]/, 1)[0] ?? ref
  return /\.(docx|docm|dotx)$/i.test(path)
}

async function readFileBytes(ref: string): Promise<Uint8Array> {
  if (!isDocxRef(ref)) fail('resource.not-found', `不是 Word 文档：${ref}`)
  try {
    return new Uint8Array(await readFile(fileURLToPath(ref)))
  } catch (error) {
    fail('resource.not-found', `读取失败：${error instanceof Error ? error.message : String(error)}`)
  }
}

function hashOf(bytes: Uint8Array): string {
  return sha256Hex(bytes)
}

let SEED: Uint8Array | null = null

function seedBytes(): Uint8Array {
  SEED ??= new Uint8Array(readFileSync(new URL('./seed.docx', import.meta.url)))
  return SEED
}

function itemOf(block: ModelBlock): ResourceItem {
  return {
    locator: `docx:block:${block.docxIndex ?? 0}`,
    index: block.docxIndex ?? 0,
    type: block.kind,
    heading: block.heading,
    text: block.text,
    hash: blockHash(block.xml),
  }
}

function parseBytes(bytes: Uint8Array): ReturnType<typeof parseDocxBytes> {
  try {
    return parseDocxBytes(bytes)
  } catch (error) {
    fail('command.invalid', error instanceof Error ? error.message : String(error))
  }
}

async function readItems(ref: string, options?: { limit?: number; offset?: number }): Promise<ResourceItem[]> {
  const bytes = await readFileBytes(ref)
  const parsed = parseBytes(bytes)
  const visible = parsed.blocks.filter((block) => !block.hidden)
  const meta: ResourceItem = {
    locator: 'docx:doc',
    name: nameOf(ref),
    size: bytes.byteLength,
    hash: hashOf(bytes),
    blockCount: visible.length,
    tableCount: visible.filter((block) => block.kind === 'tbl').length,
    tier: tierFor(parsed.weight),
  }
  const items = [meta, ...visible.slice(0, MAX_BLOCKS).map(itemOf)]
  const offset = Math.max(0, Number(options?.offset ?? 0) || 0)
  const limit = Math.max(1, Math.min(MAX_BLOCKS, Number(options?.limit ?? MAX_BLOCKS) || MAX_BLOCKS))
  return items.slice(offset, offset + limit)
}

async function grepBlocks(ref: string, query: string): Promise<ResourceItem[]> {
  const bytes = await readFileBytes(ref)
  const parsed = parseBytes(bytes)
  const needle = query.toLowerCase()
  return parsed.blocks
    .filter((block) => !block.hidden && block.text.toLowerCase().includes(needle))
    .slice(0, 200)
    .map((block) => ({ ...itemOf(block), snippet: block.text.slice(0, 240) }))
}

function leaseOf(ref: string): PageLease | null {
  // 租约不按时间失效：页面打开就是打开，关闭/崩溃由新页面 attach 接管，
  // 或客户端退出时随进程清除。updatedAt 仅用于诊断。
  return leases.get(ref) ?? null
}

function requireLease(ref: string): PageLease {
  const lease = leaseOf(ref)
  if (!lease) fail('command.page-unavailable', '文档页面未打开')
  return lease
}

/** 页面自己的命令（save/report）可在租约过期后重新登记；不同 pageId 仍拒绝。 */
function ensureLeaseForPage(ref: string, pageId: string): PageLease {
  const existing = leaseOf(ref)
  if (existing) {
    if (existing.pageId !== pageId)
      fail('command.page-unavailable', '该文档已在另一个页面打开')
    existing.updatedAt = Date.now()
    return existing
  }
  const lease: PageLease = { pageId, docRev: 0, hash: '', updatedAt: Date.now() }
  leases.set(ref, lease)
  return lease
}

function rejectPendingFor(ref: string, error: Error): void {
  for (const [id, entry] of pending) {
    if (entry.ref !== ref) continue
    clearTimeout(entry.timer)
    pending.delete(id)
    entry.reject(error)
  }
}

function requestPage(
  ref: string,
  request: PageRequestInput,
): Promise<unknown> {
  requireLease(ref)
  const current = scope
  if (!current) fail('command.page-unavailable', '插件尚未激活')
  const id = request.id ?? randomUUID()
  const full = { ...request, id } as PageRequest
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pending.delete(id)
      reject(new Error('command.page-unavailable: 页面在 30 秒内没有响应'))
    }, REQUEST_TIMEOUT_MS)
    pending.set(id, { ref, request: full, resolve, reject, timer })
    current.events.emit('request', { id, type: full.type })
  })
}

function toBytes(input: CommandInput): Uint8Array {
  const raw = input['bytes']
  if (raw instanceof Uint8Array) return raw
  if (raw instanceof ArrayBuffer) return new Uint8Array(raw)
  const data = textOf(input['data'])
  if (!data) fail('command.invalid', '缺少 bytes / data')
  try {
    return new Uint8Array(Buffer.from(data, 'base64'))
  } catch {
    fail('command.invalid', 'data 不是合法 base64')
  }
}

function parseCommands(value: unknown): unknown[] {
  if (Array.isArray(value)) return value
  if (typeof value !== 'string') fail('command.invalid', 'commands 必须是 JSON 字符串')
  let parsed: unknown
  try {
    parsed = JSON.parse(value)
  } catch {
    fail('command.invalid', 'commands 不是合法 JSON')
  }
  if (!Array.isArray(parsed)) fail('command.invalid', 'commands 必须是数组')
  if (parsed.length > 200) fail('command.too-large', '一批最多 200 条命令')
  return parsed
}

function ackPayload(value: unknown): Record<string, unknown> {
  if (typeof value !== 'string') return {}
  try {
    const parsed = JSON.parse(value) as unknown
    return parsed && typeof parsed === 'object' ? (parsed as Record<string, unknown>) : {}
  } catch {
    return {}
  }
}

function parseTier(bytes: Uint8Array): ReturnType<typeof tierFor> {
  try {
    return tierFor(parseDocxBytes(bytes).weight)
  } catch {
    return 'refused'
  }
}

const resources: Record<string, ResourceAdapter> = {
  docx: {
    read: readItems,
    grep: grepBlocks,
    open: async (ref, locator) => ({ ref, ...(locator ? { locator } : {}) }),
  },
}

const commands: Record<string, CommandDecl> = {
  'docx.guide': {
    summary: 'Word 命令目录：无参分组，group 取组内 schema，op 取单条',
    args: {
      ref: { type: 'string', required: true },
      group: { enum: ['blocks', 'editor', 'actions'] },
      op: { type: 'string' },
    },
    expose: { agent: true },
    run: async (input) => {
      const guide = buildCommandGuide({
        group: typeof input['group'] === 'string' ? input['group'] : undefined,
        op: typeof input['op'] === 'string' ? input['op'] : undefined,
      })
      return guide.summary
    },
  },

  'docx.read': {
    summary: '读 Word 文档：context / blocks / comments / notes / revisions / styles / audit',
    args: {
      ref: { type: 'string', required: true },
      what: { enum: ['context', 'blocks', 'comments', 'notes', 'revisions', 'styles', 'audit'] },
      startBlockIndex: { type: 'number' },
      endBlockIndex: { type: 'number' },
      offset: { type: 'number' },
    },
    expose: { agent: true },
    run: async (input, ctx) => {
      const ref = refOf(input)
      const params = {
        what: typeof input['what'] === 'string' ? input['what'] : 'context',
        startBlockIndex: input['startBlockIndex'],
        endBlockIndex: input['endBlockIndex'],
        offset: input['offset'],
      }
      if (leaseOf(ref)) {
        const data = await requestPage(ref, { type: 'read', what: params.what, params })
        if (typeof data === 'string') return data
        if (data && typeof data === 'object') {
          const value = data as { text?: unknown; hash?: unknown; docRev?: unknown }
          if (typeof value.text === 'string')
            return `hash=${textOf(value.hash)} docRev=${String(value.docRev ?? '')}\n${value.text}`
        }
        return JSON.stringify(data ?? '')
      }
      const bytes = await ctx.fs.read(ref)
      return `hash=${hashOf(bytes)} docRev=-\n${readMany(bytes, params)}`
    },
  },

  'docx.apply': {
    summary: '在已打开的 Word 页面上执行一批命令（commands 为 JSON 字符串）；无页面直接拒绝',
    args: {
      ref: { type: 'string', required: true },
      commands: { type: 'string', required: true },
      baseHash: { type: 'string' },
      docRev: { type: 'number' },
      dryRun: { type: 'boolean' },
    },
    expose: { agent: true },
    run: async (input) => {
      const ref = refOf(input)
      const commandsValue = parseCommands(input['commands'])
      const lease = requireLease(ref)
      if (typeof input['docRev'] === 'number' && input['docRev'] !== lease.docRev)
        fail('command.stale-index', `页面基线已变（docRev ${input['docRev']} → ${lease.docRev}）`)
      const receipt = await requestPage(ref, {
        type: 'apply',
        commands: commandsValue,
        baseHash: typeof input['baseHash'] === 'string' ? input['baseHash'] : undefined,
        docRev: lease.docRev,
        dryRun: input['dryRun'] === true,
      })
      return receipt ?? { applied: commandsValue.length }
    },
  },

  'docx.create': {
    summary: '创建一个空白 Word 文件（seed 模板，不写正文）；创建后由宿主/用户打开页面再 apply',
    args: {
      ref: { type: 'string', required: true },
    },
    expose: { agent: true },
    run: async (input, ctx) => {
      const ref = refOf(input)
      if (!isDocxRef(ref)) fail('command.invalid', `不是 Word 文档路径：${ref}`)
      let exists = false
      try {
        await ctx.fs.read(ref)
        exists = true
      } catch {
        // 目标不存在（或不可读）；后续 write 会给出真实错误。
      }
      if (exists) fail('resource.exists', `目标已存在：${ref}`)
      const bytes = seedBytes()
      await ctx.fs.write(ref, bytes)
      ctx.changed(ref, 'structure')
      return { ref, size: bytes.byteLength, hash: hashOf(bytes), created: true }
    },
  },

  'docx.load': {
    summary: '页面加载当前文档字节',
    args: {
      ref: { type: 'string', required: true },
      base64: { type: 'boolean' },
    },
    expose: { agent: false },
    run: async (input, ctx) => {
      const ref = refOf(input)
      const bytes = await ctx.fs.read(ref)
      if (bytes.byteLength > MAX_READ_BYTES) fail('command.too-large', '文档超过 64 MB')
      const base = {
        ref,
        name: nameOf(ref),
        size: bytes.byteLength,
        hash: hashOf(bytes),
        tier: parseTier(bytes),
      }
      if (input['base64'] === true)
        return { ...base, data: Buffer.from(bytes).toString('base64') }
      return { ...base, bytes }
    },
  },

  'docx.save': {
    summary: '页面保存当前 ref（页面唯一落盘点）',
    args: {
      ref: { type: 'string', required: true },
      pageId: { type: 'string', required: true },
      baseHash: { type: 'string' },
      type: { enum: ['page', 'agent'] },
      data: { type: 'string' },
    },
    expose: { agent: false },
    run: async (input, ctx) => {
      const ref = refOf(input)
      const pageId = textOf(input['pageId'])
      if (!pageId) fail('command.invalid', '缺少 pageId')
      const lease = ensureLeaseForPage(ref, pageId)
      const bytes = toBytes(input)
      if (bytes.byteLength === 0) fail('command.invalid', '空字节')
      if (bytes.byteLength > MAX_PAGE_BYTES) fail('command.too-large', '文档超过 16 MB')
      if (!isDocxPackage(bytes)) fail('command.invalid', '不是有效的 DOCX 包')
      const current = await ctx.fs.read(ref)
      if (typeof input['baseHash'] === 'string' && input['baseHash'] !== hashOf(current))
        fail('command.stale', '文件已被外部修改，请重新读取')
      await ctx.fs.write(ref, bytes)
      const hash = hashOf(bytes)
      lease.hash = hash
      lease.updatedAt = Date.now()
      ctx.changed(ref, 'content')
      return {
        ref,
        hash,
        size: bytes.byteLength,
        at: Date.now(),
        applied: input['type'] === 'agent' ? 1 : undefined,
      }
    },
  },

  'docx.attach': {
    summary: '页面登记租约',
    args: {
      ref: { type: 'string', required: true },
      pageId: { type: 'string', required: true },
    },
    expose: { agent: false },
    run: async (input, ctx) => {
      const ref = refOf(input)
      const pageId = textOf(input['pageId'])
      if (!pageId) fail('command.invalid', '缺少 pageId')
      const existing = leases.get(ref)
      // 宿主关闭页面不保证调用 detach（WebContentsView 可能直接销毁）。
      // 同一 ref 的新页面 attach 视为接管：拒绝旧页面的 pending，登记新租约。
      // 旧页面之后迟到的 detach 会因 pageId 不匹配而不会删掉新租约。
      if (existing && existing.pageId !== pageId) {
        rejectPendingFor(ref, new Error('command.page-unavailable: 页面已重新打开'))
      }
      const bytes = await ctx.fs.read(ref)
      leases.set(ref, { pageId, docRev: 0, hash: hashOf(bytes), updatedAt: Date.now() })
      return {
        hash: hashOf(bytes),
        pending: [...pending.values()]
          .filter((entry) => entry.ref === ref)
          .map((entry) => entry.request),
      }
    },
  },

  'docx.pending': {
    summary: '页面取回待处理请求正文',
    args: { id: { type: 'string', required: true } },
    expose: { agent: false },
    run: async (input) => {
      const entry = pending.get(textOf(input['id']))
      return entry ? entry.request : null
    },
  },

  'docx.report': {
    summary: '页面上报 live 版本',
    args: {
      ref: { type: 'string', required: true },
      pageId: { type: 'string', required: true },
      docRev: { type: 'number', required: true },
      hash: { type: 'string', required: true },
    },
    expose: { agent: false },
    run: async (input) => {
      const ref = refOf(input)
      const pageId = textOf(input['pageId'])
      if (!pageId) fail('command.invalid', '缺少 pageId')
      const lease = ensureLeaseForPage(ref, pageId)
      lease.docRev = Number(input['docRev'] ?? 0)
      lease.hash = textOf(input['hash']) || lease.hash
      lease.updatedAt = Date.now()
      return { ok: true }
    },
  },

  'docx.detach': {
    summary: '页面撤销租约',
    args: {
      ref: { type: 'string', required: true },
      pageId: { type: 'string', required: true },
    },
    expose: { agent: false },
    run: async (input) => {
      const ref = refOf(input)
      const lease = leases.get(ref)
      if (lease && lease.pageId === textOf(input['pageId'])) {
        leases.delete(ref)
        rejectPendingFor(ref, new Error('command.page-unavailable: 页面已关闭'))
      }
      return { ok: true }
    },
  },

  'docx.ack': {
    summary: '页面回执，结算待处理请求',
    args: {
      id: { type: 'string', required: true },
      ok: { type: 'boolean', required: true },
      payload: { type: 'string' },
    },
    expose: { agent: false },
    run: async (input) => {
      const id = textOf(input['id'])
      const entry = pending.get(id)
      if (!entry) return { ok: false }
      pending.delete(id)
      clearTimeout(entry.timer)
      const payload = ackPayload(input['payload'])
      if (input['ok'] !== true) {
        const error = payload['error'] as { code?: unknown; message?: unknown } | undefined
        entry.reject(new Error(`${textOf(error?.code) || 'command.invalid'}: ${textOf(error?.message) || '页面执行失败'}`))
        return { ok: true }
      }
      const receipt = payload['receipt']
      entry.resolve(receipt ?? payload['data'] ?? null)
      return { ok: true }
    },
  },
}

export const workstudio = definePlugin({
  resources,
  commands,
  events: true,
  lifecycle: {
    activate(value: PluginScope): void {
      scope = value
    },
    deactivate(): void {
      for (const entry of pending.values()) {
        clearTimeout(entry.timer)
        entry.reject(new Error('command.page-unavailable: 插件已停用'))
      }
      pending.clear()
      leases.clear()
      scope = null
    },
  },
})

export default workstudio
