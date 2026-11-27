// ─── 页面租约 + 请求队列 ──────────────────────────────────────────────
//
// 宿主不提供 `ctx.origin`，所以「这条命令是不是页面发的」由租约令牌回答：
//   attach 发放 token（页面持有）→ 写与结算 op 必须带 token。
// 助手能调到 `page` 命令，但拿不到 token，也抢不走已有租约（lease.busy）。
//
// 请求队列（pending）是 agent → 页面 的单次往返：引擎只把 id/type 通过事件
// 推给页面，页面再用 `pending` 取正文、`ack` 结算。

import { randomUUID } from 'node:crypto'

import { CommandError } from './errors'

export interface Lease {
  ref: string
  pageId: string
  token: string
  hash: string
  docRev: number
  structureEpoch: number
  dirty: boolean
  selection?: string
  updatedAt: number
}

/** 租约 TTL 与请求超时分开：后台标签的定时器会被 Chromium 节流。 */
export const LEASE_TTL_MS = 5 * 60_000
/** 单次页面往返的超时。 */
export const PENDING_TIMEOUT_MS = 30_000

const leases = new Map<string, Lease>()

export function leaseOf(ref: string): Lease | null {
  const lease = leases.get(ref)
  if (!lease) return null
  if (Date.now() - lease.updatedAt > LEASE_TTL_MS) {
    leases.delete(ref)
    failPending(ref, '页面租约过期，请求作废')
    return null
  }
  return lease
}

export interface AttachInput {
  pageId: string
  hash: string
  takeover?: boolean
}

export function attachLease(ref: string, input: AttachInput): Lease {
  const current = leaseOf(ref)
  if (current && current.pageId !== input.pageId && input.takeover !== true)
    throw new CommandError(
      'lease.busy',
      `页面租约已被占用（pageId=${current.pageId.slice(0, 8)}…）；页面重载请带 takeover:true`,
    )
  if (current && current.pageId !== input.pageId) failPending(ref, '页面已更换，请求作废')
  const reuse = current?.pageId === input.pageId
  const lease: Lease = {
    ref,
    pageId: input.pageId,
    token: reuse ? current!.token : randomUUID(),
    hash: input.hash,
    docRev: 0,
    structureEpoch: 0,
    dirty: false,
    updatedAt: Date.now(),
  }
  leases.set(ref, lease)
  return lease
}

export function verifyLease(ref: string, token: string): Lease {
  const lease = leaseOf(ref)
  if (!lease)
    throw new CommandError('lease.missing', '页面未登记租约（先 attach），或租约已过期')
  if (!token || token !== lease.token)
    throw new CommandError('denied', '页面令牌不匹配：该 op 只接受当前页面调用')
  return lease
}

export interface ReportInput {
  token: string
  hash?: string
  docRev?: number
  structureEpoch?: number
  dirty?: boolean
  selection?: string
}

export function reportLease(ref: string, input: ReportInput): Lease {
  const lease = verifyLease(ref, input.token)
  lease.updatedAt = Date.now()
  if (input.hash) lease.hash = input.hash
  if (input.docRev !== undefined) lease.docRev = input.docRev
  if (input.structureEpoch !== undefined) lease.structureEpoch = input.structureEpoch
  if (input.dirty !== undefined) lease.dirty = input.dirty === true
  if (input.selection !== undefined) lease.selection = input.selection
  return lease
}

export function detachLease(ref: string, token: string): void {
  const lease = verifyLease(ref, token)
  leases.delete(lease.ref)
  failPending(ref, '页面已关闭')
}

/** 诊断 / 测试：当前租约（只读拷贝）。 */
export function leaseList(): Lease[] {
  return [...leases.values()]
}

// ─── 请求队列 ─────────────────────────────────────────────────────────

export type PageRequestType = 'apply' | 'read'

export interface PageRequest {
  id: string
  ref: string
  type: PageRequestType
  commands?: unknown[]
  what?: string
  params?: Record<string, unknown>
  docRev?: number
  structureEpoch?: number
  baseHash?: string
  dryRun?: boolean
}

export interface PageAckPayload {
  applied?: number
  mutated?: boolean
  dryRun?: boolean
  receipt?: { ref: string; hash: string; size: number; at: number; applied?: number }
  data?: unknown
  error?: { code: string; message: string }
}

interface Pending {
  id: string
  ref: string
  type: PageRequestType
  request: PageRequest
  timer: NodeJS.Timeout
  resolve(payload: PageAckPayload): void
  reject(error: Error): void
}

const pending = new Map<string, Pending>()

function failPending(ref: string, message: string): void {
  for (const [id, item] of pending) {
    if (item.ref !== ref) continue
    clearTimeout(item.timer)
    pending.delete(id)
    item.reject(new CommandError('command.page-unavailable', message))
  }
}

export function enqueue(
  ref: string,
  request: Omit<PageRequest, 'id' | 'ref'>,
): Promise<PageAckPayload> {
  const lease = leaseOf(ref)
  if (!lease) return Promise.reject(new CommandError('command.page-unavailable', '页面未打开或租约已过期'))

  const id = randomUUID()
  const full: PageRequest = { id, ref, ...request }
  return new Promise<PageAckPayload>((resolve, reject) => {
    const timer = setTimeout(() => {
      pending.delete(id)
      reject(new CommandError('command.timeout', '页面请求超时（30 s）'))
    }, PENDING_TIMEOUT_MS)
    pending.set(id, { id, ref, type: full.type, request: full, timer, resolve, reject })
    try {
      notify(ref, full)
    } catch (error) {
      clearTimeout(timer)
      pending.delete(id)
      reject(error instanceof Error ? error : new Error(String(error)))
    }
  })
}

/** 由 index.ts 注入（`ctx.events.emit`）；保持 state 不依赖宿主 SDK。 */
let notify = (ref: string, request: PageRequest): void => {
  void ref
  void request
  throw new CommandError('command.unsupported', '事件通道未就绪')
}

export function bindNotifier(fn: (ref: string, request: PageRequest) => void): void {
  notify = fn
}

export function pendingList(ref: string): Array<{ id: string; type: PageRequestType }> {
  return [...pending.values()]
    .filter((item) => item.ref === ref)
    .map((item) => ({ id: item.id, type: item.type }))
}

export function pendingOf(id: string): PageRequest | null {
  return pending.get(id)?.request ?? null
}

export function settlePending(id: string, ok: boolean, payload: PageAckPayload): boolean {
  const item = pending.get(id)
  if (!item) return false
  clearTimeout(item.timer)
  pending.delete(id)
  if (ok) item.resolve(payload)
  else item.reject(new CommandError(payload.error?.code ?? 'command.invalid', payload.error?.message ?? '页面拒绝了请求'))
  return true
}

/** 测试 / deactivate：丢掉全部租约与在途请求。 */
export function resetLeases(): void {
  leases.clear()
  for (const item of pending.values()) {
    clearTimeout(item.timer)
    item.reject(new CommandError('command.page-unavailable', '插件已停用'))
  }
  pending.clear()
}
