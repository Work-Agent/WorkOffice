// ─── xlsx.page：页面通道（单入口 + 租约令牌） ───────────────────
//
// 只读 op（status/open/chunk/meta/range）不要求令牌：页面在 attach 之前
// 就要用它们建模型。写与结算 op 一律要求 attach 发放的 token —— 这是
// 「宿主不给 ctx.origin」时区分页面与助手的唯一可靠依据。
//
// 打开有两条路：
//   wasm 可用 → `meta`（open 元数据）+ `range`（read_range），页面不读整份字节；
//   wasm 不可用 → `meta` 只给薄清单，页面回退到 `open/chunk` 整份字节。

import { randomUUID } from 'node:crypto'
import { mkdir, open, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { pathToFileURL } from 'node:url'

import type { PluginContext } from '@work-studio/sdk'

import type { XlsxSaveRequest } from '../../shared/save-request'
import { CommandError, invalid, stale } from '../errors'
import { assertXlsxPackage, bytesOfPlan, bytesOfSaveRequest, planSaveRequest } from '../gateway'
import { saveViaWasm } from '../archive'
import { loadWasmEngine } from '../wasm'
import { hashOfFile, rememberHash, sha256Hex } from '../hash'
import { withLock } from '../lock'
import { recalcFormulaValues } from '../recalc'
import { assertExt, assertXlsxFile, MAX_FILE_BYTES, pathOf, tierOf, type XlsxFile } from '../ref'
import { pageMeta } from '../read'
import { getScope } from '../scope'
import { attachLease, detachLease, pendingOf, reportLease, settlePending, verifyLease } from '../state'
import { wasmRange } from '../wasm-read'

/** 单块读取上限：4 MiB。 */
const MAX_CHUNK_BYTES = 4 * 1024 * 1024

const READ_OPS = new Set(['status', 'open', 'chunk', 'meta', 'range', 'draftStatus', 'draftLoad'])
const SAVE_OPS = new Set(['save', 'saveAs', 'saveBegin', 'saveChunk', 'saveCommit'])
const TOKEN_OPS = new Set(['report', 'detach', 'pending', 'ack', 'draft', 'draftClear'])

type AnyRecord = Record<string, unknown>

function argsOf(raw: unknown): AnyRecord {
  if (raw === undefined || raw === null || raw === '') return {}
  if (typeof raw === 'object') return raw as AnyRecord
  if (typeof raw !== 'string') invalid('input 必须是 JSON 字符串')
  try {
    const parsed = JSON.parse(raw) as unknown
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed))
      invalid('input 必须是 JSON 对象')
    return parsed as AnyRecord
  } catch (error) {
    if (error instanceof CommandError) throw error
    invalid(`input 不是合法 JSON：${error instanceof Error ? error.message : String(error)}`)
  }
}

function numberArg(value: unknown, name: string): number {
  const parsed = Number(value)
  if (!Number.isFinite(parsed) || parsed < 0) invalid(`${name} 必须是非负数字`)
  return Math.floor(parsed)
}

function asBytes(value: unknown): Uint8Array | null {
  if (value instanceof Uint8Array) return value
  if (value instanceof ArrayBuffer) return new Uint8Array(value)
  if (Array.isArray(value)) return new Uint8Array(value as number[])
  return null
}

async function readSlice(path: string, offset: number, length: number): Promise<Uint8Array> {
  const handle = await open(path, 'r')
  try {
    const buffer = Buffer.alloc(length)
    const { bytesRead } = await handle.read(buffer, 0, length, offset)
    return new Uint8Array(buffer.buffer, buffer.byteOffset, bytesRead)
  } finally {
    await handle.close()
  }
}

export async function page(input: AnyRecord, ctx?: PluginContext): Promise<unknown> {
  const ref = String(input.ref ?? '')
  const op = String(input.op ?? '')
  if (!ref) invalid('缺少 ref')
  if (!op) invalid('缺少 op')

  const args = argsOf(input.input)
  const file = await assertXlsxFile(ref)

  if (op === 'attach') return attachOp(input, file, args)
  if (READ_OPS.has(op)) return readOp(op, file, args)
  if (SAVE_OPS.has(op)) return saveOp(op, file, input, args, ctx)
  if (TOKEN_OPS.has(op)) return settleOp(op, file, input, args)
  throw new CommandError('command.unsupported', `未知页面 op：${op}`)
}

async function attachOp(input: AnyRecord, file: XlsxFile, args: AnyRecord): Promise<unknown> {
  const pageId = String(input.pageId ?? args.pageId ?? '')
  if (!pageId) invalid('attach 缺少 pageId')
  const hash = await hashOfFile(file)
  const lease = attachLease(file.ref, {
    pageId,
    hash,
    ...(args.takeover === true ? { takeover: true } : {}),
  })
  return {
    ref: file.ref,
    name: file.name,
    size: file.size,
    token: lease.token,
    hash: lease.hash,
    docRev: lease.docRev,
    structureEpoch: lease.structureEpoch,
    pending: [],
  }
}

async function readOp(op: string, file: XlsxFile, args: AnyRecord): Promise<unknown> {
  if (op === 'status' || op === 'open') {
    const hash = await hashOfFile(file)
    return {
      ref: file.ref,
      name: file.name,
      size: file.size,
      ...(op === 'status' ? { mtime: file.mtime } : {}),
      hash,
      tier: tierOf(file.size),
    }
  }

  if (op === 'chunk') {
    const expected = String(args.hash ?? '')
    if (!expected) invalid('chunk 缺少 hash')
    const actual = await hashOfFile(file)
    if (actual !== expected)
      stale(`文件已变化（hash ${expected.slice(0, 8)}… → ${actual.slice(0, 8)}…），请重新 open`)
    const offset = numberArg(args.offset ?? 0, 'offset')
    const length = Math.min(numberArg(args.length ?? 1024 * 1024, 'length') || 1, MAX_CHUNK_BYTES)
    if (offset > file.size) invalid(`offset 越界：${offset} > ${file.size}`)
    const bytes = await readSlice(file.path, offset, Math.min(length, file.size - offset))
    const next = offset + bytes.byteLength
    return { ref: file.ref, hash: actual, bytes, offset, next, eof: next >= file.size, size: file.size }
  }

  const hash = await hashOfFile(file)
  if (op === 'meta') return pageMeta(file, hash)

  if (op === 'draftStatus' || op === 'draftLoad') {
    const draftPath = recoveryPath(file.ref)
    const info = await stat(draftPath).catch(() => null)
    if (!info?.isFile()) return { exists: false }
    return {
      exists: true,
      at: info.mtimeMs,
      size: info.size,
      ...(op === 'draftLoad' ? { bytes: new Uint8Array(await readFile(draftPath)) } : {}),
    }
  }

  const baseHash = typeof args.baseHash === 'string' ? args.baseHash : ''
  if (baseHash && baseHash !== hash)
    stale(`文件已变化（hash ${baseHash.slice(0, 8)}… → ${hash.slice(0, 8)}…），请重新 meta`)
  return wasmRange(file, hash, args as never)
}

// ─── 写 ───────────────────────────────────────────────────────────────

/** 分块保存的在途缓冲：一次保存一个 transferId，完成后立刻释放。 */
const transfers = new Map<string, { size: number; chunks: Map<number, Uint8Array> }>()

async function saveOp(
  op: string,
  file: XlsxFile,
  input: AnyRecord,
  args: AnyRecord,
  ctx: PluginContext | undefined,
): Promise<unknown> {
  const token = String(input.token ?? args.token ?? '')
  verifyLease(file.ref, token)
  if (!ctx?.fs) throw new CommandError('command.unsupported', '宿主未提供写入通道')

  if (op === 'saveBegin') {
    const transferId = randomUUID()
    transfers.set(transferId, { size: numberArg(args.size ?? 0, 'size'), chunks: new Map() })
    return { transferId }
  }

  if (op === 'saveChunk') {
    const transferId = String(args.transferId ?? '')
    const transfer = transfers.get(transferId)
    if (!transfer) invalid(`保存批次不存在：${transferId}`)
    const seq = numberArg(args.seq ?? 0, 'seq')
    const bytes = asBytes(args.bytes)
    if (!bytes) invalid('saveChunk 缺少 bytes')
    transfer.chunks.set(seq, bytes)
    return { transferId, received: transfer.chunks.size }
  }

  if (op === 'saveCommit') {
    const transferId = String(args.transferId ?? '')
    const transfer = transfers.get(transferId)
    if (!transfer) invalid(`保存批次不存在：${transferId}`)
    transfers.delete(transferId)
    const seqs = [...transfer.chunks.keys()].sort((a, b) => a - b)
    if (seqs.length !== transfer.chunks.size) invalid('保存批次序号不连续')
    const total = [...transfer.chunks.values()].reduce((sum, chunk) => sum + chunk.byteLength, 0)
    if (transfer.size > 0 && total !== transfer.size)
      invalid(`保存数据不完整：收到 ${total} 字节，期望 ${transfer.size}`)
    const bytes = new Uint8Array(total)
    let at = 0
    for (const seq of seqs) {
      const chunk = transfer.chunks.get(seq)!
      bytes.set(chunk, at)
      at += chunk.byteLength
    }
    return writeBytes(file, bytes, args, ctx, 1)
  }

  if (op === 'saveAs') {
    const saveToken = String(args.saveToken ?? '')
    if (!saveToken.startsWith('save:')) invalid('缺少 save: token（另存为先用 shell/native.saveFile 取 token）')
    const target = getScope().files.resolveSave(saveToken)
    assertExt(target)
    const ref = pathToFileURL(target).href
    const provided = asBytes(args.bytes)
    if (!provided) invalid('saveAs 需要 bytes')
    const bytes = provided
    checkBytes(bytes)
    const targetPath = pathOf(ref)
    await ctx.fs.write(ref, bytes)
    const info = await stat(targetPath)
    const hash = sha256Hex(bytes)
    rememberHash(targetPath, info.size, info.mtimeMs, hash)
    ctx.changed?.(ref, 'content')
    return { ref, hash, size: bytes.byteLength, at: Date.now(), applied: 1 }
  }

  const provided = asBytes(args.bytes)
  let bytes = provided
  let applied = 1
  let notices: string[] = []
  if (!bytes) {
    const request = (args.request ?? null) as XlsxSaveRequest | null
    if (!request) invalid('save 需要 bytes 或 request')
    const baseline = new Uint8Array(await readFile(file.path))

    // 先重算，再把 stale 缓存值和新写的公式值一起交给 gateway（顺序不能反）。
    let formulaValues = request.formulaValues ?? []
    try {
      const recalculated = await recalcFormulaValues(baseline, request)
      notices = recalculated.notices
      if (recalculated.formulaValues) formulaValues = recalculated.formulaValues
    } catch (error) {
      notices = [`recalc-failed:${error instanceof Error ? error.message.slice(0, 160) : String(error)}`]
    }

    const plan = await planSaveRequest(baseline, { ...request, formulaValues })
    const engine = await loadWasmEngine(getScope().tempDir)
    if (engine) {
      try {
        bytes = await saveViaWasm(engine, baseline, plan)
      } catch (error) {
        notices.push(`save-archive-failed:${error instanceof Error ? error.message.slice(0, 120) : String(error)}`)
        bytes = (await bytesOfPlan(baseline, plan)).bytes
      }
    } else {
      bytes = (await bytesOfPlan(baseline, plan)).bytes
    }
    applied = Array.isArray(request.edits) ? request.edits.length : 1
  }
  return writeBytes(file, bytes, args, ctx, applied, notices)
}

function checkBytes(bytes: Uint8Array): void {
  if (bytes.byteLength > MAX_FILE_BYTES)
    invalid(`保存产物 ${bytes.byteLength} 字节超过 ${MAX_FILE_BYTES} 上限`)
  assertXlsxPackage(bytes)
}

async function writeBytes(
  file: XlsxFile,
  bytes: Uint8Array,
  args: AnyRecord,
  ctx: PluginContext,
  applied: number,
  notices: string[] = [],
): Promise<unknown> {
  checkBytes(bytes)
  const baseHash = typeof args.baseHash === 'string' ? args.baseHash : ''
  return withLock(file.ref, async () => {
    const current = await hashOfFile(file)
    if (baseHash && baseHash !== current)
      stale(`文件已变化（hash ${baseHash.slice(0, 8)}… → ${current.slice(0, 8)}…），请重新打开`)

    await ctx.fs.write(file.ref, bytes)
    const info = await stat(file.path)
    const hash = sha256Hex(bytes)
    rememberHash(file.path, info.size, info.mtimeMs, hash)
    ctx.changed?.(file.ref, 'content')
    return {
      ref: file.ref,
      hash,
      size: bytes.byteLength,
      at: Date.now(),
      applied,
      ...(notices.length > 0 ? { notices } : {}),
    }
  })
}

/** 恢复副本路径：插件数据目录，按 ref 哈希分文件（一个文件一份草稿）。 */
function recoveryPath(ref: string): string {
  const key = sha256Hex(new TextEncoder().encode(ref))
  return join(getScope().dataDir, 'recovery', `${key}.xlsx`)
}

async function settleOp(op: string, file: XlsxFile, input: AnyRecord, args: AnyRecord): Promise<unknown> {
  const token = String(input.token ?? args.token ?? '')
  if (op === 'detach') {
    detachLease(file.ref, token)
    return { ok: true }
  }

  if (op === 'pending') {
    verifyLease(file.ref, token)
    const id = String(args.id ?? '')
    const request = pendingOf(id)
    if (!request) throw new CommandError('command.not-found', `pending 不存在：${id}`)
    return request
  }

  if (op === 'ack') {
    verifyLease(file.ref, token)
    const id = String(args.id ?? '')
    const ok = args.ok === true
    let payload: Record<string, unknown> = {}
    const raw = args.payload
    if (typeof raw === 'string' && raw) {
      try {
        payload = JSON.parse(raw) as Record<string, unknown>
      } catch {
        invalid('ack.payload 不是合法 JSON')
      }
    } else if (raw && typeof raw === 'object') {
      payload = raw as Record<string, unknown>
    }
    if (!settlePending(id, ok, payload as never))
      throw new CommandError('command.not-found', `pending 不存在：${id}`)
    return { ok: true }
  }

  if (op === 'draft' || op === 'draftClear') {
    verifyLease(file.ref, token)
    const draftPath = recoveryPath(file.ref)
    if (op === 'draftClear') {
      await rm(draftPath, { force: true })
      return { ok: true }
    }
    const provided = asBytes(args.bytes)
    const bytes = provided ?? (await (async () => {
      const request = (args.request ?? null) as XlsxSaveRequest | null
      if (!request) invalid('draft 需要 bytes 或 request')
      const baseline = new Uint8Array(await readFile(file.path))
      return (await bytesOfSaveRequest(baseline, request)).bytes
    })())
    checkBytes(bytes)
    await mkdir(dirname(draftPath), { recursive: true })
    await writeFile(draftPath, bytes)
    const info = await stat(draftPath)
    return { ok: true, at: info.mtimeMs, size: info.size }
  }
  const lease = reportLease(file.ref, {
    token,
    ...(typeof args.hash === 'string' ? { hash: args.hash } : {}),
    ...(typeof args.docRev === 'number' ? { docRev: args.docRev } : {}),
    ...(typeof args.structureEpoch === 'number' ? { structureEpoch: args.structureEpoch } : {}),
    ...(typeof args.dirty === 'boolean' ? { dirty: args.dirty } : {}),
    ...(typeof args.selection === 'string' ? { selection: args.selection } : {}),
  })
  return {
    ok: true,
    hash: lease.hash,
    docRev: lease.docRev,
    structureEpoch: lease.structureEpoch,
    dirty: lease.dirty,
  }
}

export { MAX_CHUNK_BYTES }
