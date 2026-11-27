// ─── xlsx.apply：一批命令（≤200） ────────────────────────────────
//
// 有页面（租约在）→ 入队给页面：页面校验 op 级令牌 → Univer 事务 → 保存 → ack。
// 无页面 → 引擎自己展开成保存请求 → gateway 产字节 → 加锁写入（baseHash CAS）。
// `objects` 组（图表/透视表/条件格式…）只能走页面。

import { readFile } from 'node:fs/promises'

import {
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'

import type { PluginContext } from '@work-studio/sdk'

import { groupOf, isHeadlessOp, specOf } from '../../tools/catalog'
import { HeadlessOpError, buildSaveRequest } from '../../shared/headless'
import { bytesOfPlan, planSaveRequest } from '../gateway'
import { saveViaWasm } from '../archive'
import { loadWasmEngine } from '../wasm'
import { getScope } from '../scope'
import { CommandError, invalid, stale } from '../errors'
import { hashOfFile, rememberHash, sha256Hex } from '../hash'
import { withLock } from '../lock'
import { recalcFormulaValues } from '../recalc'
import { assertXlsxFile, MAX_FILE_BYTES, type XlsxFile } from '../ref'
import { loadView } from '../read'
import { enqueue, leaseOf } from '../state'
import { stat } from 'node:fs/promises'

const MAX_COMMANDS = 200

type AnyRecord = Record<string, unknown>

function commandsOf(input: AnyRecord): unknown[] {
  const raw =
    input['commands'] ?? input['ops'] ?? input['commandsJson']
  if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw) as unknown
      if (!Array.isArray(parsed)) invalid('commands 必须是数组')
      return parsed
    } catch (error) {
      if (error instanceof CommandError) throw error
      invalid(`commands 不是合法 JSON：${error instanceof Error ? error.message : String(error)}`)
    }
  }
  if (!Array.isArray(raw)) invalid('缺少 commands 数组')
  return raw
}

/** 校验并把 `sheet` 归一到 `sheetId`（引擎侧不信任调用方）。 */
function normalize(raw: unknown[]): AnyRecord[] {
  const out: AnyRecord[] = []
  for (const [index, command] of raw.entries()) {
    if (!command || typeof command !== 'object') invalid(`commands[${index}] 不是对象`)
    const record = { ...(command as AnyRecord) }
    const op = String(record['op'] ?? '')
    if (!op) invalid(`commands[${index}] 缺少 op`)
    if (!specOf(op)) invalid(`commands[${index}] 未知 op：${op}（用 xlsx.guide 取可用 op）`)
    if (typeof record['sheet'] === 'string' && record['sheetId'] === undefined)
      record['sheetId'] = record['sheet']
    delete record['sheet']
    const parsed = workbookOperationSchema.safeParse(record)
    if (!parsed.success)
      invalid(
        `commands[${index}] 非法：` +
          parsed.error.issues.map((issue) => `${issue.path.join('.')}: ${issue.message}`).join('; '),
      )
    out.push(parsed.data as AnyRecord)
  }
  return out
}

export interface ApplyResult {
  requested: number
  applied: number
  dryRun?: boolean
  wouldApply?: number
  mutated?: boolean
  receipt?: { ref: string; hash: string; size: number; at: number; applied?: number }
  /** 写完后页面的新鲜度（下一条命令直接可用，不必重新 read）。 */
  docRev?: number
  structureEpoch?: number
  notices?: string[]
}

export async function apply(input: AnyRecord, ctx: PluginContext): Promise<ApplyResult> {
  const ref = String(input['ref'] ?? '')
  if (!ref) invalid('apply 缺少 ref')
  const raw = commandsOf(input)
  if (raw.length === 0) invalid('commands 不能为空')
  if (raw.length > MAX_COMMANDS) invalid(`批上限 ${MAX_COMMANDS} 条（收到 ${raw.length}）`)

  const file = await assertXlsxFile(ref)
  const normalized = normalize(raw)
  const dryRun = input['dryRun'] === true
  const baseHash = typeof input['baseHash'] === 'string' ? input['baseHash'] : ''

  const lease = leaseOf(ref)
  if (lease) {
    if (baseHash && baseHash !== lease.hash)
      stale(`文件已变化（hash ${baseHash.slice(0, 8)}… → ${lease.hash.slice(0, 8)}…），请重新读取`)
    const payload = await enqueue(ref, {
      type: 'apply',
      commands: normalized,
      ...(typeof input['docRev'] === 'number' ? { docRev: input['docRev'] } : {}),
      ...(typeof input['structureEpoch'] === 'number'
        ? { structureEpoch: input['structureEpoch'] }
        : {}),
      ...(dryRun ? { dryRun: true } : {}),
      ...(baseHash ? { baseHash } : {}),
    })
    if (payload.error)
      throw new CommandError(payload.error.code, payload.error.message)
    // 页面在 ack 之前已经 report 过，这里的令牌就是写完后的值。
    const after = leaseOf(ref)
    return {
      requested: normalized.length,
      applied: payload.applied ?? 0,
      ...(payload.dryRun ? { dryRun: true } : {}),
      ...(payload.mutated === false ? { mutated: false } : {}),
      ...(payload.receipt ? { receipt: payload.receipt } : {}),
      ...(after
        ? { docRev: after.docRev, structureEpoch: after.structureEpoch }
        : {}),
    }
  }

  // 无页面：objects 组必须打开文件。
  for (const [index, command] of normalized.entries()) {
    const op = String(command['op'])
    if (!isHeadlessOp(op))
      throw new CommandError(
        'command.page-unavailable',
        `commands[${index}] 的 op「${op}」属于 ${groupOf(op)} 组，需要打开工作簿页面；` +
          '没有页面时只能用 cells / structure / sheets 组的 op。',
      )
  }

  return headless(file, normalized as unknown as WorkbookOperation[], ctx, { dryRun, baseHash, notices: [] })
}

async function headless(
  file: XlsxFile,
  operations: WorkbookOperation[],
  ctx: PluginContext,
  options: { dryRun: boolean; baseHash: string; notices: string[] },
): Promise<ApplyResult> {
  return withLock(file.ref, async () => {
    const baseline = new Uint8Array(await readFile(file.path))
    const current = sha256Hex(baseline)
    if (options.baseHash && options.baseHash !== current)
      stale(`文件已变化（hash ${options.baseHash.slice(0, 8)}… → ${current.slice(0, 8)}…），请重新读取`)

    const entry = await loadView(file)
    let built
    try {
      built = buildSaveRequest(entry.view, operations)
    } catch (error) {
      if (error instanceof HeadlessOpError) throw new CommandError(error.code, error.message)
      throw error
    }

    if (options.dryRun)
      return {
        requested: operations.length,
        applied: 0,
        dryRun: true,
        wouldApply: built.changedCount,
        notices: built.warnings,
      }

    let formulaValues = built.request.formulaValues ?? []
    try {
      const recalculated = await recalcFormulaValues(baseline, built.request)
      options.notices.push(...recalculated.notices)
      if (recalculated.formulaValues) formulaValues = recalculated.formulaValues
    } catch (error) {
      options.notices.push(
        `recalc-failed:${error instanceof Error ? error.message.slice(0, 160) : String(error)}`,
      )
    }

    const plan = await planSaveRequest(baseline, { ...built.request, formulaValues })
    const engine = await loadWasmEngine(getScope().tempDir)
    let bytes: Uint8Array
    if (engine) {
      try {
        bytes = await saveViaWasm(engine, baseline, plan)
      } catch (error) {
        options.notices.push(
          `save-archive-failed:${error instanceof Error ? error.message.slice(0, 120) : String(error)}`,
        )
        bytes = (await bytesOfPlan(baseline, plan)).bytes
      }
    } else {
      bytes = (await bytesOfPlan(baseline, plan)).bytes
    }
    if (bytes.byteLength > MAX_FILE_BYTES)
      throw new CommandError('command.too-large', `保存产物 ${bytes.byteLength} 字节超过上限`)
    await ctx.fs.write(file.ref, bytes)
    const info = await stat(file.path)
    const hash = sha256Hex(bytes)
    rememberHash(file.path, info.size, info.mtimeMs, hash)
    ctx.changed?.(file.ref, 'content')
    return {
      requested: operations.length,
      applied: operations.length,
      receipt: { ref: file.ref, hash, size: bytes.byteLength, at: Date.now(), applied: operations.length },
      ...(options.notices.length > 0 || built.warnings.length > 0
        ? { notices: [...built.warnings, ...options.notices] }
        : {}),
    }
  })
}
