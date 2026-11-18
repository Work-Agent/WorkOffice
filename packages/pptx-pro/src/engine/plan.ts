// ─── engine.plan：命令批 → TxnRequest（执行点 / 无页面路径） ───────────
//
// - 只读工具返回 summary，不产字节。
// - 写工具在无页面时校验 op 名 / 页面能力，并用 runTxn(dryRun) 预检。
// - 有页面时宿主把命令转发给页面代理，本文件不参与。

import { runTxn } from '@genoffice/pptx-ops'
import { guideText, isKnownOp, isPageOnly } from '../tools/catalog'
import { normalizeOps, type PptxOp } from '../shared/ops'
import { DeckSpecError, normalizeSpec } from '../shared/spec'
import { coded, messageOf, type EngineCtx, type PlanOutcome, type PlanRequest, type Rejection } from './abi'
import { getOrOpen, sha256Hex } from './state'
import * as read from './read'

const MAX_COMMANDS = 200

function reject(
  index: number,
  code: string,
  message: string,
  summary?: string,
): PlanOutcome {
  const rejected: Rejection[] = [{ index, code, message }]
  return {
    plan: { kind: 'reject', applied: 0, rejected },
    summary: summary ?? `rejected[${index}] ${code}: ${message}`,
  }
}

export async function plan(req: PlanRequest, ctx: EngineCtx): Promise<PlanOutcome> {
  const params = (req.params ?? {}) as Record<string, unknown>
  const tool = String(req.tool ?? '')

  if (tool === 'pptx.guide' || tool === 'pptx_guide') {
    return {
      plan: { kind: 'noop' },
      summary: guideText({
        ...(typeof params['group'] === 'string' ? { group: params['group'] } : {}),
        ...(typeof params['op'] === 'string' ? { op: params['op'] } : {}),
        ...(req.baseVersion ? { version: req.baseVersion } : {}),
        ...(typeof params['epoch'] === 'number' ? { epoch: params['epoch'] } : {}),
      }),
    }
  }

  if (tool === 'pptx_create' || tool === 'pptx.create') {
    if (!req.missing)
      return reject(0, 'resource.exists', '目标已存在；修改已有演示文稿请用 pptx_apply')
    try {
      const spec = normalizeSpec(params)
      return {
        plan: { kind: 'new', spec, applied: spec.slides.length },
        summary: `create: ${req.ref} slides=${spec.slides.length}`,
      }
    } catch (error) {
      if (error instanceof DeckSpecError)
        return reject(0, 'command.invalid', error.message)
      throw error
    }
  }

  if (req.missing)
    throw coded('resource.not-found', `文件不存在：${req.ref}；新建请用 pptx_create`)

  const bytes = req.bytes ?? (await ctx.readBytes()).bytes
  const version = req.baseVersion ?? sha256Hex(bytes)

  if (tool === 'pptx.read' || tool === 'pptx_read') {
    const opened = (await getOrOpen(req.ref, version, bytes)).opened
    try {
      return {
        plan: { kind: 'noop' },
        summary: read.readMany(opened, { ...params, version }),
      }
    } catch (error) {
      if ((error as { code?: string }).code) throw error
      throw coded('command.invalid', messageOf(error))
    }
  }

  if (tool === 'pptx.apply' || tool === 'pptx_apply') {
    return applyOutcome(req, params, bytes, version)
  }

  throw coded('command.not-found', `不支持的工具：${tool}`)
}

async function applyOutcome(
  req: PlanRequest,
  params: Record<string, unknown>,
  bytes: Uint8Array,
  version: string,
): Promise<PlanOutcome> {
  const raw = Array.isArray(req.commands)
    ? req.commands
    : Array.isArray(params['commands'])
      ? params['commands']
      : null
  if (!raw) return reject(0, 'command.invalid', '缺少 commands 数组')
  if (raw.length > MAX_COMMANDS)
    return reject(0, 'command.too-large', `批上限 ${MAX_COMMANDS} 条（收到 ${raw.length}）`)
  if (raw.length === 0) return reject(0, 'command.invalid', 'commands 为空')

  const normalized = normalizeOps(raw)
  const ops: PptxOp[] = normalized.ops
  for (const [index, op] of ops.entries()) {
    if (!isKnownOp(op.op))
      return reject(
        index,
        'command.not-found',
        `未知 op：${op.op || '(空)'}（用 pptx_guide 取可用 op）`,
      )
    if (isPageOnly(op.op))
      return reject(
        index,
        'command.page-unavailable',
        `op「${op.op}」需要打开演示文稿页面（图片/媒体/剪贴板/3D/旁白/动画时间线）；无页面时请改用无页面可执行的 op。`,
      )
  }

  const isolation = params['isolation'] === 'per_op' ? 'per_op' : 'atomic'
  const dryRun = params['dryRun'] === true

  // 预检：打开基线并 dry-run，避免把明显的 ID / 参数错误留到 exec。
  let opened
  try {
    opened = (await getOrOpen(req.ref, version, bytes)).opened
  } catch (error) {
    throw coded('command.invalid', `PPTX 解析失败：${messageOf(error)}`)
  }

  const precheck = runTxn(opened, { ops, dryRun: true, isolation })
  if (precheck.failures?.length) {
    return {
      plan: {
        kind: 'reject',
        applied: 0,
        rejected: precheck.failures.map((failure) => ({
          index: failure.index,
          code: 'command.invalid',
          message: failure.error,
        })),
      },
      summary: precheck.failures
        .map((failure) => `rejected[${failure.index}] ${failure.error}`)
        .join('\n'),
    }
  }

  return {
    plan: {
      kind: 'txn',
      txn: { ops, isolation, ...(dryRun ? { dryRun: true } : {}) },
      applied: ops.length,
    },
    summary: `plan: ${ops.length} ops · isolation=${isolation} · ${req.baseVersion ? `version=${req.baseVersion.slice(0, 12)}` : 'version=-'}`,
  }
}

/** 只读投影入口（宿主在 `engine.plan` 缺失时才用 read.project）。 */
export async function project(req: PlanRequest, ctx: EngineCtx): Promise<string> {
  return read.project(req, ctx)
}
