// ─── engine.exec：计划 → PPTX 字节（唯一字节产生者） ──────────────────
//
// - txn：runTxn(cached) → savePptxToFile(ctx.tempDir) → artifact + modelToken
// - new：seed.pptx → runTxn → 同上
// - undo / redo：恢复快照 → 同上
// - commitModel / abortModel：两阶段模型提交
//
// 当前宿主 artifact 仍会被整份读入并受 maxCommitBytes 约束，所以 P0 上限 64 MB。

import { createHash } from 'node:crypto'
import { strToU8 } from 'fflate'
import { readFileSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import {
  commitSaved,
  createBlankPptx,
  openPptx,
  savePptxToFile,
  type OpenedPptx,
} from '@genoffice/pptx-engine'
import { runTxn } from '@genoffice/pptx-ops'
import { buildCreateOps } from './create'
import { coded, messageOf, type EngineCtx, type ExecOutcome, type ExecRequest, type PptxOp, type PptxProPlan } from './abi'
import * as state from './state'
import { normalizeSpec, type DeckSpec } from '../shared/spec'
import { MAX_COMMIT_BYTES } from '../shared/runtime'

let SEED: Uint8Array | null = null

function seedBytes(): Uint8Array {
  SEED ??= new Uint8Array(readFileSync(new URL('../seed.pptx', import.meta.url)))
  return SEED
}

function guardSize(bytes: number, ctx: EngineCtx): void {
  const limit = Math.min(ctx.budget.maxBytes || MAX_COMMIT_BYTES, MAX_COMMIT_BYTES)
  if (bytes > limit)
    throw coded('command.too-large', `产出超过提交上限：${bytes} > ${limit}`)
}

function assertSafePartPath(path: string): void {
  if (!path || path.startsWith('/') || path.includes('\\') || path.split('/').includes('..'))
    throw coded('command.invalid', `非法 part 路径：${path}`)
}

function assertSafeSlidePath(path: string): void {
  assertSafePartPath(path)
  if (!/^ppt\/slides\/slide\d+\.xml$/.test(path))
    throw coded('command.invalid', `非法 slide 路径：${path}`)
}

async function artifactOf(
  opened: OpenedPptx,
  ctx: EngineCtx,
): Promise<{ artifact: NonNullable<ExecOutcome['artifact']>; sha256: string }> {
  const tempPath = join(
    ctx.tempDir || tmpdir(),
    `pptx-pro-${Date.now()}-${Math.random().toString(16).slice(2)}.pptx`,
  )
  await savePptxToFile(opened, tempPath)
  const bytes = new Uint8Array(await readFile(tempPath))
  guardSize(bytes.byteLength, ctx)
  const sha256 = createHash('sha256').update(bytes).digest('hex')
  return { artifact: { tempPath, size: bytes.byteLength, sha256 }, sha256 }
}

function syncVersion(): { kind: 'version' } {
  return { kind: 'version' }
}

export async function exec(req: ExecRequest, ctx: EngineCtx): Promise<ExecOutcome> {
  const plan = req.plan as PptxProPlan | undefined
  if (!plan || typeof plan !== 'object')
    throw coded('command.invalid', '计划不是对象')

  if (req.command === 'undo' || req.command === 'redo') {
    const bytes = req.bytes ?? (await ctx.readBytes()).bytes
    const entry = await state.getOrOpen(req.ref, req.baseVersion, bytes)
    const token = state.beginApply(entry, req.command)
    const changed = req.command === 'undo' ? state.undo(entry) : state.redo(entry)
    if (!changed) {
      state.abort(token.token)
      return {
        result: {
          applied: 0,
          warnings: [`没有可${req.command === 'undo' ? '撤销' : '重做'}的事务`],
          sync: syncVersion(),
        },
      }
    }
    if (req.dryRun) {
      state.abort(token.token)
      return { result: { applied: 1, sync: syncVersion() } }
    }
    const saved = await artifactOf(entry.opened, ctx)
    state.markArtifact(token.token, saved.sha256)
    return {
      artifact: saved.artifact,
      modelToken: token.token,
      result: { applied: 1, sync: syncVersion() },
    }
  }

  if (plan.kind === 'noop')
    return { result: { applied: plan.applied ?? 0, sync: { kind: 'none' } } }

  if (plan.kind === 'reject')
    return {
      result: {
        applied: 0,
        rejected: plan.rejected,
        ...(plan.warnings ? { warnings: plan.warnings } : {}),
        sync: { kind: 'none' },
      },
    }

  if (plan.kind === 'new') return execNew(req, plan.spec, ctx)

  if (plan.kind === 'pageSave') {
    const bytes = req.bytes ?? (await ctx.readBytes()).bytes
    const opened = await openPptx(bytes)
    const entries = opened.archive.entries
    for (const part of plan.parts) {
      assertSafePartPath(part.path)
      entries.set(part.path, part.bytes)
    }
    for (const path of plan.deletedParts) {
      assertSafePartPath(path)
      entries.delete(path)
    }
    for (const slide of plan.slides) {
      assertSafeSlidePath(slide.path)
      entries.set(slide.path, strToU8(slide.xml))
    }
    if (req.dryRun)
      return {
        result: {
          applied: 0,
          audit: { slides: plan.slides.length, parts: plan.parts.length },
          sync: { kind: 'none' },
        },
      }
    const saved = await artifactOf(opened, ctx)
    state.evictRef(req.ref)
    return {
      artifact: saved.artifact,
      result: {
        applied: plan.changedCount || plan.slides.length + plan.parts.length + plan.deletedParts.length,
        sync: syncVersion(),
      },
    }
  }

  if (plan.kind !== 'txn')
    throw coded('command.invalid', `未知计划类型：${String((plan as { kind?: unknown }).kind)}`)

  const bytes = req.bytes ?? (await ctx.readBytes()).bytes
  const entry = await state.getOrOpen(req.ref, req.baseVersion, bytes)
  const txn = plan.txn
  const batches: PptxOp[][] =
    Array.isArray(txn?.batches) && txn.batches.length > 0
      ? txn.batches
      : Array.isArray(txn?.ops) && txn.ops.length > 0
        ? [txn.ops]
        : []
  if (batches.length === 0)
    return {
      result: {
        applied: 0,
        rejected: [{ index: 0, code: 'command.invalid', message: '空命令批' }],
        sync: { kind: 'none' },
      },
    }

  if (req.dryRun || txn.dryRun) {
    // 在临时模型上按批真实应用，避免“先建页、再编辑新页”在 dry-run 里
    // 被 runTxn 的整批预校验误拒。临时模型不落盘、不进缓存。
    const temp = await openPptx(bytes)
    const rejected: Array<{ index: number; code: string; message: string }> = []
    const lines: string[] = []
    for (const [batchIndex, batch] of batches.entries()) {
      const result = runTxn(temp, {
        ops: batch,
        ...(txn.isolation ? { isolation: txn.isolation } : {}),
      })
      if (!result.applied) {
        for (const failure of result.failures ?? [])
          rejected.push({
            index: failure.index,
            code: 'command.invalid',
            message: `batch ${batchIndex}: ${failure.error}`,
          })
        break
      }
      lines.push(...(result.records ?? []).map((_, i) => `batch ${batchIndex} op ${i}`))
    }
    return {
      result: {
        applied: 0,
        ...(rejected.length ? { rejected } : {}),
        audit: lines,
        sync: { kind: 'none' },
      },
    }
  }

  const token = state.beginApply(entry)
  let applied = 0
  for (const [batchIndex, batch] of batches.entries()) {
    const result = runTxn(entry.opened, {
      ops: batch,
      ...(txn.isolation ? { isolation: txn.isolation } : {}),
    })
    if (!result.applied) {
      state.abort(token.token)
      return {
        result: {
          applied: 0,
          rejected: (result.failures ?? []).map((failure) => ({
            index: failure.index,
            code: 'command.invalid',
            message: `batch ${batchIndex}: ${failure.error}`,
          })),
          sync: { kind: 'none' },
        },
      }
    }
    applied += result.records?.length ?? batch.length
  }

  const saved = await artifactOf(entry.opened, ctx)
  state.markArtifact(token.token, saved.sha256)
  return {
    artifact: saved.artifact,
    modelToken: token.token,
    result: {
      applied,
      sync: syncVersion(),
    },
  }
}

async function execNew(req: ExecRequest, specInput: unknown, ctx: EngineCtx): Promise<ExecOutcome> {
  const spec = specInput as DeckSpec
  const seed = seedBytes()
  const opened = await openPptx(seed)
  const { setup, content } = buildCreateOps(opened, spec)
  const total = setup.length + content.length

  if (req.dryRun) {
    // 在临时模型上真实跑 setup（不落盘），再 dry-run content；这样能校验
    // “先建页、再往新页写内容”的坐标。
    if (setup.length > 0) {
      const setupResult = runTxn(opened, { ops: setup, dryRun: true })
      if (setupResult.failures?.length)
        return {
          result: {
            applied: 0,
            rejected: setupResult.failures.map((failure) => ({
              index: failure.index,
              code: 'command.invalid',
              message: failure.error,
            })),
            sync: { kind: 'none' },
          },
        }
      // dry-run 不能产生新页；用真实 apply 到临时模型后继续校验 content。
      const applied = runTxn(opened, { ops: setup })
      if (!applied.applied)
        return {
          result: {
            applied: 0,
            rejected: (applied.failures ?? []).map((failure) => ({
              index: failure.index,
              code: 'command.invalid',
              message: failure.error,
            })),
            sync: { kind: 'none' },
          },
        }
    }
    const checked = content.length > 0 ? runTxn(opened, { ops: content, dryRun: true }) : null
    return {
      result: {
        applied: 0,
        ...(checked?.failures?.length
          ? {
              rejected: checked.failures.map((failure) => ({
                index: failure.index,
                code: 'command.invalid',
                message: failure.error,
              })),
            }
          : {}),
        audit: checked?.plan,
        sync: { kind: 'none' },
      },
    }
  }

  for (const [label, ops] of [
    ['setup', setup],
    ['content', content],
  ] as const) {
    if (ops.length === 0) continue
    const result = runTxn(opened, { ops })
    if (!result.applied) {
      throw coded(
        'command.invalid',
        `新建 ${label} 事务失败：${(result.failures ?? []).map((f) => f.error).join('; ')}`,
      )
    }
  }
  const saved = await artifactOf(opened, ctx)
  const token = `pptx-pro-new-${Date.now()}-${Math.random().toString(16).slice(2)}`
  // 新建没有缓存 entry；token 交给 commitModel 时只做空操作。下一次请求走冷路径。
  return {
    artifact: saved.artifact,
    modelToken: token,
    result: { applied: total, sync: syncVersion() },
  }
}

/** 宿主写成功后调用。 */
export async function commitModel(token: string): Promise<void> {
  state.commit(token)
}

/** 宿主写失败 / 取消时调用。 */
export async function abortModel(token: string): Promise<void> {
  state.abort(token)
}

/** 只为测试 / 诊断。 */
export function resetState(): void {
  state.clear()
}

export { commitSaved }
