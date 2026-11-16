// ─── engine.exec：计划 → DOCX 字节（唯一字节产生者） ────────────────────
//
// 契约 §6.2：exec 只写 `ctx.tempDir`，目标路径由权威层 rename；返回体不带
// `version`（由宿主算出）。docx 没有常驻写侧模型，因此没有 modelToken。
//
// 两种计划来源，一个字节出口：
//   1. `save`——执行点解析基线字节、按块切片拼接 `word/document.xml`，
//      其余 zip 条目逐字节拷贝（无页面路径；未改块保真）。
//   2. `bytes`——页面（完整编辑器）序列化出的整份字节。exec 仍是唯一字节
//      产生者：校验包结构、体积预算，然后原样输出。

import { readFileSync } from 'node:fs'
import { strToU8 } from 'fflate'
import {
  DOCUMENT_PART,
  DocxParseError,
  parseDocumentXml,
  parseDocxBytes,
} from '../shared/document'
import { PlanShapeError, assertPlanShape, baselineSlices } from '../shared/plan'
import { buildManifest, spliceBody } from '../shared/sync'
import { runtimeReport } from '../shared/runtime'
import { sha256Hex } from '../shared/sha256'
import { coded, messageOf, type EngineCtx, type ExecOutcome, type ExecRequest } from './abi'
import { isDocxPackage } from './read'
import { ZipShapeError, rewriteZipPart } from './zip'

export async function exec(req: ExecRequest, ctx: EngineCtx): Promise<ExecOutcome> {
  if (req.command === 'undo' || req.command === 'redo')
    throw coded(
      'command.unsupported',
      'docx 的撤销是页面模型历史（I16）；无页面路径不提供 undo',
    )

  const plan = req.plan
  if (!plan || typeof plan !== 'object')
    throw coded('command.invalid', '计划不是对象')

  const kind = (plan as { kind?: unknown }).kind

  if (kind === 'noop')
    return {
      result: {
        applied: Number((plan as { applied?: number }).applied ?? 0),
        sync: { kind: 'version' },
      },
    }

  if (kind === 'bytes')
    return execBytes(req, plan as { kind: 'bytes'; bytes?: unknown; applied?: number }, ctx)

  if (kind === 'new')
    return execNew(
      req,
      plan as { kind: 'new'; documentXml?: unknown; applied?: number },
      ctx,
    )

  // ─── 块级切片计划（无页面路径） ─────────────────────────────────
  const baselineBytes = req.bytes ?? (await ctx.readBytes()).bytes
  let parsed
  try {
    parsed = parseDocxBytes(baselineBytes)
  } catch (error) {
    if (error instanceof DocxParseError) throw coded('command.invalid', messageOf(error))
    throw error
  }
  const slices = baselineSlices(parsed)

  let checked
  try {
    checked = assertPlanShape(plan, slices)
  } catch (error) {
    if (error instanceof PlanShapeError) throw coded('command.invalid', error.message)
    throw error
  }

  if (checked.kind === 'reject')
    return {
      result: {
        applied: 0,
        rejected: checked.rejected ?? [],
        ...(checked.warnings ? { warnings: checked.warnings } : {}),
      },
    }

  if (checked.kind === 'check')
    return {
      result: { applied: 0, ...(checked.audit !== undefined ? { audit: checked.audit } : {}) },
    }

  // 拼接 → 重新解析校验（良构 / sectPr 保序），再重写 zip 条目。
  let documentXml: string
  let nextParsed
  try {
    documentXml = spliceBody(parsed, checked)
    nextParsed = parseDocumentXml(documentXml)
  } catch (error) {
    if (error instanceof DocxParseError)
      throw coded('command.invalid', `计划产出的 document.xml 非法：${messageOf(error)}`)
    throw error
  }
  const manifest = buildManifest(parsed, checked)
  const warnings = [...(checked.warnings ?? [])]
  if (nextParsed.blocks.every((block) => block.hidden)) warnings.push('document-empty')

  if (req.dryRun || checked.dryRun)
    return {
      result: {
        applied: checked.applied ?? checked.changedCount,
        ...(warnings.length > 0 ? { warnings } : {}),
        audit: { changedCount: checked.changedCount, deletedCount: checked.deletedCount },
        sync: { kind: 'manifest', blocks: manifest.blocks },
      },
    }

  let bytes: Uint8Array
  try {
    bytes = rewriteZipPart(baselineBytes, DOCUMENT_PART, strToU8(documentXml))
  } catch (error) {
    if (error instanceof ZipShapeError) throw coded('command.unsupported', messageOf(error))
    throw error
  }
  guardOutput(bytes, ctx)

  return {
    bytes,
    result: {
      applied: checked.applied ?? checked.changedCount,
      ...(warnings.length > 0 ? { warnings } : {}),
      sync: {
        kind: 'manifest',
        blocks: manifest.blocks,
        ...(manifest.hidden ? { hidden: manifest.hidden } : {}),
      },
      audit: {
        changedCount: checked.changedCount,
        deletedCount: checked.deletedCount,
        resultDocHash: sha256Hex(bytes),
      },
    },
  }
}

/** 页面序列化出的整份字节：校验后原样输出（exec 仍是唯一字节产生者）。 */
function execBytes(
  req: ExecRequest,
  plan: { bytes?: unknown; applied?: number },
  ctx: EngineCtx,
): ExecOutcome {
  const value = plan.bytes
  const bytes =
    value instanceof Uint8Array
      ? value
      : Array.isArray(value)
        ? new Uint8Array(value as number[])
        : null
  if (!bytes) throw coded('command.invalid', 'bytes 计划缺少字节')
  if (bytes.byteLength === 0) throw coded('command.invalid', '页面提交了空字节')
  if (!isDocxPackage(bytes))
    throw coded('command.invalid', '页面提交的不是有效的 DOCX 包')

  if (req.dryRun)
    return { result: { applied: plan.applied ?? 1, sync: { kind: 'version' } } }

  guardOutput(bytes, ctx)
  return {
    bytes,
    result: {
      applied: plan.applied ?? 1,
      sync: { kind: 'version' },
      audit: { resultDocHash: sha256Hex(bytes), source: 'page' },
    },
  }
}

let SEED: Uint8Array | null = null

/** 插件资产：dist/seed.docx（build 时从 assets/seed.docx 复制）。 */
function seedBytes(): Uint8Array {
  SEED ??= new Uint8Array(readFileSync(new URL('../seed.docx', import.meta.url)))
  return SEED
}

/** 新建：document.xml 已由 plan 生成，这里只把它换进 seed 包。 */
function execNew(
  req: ExecRequest,
  plan: { documentXml?: unknown; applied?: number },
  ctx: EngineCtx,
): ExecOutcome {
  const xml = plan.documentXml
  if (typeof xml !== 'string' || xml.length === 0)
    throw coded('command.invalid', '新建计划缺少 document.xml')

  let parsed
  try {
    parsed = parseDocumentXml(xml)
  } catch (error) {
    if (error instanceof DocxParseError)
      throw coded('command.invalid', `新建文档 XML 非法：${messageOf(error)}`)
    throw error
  }

  const warnings: string[] = []
  if (parsed.blocks.every((block) => block.hidden)) warnings.push('document-empty')

  if (req.dryRun)
    return {
      result: {
        applied: plan.applied ?? 0,
        ...(warnings.length > 0 ? { warnings } : {}),
        sync: { kind: 'version' },
      },
    }

  let bytes: Uint8Array
  try {
    bytes = rewriteZipPart(seedBytes(), DOCUMENT_PART, strToU8(xml))
  } catch (error) {
    if (error instanceof ZipShapeError)
      throw coded('command.unsupported', messageOf(error))
    throw error
  }

  if (!isDocxPackage(bytes))
    throw coded('command.invalid', '新建产出不是有效的 DOCX 包')

  guardOutput(bytes, ctx)

  return {
    bytes,
    result: {
      applied: plan.applied ?? 0,
      ...(warnings.length > 0 ? { warnings } : {}),
      sync: { kind: 'version' },
      audit: { resultDocHash: sha256Hex(bytes), source: 'create' },
    },
  }
}

function guardOutput(bytes: Uint8Array, ctx: EngineCtx): void {
  if (bytes.byteLength > ctx.budget.maxBytes)
    throw coded(
      'command.too-large',
      `产出超过执行点上限：${bytes.byteLength} > ${ctx.budget.maxBytes}`,
    )
  if (bytes.byteLength > runtimeReport().limits.maxCommitBytes)
    throw coded('command.too-large', '产出超过插件提交上限')
}

// docx 无常驻写侧模型：两阶段提交为空实现（宿主只在 modelToken 存在时调用）。
export async function commitModel(): Promise<void> {}
export async function abortModel(): Promise<void> {}
