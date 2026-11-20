// ─── engine.plan：命令批 → 计划（执行点 / 无页面路径） ─────────────────
//
// 契约 §6.2 / §7.2：
// - 只读工具（readOnly）永远进执行点：返回 `summary`（模型能读到的文本）。
// - 写工具：有页面时宿主转发页面代理（本文件不参与）；无页面时走到这里。
//   `docx_apply` 只有 blocks 组能在这里执行（XML 切片）；editor / actions 组
//   明确回 `command.page-unavailable` —— 契约禁止回退 headless 写第二份模型。
//
// 页面与执行点共用 `applyCommands` + `planFromModel`（页面把同样的块命令映射
// 到 PM 事务），两条路径的命令语义来自同一份目录。

import {
  DocxParseError,
  modelFromParsed,
  parseDocxBytes,
} from '../shared/document'
import { applyCommands } from '../shared/ops'
import {
  baselineSlices as slicesOf,
  checkPlan,
  planFromModel,
  rejectedPlan,
} from '../shared/plan'
import { runtimeReport } from '../shared/runtime'
import { CommandRejected, type DocxCommand } from '../shared/types'
import { buildCommandGuide, groupOf } from '../tools/catalog'
import { coded, messageOf, type EngineCtx, type PlanOutcome, type PlanRequest } from './abi'
import { markdownToDocumentXml } from './create'
import * as read from './read'

const MAX_COMMANDS = 200

/** 旧版工具名（`--legacy-tools` 构建 + 兼容期）→ 字节侧实现。 */
const LEGACY_READERS: Record<
  string,
  (bytes: Uint8Array, params: Record<string, unknown>) => string
> = {
  get_document_context: (bytes, params) => contextOf(bytes, params),
  read_blocks: (bytes, params) => read.readBlocks(read.parse(bytes), params),
  read_comments: (bytes) => read.readComments(bytes),
  read_notes: (bytes) => read.readNotes(bytes),
  read_revisions: (bytes) => read.readRevisions(bytes, read.parse(bytes)),
  list_styles: (bytes) => read.readStyles(bytes),
  docx_check: (bytes) => read.audit(read.parse(bytes)),
}

const READ_WHAT = read.READ_WHAT

export async function plan(req: PlanRequest, ctx: EngineCtx): Promise<PlanOutcome> {
  const params = (req.params ?? {}) as Record<string, unknown>
  const tool = String(req.tool ?? '')

  if (tool === 'docx.guide' || tool === 'docx_guide') return guideOutcome(req, params)

  // 新建：目标必须不存在；Markdown → document.xml，exec 再装进 seed.docx。
  if (tool === 'docx_create' || tool === 'docx.create') {
    if (!req.missing)
      throw coded('command.invalid', '目标已存在；修改已有文档请用 docx_apply')

    const content = String(params['content'] ?? '')
    if (content.length > 2_000_000)
      throw coded('command.too-large', '新建文档内容超过 2,000,000 字符')

    return {
      plan: {
        kind: 'new',
        documentXml: markdownToDocumentXml(content),
        applied: 1,
      },
      summary: `create: ${req.ref}`,
    }
  }

  // 其他写工具遇到缺失：由插件明确拒绝，不猜测要不要新建。
  if (req.missing)
    throw coded(
      'resource.not-found',
      `文件不存在：${req.ref}；新建请用 docx_create`,
    )

  const bytes = req.bytes ?? (await ctx.readBytes()).bytes

  if (tool === 'docx.read' || tool === 'docx_read') return readOutcome(bytes, params)

  const legacyReader = LEGACY_READERS[tool]
  if (legacyReader) {
    try {
      return { plan: { kind: 'noop' }, summary: legacyReader(bytes, params) }
    } catch (error) {
      if (error instanceof DocxParseError) throw coded('command.invalid', messageOf(error))
      throw error
    }
  }

  if (tool === 'docx.apply' || tool === 'docx_apply')
    return applyOutcome(req, params, bytes)

  // 兼容期：旧版 docx_edit 与 35 个逐工具声明。
  if (tool === 'docx.edit' || tool === 'docx_edit') return blockPlan(req, params, bytes)

  throw coded(
    'command.page-unavailable',
    `${tool} 需要打开文档页面（编辑器内工具）。没有页面时用 docx_apply 的 blocks 组命令，或先让用户打开文档。`,
  )
}

// ─── docx_read ─────────────────────────────────────────────────────────

function readOutcome(bytes: Uint8Array, params: Record<string, unknown>): PlanOutcome {
  try {
    return text(read.readMany(bytes, params))
  } catch (error) {
    if (error instanceof DocxParseError) throw coded('command.invalid', messageOf(error))
    throw error
  }
}

function text(summary: string): PlanOutcome {
  return { plan: { kind: 'noop' }, summary }
}

function contextOf(bytes: Uint8Array, params: Record<string, unknown>): string {
  return read.documentContext(read.parse(bytes), {
    ...(typeof params['name'] === 'string' ? { name: params['name'] } : {}),
    ...(typeof params['version'] === 'string' ? { version: params['version'] } : {}),
    ...(typeof params['epoch'] === 'number' ? { epoch: params['epoch'] } : {}),
  })
}

// ─── docx_apply ────────────────────────────────────────────────────────

function applyOutcome(
  req: PlanRequest,
  params: Record<string, unknown>,
  bytes: Uint8Array,
): PlanOutcome {
  const rawCommands = Array.isArray(req.commands)
    ? req.commands
    : Array.isArray(params['commands'])
      ? params['commands']
      : null
  if (!rawCommands) throw coded('command.invalid', '缺少 commands 数组')
  if (rawCommands.length > MAX_COMMANDS)
    return {
      plan: rejectedPlan(
        new CommandRejected(
          0,
          'command.too-large',
          `批上限 ${MAX_COMMANDS} 条（收到 ${rawCommands.length}）`,
        ),
        null,
      ),
      summary: `rejected: 批上限 ${MAX_COMMANDS} 条`,
    }

  // 无页面：只有 blocks 组能在这里执行。
  for (const [index, command] of rawCommands.entries()) {
    const op = String((command as { op?: unknown } | null)?.op ?? '')
    const group = groupOf(op)
    if (!group)
      return {
        plan: rejectedPlan(
          new CommandRejected(
            index,
            'command.not-found',
            `未知 op：${op || '(空)'}（用 docx_guide 取可用 op）`,
          ),
          null,
        ),
        summary: `rejected[${index}] command.not-found: 未知 op「${op || '(空)'}」`,
      }
    if (group !== 'blocks')
      throw coded(
        'command.page-unavailable',
        `commands[${index}] 的 op「${op}」属于 ${group} 组，需要打开文档页面；` +
          '没有页面时只能用 blocks 组（setText / insertParagraph / setHeadingLevel / setTextStyle / deleteBlock / moveBlock / findReplace）。',
      )
  }

  return blockPlan(req, params, bytes)
}

// ─── blocks 组（无页面路径：XML 切片） ──────────────────────────────────

function blockPlan(
  req: PlanRequest,
  params: Record<string, unknown>,
  bytes: Uint8Array,
): PlanOutcome {
  const rawCommands = Array.isArray(req.commands)
    ? req.commands
    : Array.isArray(params['commands'])
      ? params['commands']
      : null
  if (!rawCommands) throw coded('command.invalid', '缺少 commands 数组')

  let parsed
  try {
    parsed = parseDocxBytes(bytes)
  } catch (error) {
    if (error instanceof DocxParseError) throw coded('command.invalid', messageOf(error))
    throw error
  }

  if (rawCommands.length > MAX_COMMANDS)
    return {
      plan: rejectedPlan(
        new CommandRejected(
          0,
          'command.too-large',
          `批上限 ${MAX_COMMANDS} 条（收到 ${rawCommands.length}）`,
        ),
        parsed,
      ),
      summary: `rejected: 批上限 ${MAX_COMMANDS} 条`,
    }

  const slices = slicesOf(parsed)
  try {
    const commands = rawCommands as DocxCommand[]
    const checks = commands.filter((command) => command.op === 'check')
    if (checks.length > 0) {
      if (commands.length !== 1)
        return {
          plan: rejectedPlan(
            new CommandRejected(0, 'command.invalid', 'check 必须单独成批（不与其他命令混排）'),
            parsed,
          ),
          summary: 'rejected: check 必须单独成批',
        }
      const checked = checkPlan(parsed)
      if (params['dryRun'] === true) checked.dryRun = true
      return { plan: checked, summary: read.audit(parsed) }
    }
    if (commands.length === 0) return { plan: checkPlan(parsed), summary: read.audit(parsed) }

    const model = modelFromParsed(parsed)
    const result = applyCommands(model, commands)
    const next = planFromModel(result.model, slices, result.applied)
    if (params['dryRun'] === true) next.dryRun = true
    return {
      plan: next,
      summary: `plan: ${result.applied} 条命令 · 改 ${next.changedCount} 块 · 删 ${next.deletedCount} 块 · ${shortVersion(req.baseVersion)}`,
    }
  } catch (error) {
    if (error instanceof CommandRejected) {
      return {
        plan: rejectedPlan(error, parsed),
        summary: `rejected[${error.index}] ${error.code}: ${error.message}`,
      }
    }
    throw error
  }
}

// ─── guide ─────────────────────────────────────────────────────────────

function guideOutcome(req: PlanRequest, params: Record<string, unknown>): PlanOutcome {
  const guide = buildCommandGuide({
    version: req.baseVersion,
    group: typeof params['group'] === 'string' ? params['group'] : undefined,
    op: typeof params['op'] === 'string' ? params['op'] : undefined,
    epoch: typeof params['epoch'] === 'number' ? params['epoch'] : undefined,
  })
  const report = runtimeReport()
  return {
    plan: { kind: 'noop' },
    summary: [
      guide.summary,
      '',
      `runtime: write=${report.write} execution=${report.execution} undo=${report.features.undo} sync=${report.features.tokens?.baselineSync}`,
    ].join('\n'),
  }
}

function shortVersion(version?: string): string {
  if (!version) return 'version=-'
  return `version=${version.slice(0, 12)}`
}

/** 只读投影入口（宿主在 `engine.plan` 缺失时才用 read.project）。 */
export async function project(req: PlanRequest, ctx: EngineCtx): Promise<string> {
  const bytes = req.bytes ?? (await ctx.readBytes()).bytes
  const params = (req.params ?? {}) as Record<string, unknown>
  return contextOf(bytes, {
    ...params,
    ...(req.baseVersion ? { version: req.baseVersion } : {}),
  })
}
