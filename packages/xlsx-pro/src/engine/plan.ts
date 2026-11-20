// ─── engine.plan：命令 → 计划（无页面路径） ───────────────────────────
//
// 只读工具永远进执行点；xlsx_apply 只有 cells / structure / sheets 组能在
// 这里执行，objects 组明确回 command.page-unavailable。页面路径由宿主转发
// 页面代理，本文件不参与。

import { workbookOperationSchema, type WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { buildGuide, groupOf, isHeadlessOp, specOf } from '../tools/catalog'
import type { NewSheetInput, Rejection, XlsxProPlan } from '../shared/types'
import { parseWorkbook, type WorkbookView } from '../shared/workbook'
import { HeadlessOpError, buildSaveRequest } from '../shared/headless'
import { coded, messageOf, type EngineCtx, type PlanOutcome, type PlanRequest } from './abi'
import * as read from './read'
import { readMetaForPage, readRangeForPage } from './page-read'

const MAX_COMMANDS = 200

type AnyRecord = Record<string, unknown>

export async function plan(req: PlanRequest, ctx: EngineCtx): Promise<PlanOutcome> {
  const params = (req.params ?? {}) as AnyRecord
  const tool = String(req.tool ?? '')

  if (tool === 'xlsx_guide' || tool === 'xlsx.guide') {
    const guide = buildGuide({
      ...(typeof params['group'] === 'string' ? { group: params['group'] } : {}),
      ...(typeof params['op'] === 'string' ? { op: params['op'] } : {}),
      ...(typeof params['version'] === 'string' ? { version: params['version'] } : {}),
      ...(typeof params['epoch'] === 'number' ? { epoch: params['epoch'] } : {}),
    })
    return text(guide.summary)
  }

  if (tool === 'xlsx_create' || tool === 'xlsx.create')
    return createOutcome(req, params)

  if (req.missing)
    throw coded('resource.not-found', `文件不存在：${req.ref}；新建请用 xlsx_create`)

  // 工作簿元数据（页面 invoke）：重建 file 模型用，不解析整包。
  if ((tool === 'xlsx_read' || tool === 'xlsx.read') && params['what'] === 'meta') {
    const meta = await readMetaForPage(req, ctx)
    return {
      plan: { kind: 'read', data: meta.data, audit: { sessionId: meta.sessionId, cached: meta.cached } },
      summary: `meta: ${(meta.data as { sheets?: unknown[] })?.sheets?.length ?? 0} sheet(s) (session=${meta.cached ? 'cached' : 'opened'})`,
    }
  }

  // 结构化范围读取（页面 invoke）：不解析整包，直接问 wasm 引擎。
  if ((tool === 'xlsx_read' || tool === 'xlsx.read') && params['what'] === 'cells') {
    const read = await readRangeForPage(req, params, ctx)
    return {
      plan: {
        kind: 'read',
        data: read.data,
        audit: { sessionId: read.sessionId, cached: read.cached },
      },
      summary: `cells: ${read.sheetId} ${String(params['sheetId'] ?? params['sheet'] ?? '')} (session=${read.cached ? 'cached' : 'opened'})`,
    }
  }

  const source = req.bytes
    ? {
        bytes: req.bytes,
        // 宿主默认会把基线版本一起送来；没送就自己读一次（旧宿主 / 测试）。
        version: req.baseVersion ?? (await ctx.readBytes()).version,
      }
    : await ctx.readBytes()
  const bytes = source.bytes

  if (tool === 'xlsx_read' || tool === 'xlsx.read') return readOutcome(bytes, params, source.version)
  if (tool === 'xlsx_apply' || tool === 'xlsx.apply') return applyOutcome(req, params, bytes)

  throw coded(
    'command.page-unavailable',
    `${tool} 需要打开工作簿页面（页面内工具）。没有页面时用 xlsx_apply 的 cells / structure / sheets 组，或先让用户打开文件。`,
  )
}

function text(summary: string): PlanOutcome {
  // 只读结果走 `text`，与「命令没产生修改」的 `noop` 分开（设计稿 §3.1）。
  return { plan: { kind: 'text', text: summary }, summary }
}

async function readOutcome(
  bytes: Uint8Array,
  params: AnyRecord,
  version?: string,
): Promise<PlanOutcome> {
  let view: WorkbookView
  try {
    view = await parseWorkbook(bytes)
  } catch (error) {
    throw coded('command.invalid', messageOf(error))
  }
  const what = String(params['what'] ?? 'context')
  const name = String(params['name'] ?? 'workbook.xlsx')

  if (what === 'context') return withVersion(text(requireRead(() => read.context(view, name))), version)
  if (what === 'sheets') return withVersion(text(requireRead(() => read.sheets(view))), version)
  if (what === 'range') return withVersion(text(requireRead(() => read.range(view, params))), version)
  if (what === 'formulas') return withVersion(text(requireRead(() => read.formulas(view, params))), version)
  if (what === 'find') return withVersion(text(requireRead(() => read.find(view, params))), version)
  if (what === 'audit') return withVersion(text(await read.audit(bytes, view)), version)

  throw coded('command.invalid', `what 非法：${what}（可用：context / sheets / range / formulas / find / audit）`)
}

/**
 * 只读结果头部带上基线版本。
 *
 * 无页面写（`xlsx_apply`）要求 `version` 必填（宿主 `requiresBaseVersion` 把关），
 * 而它唯一的来源就是一次读。读不回版本 = 写永远发不出去（"需要基线版本"）。
 */
function withVersion(outcome: PlanOutcome, version?: string): PlanOutcome {
  if (!version) return outcome
  const plan = outcome.plan
  if (plan.kind !== 'text') return outcome
  const body = `version=${version}\n${plan.text}`
  return { ...outcome, plan: { ...plan, text: body }, summary: body }
}

function requireRead(run: () => string): string {
  try {
    return run()
  } catch (error) {
    if (error instanceof HeadlessOpError) throw coded(error.code, error.message)
    throw error
  }
}

function createOutcome(req: PlanRequest, params: AnyRecord): PlanOutcome {
  if (!req.missing)
    throw coded('command.invalid', '目标已存在；修改已有工作簿请用 xlsx_apply')

  const sheets = normalizeNewSheets(params)
  if (sheets.length === 0) sheets.push({ name: 'Sheet1', rows: [] })
  for (const sheet of sheets) {
    if (!sheet.name || sheet.name.length > 31)
      throw coded('command.invalid', `工作表名非法：${sheet.name || '(空)'}`)
  }
  return {
    plan: { kind: 'new', sheets, applied: sheets.length },
    summary: `create: ${req.ref} · sheets=${sheets.length}`,
  }
}

function normalizeNewSheets(params: AnyRecord): NewSheetInput[] {
  if (Array.isArray(params['sheets'])) {
    return (params['sheets'] as unknown[]).map((entry, index) => {
      if (!entry || typeof entry !== 'object') throw coded('command.invalid', `sheets[${index}] 不是对象`)
      const record = entry as AnyRecord
      const name = String(record['name'] ?? (index === 0 ? 'Sheet1' : `Sheet${index + 1}`))
      const rows = Array.isArray(record['rows'])
        ? (record['rows'] as unknown[][]).map((row) => (Array.isArray(row) ? row : []))
        : []
      return { name, rows: rows as unknown as NewSheetInput['rows'] }
    })
  }
  const csv = typeof params['csv'] === 'string' ? params['csv'] : null
  if (csv !== null) return [{ name: String(params['name'] ?? 'Sheet1'), rows: parseCsv(csv) as unknown as NewSheetInput['rows'] }]
  const rows = Array.isArray(params['rows'])
    ? (params['rows'] as unknown[][]).map((row) => (Array.isArray(row) ? row : []))
    : []
  return [{ name: String(params['name'] ?? 'Sheet1'), rows: rows as unknown as NewSheetInput['rows'] }]
}

function parseCsv(text: string): unknown[][] {
  const rows: unknown[][] = []
  let row: unknown[] = []
  let field = ''
  let quoted = false
  for (let index = 0; index < text.length; index += 1) {
    const ch = text[index]!
    if (quoted) {
      if (ch === '"') {
        if (text[index + 1] === '"') {
          field += '"'
          index += 1
        } else {
          quoted = false
        }
      } else {
        field += ch
      }
      continue
    }
    if (ch === '"') {
      quoted = true
    } else if (ch === ',') {
      row.push(cellValue(field))
      field = ''
    } else if (ch === '\n') {
      row.push(cellValue(field))
      rows.push(row)
      row = []
      field = ''
    } else if (ch !== '\r') {
      field += ch
    }
  }
  if (field !== '' || row.length > 0) {
    row.push(cellValue(field))
    rows.push(row)
  }
  return rows
}

function cellValue(text: string): string | number | boolean | null {
  const trimmed = text.trim()
  if (trimmed === '') return null
  if (trimmed === 'TRUE') return true
  if (trimmed === 'FALSE') return false
  if (/^-?\d+(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(trimmed)) return Number(trimmed)
  return text
}

async function applyOutcome(
  req: PlanRequest,
  params: AnyRecord,
  bytes: Uint8Array,
): Promise<PlanOutcome> {
  const raw = Array.isArray(req.commands)
    ? req.commands
    : Array.isArray(params['commands'])
      ? params['commands']
      : Array.isArray(params['ops'])
        ? params['ops']
        : null
  if (!raw) throw coded('command.invalid', '缺少 commands 数组')

  if (raw.length === 0) {
    return { plan: { kind: 'check', applied: 0 }, summary: 'no commands' }
  }
  if (raw.length > MAX_COMMANDS)
    return {
      plan: {
        kind: 'reject',
        applied: 0,
        rejected: [{ index: -1, code: 'command.too-large', message: `批上限 ${MAX_COMMANDS} 条（收到 ${raw.length}）` }],
      },
      summary: `rejected: 批上限 ${MAX_COMMANDS} 条`,
    }

  const normalized: AnyRecord[] = []
  for (const [index, command] of raw.entries()) {
    if (!command || typeof command !== 'object')
      return rejected(index, 'command.invalid', '命令不是对象')
    const record = { ...(command as AnyRecord) }
    const op = String(record['op'] ?? '')
    if (!op) return rejected(index, 'command.not-found', '缺少 op')
    if (!specOf(op)) return rejected(index, 'command.not-found', `未知 op：${op}（用 xlsx_guide 取可用 op）`)
    if (!isHeadlessOp(op))
      throw coded(
        'command.page-unavailable',
        `commands[${index}] 的 op「${op}」属于 ${groupOf(op)} 组，需要打开工作簿页面；` +
          '没有页面时只能用 cells / structure / sheets 组的无页面 op。',
      )
    if (typeof record['sheet'] === 'string' && record['sheetId'] === undefined)
      record['sheetId'] = record['sheet']
    delete record['sheet']
    const parsed = workbookOperationSchema.safeParse(record)
    if (!parsed.success)
      return rejected(index, 'command.invalid', parsed.error.issues.map((issue) => `${issue.path.join('.')}: ${issue.message}`).join('; '))
    normalized.push(parsed.data as AnyRecord)
  }

  return buildOutcome(bytes, normalized as unknown as WorkbookOperation[], params)
}

async function buildOutcome(
  bytes: Uint8Array,
  operations: WorkbookOperation[],
  params: AnyRecord,
): Promise<PlanOutcome> {
  let view: WorkbookView
  try {
    view = await parseWorkbook(bytes)
  } catch (error) {
    throw coded('command.invalid', messageOf(error))
  }
  try {
    const built = buildSaveRequest(view, operations)
    const plan: XlsxProPlan = {
      kind: 'save',
      request: built.request,
      changedCount: built.changedCount,
      applied: operations.length,
      ...(params['dryRun'] === true ? { dryRun: true } : {}),
      ...(built.warnings.length > 0 ? { warnings: built.warnings } : {}),
    }
    return {
      plan,
      summary:
        `plan: ${operations.length} command(s), edits=${built.request.edits.length}, ` +
        `structural=${built.request.structuralOps.reduce((sum, entry) => sum + entry.ops.length, 0)}` +
        `${built.request.sheetPlan ? ', sheet-plan=yes' : ''}`,
    }
  } catch (error) {
    if (error instanceof HeadlessOpError)
      // 失败发生在整批展开阶段，不是某一条命令：index 用 -1 表示整批级。
      return rejected(-1, error.code, error.message)
    throw coded('command.invalid', messageOf(error))
  }
}

function rejected(index: number, code: string, message: string): PlanOutcome {
  const rejection: Rejection = { index, code, message }
  return {
    plan: { kind: 'reject', applied: 0, rejected: [rejection] },
    summary: `rejected[${index}] ${code}: ${message}`,
  }
}
