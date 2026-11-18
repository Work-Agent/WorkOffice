// ─── 页面代理：agent-request → live 模型读 / Univer 事务 → commit ─────
//
// 页面是编辑权威（I10）：有页面时 agent 的**读与写都进这里**。
// 读不落磁盘——页面有未保存修改时，磁盘内容与模型不一致，用磁盘回答
// 会让 agent 拿旧值写回，且令牌校验挡不住（那些修改发生在它读之前）。
//
// 新鲜度按 op 分级（设计稿 §6.4，shared/op-tiers.ts）：
//   L1 读内容型 → docRev        L2 绝对坐标 → structureEpoch
// L0（按名称的全局语义）不校验，在当前模型上重新解析即可。

import { z } from 'zod'
import {
  describeOperationErrors,
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { formatAddress, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import { normalizeSheetRefs } from '../editor/mcp-sheet-refs'
import { opTier, tokenNeeds } from '../shared/op-tiers'
import type { AgentContextRequest, Receipt } from './env'

export interface AgentBridgeDeps {
  sessionId: string
}

type AnyRecord = Record<string, unknown>

/** 页面读投影的单次上限：超过就截断并明说，不静默。 */
const MAX_READ_CELLS = 2_000

function coded(code: string, message: string): Error & { code: string } {
  const error = new Error(message) as Error & { code: string }
  error.code = code
  return error
}

function bridge() {
  const value = window.__xlsxPro
  if (!value) throw coded('command.page-unavailable', '工作簿页面尚未就绪')
  return value
}

function host() {
  const value = window.__xlsxProHost
  if (!value) throw coded('command.page-unavailable', '页面保存通道尚未就绪')
  return value
}

function tokens(page: ReturnType<typeof bridge>) {
  return { docRev: page.docRev(), structureEpoch: page.structureEpoch() }
}

export function installAgentBridge(deps: AgentBridgeDeps): () => void {
  const unsubscribe = window.extension.bus.on((raw) => {
    if (!raw || raw.type !== 'agent-request') return
    const request = raw as unknown as AgentContextRequest
    void handle(deps, request).then(
      (result) =>
        void window.extension.ura.page.agentResult(deps.sessionId, {
          type: 'agent-result',
          id: request.id,
          ok: true,
          ...result,
        }),
      (error: unknown) =>
        void window.extension.ura.page.agentResult(deps.sessionId, {
          type: 'agent-result',
          id: request.id,
          ok: false,
          error: {
            code: (error as { code?: string }).code ?? 'command.invalid',
            message: error instanceof Error ? error.message : String(error),
          },
        }),
    )
  })
  return unsubscribe
}

async function handle(
  deps: AgentBridgeDeps,
  request: AgentContextRequest,
): Promise<{ context?: unknown; receipt?: Receipt } & Record<string, unknown>> {
  const page = bridge()
  const h = host()

  if (request.kind === 'context') {
    const context = (page.context() ?? {}) as AnyRecord
    return {
      context: { ...context, version: h.getVersion(), epoch: h.getEpoch(), ...tokens(page) },
    }
  }

  if (request.kind === 'read') return { context: readFromPage(page, request.read) }

  return applyOnPage(page, h, request)
}

async function applyOnPage(
  page: ReturnType<typeof bridge>,
  h: ReturnType<typeof host>,
  request: AgentContextRequest,
): Promise<{ receipt?: Receipt; requested: number; applied: number; dryRun?: boolean; wouldApply?: number }> {
  const commands = Array.isArray(request.commands) ? (request.commands as AnyRecord[]) : []
  if (commands.length === 0) throw coded('command.invalid', '缺少 commands')

  // 先把所有命令拆成 op 批（工具 → ops），再做令牌校验：令牌是 op 级的。
  const allOps: WorkbookOperation[] = []
  for (const [index, command] of commands.entries()) {
    const tool = String(command['tool'] ?? '')
    const params = (command['params'] ?? {}) as AnyRecord
    if (tool !== 'xlsx_apply' && tool !== 'xlsx.apply')
      throw coded('command.page-unavailable', `页面代理不支持工具 ${tool}；写操作请用 xlsx_apply`)
    const rawOps = Array.isArray(params['commands'])
      ? params['commands']
      : Array.isArray(params['ops'])
        ? params['ops']
        : []
    if (rawOps.length === 0) throw coded('command.invalid', `commands[${index}] 缺少 ops`)
    const named = normalizeSheetRefs(rawOps, page.sheets())
    if (!named.ok) throw coded('command.invalid', named.error)
    const parsed = z.array(workbookOperationSchema).safeParse(named.ops)
    if (!parsed.success)
      throw coded('command.invalid', describeOperationErrors(named.ops, parsed.error))
    allOps.push(...(parsed.data as WorkbookOperation[]))
  }

  const needs = tokenNeeds(allOps)
  if (needs.structureEpoch) {
    const current = page.structureEpoch()
    if (request.structureEpoch !== undefined && request.structureEpoch !== current)
      throw coded(
        'command.stale-index',
        `工作表结构已变（structureEpoch ${request.structureEpoch} → ${current}）；` +
          `批内含 A1 绝对坐标命令，重新读取区域后再发命令`,
      )
    if (request.structureEpoch === undefined)
      throw coded(
        'command.stale-index',
        '批内含 A1 绝对坐标命令，必须带上读到上下文时的 structureEpoch',
      )
  }
  if (needs.docRev) {
    const current = page.docRev()
    if (request.docRev !== undefined && request.docRev !== current)
      throw coded(
        'command.stale-index',
        `页面内容已变（docRev ${request.docRev} → ${current}）；` +
          `批内含读内容型命令（fill/copy/sort/find），重新读取后再发命令`,
      )
    if (request.docRev === undefined)
      throw coded('command.stale-index', '批内含读内容型命令，必须带上读到上下文时的 docRev')
  }

  const requested = allOps.length

  if (request.dryRun === true) {
    const planned = (await page.applyOps(allOps, true)) as AnyRecord
    if (planned['ok'] !== true)
      throw coded('command.invalid', String(planned['reason'] ?? 'dry-run 失败'))
    return {
      dryRun: true,
      requested,
      applied: 0,
      wouldApply: Number(planned['cellChangeCount'] ?? 0) +
        (Array.isArray(planned['structuralChanges']) ? planned['structuralChanges'].length : 0),
    }
  }

  // 所有命令合成一个批次：页面执行器只创建一个撤销项；失败时不会只落一半。
  const depthBefore = page.undoDepth()
  const outcome = (await page.applyOps(allOps, false)) as AnyRecord
  if (outcome['ok'] !== true)
    throw coded('command.invalid', String(outcome['reason'] ?? '应用失败'))

  h.setOrigin('agent')
  let saved
  try {
    saved = await page.save()
  } finally {
    h.setOrigin('page')
  }

  if (!saved?.ok) {
    // 严格策略（设计稿 §5.4）：落盘失败就把这一批撤销掉，页面保持 clean。
    // 否则页面已脏、磁盘未变，agent 之后的读会看到「已应用」而文件没有。
    // 只撤到自己那一批为止：保存是异步的，期间用户可能又改了一笔。
    let rolledBack = false
    for (let guard = 0; guard < 3 && page.undoDepth() > depthBefore; guard += 1)
      rolledBack = page.undo() || rolledBack
    throw coded(
      'command.invalid',
      `${saved?.error ?? '保存失败：工作簿未落盘'}` +
        (rolledBack ? '（已回滚本批修改）' : '（回滚失败：页面已有未保存修改，请人工确认）'),
    )
  }

  const receipt = h.getReceipt()
  if (!receipt) throw coded('command.invalid', '提交未返回回执')
  return { receipt: { ...receipt, applied: requested }, requested, applied: requested }
}

// ─── 读投影 ────────────────────────────────────────────────────────────

function readFromPage(
  page: ReturnType<typeof bridge>,
  read: AgentContextRequest['read'],
): unknown {
  const request = read ?? { what: 'sheets' as const }
  const ctx = (page.context() ?? {}) as AnyRecord
  const base = {
    source: 'page' as const,
    version: window.__xlsxProHost?.getVersion() ?? '',
    epoch: window.__xlsxProHost?.getEpoch() ?? 0,
    ...tokens(page),
  }

  if (request.what === 'sheets') {
    const sheets = (ctx['sheets'] as Array<AnyRecord> | undefined) ?? []
    const lines = ['# 工作表', ...sheets.map((sheet) =>
      `- ${String(sheet['name'])}${sheet['rows'] ? ` · ${String(sheet['rows'])} 行` : ''}${
        sheet['columns'] ? ` × ${String(sheet['columns'])} 列` : ''
      }`,
    )]
    return { ...base, text: lines.join('\n'), sheets: page.sheets() }
  }

  if (request.what === 'audit') {
    const sheets = page.sheets()
    const ctxLines = [
      '# 页面审计',
      `version=${base.version.slice(0, 12)} epoch=${base.epoch} docRev=${base.docRev} structureEpoch=${base.structureEpoch}`,
      `sheets=${sheets.length}（${sheets.map((sheet) => sheet.name).join(' / ')}）`,
      '包结构级诊断（part / rels / 加密 / zip64）需要关闭页面后走无页面 audit。',
    ]
    return { ...base, text: ctxLines.join('\n'), sheets }
  }

  // range / formulas / find 都基于 live 单元格
  const wanted = request.what
  const bounds = rangeBoundsOf(request, ctx)
  const sheetName = bounds.sheet
  const sheet = page.sheets().find((entry) => entry.name === sheetName) ??
    page.sheets().find((entry) => entry.name.toLowerCase() === sheetName.toLowerCase())
  if (!sheet) throw coded('command.not-found', `工作表不存在：${sheetName}`)

  const addresses: string[] = []
  let truncated = false
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      if (addresses.length >= MAX_READ_CELLS) {
        truncated = true
        break
      }
      addresses.push(formatAddress(row, column))
    }
    if (truncated) break
  }

  const cells = (page.readCells(addresses, sheet.id) ?? {}) as Record<
    string,
    { value?: unknown; display?: unknown; formula?: string }
  >

  if (wanted === 'find') {
    const query = String(request.query ?? '')
    if (!query) throw coded('command.invalid', 'find 缺少 query')
    const matcher = request.regex === true
      ? safeRegExp(query, request.matchCase === true)
      : null
    const matches: string[] = []
    const limit = Math.max(1, Math.min(Number(request.limit ?? 100), 1_000))
    for (const address of addresses) {
      const cell = cells[address]
      if (!cell) continue
      const text = String(cell.formula ?? cell.value ?? '')
      const hit = matcher
        ? matcher.test(text)
        : (request.matchCase === true ? text.includes(query) : text.toLowerCase().includes(query.toLowerCase()))
      if (!hit) continue
      matches.push(`${sheetName}!${address}\t${text}`)
      if (matches.length >= limit) break
    }
    return {
      ...base,
      text: [`# ${sheetName}!${bounds.range} 命中 ${matches.length}`, ...matches].join('\n'),
      matches,
      truncated,
    }
  }

  if (wanted === 'formulas') {
    const lines: string[] = []
    for (const address of addresses) {
      const cell = cells[address]
      if (!cell?.formula) continue
      lines.push(`${sheetName}!${address}\t${cell.formula}`)
    }
    return {
      ...base,
      text: [`# ${sheetName} 公式 ${lines.length}`, ...lines].join('\n'),
      truncated,
    }
  }

  // range
  const lines: string[] = [`# ${sheetName}!${bounds.range}`]
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    const rowCells: string[] = []
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      const cell = cells[formatAddress(row, column)]
      const shown = cell?.value ?? cell?.display ?? ''
      rowCells.push(String(shown).replace(/\t/g, ' ').replace(/\n/g, ' '))
    }
    lines.push(`${row + 1}\t${rowCells.join('\t')}`)
  }
  if (truncated) lines.push(`…（超过 ${MAX_READ_CELLS} 格，已截断）`)
  return { ...base, text: lines.join('\n'), truncated }
}

function rangeBoundsOf(
  request: NonNullable<AgentContextRequest['read']>,
  ctx: AnyRecord,
): {
  sheet: string
  range: string
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
} {
  const sheets = (ctx['sheets'] as Array<AnyRecord> | undefined) ?? []
  const sheetName = String(
    request.sheet ?? ctx['sheetName'] ?? (sheets[0]?.['name'] as string | undefined) ?? '',
  )
  if (!sheetName) throw coded('command.invalid', '缺少 sheet')
  const meta = sheets.find((entry) => entry['name'] === sheetName)
  const text = request.range ?? `${request.start ?? 'A1'}:${request.end ?? ''}`
  if (text && !text.endsWith(':')) {
    const bounds = parseRange(text)
    return { sheet: sheetName, range: text, ...bounds }
  }
  // 没给范围时用页面已知的表尺寸（缺失时退回 A1:H50），避免扫整张表。
  const rows = Number(meta?.['rows'] ?? 50)
  const columns = Number(meta?.['columns'] ?? 8)
  const range = `A1:${formatAddress(Math.max(0, rows - 1), Math.max(0, columns - 1))}`
  return { sheet: sheetName, range, ...parseRange(range) }
}

function safeRegExp(query: string, matchCase: boolean): RegExp | null {
  try {
    return new RegExp(query, matchCase ? '' : 'i')
  } catch {
    return null
  }
}

/** 供 guide / 报错信息共用的层级名。 */
export function tierOf(op: string): string {
  return opTier(op)
}
