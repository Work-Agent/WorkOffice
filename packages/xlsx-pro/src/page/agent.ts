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
import { opTier, resolveTokens, tokenNeeds } from '../shared/op-tiers'
import type { AgentContextRequest, Receipt } from './env'

export interface AgentBridgeDeps {
  sessionId: string
}

type AnyRecord = Record<string, unknown>

/** 页面读投影的单次上限：超过就截断并明说，不静默。 */
const MAX_READ_CELLS = 2_000

/** 页面读投影的文本上限（与无页面 `xlsx_read` 的 200 KB 对齐，留一点余量）。 */
const MAX_READ_TEXT = 120_000

/**
 * 真正需要页面 live 模型的只读工具。
 *
 * `xlsx_guide` 是静态命令目录：进页面只会拿到一份与本意无关的投影
 * （之前 guide 就掉进了 range 分支）。回 `command.unsupported`，
 * 宿主就会走执行点那份唯一事实源。
 */
const PAGE_READ_TOOLS = new Set(['xlsx_read', 'xlsx.read'])

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

  if (request.kind === 'read') {
    if (request.tool && !PAGE_READ_TOOLS.has(request.tool))
      throw coded(
        'command.unsupported',
        `${request.tool} 不需要页面模型：请由宿主走执行点（静态目录 / 结构化读）`,
      )
    return { context: clipContext(readFromPage(page, request.read)) }
  }

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
  // 令牌来源：消息顶层（宿主透传）→ 命令入参（旧名 epoch / revision）→ 报错。
  // 宿主曾经从来不透传，而入参里叫 `epoch`、校验点只认 `structureEpoch`，
  // 于是含 A1 坐标的批次 100% 被拒。
  const supplied = resolveTokens({
    message: { docRev: request.docRev, structureEpoch: request.structureEpoch },
    commands,
  })
  if (needs.structureEpoch) {
    const current = page.structureEpoch()
    if (supplied.structureEpoch !== undefined && supplied.structureEpoch !== current)
      throw coded(
        'command.stale-index',
        `工作表结构已变（structureEpoch ${supplied.structureEpoch} → ${current}）；` +
          `批内含 A1 绝对坐标命令，重新读取区域后再发命令`,
      )
    if (supplied.structureEpoch === undefined)
      throw coded(
        'command.stale-index',
        '批内含 A1 绝对坐标命令，必须带上读到上下文时的 structureEpoch',
      )
  }
  if (needs.docRev) {
    const current = page.docRev()
    if (supplied.docRev !== undefined && supplied.docRev !== current)
      throw coded(
        'command.stale-index',
        `页面内容已变（docRev ${supplied.docRev} → ${current}）；` +
          `批内含读内容型命令（fill/copy/sort/find），重新读取后再发命令`,
      )
    if (supplied.docRev === undefined)
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
      // dry-run 不改模型：令牌原样带回，宿主可以直接用于下一条真写。
      ...tokens(page),
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
  // 回执 + 新令牌：写完 docRev / structureEpoch 都变了，不回令牌的话
  // 下一条命令只能靠“重新 read”才能拿到（之前只回 receipt，槽位永远落后）。
  return { receipt: { ...receipt, applied: requested }, requested, applied: requested, ...tokens(page) }
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

  // 缺省 what 必须是 sheets：之前只匹配 'sheets' 字面量，undefined 就直接
  // 掉进 range 分支，于是“读概览”变成一张 1000×52 的空表 dump。
  const wanted = String(request.what ?? 'sheets')

  if (wanted === 'sheets') {
    const sheets = (ctx['sheets'] as Array<AnyRecord> | undefined) ?? []
    const lines = ['# 工作表', ...sheets.map((sheet) =>
      `- ${String(sheet['name'])}${sheet['rows'] ? ` · ${String(sheet['rows'])} 行` : ''}${
        sheet['columns'] ? ` × ${String(sheet['columns'])} 列` : ''
      }`,
    )]
    return { ...base, text: lines.join('\n'), sheets: page.sheets() }
  }

  if (wanted === 'audit') {
    const sheets = page.sheets()
    const ctxLines = [
      '# 页面审计',
      `version=${base.version.slice(0, 12)} epoch=${base.epoch} docRev=${base.docRev} structureEpoch=${base.structureEpoch}`,
      `sheets=${sheets.length}（${sheets.map((sheet) => sheet.name).join(' / ')}）`,
      '包结构级诊断（part / rels / 加密 / zip64）需要关闭页面后走无页面 audit。',
    ]
    return { ...base, text: ctxLines.join('\n'), sheets }
  }

  // 概览 / 元数据：页面已经持有这些信息，不必再扫单元格。
  // `meta` 额外带上原始 context（页面重建 file 模型用）。
  if (wanted === 'context' || wanted === 'meta') {
    const text = contextTextOf(ctx)
    return wanted === 'meta'
      ? { ...base, text, sheets: page.sheets(), data: ctx }
      : { ...base, text, sheets: page.sheets() }
  }

  // cells / formulas 之外的结构化读和未知值一律交回宿主：
  // 宿主收到 command.unsupported 会落回执行点（wasm 结构化读 / 静态目录）。
  if (wanted !== 'range' && wanted !== 'formulas' && wanted !== 'find')
    throw coded(
      'command.unsupported',
      `页面读不支持 what=${wanted}（页面支持：context / meta / sheets / range / formulas / find / audit）`,
    )

  // range / formulas / find 都基于 live 单元格
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

  // range：只打印真正取到地址的行（之前无论 truncated 与否都把整个 bounds
  // 走一遍，A1:AZ1000 会输出 5 万多格几乎全空的行）。
  const lines: string[] = [`# ${sheetName}!${bounds.range}`]
  const width = bounds.endColumn - bounds.startColumn + 1
  for (let offset = 0; offset < addresses.length; offset += width) {
    const rowCells = addresses.slice(offset, offset + width).map((address) => {
      const cell = cells[address]
      const shown = cell?.value ?? cell?.display ?? ''
      return String(shown).replace(/\t/g, ' ').replace(/\n/g, ' ')
    })
    lines.push(`${bounds.startRow + Math.floor(offset / width) + 1}\t${rowCells.join('\t')}`)
  }
  if (truncated) lines.push(`…（超过 ${MAX_READ_CELLS} 格，已截断；用更小的 range）`)
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
  const requested =
    request.range ??
    (request.start || request.end
      ? `${request.start ?? 'A1'}:${request.end ?? ''}`
      : null)
  // 未给范围（或只给了起点）时，优先用页面已知的已加载范围 / 当前选区，
  // 而不是页表尺寸（1000×52）——那份 dump 既浪费上下文也不代表 used range。
  const text = requested && !requested.endsWith(':') ? requested : defaultRangeOf(ctx)
  if (text) {
    const bounds = parseRange(text)
    return { sheet: sheetName, range: text, ...bounds }
  }
  // 连已加载范围都不知道：退回表尺寸（仍受 MAX_READ_CELLS 限制）。
  const rows = Number(meta?.['rows'] ?? 50)
  const columns = Number(meta?.['columns'] ?? 8)
  const range = `A1:${formatAddress(Math.max(0, rows - 1), Math.max(0, columns - 1))}`
  return { sheet: sheetName, range, ...parseRange(range) }
}

/**
 * 页面已知的“有意义范围”：已加载区 → 当前选区。
 * 两者都没有时返回 null，由调用方决定退路。
 */
function defaultRangeOf(ctx: AnyRecord): string | null {
  const loaded = ctx['loadedRange']
  if (typeof loaded === 'string' && loaded) return loaded
  const selection = ctx['selection']
  if (typeof selection === 'string' && selection) {
    const part = selection.includes('!')
      ? selection.slice(selection.indexOf('!') + 1)
      : selection
    if (part.includes(':')) return part
  }
  return null
}

/** 页面概览（与无页面 `contextText` 同构，但只报页面真知道的事）。 */
function contextTextOf(ctx: AnyRecord): string {
  const sheets = (ctx['sheets'] as Array<AnyRecord> | undefined) ?? []
  const lines = ['# workbook.xlsx', `sheets=${sheets.length}`]
  for (const sheet of sheets) {
    const rows = sheet['rows']
    const columns = sheet['columns']
    const extent =
      typeof rows === 'number' && typeof columns === 'number'
        ? ` rows=${rows} columns=${columns}`
        : ''
    lines.push(
      `- ${String(sheet['name'])} id=${String(sheet['id'])}${extent}` +
        (sheet['readOnlyOversized'] === true ? ' readOnlyOversized' : ''),
    )
  }
  const active = ctx['sheetName']
  if (typeof active === 'string' && active)
    lines.push(
      `active=${active}` +
        (typeof ctx['selection'] === 'string' && ctx['selection'] ? ` selection=${ctx['selection']}` : ''),
    )
  const loaded = ctx['loadedRange']
  if (typeof loaded === 'string' && loaded) lines.push(`loaded=${loaded}`)
  const merges = ctx['merges']
  if (Array.isArray(merges) && merges.length > 0)
    lines.push(
      `merges=${merges.length}（${merges.slice(0, 12).map((entry) => String(entry)).join(' ')}` +
        `${merges.length > 12 ? ' …' : ''}）`,
    )
  lines.push('（页面 live 模型：含未保存修改；逐格内容用 what=range）')
  return lines.join('\n')
}

/**
 * 读投影太长就截断并明说：一屏上下文比“半个 A1:AZ1000”有用。
 */
function clipContext(value: unknown): unknown {
  if (!value || typeof value !== 'object') return value
  const record = value as AnyRecord
  const text = record['text']
  if (typeof text !== 'string' || text.length <= MAX_READ_TEXT) return value
  return {
    ...record,
    text: `${text.slice(0, MAX_READ_TEXT)}\n…（已截断：超过 ${MAX_READ_TEXT} 字符，用 range 缩小范围）`,
  }
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
