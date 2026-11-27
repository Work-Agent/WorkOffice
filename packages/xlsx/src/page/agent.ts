/**
 * 页面代理：`xlsx/events` → pending → 编辑器事务 → 保存 → ack。
 *
 * 页面是唯一执行者：引擎把请求入队后只推 id/type，正文由页面取；
 * 令牌（docRev / structureEpoch）在这里按 op 级校验。
 */

import { opTier, tokenNeeds } from '../shared/op-tiers'

import type { EditorBridge } from './bridge'
import type { Channel } from './channel'

interface PageRequest {
  id: string
  type: 'apply' | 'read'
  commands?: unknown[]
  what?: string
  params?: Record<string, unknown>
  docRev?: number
  structureEpoch?: number
  dryRun?: boolean
}

interface EditorHandle {
  context(): unknown
  readCells(addresses: string[], sheetId?: string): Record<string, unknown>
  sheets(): ReadonlyArray<{ id: string; name: string }>
  applyOps(ops: unknown[], dryRun: boolean): Promise<Record<string, unknown>>
  save(): Promise<{ ok: boolean; error?: string } | undefined>
  docRev(): number
  structureEpoch(): number
  undo(): boolean
  undoDepth(): number
}

const MAX_READ_CELLS = 2_000

type AnyRecord = Record<string, unknown>

function editor(): EditorHandle | null {
  return (window.__xlsxPro as EditorHandle | undefined) ?? null
}

/** 请求可能早于 App 的 effect 到达：最多等 5 s。 */
async function waitForEditor(): Promise<EditorHandle | null> {
  const deadline = Date.now() + 5_000
  for (;;) {
    const handle = editor()
    if (handle) return handle
    if (Date.now() >= deadline) return null
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
}

function errorOf(error: unknown): { code: string; message: string } {
  const code = (error as { code?: unknown } | null)?.code
  return {
    code: typeof code === 'string' ? code : 'command.invalid',
    message: error instanceof Error ? error.message : String(error),
  }
}

export function installAgent(channel: Channel, bridge: EditorBridge): () => void {
  const bridgeApi = window.workStudio
  if (!bridgeApi) return () => {}

  const off = bridgeApi.onUpdate((update) => {
    if (update.domain !== 'xlsx' || update.resource !== 'events') return
    if (update.type !== 'event' || update.event !== 'request') return
    const data = update.data as { id?: unknown; type?: unknown } | null
    const id = typeof data?.id === 'string' ? data.id : ''
    const type = typeof data?.type === 'string' ? data.type : ''
    if (id) void handleRequest(channel, bridge, id, type)
  })
  void bridgeApi.subscribe('xlsx', 'events').catch(() => undefined)
  return off
}

async function handleRequest(
  channel: Channel,
  bridge: EditorBridge,
  id: string,
  type: string,
): Promise<void> {
  const ack = async (ok: boolean, payload: Record<string, unknown>): Promise<void> => {
    try {
      await channel.call('ack', { id, ok, payload })
    } catch {
      /* pending 已超时：无从补救 */
    }
  }

  try {
    const request = await channel.call<PageRequest>('pending', { id })
    if (type === 'read') {
      await ack(true, { data: readFromPage(channel, bridge, request) })
      return
    }
    if (type !== 'apply') {
      await ack(false, { error: { code: 'command.unsupported', message: `不支持的请求：${type}` } })
      return
    }

    const handle = await waitForEditor()
    if (!handle) throw Object.assign(new Error('编辑器尚未就绪'), { code: 'command.page-unavailable' })
    const commands = Array.isArray(request.commands) ? request.commands : []
    if (commands.length === 0) throw new Error('请求没有 commands')

    // op 级令牌：L1 → docRev，L2 → structureEpoch。
    const needs = tokenNeeds(commands)
    const docRev = handle.docRev()
    const epoch = handle.structureEpoch()
    if (needs.structureEpoch) {
      if (request.structureEpoch === undefined)
        throw Object.assign(new Error('批内含 A1 绝对坐标命令，必须带上 structureEpoch'), {
          code: 'command.stale-index',
        })
      if (request.structureEpoch !== epoch)
        throw Object.assign(
          new Error(`工作表结构已变（structureEpoch ${request.structureEpoch} → ${epoch}）；重新读取后再发命令`),
          { code: 'command.stale-index' },
        )
    }
    if (needs.docRev) {
      if (request.docRev === undefined)
        throw Object.assign(new Error('批内含读内容型命令，必须带上 docRev'), {
          code: 'command.stale-index',
        })
      if (request.docRev !== docRev)
        throw Object.assign(
          new Error(`页面内容已变（docRev ${request.docRev} → ${docRev}）；重新读取后再发命令`),
          { code: 'command.stale-index' },
        )
    }

    const depthBefore = handle.undoDepth()
    const outcome = await handle.applyOps(commands, request.dryRun === true)
    if (outcome['ok'] !== true)
      throw new Error(String(outcome['reason'] ?? '命令执行失败'))
    if (request.dryRun === true) {
      await ack(true, {
        applied: 0,
        mutated: false,
        dryRun: true,
        wouldApply:
          Number(outcome['cellChangeCount'] ?? 0) +
          (Array.isArray(outcome['structuralChanges']) ? (outcome['structuralChanges'] as unknown[]).length : 0),
      })
      return
    }

    const saved = await handle.save()
    if (!saved?.ok) {
      // 落盘失败：把本批撤销掉，页面保持 clean（否则 agent 之后的读会看到「已应用」）。
      let rolledBack = false
      for (let guard = 0; guard < 3 && handle.undoDepth() > depthBefore; guard += 1)
        rolledBack = handle.undo() || rolledBack
      throw new Error(
        `${saved?.error ?? '保存失败：工作簿未落盘'}` +
          (rolledBack ? '（已回滚本批修改）' : '（回滚失败：请人工确认）'),
      )
    }

    // 先把新令牌写回租约，再 ack —— 引擎在 ack 解析后读到的是写完后的值。
    await channel
      .report({ hash: bridge.hash(), docRev: handle.docRev(), structureEpoch: handle.structureEpoch(), dirty: false })
      .catch(() => undefined)

    const receipt = window.__xlsxProHost?.getReceipt?.() as
      | { ref?: string; hash?: string; size?: number; at?: number }
      | null
    await ack(true, {
      applied: commands.length,
      mutated: true,
      ...(receipt ? { receipt: { ref: receipt.ref ?? bridge.ref, hash: receipt.hash ?? bridge.hash(), size: receipt.size ?? 0, at: receipt.at ?? Date.now(), applied: commands.length } } : {}),
    })
  } catch (error) {
    await ack(false, { error: errorOf(error) })
  }
}

// ─── 读投影（页面 live 模型） ─────────────────────────────────────────

function readFromPage(
  channel: Channel,
  bridge: EditorBridge,
  request: PageRequest,
): AnyRecord {
  const handle = editor()
  if (!handle) throw Object.assign(new Error('编辑器尚未就绪'), { code: 'command.page-unavailable' })
  const ctx = (handle.context() ?? {}) as AnyRecord
  const base = {
    source: 'page',
    version: bridge.hash(),
    docRev: handle.docRev(),
    structureEpoch: handle.structureEpoch(),
  }
  const wanted = String(request.what ?? 'context')

  if (wanted === 'context' || wanted === 'meta') {
    const sheets = (ctx['sheets'] as Array<AnyRecord> | undefined) ?? []
    const lines = ['# workbook.xlsx', `sheets=${sheets.length}`]
    for (const sheet of sheets) {
      const extent =
        typeof sheet['rows'] === 'number' && typeof sheet['columns'] === 'number'
          ? ` rows=${String(sheet['rows'])} columns=${String(sheet['columns'])}`
          : ''
      lines.push(`- ${String(sheet['name'])} id=${String(sheet['id'])}${extent}`)
    }
    if (typeof ctx['sheetName'] === 'string')
      lines.push(`active=${ctx['sheetName']}${typeof ctx['selection'] === 'string' ? ` selection=${ctx['selection']}` : ''}`)
    if (typeof ctx['loadedRange'] === 'string') lines.push(`loaded=${ctx['loadedRange']}`)
    lines.push('（页面 live 模型：含未保存修改）')
    return {
      ...base,
      text: lines.join('\n'),
      sheets: handle.sheets(),
      ...(wanted === 'meta' ? { data: ctx } : {}),
    }
  }

  if (wanted !== 'range' && wanted !== 'formulas' && wanted !== 'find')
    throw Object.assign(
      new Error(`页面读不支持 what=${wanted}（支持：context / meta / range / formulas / find）`),
      { code: 'command.unsupported' },
    )

  const params = request.params ?? {}
  const sheets = (ctx['sheets'] as Array<AnyRecord> | undefined) ?? []
  const sheetName = String(params['sheet'] ?? ctx['sheetName'] ?? (sheets[0]?.['name'] as string) ?? '')
  const sheet = handle.sheets().find((entry) => entry.name === sheetName)
  if (!sheet) throw Object.assign(new Error(`工作表不存在：${sheetName}`), { code: 'command.not-found' })

  const text =
    typeof params['range'] === 'string' && params['range']
      ? params['range']
      : typeof ctx['loadedRange'] === 'string' && ctx['loadedRange']
        ? ctx['loadedRange']
        : 'A1:F30'
  const bounds = parseBounds(text)

  const addresses: string[] = []
  let truncated = false
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      if (addresses.length >= MAX_READ_CELLS) {
        truncated = true
        break
      }
      addresses.push(`${columnLabel(column)}${row + 1}`)
    }
    if (truncated) break
  }
  const cells = handle.readCells(addresses, sheet.id) ?? {}

  if (wanted === 'find') {
    const query = String(params['query'] ?? '')
    if (!query) throw new Error('find 缺少 query')
    const limit = Math.max(1, Math.min(Number(params['limit'] ?? 100), 1_000))
    const matches: string[] = []
    for (const address of addresses) {
      const cell = cells[address] as AnyRecord | undefined
      if (!cell) continue
      const value = String(cell['formula'] ?? cell['value'] ?? '')
      if (!value.toLowerCase().includes(query.toLowerCase())) continue
      matches.push(`${sheetName}!${address}\t${value}`)
      if (matches.length >= limit) break
    }
    return {
      ...base,
      text: [`# ${sheetName}!${text} 命中 ${matches.length}`, ...matches].join('\n'),
      matches,
      truncated,
    }
  }

  if (wanted === 'formulas') {
    const lines: string[] = []
    for (const address of addresses) {
      const cell = cells[address] as AnyRecord | undefined
      if (!cell?.['formula']) continue
      lines.push(`${sheetName}!${address}\t${String(cell['formula'])}`)
    }
    return { ...base, text: [`# ${sheetName} 公式 ${lines.length}`, ...lines].join('\n'), truncated }
  }

  const lines: string[] = [`# ${sheetName}!${text}`]
  const width = bounds.endColumn - bounds.startColumn + 1
  for (let offset = 0; offset < addresses.length; offset += width) {
    const row = addresses.slice(offset, offset + width).map((address) => {
      const cell = cells[address] as AnyRecord | undefined
      return String(cell?.['value'] ?? cell?.['display'] ?? '').replace(/\t/g, ' ').replace(/\n/g, ' ')
    })
    lines.push(`${bounds.startRow + Math.floor(offset / width) + 1}\t${row.join('\t')}`)
  }
  if (truncated) lines.push(`…（超过 ${MAX_READ_CELLS} 格，已截断）`)
  return { ...base, text: lines.join('\n'), truncated }
}

function parseBounds(text: string): {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
} {
  const [start, end = start] = text.split(':')
  const a = parseAddress(start ?? 'A1')
  const b = parseAddress(end ?? 'A1')
  return {
    startRow: Math.min(a.row, b.row),
    endRow: Math.max(a.row, b.row),
    startColumn: Math.min(a.column, b.column),
    endColumn: Math.max(a.column, b.column),
  }
}

function parseAddress(address: string): { row: number; column: number } {
  const match = /^([A-Z]+)(\d+)$/i.exec(address.trim())
  if (!match) return { row: 0, column: 0 }
  let column = 0
  for (const char of match[1]!.toUpperCase()) column = column * 26 + char.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, column: column - 1 }
}

function columnLabel(index: number): string {
  let label = ''
  let value = index + 1
  while (value > 0) {
    const rem = (value - 1) % 26
    label = String.fromCharCode(65 + rem) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

export { opTier }
