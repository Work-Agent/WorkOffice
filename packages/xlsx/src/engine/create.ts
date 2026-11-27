// ─── xlsx.create：从空白模板建一个新工作簿 ──────────────────────
//
// 单文件范围：目标 ref 必须不存在、父目录必须已存在、必须落在当前工作区
// （`ctx.fs.write` 的边界）。内容 = seed.xlsx + 初始工作表 / 单元格。

import { readFileSync } from 'node:fs'

import type { PluginContext } from '@work-studio/sdk'
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'

import type { XlsxSaveRequest } from '../shared/save-request'
import { CommandError } from './errors'
import { bytesOfSaveRequest } from './gateway'
import { sha256Hex } from './hash'
import { withLock } from './lock'
import { assertExt, pathOf } from './ref'
import { dirname } from 'node:path'
import { stat } from 'node:fs/promises'

export interface NewSheetInput {
  name: string
  rows?: unknown[][]
}

let SEED: Uint8Array | null = null

export function seedBytes(): Uint8Array {
  SEED ??= new Uint8Array(readFileSync(new URL('./seed.xlsx', import.meta.url)))
  return SEED
}

function cellEdit(
  sheetName: string,
  address: string,
  cell: { value: string | number | boolean | null; formula?: string },
) {
  const match = /^([A-Z]+)(\d+)$/.exec(address)
  if (!match) throw new CommandError('command.invalid', `非法地址：${address}`)
  let column = 0
  for (const char of match[1]!) column = column * 26 + char.charCodeAt(0) - 64
  return {
    sheetName,
    row: Number(match[2]) - 1,
    column: column - 1,
    writeValue: true,
    cell,
  }
}

/** 解析 `sheets` / `rows` / `csv` 三种入参，统一成工作表列表。 */
export function normalizeSheets(input: Record<string, unknown>): NewSheetInput[] {
  // 助手 CLI 只能传字符串：允许 JSON 形式的 sheets / rows。
  const rawSheets = typeof input['sheets'] === 'string' ? parseJson(input['sheets'], 'sheets') : input['sheets']
  const rawRows = typeof input['rows'] === 'string' ? parseJson(input['rows'], 'rows') : input['rows']
  const normalized = { ...input, sheets: rawSheets, rows: rawRows }
  return sheetsOf(normalized)
}

function parseJson(text: string, label: string): unknown {
  try {
    return JSON.parse(text) as unknown
  } catch (error) {
    throw new CommandError('command.invalid', `${label} 不是合法 JSON：${error instanceof Error ? error.message : String(error)}`)
  }
}

function sheetsOf(input: Record<string, unknown>): NewSheetInput[] {
  if (Array.isArray(input['sheets'])) {
    return (input['sheets'] as unknown[]).map((entry, index) => {
      if (!entry || typeof entry !== 'object')
        throw new CommandError('command.invalid', `sheets[${index}] 不是对象`)
      const record = entry as Record<string, unknown>
      const name = String(record['name'] ?? (index === 0 ? 'Sheet1' : `Sheet${index + 1}`))
      const rows = Array.isArray(record['rows']) ? (record['rows'] as unknown[][]) : []
      return { name, rows: rows.map((row) => (Array.isArray(row) ? row : [])) }
    })
  }
  const csv = typeof input['csv'] === 'string' ? input['csv'] : null
  if (csv !== null)
    return [{ name: String(input['name'] ?? 'Sheet1'), rows: parseCsv(csv) }]
  const rows = Array.isArray(input['rows']) ? (input['rows'] as unknown[][]) : []
  return [{ name: String(input['name'] ?? 'Sheet1'), rows: rows.map((row) => (Array.isArray(row) ? row : [])) }]
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
        } else quoted = false
      } else field += ch
      continue
    }
    if (ch === '"') quoted = true
    else if (ch === ',') {
      row.push(cellValue(field))
      field = ''
    } else if (ch === '\n') {
      row.push(cellValue(field))
      rows.push(row)
      row = []
      field = ''
    } else if (ch !== '\r') field += ch
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

export interface CreateResult {
  ref: string
  hash: string
  size: number
  at: number
  applied: number
}

export async function create(
  input: Record<string, unknown>,
  ctx: PluginContext,
): Promise<CreateResult> {
  const ref = String(input['ref'] ?? '')
  if (!ref) throw new CommandError('command.invalid', 'create 缺少 ref')
  const path = pathOf(ref)
  assertExt(path)

  const existing = await stat(path).catch(() => null)
  if (existing)
    throw new CommandError('command.invalid', '目标已存在；修改已有工作簿请用 xlsx.apply')

  const sheets = normalizeSheets(input)
  if (sheets.length === 0) sheets.push({ name: 'Sheet1', rows: [] })
  for (const sheet of sheets) {
    if (!sheet.name || sheet.name.length > 31)
      throw new CommandError('command.invalid', `工作表名非法：${sheet.name || '(空)'}`)
  }

  const seed = seedBytes()
  const first = sheets[0]!
  const request: XlsxSaveRequest = { edits: [], structuralOps: [] }
  const renames = first.name === 'Sheet1' ? [] : [{ sheetName: 'Sheet1', newName: first.name }]
  const additions = sheets.slice(1).map((sheet) => ({ name: sheet.name }))
  const order = [first.name, ...sheets.slice(1).map((sheet) => sheet.name)]
  if (renames.length > 0 || additions.length > 0) {
    request.sheetPlan = {
      renames,
      additions,
      removals: [],
      order,
      orderChanged: additions.length > 0,
    }
  }

  let applied = 0
  for (const sheet of sheets) {
    const targetSheet = sheet === first ? 'Sheet1' : sheet.name
    ;(sheet.rows ?? []).forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        const address = formatAddress(rowIndex, columnIndex)
        applied += 1
        if (typeof value === 'string' && value.startsWith('='))
          request.edits.push(cellEdit(targetSheet, address, { value: null, formula: value }))
        else
          request.edits.push(
            cellEdit(targetSheet, address, {
              value: (value ?? null) as string | number | boolean | null,
            }),
          )
      })
    })
  }

  const bytes =
    request.edits.length === 0 && !request.sheetPlan
      ? seed
      : (await bytesOfSaveRequest(seed, request)).bytes

  return withLock(ref, async () => {
    // 父目录必须已存在：本插件不做目录操作。
    const parent = await stat(dirname(path)).catch(() => null)
    if (!parent?.isDirectory())
      throw new CommandError('command.invalid', `父目录不存在：${dirname(path)}`)

    await ctx.fs.write(ref, bytes)
    const hash = sha256Hex(bytes)
    ctx.changed?.(ref, 'content')
    return { ref, hash, size: bytes.byteLength, at: Date.now(), applied }
  })
}
