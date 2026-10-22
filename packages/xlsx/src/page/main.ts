// ─── XLSX 网格页 ────────────────────────────────────────────────────
//
// 页面只消费 codec 投影出的 sheet/cell 快照：标签、选区、滚动、主题
// 全部由 WorkStudio 侧实现；不引入 GenOffice 的 Univer 外壳。

import type { PageSessionInfo } from './env'
import type { SheetLike } from '../ura-types'

interface XlsxPageData {
  name: string
  ref: string
  version: string
  sheets: SheetLike[]
}

const MAX_ROWS = 200
const MAX_COLS = 40

const root = document.getElementById('root')!

function fail(message: string): void {
  root.innerHTML = ''
  const boot = document.createElement('div')
  boot.className = 'boot'
  boot.dataset.error = '1'
  boot.textContent = message
  root.append(boot)
}

function applyTheme(theme: 'light' | 'dark'): void {
  document.documentElement.dataset.theme = theme
  document.documentElement.style.colorScheme = theme
}

async function currentTheme(): Promise<'light' | 'dark'> {
  try {
    const ui = await window.extension.invoke<{ theme?: 'light' | 'dark' }>(
      'app:ui.get',
    )
    if (ui?.theme === 'light' || ui?.theme === 'dark') return ui.theme
  } catch {
    /* 系统主题兜底 */
  }
  return window.matchMedia('(prefers-color-scheme: dark)').matches
    ? 'dark'
    : 'light'
}

function columnLabel(index: number): string {
  let label = ''
  let value = index
  while (value > 0) {
    const rem = (value - 1) % 26
    label = String.fromCharCode(65 + rem) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

function addressOf(row: number, col: number): string {
  return `${columnLabel(col)}${row}`
}

function parseAddress(address: string): { row: number; col: number } | null {
  const match = /^([A-Z]+)(\d+)$/.exec(address)
  if (!match) return null
  let col = 0
  for (const ch of match[1]!) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { row: Number(match[2]), col }
}

function displayCell(cell: { value?: unknown; formula?: string } | undefined): {
  text: string
  type: 'text' | 'num' | 'bool'
} {
  if (!cell) return { text: '', type: 'text' }
  if (cell.formula) return { text: cell.formula, type: 'text' }
  const value = cell.value
  if (value === null || value === undefined) return { text: '', type: 'text' }
  if (typeof value === 'number') return { text: String(value), type: 'num' }
  if (typeof value === 'boolean') {
    return { text: value ? 'TRUE' : 'FALSE', type: 'bool' }
  }
  return { text: String(value), type: 'text' }
}

function boundsOf(sheet: SheetLike): { rows: number; cols: number; clipped: boolean } {
  let maxRow = 1
  let maxCol = 1
  for (const address of Object.keys(sheet.cells ?? {})) {
    const parsed = parseAddress(address)
    if (!parsed) continue
    maxRow = Math.max(maxRow, parsed.row)
    maxCol = Math.max(maxCol, parsed.col)
  }
  const rows = Math.min(Math.max(maxRow + 2, 20), MAX_ROWS)
  const cols = Math.min(Math.max(maxCol + 1, 8), MAX_COLS)
  return {
    rows,
    cols,
    clipped: maxRow + 2 > MAX_ROWS || maxCol + 1 > MAX_COLS,
  }
}

function buildSheetView(
  sheet: SheetLike,
  selected: string | null,
  onSelect: (address: string) => void,
): HTMLElement {
  const { rows, cols, clipped } = boundsOf(sheet)
  const wrap = document.createElement('div')
  wrap.className = 'grid-wrap'

  const table = document.createElement('table')
  const thead = document.createElement('thead')
  const headRow = document.createElement('tr')
  const corner = document.createElement('th')
  corner.className = 'corner row-head'
  headRow.append(corner)
  for (let col = 1; col <= cols; col++) {
    const th = document.createElement('th')
    th.textContent = columnLabel(col)
    headRow.append(th)
  }
  thead.append(headRow)
  table.append(thead)

  const tbody = document.createElement('tbody')
  for (let row = 1; row <= rows; row++) {
    const tr = document.createElement('tr')
    const head = document.createElement('th')
    head.className = 'row-head'
    head.textContent = String(row)
    tr.append(head)
    for (let col = 1; col <= cols; col++) {
      const address = addressOf(row, col)
      const td = document.createElement('td')
      const cell = sheet.cells?.[address]
      const display = displayCell(cell)
      td.textContent = display.text
      if (display.type !== 'text') td.classList.add(display.type)
      if (address === selected) td.classList.add('sel')
      td.dataset.address = address
      td.addEventListener('click', () => onSelect(address))
      const style = sheet.styles?.[address] as
        | {
            bold?: boolean
            italic?: boolean
            fontColor?: string
            fillColor?: string
          }
        | undefined
      if (style?.bold) td.style.fontWeight = '600'
      if (style?.italic) td.style.fontStyle = 'italic'
      if (style?.fontColor) td.style.color = style.fontColor.startsWith('#') ? style.fontColor : `#${style.fontColor}`
      if (style?.fillColor) td.style.background = style.fillColor.startsWith('#') ? style.fillColor : `#${style.fillColor}`
      tr.append(td)
    }
    tbody.append(tr)
  }
  table.append(tbody)
  wrap.append(table)

  if (clipped) {
    const note = document.createElement('div')
    note.className = 'truncated'
    note.textContent = `只显示前 ${MAX_ROWS} 行 × ${MAX_COLS} 列；完整数据仍在文件中。`
    wrap.append(note)
  }
  return wrap
}

function mount(data: XlsxPageData): void {
  root.innerHTML = ''
  const app = document.createElement('div')
  app.className = 'app'

  const toolbar = document.createElement('div')
  toolbar.className = 'toolbar'
  const name = document.createElement('div')
  name.className = 'file-name'
  name.textContent = data.name
  name.title = data.ref
  const badge = document.createElement('span')
  badge.className = 'badge'
  badge.textContent = '只读预览'
  const spacer = document.createElement('span')
  spacer.className = 'spacer'
  toolbar.append(name, badge, spacer)

  const formula = document.createElement('div')
  formula.className = 'formula'
  const address = document.createElement('div')
  address.className = 'address'
  address.textContent = 'A1'
  const value = document.createElement('div')
  value.className = 'value'
  formula.append(address, value)

  const tabs = document.createElement('div')
  tabs.className = 'tabs'

  const content = document.createElement('div')
  content.style.cssText = 'flex:1;min-height:0;display:flex;flex-direction:column'

  app.append(toolbar, formula, tabs, content)
  root.append(app)

  let activeId = data.sheets[0]?.id ?? ''
  let selected: string | null = null

  const sheetOf = (id: string): SheetLike | undefined =>
    data.sheets.find((sheet) => sheet.id === id)

  const updateFormula = (): void => {
    const sheet = sheetOf(activeId)
    const cell = selected && sheet ? sheet.cells?.[selected] : undefined
    address.textContent = selected ?? '—'
    value.textContent = cell
      ? (cell.formula ?? String(cell.value ?? ''))
      : ''
  }

  const render = (): void => {
    const sheet = sheetOf(activeId)
    content.replaceChildren()
    if (!sheet) {
      const empty = document.createElement('div')
      empty.className = 'boot'
      empty.textContent = '工作簿没有可读工作表'
      content.append(empty)
      updateFormula()
      return
    }
    content.append(
      buildSheetView(sheet, selected, (next) => {
        selected = next
        updateFormula()
        for (const td of content.querySelectorAll('td.sel')) {
          td.classList.remove('sel')
        }
        content
          .querySelector(`td[data-address="${next}"]`)
          ?.classList.add('sel')
      }),
    )
    updateFormula()
  }

  for (const sheet of data.sheets) {
    const tab = document.createElement('button')
    tab.className = 'tab'
    tab.textContent = sheet.name
    tab.classList.toggle('on', sheet.id === activeId)
    tab.addEventListener('click', () => {
      activeId = sheet.id
      selected = null
      for (const node of tabs.querySelectorAll('.tab')) {
        node.classList.toggle('on', node === tab)
      }
      render()
    })
    tabs.append(tab)
  }

  render()

  window.addEventListener('keydown', (event) => {
    if (!selected) return
    const parsed = parseAddress(selected)
    if (!parsed) return
    const delta =
      event.key === 'ArrowUp'
        ? [0, -1]
        : event.key === 'ArrowDown'
          ? [0, 1]
          : event.key === 'ArrowLeft'
            ? [-1, 0]
            : event.key === 'ArrowRight'
              ? [1, 0]
              : null
    if (!delta) return
    event.preventDefault()
    const row = Math.max(1, parsed.row + delta[1]!)
    const col = Math.max(1, parsed.col + delta[0]!)
    const next = addressOf(row, col)
    selected = next
    const td = content.querySelector(`td[data-address="${next}"]`)
    if (td) {
      for (const node of content.querySelectorAll('td.sel')) {
        node.classList.remove('sel')
      }
      td.classList.add('sel')
      td.scrollIntoView({ block: 'nearest', inline: 'nearest' })
    }
    updateFormula()
  })
}

async function boot(): Promise<void> {
  const sessionId = new URLSearchParams(location.search).get('session')
  if (!sessionId) {
    fail('缺少页面会话：请从 WorkStudio 资源面板打开 XLSX 工作簿')
    return
  }
  try {
    applyTheme(await currentTheme())
    const info = await window.extension.ura.page.get<PageSessionInfo>(sessionId)
    const data = info.data as XlsxPageData | null
    if (!data?.sheets) throw new Error('XLSX 页面数据缺少工作表快照')
    mount(data)
    window.extension.bus.on((message) => {
      if (message.type === 'theme') {
        applyTheme(message.theme === 'dark' ? 'dark' : 'light')
      }
    })
    window.extension.bus.loaded()
    window.extension.bus.send({ type: 'dirty', dirty: false })
  } catch (error) {
    fail(
      `XLSX 打开失败：${error instanceof Error ? error.message : String(error)}`,
    )
  }
}

void boot()
