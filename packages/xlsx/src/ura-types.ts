// ─── URA 契约子集（插件侧） ──────────────────────────────────────────
//
// 与 WorkStudio src/shared/ura.ts 的结构保持一致；插件包不依赖宿主源码，
// 只依赖这份最小结构契约（字段名以宿主为准）。

export type DocType = `${string}@${number}`
export type PatchKind = 'text' | 'bytes'

export interface RawContent {
  bytes: Uint8Array
  version: string
  formatHints: string[]
  mime?: string
}

export type Patch =
  | { type: 'replace-all'; bytes: Uint8Array }
  | {
      type: 'replace-range'
      start: number
      end: number
      bytes: Uint8Array
      expected?: string
    }

export type Anchor = { start: number; end: number; expected?: string }

export type Change =
  | { type: 'replace-all'; content: string | Uint8Array }
  | { type: 'replace-range'; anchor: Anchor; content: string | Uint8Array }

export interface DecodedDoc {
  docType: DocType
  doc: unknown
  projection: unknown
  outline?: unknown
}

export interface Span {
  start: number
  end: number
}

export interface Codec {
  readonly hints: string[]
  readonly extensions?: string[]
  readonly docType: DocType
  readonly patchType: PatchKind
  decode(raw: RawContent, ref: string): Promise<DecodedDoc> | DecodedDoc
  locate(decoded: DecodedDoc, anchor: Anchor): Span
  encode(decoded: DecodedDoc, change: Change): Patch
}

export interface AgentProjection {
  type: 'text'
  content: string
  truncated?: boolean
}

export interface PageProjection {
  type: 'page'
  surface: string
  owner?: string
  page: string
  data: unknown
  actions?: string[]
}

export interface BrowserProjection {
  type: 'browser'
  surface: string
  url: string
}

export type EditorProjection = PageProjection | BrowserProjection
export type Projection = AgentProjection | EditorProjection

export interface Consumer {
  readonly docType: DocType
  readonly type: 'agent' | 'editor'
  readonly requires?: string[]
  readonly surface?: string
  project(
    decoded: DecodedDoc,
    raw: RawContent,
    ref: string,
  ): Promise<Projection> | Projection
}

export interface Query {
  select?: string[]
  where?: Record<string, unknown>
  limit?: number
}

export interface Store {
  readonly docType: DocType
  readonly schema: unknown
  query(decoded: DecodedDoc, query: Query): Promise<unknown> | unknown
}

export interface Transport {
  readonly scheme: string
  readonly capabilities?: string[]
  stat(ref: string): Promise<unknown>
  read(ref: string): Promise<RawContent>
  write(ref: string, patch: Patch, opts: unknown): Promise<{ version: string }>
  list(ref: string): Promise<unknown[]>
  watch?(ref: string, emit: (event: unknown) => void): () => void
}

export interface CollectionProvider {
  readonly id: string
  readonly label: string
  readonly icon?: string
  list(ctx: unknown): unknown
}

export interface ViewAction {
  readonly id: string
  readonly label: string
  readonly icon?: string
  readonly target: string
  readonly toggle?: boolean
  run(ctx: unknown): Promise<void> | void
}

// ─── GenOffice docx-engine 的最小结构（避免插件类型依赖宿主/源码路径） ───

export interface RunLike {
  text: string
}

export interface TableCellLike {
  paras: string[]
}

export interface TableLike {
  rows: TableCellLike[][]
}

export interface BlockLike {
  type: 'paragraph' | 'heading' | 'listItem' | 'table' | 'image' | 'passthrough'
  level?: number
  runs?: RunLike[]
  table?: TableLike
  previewText?: string
}

export interface ParsedDocLike {
  blocks: BlockLike[]
}

// ─── GenOffice xlsx-gateway 的最小结构 ───────────────────────────────

export type CellValue = string | number | boolean | null | undefined

export interface CellStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontColor?: string
  fillColor?: string
  fontSize?: number
  fontFamily?: string
  align?: 'left' | 'center' | 'right'
  wrap?: boolean
  numberFormat?: string
}

export interface CellLike {
  value: CellValue
  formula?: string | undefined
  rawValue?: CellValue
  style?: CellStyle
}

export interface SheetLike {
  id: string
  name: string
  /** ZIP 内的 worksheet 路径（编辑写回用） */
  path?: string
  cells: Readonly<Record<string, CellLike>>
  gridRows?: number
  gridColumns?: number
  styles?: Readonly<Record<string, unknown>>
  /** 合并区域，A1:C1 形式 */
  merges?: readonly string[]
  /** 列宽（px，按 A/B/C 键） */
  colWidths?: Readonly<Record<string, number>>
  /** 行高（px，按 1/2/3 键） */
  rowHeights?: Readonly<Record<string, number>>
}
