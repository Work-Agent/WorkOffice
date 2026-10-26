// ─── WorkStudio URA 契约子集（插件侧） ───────────────────────────────
//
// 与 WorkStudio src/shared/ura.ts 对齐：插件包不依赖宿主源码，只依赖这份
// 结构化契约。注册对象最终由宿主注册表校验，字段名以宿主为准。
//
// Office 插件当前使用 codec / presenter / sidecar 三个注册面；其余类型
// 按同一契约保留，后续扩展 transport / store / collection / action 时
// 不需要再改契约文件。

export type Mime = string
export type TransportId = string
export type PatchKind = 'text' | 'bytes'

export const MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
} as const

// ─── 来源 ────────────────────────────────────────────────────────────

export type ResourceType = 'file' | 'dir' | 'other'

export interface ResourceStat {
  version: string
  type: ResourceType
  size: number
  mtimeMs?: number
}

export interface RawContent extends ResourceStat {
  bytes: Uint8Array
  mime?: Mime
}

export interface ListEntry {
  name: string
  ref: string
  type: ResourceType
  size?: number
  version?: string
}

export type TransportWriteOptions = Record<string, never>

export interface TransportWriteResult {
  version: string
}

export type ChangeType = 'content' | 'structure' | 'state'

export interface ChangeEvent {
  ref: string
  type: ChangeType
  version?: string
  at: number
}

export interface Transport {
  readonly id?: TransportId
  readonly scheme: string
  readonly capabilities?: readonly string[]
  stat(ref: string): Promise<ResourceStat>
  read(ref: string): Promise<RawContent>
  write(
    ref: string,
    patch: Patch,
    opts: TransportWriteOptions,
  ): Promise<TransportWriteResult>
  list(ref: string): Promise<ListEntry[]>
  watch?(ref: string, emit: (event: ChangeEvent) => void): () => void
}

// ─── 传输层补丁 ──────────────────────────────────────────────────────

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

// ─── 段键 ────────────────────────────────────────────────────────────
//
// codec 认 (transport, mime)；store / presenter / sidecar 认
// (transport, mime, codec)。'*' 表示不限；codec 空串表示不经解码。

export interface CodecKey {
  transport: TransportId | '*'
  mime: Mime
}

export interface OutletKey {
  transport: TransportId | '*'
  mime: Mime | '*'
  /** codec id；空串 = 不经解码。 */
  codec: string
}

// ─── 格式 ────────────────────────────────────────────────────────────

export interface DecodedDoc {
  /** 结构化文档（插件自定义）。 */
  doc: unknown
  /** 锚点映射：供 locate 反查锚点，同时作为 grep 的文本来源。 */
  anchorMap: unknown
  outline?: unknown
}

export interface Span {
  start: number
  end: number
}

export interface Codec {
  /** 实现身份：store / presenter / sidecar 都以它作键的一部分。 */
  readonly id: string
  /** 能解析的 (transport, mime) 组合。 */
  readonly keys: readonly CodecKey[]
  /** 扩展名（不含点；用于资源地图的格式标注）。 */
  readonly extensions?: string[]
  readonly patchType: PatchKind
  decode(raw: RawContent, ref: string): Promise<DecodedDoc> | DecodedDoc
  locate(decoded: DecodedDoc, anchor: Anchor): Span
  encode(decoded: DecodedDoc, change: Change): Patch
}

// ─── 数据 ────────────────────────────────────────────────────────────

export interface Query {
  select?: string[]
  where?: Record<string, unknown>
  limit?: number
}

export interface Store {
  readonly key: OutletKey
  readonly schema: unknown
  query(decoded: DecodedDoc, query: Query): Promise<unknown> | unknown
}

// ─── 出口 ────────────────────────────────────────────────────────────

export interface Sidecar {
  type: 'text'
  content: string
  truncated?: boolean
}

export interface PagePresentation {
  type: 'page'
  surface: string
  /** 渲染器归属插件（页面与动作的作用域）。 */
  owner?: string
  /** 页面入口（由插件清单解析；raw 降级时为空）。 */
  page: string
  data: unknown
  actions?: string[]
}

export interface Browser {
  type: 'browser'
  surface: string
  url: string
}

export type Presentation = PagePresentation | Browser

/** 表现层实现：资源页（渲染 / 编辑）。 */
export interface Presenter {
  readonly key: OutletKey
  readonly type: 'page' | 'browser'
  readonly surface?: string
  readonly requires?: string[]
  present(
    decoded: DecodedDoc,
    raw: RawContent,
    ref: string,
  ): Promise<Presentation> | Presentation
}

/** agent 读取面：只负责“agent 读到什么”。 */
export interface SidecarProvider {
  readonly key: OutletKey
  sidecar(
    decoded: DecodedDoc,
    raw: RawContent,
    ref: string,
  ): Promise<Sidecar> | Sidecar
}

// ─── 集合 / 动作（注册面保留） ────────────────────────────────────────

export interface CollectionItem {
  ref: string
  title: string
}

export interface CollectionGroup {
  id: string
  title: string
  icon?: string
  order?: number
  items: CollectionItem[]
}

export interface TreeItem {
  ref: string
  name: string
  type: ResourceType
  ext: string
  size?: number
}

export interface CollectionContext {
  workspace: string
  tree(ref: string, limit?: number): Promise<TreeItem[]>
  find(
    pattern: string,
    options?: { scope?: string; limit?: number },
  ): Promise<ListEntry[]>
  settings: {
    get<T>(key: string, fallback: T): Promise<T>
  }
}

export interface CollectionView {
  source: string
  owner: string
  label: string
  icon?: string
  groups: CollectionGroup[]
}

export interface CollectionProvider {
  readonly id: string
  readonly label: string
  readonly icon?: string
  list(ctx: CollectionContext): CollectionGroup[] | Promise<CollectionGroup[]>
}

export interface ViewActionContext {
  viewId: string
  type: 'browser' | 'extension'
  /** 本次切换后的状态（toggle 动作）。 */
  on: boolean
  setState(state: string): void
  execute<T = unknown>(script: string): Promise<T>
}

export interface ViewAction {
  readonly id: string
  readonly label: string
  readonly icon?: string
  readonly target: string
  readonly toggle?: boolean
  run(ctx: ViewActionContext): Promise<void> | void
}

// ─── XLSX 页面 / codec 共用的单元格结构 ──────────────────────────────

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
  /** ZIP 内的 worksheet 路径（编辑写回用）。 */
  path?: string
  cells: Readonly<Record<string, CellLike>>
  gridRows?: number
  gridColumns?: number
  styles?: Readonly<Record<string, unknown>>
  /** 合并区域，A1:C1 形式。 */
  merges?: readonly string[]
  /** 列宽（px，按 A/B/C 键）。 */
  colWidths?: Readonly<Record<string, number>>
  /** 行高（px，按 1/2/3 键）。 */
  rowHeights?: Readonly<Record<string, number>>
}
