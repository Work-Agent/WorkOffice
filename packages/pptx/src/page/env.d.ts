/// <reference types="vite/client" />

/** 页面桥返回：宿主 preload 不抛错，统一用 CallResult。 */
export interface CallError {
  code: string
  message: string
}

export type CallResult<T = unknown> = { ok: true; data: T } | { ok: false; error: CallError }

export type DomainUpdate =
  | { type: 'snapshot'; domain: string; resource: string; data: unknown }
  | { type: 'patch'; domain: string; resource: string; data: Record<string, unknown> }
  | { type: 'event'; domain: string; resource: string; event: string; data: unknown }

export interface WorkStudioBridge {
  readonly workspaceId: string | null
  command<T = unknown>(
    domain: string,
    resource: string,
    action: string,
    input?: unknown,
  ): Promise<CallResult<T>>
  subscribe(domain: string, resource: string): Promise<unknown>
  unsubscribe(domain: string, resource: string): void
  onUpdate(listener: (update: DomainUpdate) => void): () => void
}

export interface SaveReceipt {
  ref: string
  hash: string
  size: number
  at: number
  applied?: number
  warnings?: string[]
}

export interface LoadResult {
  ref: string
  name: string
  size: number
  hash: string
  bytes: Uint8Array
}

export type PageRequest =
  | {
      type: 'apply'
      id: string
      commands: unknown[]
      baseHash?: string
      docRev?: number
      dryRun?: boolean
    }
  | { type: 'read'; id: string; what: string; params: Record<string, unknown> }

export interface AttachResult {
  hash: string
  docRev: number
  pending: PageRequest[]
}

/** 页面启动数据：page/main.ts 注入，editor 只读。 */
export interface PptxProBoot {
  ref: string
  name: string
  pageId: string
  hash: string
  bytes: Uint8Array
}

export interface PptxProUi {
  getState(): {
    current: number
    selectedIds: string[]
    editing: unknown
    slides: unknown[]
  }
}

export interface PptxProHost {
  getHash(): string
  getDocRev(): number
  getBytes(): Uint8Array
  context(): unknown
  /** 页面 live 读投影（pptx.read 语义，与引擎同源）。 */
  read(params: Record<string, unknown>): unknown
  /** 命令目录（pptx.guide，不依赖文档内容）。 */
  guide(params: Record<string, unknown>): string
  /** 写前检查（pptx.check：审计 / 骨架读活模型）。 */
  check(params: Record<string, unknown>): Promise<string>
  applyBatch(commands: unknown[], dryRun?: boolean): Promise<{
    ok: boolean
    dryRun?: boolean
    plan?: unknown
    failures?: Array<{ index: number; error: string }>
    applied?: number
    error?: string
    slides?: unknown[]
  }>
  save(): Promise<{ ok: boolean; path?: string; error?: string; receipt?: SaveReceipt }>
  undo(): Promise<unknown[] | null>
  redo(): Promise<unknown[] | null>
}

declare global {
  interface Window {
    workStudio: WorkStudioBridge
    __pptxProBoot?: PptxProBoot
    __pptxProHost?: PptxProHost
    __pptxProUi?: PptxProUi
  }
}

export {}
