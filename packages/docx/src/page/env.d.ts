/// <reference types="vite/client" />

export interface Receipt {
  ref: string
  hash: string
  size: number
  at: number
  applied?: number
  warnings?: string[]
}

export interface CallError {
  code: string
  message: string
}

export type CallResult<T = unknown> =
  | { ok: true; data: T }
  | { ok: false; error: CallError }

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

export interface LoadResult {
  ref: string
  name: string
  size: number
  hash: string
  tier: 'normal' | 'lite' | 'readOnly' | 'refused'
  bytes?: Uint8Array
  data?: string
}

export interface PageRequestApply {
  type: 'apply'
  id: string
  commands: unknown[]
  baseHash?: string
  docRev?: number
  dryRun?: boolean
}

export interface PageRequestRead {
  type: 'read'
  id: string
  what: string
  params: Record<string, unknown>
}

export type PageRequest = PageRequestApply | PageRequestRead

export interface DocxProBridge {
  getCtx(): unknown
  getEditor(): unknown
  serialize(): Promise<Uint8Array | null>
  reload(bytes: Uint8Array, meta: { name: string; hash: string }): Promise<void>
  replaceDocument(bytes: Uint8Array): Promise<void>
}

declare global {
  interface Window {
    workStudio: WorkStudioBridge
    __docxPro?: DocxProBridge
  }
}

export {}
