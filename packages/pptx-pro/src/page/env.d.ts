/// <reference types="vite/client" />

export interface PptxPageMeta {
  kind: 'pptx'
  name: string
  ref: string
  size: number
  version: string
  slideCount: number
  runtime?: unknown
}

export interface PageSessionInfo {
  id: string
  ref: string
  version: string
  epoch: number
  codec: string | null
  surface: string
  owner?: string
  page: string
  data: PptxPageMeta
  degraded: boolean
  write: 'none' | 'bytes' | 'plan'
  size?: number
  tools?: string[]
  tier?: 'normal' | 'lite' | 'readOnly' | 'refused'
}

export interface Receipt {
  ref: string
  txnId: string
  version: string
  epoch: number
  applied: number
  rejected?: Array<{ index: number; code: string; message: string }>
  warnings?: string[]
  notices?: string[]
  audit?: unknown
  sync?: {
    kind: 'none' | 'version' | 'manifest' | 'bytes'
    blocks?: unknown[]
    dataUrl?: string
  }
}

export interface PageCommitInput {
  txnId: string
  baseVersion?: string
  payload:
    | { type: 'bytes'; bytes: Uint8Array }
    | {
        type: 'plan'
        plan?: unknown
        tool?: string
        params?: unknown
        options?: unknown
        command?: 'undo' | 'redo'
      }
  origin?: 'page' | 'agent'
  expect?: { docRev?: string; epoch?: number }
  dryRun?: boolean
}

export interface AgentContextRequest {
  type: 'agent-request'
  id: string
  kind: 'context' | 'read' | 'apply'
  txnId: string
  version?: string
  epoch?: number
  /** kind:'read' 的工具名（宿主透传）；缺省时按参数形状判断。 */
  tool?: string
  /** kind:'read' 的读请求 */
  read?: Record<string, unknown>
  commands?: unknown[]
  dryRun?: boolean
}

export interface AgentResultMessage {
  type: 'agent-result'
  id: string
  ok: boolean
  context?: unknown
  receipt?: Receipt
  error?: { code: string; message: string }
}

export interface ExtensionBridge {
  ura: {
    page: {
      get<T = unknown>(id: string): Promise<T>
      bytes(id: string, window?: { start?: number; end?: number }): Promise<Uint8Array>
      commit<T = unknown>(id: string, request: PageCommitInput): Promise<T>
      agentResult(id: string, message: AgentResultMessage): Promise<unknown>
      close(id: string): Promise<unknown>
    }
    watch(ref?: string): Promise<{ ref: string } | null>
    onChange(cb: (event: { ref: string; type: string; version?: string; at: number }) => void): () => void
  }
  bus: {
    on(cb: (message: { type: string; [key: string]: unknown }) => void): () => void
    send(message: { type: string; [key: string]: unknown }): void
    loaded(): void
  }
}

export interface PptxProBoot {
  sessionId: string
  data: PptxPageMeta
  bytes: Uint8Array
  version: string
  epoch: number
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
  getVersion(): string
  getEpoch(): number
  getDocRev(): number
  getReceipt(): Receipt | null
  getBytes(): Uint8Array
  setOrigin(origin: 'page' | 'agent'): void
  context(): unknown
  /** 页面 live 读投影（pptx_read 语义，与执行点同源）。 */
  read(params: Record<string, unknown>): unknown
  /** 命令目录（pptx_guide，不依赖文档内容）。 */
  guide(params: Record<string, unknown>): string
  /** 设计稿检查（pptx_check：审计 / 骨架读活模型，outline / pages 对账不看文档）。 */
  check(params: Record<string, unknown>): Promise<string>
  applyAgentTxn(batches: unknown[][], dryRun?: boolean): Promise<{
    ok: boolean
    dryRun?: boolean
    plan?: unknown
    failures?: Array<{ index: number; error: string }>
    applied?: number
    receipt?: Receipt
    error?: string
    slides?: unknown[]
  }>
  save(): Promise<{ ok: boolean; path?: string; error?: string; slides?: unknown[] }>
  undo(): Promise<unknown[] | null>
  redo(): Promise<unknown[] | null>
}

declare global {
  interface Window {
    extension: ExtensionBridge
    appBridge?: {
      available(): Promise<boolean>
      call<T = unknown>(path: string, input?: unknown): Promise<T>
      peek<T = unknown>(path: string): Promise<T | null>
    }
    __pptxProBoot?: PptxProBoot
    __pptxProHost?: PptxProHost
    __pptxProUi?: PptxProUi
  }
}

export {}
