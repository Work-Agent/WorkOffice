/// <reference types="vite/client" />

export interface PageSessionInfo {
  id: string
  ref: string
  version: string
  /** 提交边界的基线代数（契约 §8.3）；页面只镜像，不自增。 */
  epoch: number
  /** 命中的 codec id；null = 未经解码（raw）。 */
  codec: string | null
  surface: string
  owner?: string
  page: string
  data: unknown
  degraded: boolean
  /** 声明位推导的写法：页面据此决定调 save 还是 commit。 */
  write: 'none' | 'bytes' | 'plan'
  /** 资源体积（`bytes` 分块读取用）。 */
  size?: number
  /** 页面内可调的只读工具（所属插件声明）。 */
  tools?: string[]
  tier?: 'normal' | 'lite' | 'readOnly' | 'refused'
}

/** 宿主提交回执（契约 §8.1）。 */
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
    | { type: 'plan'; plan?: unknown; tool?: string; params?: unknown }
  origin?: 'page' | 'agent'
  expect?: { docRev?: string; epoch?: number }
  dryRun?: boolean
}

export interface UraChangeEvent {
  ref: string
  type: 'content' | 'structure' | 'state'
  version?: string
  at: number
}

export interface AppBridgeApi {
  available(): Promise<boolean>
  call<T = unknown>(path: string, input?: unknown): Promise<T>
  peek<T = unknown>(path: string): Promise<T | null>
}

/** 主 → 页的 agent 请求（契约 §7.5）。 */
export interface AgentContextRequest {
  type: 'agent-request'
  id: string
  kind: 'context' | 'apply'
  txnId: string
  version?: string
  epoch?: number
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
      open<T = unknown>(ref: string): Promise<T>
      save<T = unknown>(id: string, content: string): Promise<T>
      saveBytes<T = unknown>(id: string, data: Uint8Array): Promise<T>
      /** 计划型提交：页面产 plan，宿主锁内 exec → 落盘 → 回执（含 epoch / sync）。 */
      commit<T = unknown>(id: string, request: PageCommitInput): Promise<T>
      /** 分块读（单块 ≤ 1 MB）：大内容不进 get 的 data。 */
      bytes(
        id: string,
        window?: { start?: number; end?: number },
      ): Promise<Uint8Array>
      /** 只读调用：只能调本插件声明为 readOnly 的工具。 */
      invoke<T = unknown>(id: string, tool: string, input?: unknown): Promise<T>
      /** agent-request 回执（页面代理；失败也要如实回报）。 */
      agentResult(id: string, message: AgentResultMessage): Promise<unknown>
      reload<T = unknown>(id: string): Promise<T>
      close(id: string): Promise<unknown>
    }
    watch(ref?: string): Promise<{ ref: string } | null>
    onChange(cb: (event: UraChangeEvent) => void): () => void
  }
  bus: {
    on(
      cb: (message: { type: string; [key: string]: unknown }) => void,
    ): () => void
    send(message: { type: string; [key: string]: unknown }): void
    loaded(): void
  }
}

declare global {
  interface Window {
    extension: ExtensionBridge
    appBridge?: AppBridgeApi
    /** docs 编辑器副本暴露给页面桥的句柄（见 src/editor/App.tsx）。 */
    __docxPro?: DocxProBridge
  }
}

/** App.tsx 注册的编辑器句柄：agent 桥唯一的取数入口。 */
export interface DocxProBridge {
  /** live FileActionContext（每个 render 刷新）。 */
  getCtx(): unknown
  /** 当前编辑器实例。 */
  getEditor(): unknown
  /** 序列化当前文档（纯函数，不改状态）。 */
  serialize(): Promise<Uint8Array | null>
}

export {}
