/// <reference types="vite/client" />

import type { XlsxPageData, XlsxProReceipt } from '../editor/workstudio-host'

export interface PageSessionInfo {
  id: string
  ref: string
  version: string
  epoch: number
  codec: string | null
  surface: string
  owner?: string
  page: string
  data: unknown
  degraded: boolean
  write: 'none' | 'bytes' | 'plan'
  size?: number
  tools?: string[]
  tier?: 'normal' | 'lite' | 'readOnly' | 'refused'
}

export interface Receipt extends XlsxProReceipt {}

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

export interface AgentContextRequest {
  type: 'agent-request'
  id: string
  /** context = 概览；read = 只读取投影；apply = 写入 */
  kind: 'context' | 'read' | 'apply'
  txnId: string
  version?: string
  epoch?: number
  /** apply 的新鲜度令牌（按 op 分级校验，见 shared/op-tiers） */
  docRev?: number
  structureEpoch?: number
  /** kind:'read' 的读请求 */
  read?: {
    what: 'sheets' | 'range' | 'formulas' | 'find' | 'audit'
    sheet?: string
    range?: string
    start?: string
    end?: string
    query?: string
    regex?: boolean
    matchCase?: boolean
    limit?: number
  }
  commands?: unknown[]
  dryRun?: boolean
}

export interface AgentResultMessage {
  type: 'agent-result'
  id: string
  ok: boolean
  context?: unknown
  receipt?: Receipt
  /** 请求条数与应用条数分开：整批级失败时 applied = 0 */
  requested?: number
  applied?: number
  wouldApply?: number
  dryRun?: boolean
  error?: { code: string; message: string }
}

export interface ExtensionBridge {
  ura: {
    page: {
      get<T = unknown>(id: string): Promise<T>
      open<T = unknown>(ref: string): Promise<T>
      save<T = unknown>(id: string, content: string): Promise<T>
      saveBytes<T = unknown>(id: string, data: Uint8Array): Promise<T>
      commit<T = unknown>(id: string, request: PageCommitInput): Promise<T>
      bytes(id: string, window?: { start?: number; end?: number }): Promise<Uint8Array>
      invoke<T = unknown>(id: string, tool: string, input?: unknown): Promise<T>
      agentResult(id: string, message: AgentResultMessage): Promise<unknown>
      reload<T = unknown>(id: string): Promise<T>
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

export interface XlsxProBridge {
  hasWorkbook(): boolean
  context(): unknown
  readCells(addresses: string[], sheetId?: string): unknown
  sheets(): ReadonlyArray<{ id: string; name: string }>
  focusSheet(sheetId: string, address?: string): void
  applyOps(ops: unknown[], dryRun: boolean): Promise<unknown>
  save(): Promise<{ ok: boolean; path?: string; error?: string }>
  /** 页面模型事务计数（L1 内容型 op 的新鲜度） */
  docRev(): number
  /** 结构修订号（L2 绝对坐标 op 的新鲜度） */
  structureEpoch(): number
  /** 回滚上一次撤销单元（agent 批次落盘失败时使用） */
  undo(): boolean
  /** 当前撤销栈深度（回滚时确认只撤掉 agent 自己那一批） */
  undoDepth(): number
}

export interface XlsxProHost {
  getVersion(): string
  getEpoch(): number
  getReceipt(): Receipt | null
  getFile(): unknown
  getBytes(): Uint8Array
  setOrigin(origin: 'page' | 'agent'): void
  commitBytes(bytes: Uint8Array, origin?: 'page' | 'agent'): Promise<Receipt>
}

export interface XlsxProBoot {
  sessionId: string
  data: XlsxPageData
  epoch: number
  /** 启动字节（同进程直传，不经 base64）；旧类型只给 data.bytes。 */
  bytes?: Uint8Array
}

declare global {
  interface Window {
    extension: ExtensionBridge
    appBridge?: {
      available(): Promise<boolean>
      call<T = unknown>(path: string, input?: unknown): Promise<T>
      peek<T = unknown>(path: string): Promise<T | null>
    }
    __xlsxProBoot?: XlsxProBoot
    __xlsxProHost?: XlsxProHost
    __xlsxPro?: XlsxProBridge
  }
}

export {}
