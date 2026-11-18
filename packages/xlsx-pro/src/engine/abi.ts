import type { Rejection, XlsxProPlan } from '../shared/types'

export interface PlanRequest {
  tool?: string
  params?: unknown
  commands?: unknown
  ref: string
  baseVersion?: string
  bytes?: Uint8Array
  missing?: boolean
}

export interface PlanOutcome {
  plan: XlsxProPlan
  options?: unknown
  summary?: string
}

export interface ExecRequest {
  ref: string
  baseVersion: string
  plan: unknown
  options?: unknown
  command?: 'undo' | 'redo'
  dryRun?: boolean
  bytes?: Uint8Array
  missing?: boolean
}

export interface ExecResult {
  applied: number
  rejected?: Rejection[]
  warnings?: string[]
  notices?: string[]
  audit?: unknown
  sync?: {
    kind: 'none' | 'version' | 'manifest' | 'bytes'
    blocks?: unknown[]
    sheets?: unknown[]
    dataUrl?: string
    /** 保存产物的摘要（`kind:'manifest'`）：页面据此重建 file 模型，不回读整包。 */
    sha256?: string
    /** 保存产物的字节数：页面据此刷新分块读取的 size 提示。 */
    size?: number
  }
}

export interface ExecOutcome {
  bytes?: Uint8Array
  artifact?: { tempPath: string; size: number; sha256: string }
  result: ExecResult
}

export interface EngineCtx {
  readBytes(): Promise<{ bytes: Uint8Array; version: string }>
  /** 宿主可选的权威路径；仅用于诊断（guest 只能看 preopen 的临时区）。 */
  filePath?: string
  /**
   * 稳定的**按插件会话目录**：常驻引擎 / wasm 实例 / 工作簿会话的落脚点，
   * 跨调用存活。与 `tempDir`（每次调用一个、用完即删）是两回事。
   */
  sessionDir?: string
  tempDir: string
  signal?: AbortSignal
  importEpoch: number
  budget: { maxBytes: number; deadlineMs: number }
}

export interface ErrorWithCode extends Error {
  code?: string
}

export function coded(code: string, message: string): ErrorWithCode {
  const error = new Error(message) as ErrorWithCode
  error.code = code
  return error
}

export function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
