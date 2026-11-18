// ─── 执行点 ABI（与 WorkStudio host-plugin-contract 对齐） ─────────────
//
// 本插件是计划型 + 常驻模型：
//   - plan 返回 TxnRequest
//   - exec 返回 artifact + modelToken
//   - 宿主写成功后 commitModel，失败 abortModel
//   - 不返回目标路径；exec 只写 ctx.tempDir

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
  plan: PptxProPlan
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

export interface ExecOutcome {
  artifact?: { tempPath: string; size: number; sha256: string }
  modelToken?: string
  result: ExecResult
}

export interface EngineCtx {
  readBytes(): Promise<{ bytes: Uint8Array; version: string }>
  tempDir: string
  signal?: AbortSignal
  importEpoch: number
  budget: { maxBytes: number; deadlineMs: number }
}

export interface Rejection {
  index: number
  code: string
  message: string
}

export type PptxOp = {
  op: string
  target?: { slide?: number | string; el?: string; part?: string }
  [key: string]: unknown
}

export type PptxProPlan =
  | {
      kind: 'txn'
      txn: {
        ops?: PptxOp[]
        /** 页面按 UI 事务分批；exec 按顺序应用，失败整批回滚。 */
        batches?: PptxOp[][]
        isolation?: 'atomic' | 'per_op'
        dryRun?: boolean
      }
      applied?: number
    }
  | {
      kind: 'pageSave'
      slides: Array<{ path: string; xml: string }>
      parts: Array<{ path: string; bytes: Uint8Array }>
      deletedParts: string[]
      changedCount: number
      metaDirty?: boolean
    }
  | { kind: 'new'; spec: unknown; applied?: number; dryRun?: boolean }
  | { kind: 'reject'; applied?: number; rejected: Rejection[]; warnings?: string[] }
  | { kind: 'noop'; applied?: number }

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
