// ─── 执行点 ABI（契约 §6.2） ───────────────────────────────────────────
//
// 宿主契约的唯一事实源是 WorkStudio `src/shared/plugin-contract.ts`；执行点是
// 独立进程（按文件路径 import 本包产物），不能 import 宿主源码。这里只声明
// 本插件用到的字段，形状与契约 §6.2 一致。

import type { Rejection, SavePlan } from '../shared/types'

export interface PlanRequest {
  tool?: string
  params?: unknown
  commands?: unknown
  ref: string
  baseVersion?: string
  bytes?: Uint8Array
  /** 目标文件不存在：新建工具据此生成计划。 */
  missing?: boolean
}

export interface PlanOutcome {
  plan: DocxProPlan
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
    dataUrl?: string
  }
}

export interface ExecOutcome {
  bytes?: Uint8Array
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

/**
 * 本插件的计划词汇。一个 `write: 'plan'` 插件必须能同时接住两条路径：
 *
 * - `save`：执行点解析字节 → 块级 XML 切片 → 只重写 `word/document.xml`
 *   （无页面路径；未改块逐字节不变）。
 * - `bytes`：页面（完整编辑器）自己序列化出的整份字节。`exec` 仍是唯一
 *   字节产生者：它校验包结构、体积预算与 `document.xml` 存在性，再原样落盘。
 * - `noop` / `check` / `reject`：只读诊断与被拒批，不产字节。
 */
export type DocxProPlan =
  | (SavePlan & { kind: 'save' | 'check' | 'reject' })
  | { kind: 'bytes'; bytes: Uint8Array; applied?: number }
  | { kind: 'new'; documentXml: string; applied?: number }
  | { kind: 'noop'; applied?: number }

export interface ErrorWithCode extends Error {
  code?: string
}

/** 带契约错误码的异常（宿主原样透出给调用方）。 */
export function coded(code: string, message: string): ErrorWithCode {
  const error = new Error(message) as ErrorWithCode
  error.code = code
  return error
}

export function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
