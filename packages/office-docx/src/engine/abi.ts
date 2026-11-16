// ─── 执行点 ABI 的本地声明 ─────────────────────────────────────────────
//
// 宿主契约的唯一事实源是 WorkStudio `src/shared/plugin-contract.ts`；执行点是
// 独立进程（按文件路径 import 本包产物），不能 import 宿主源码。这里只声明
// 本插件用到的字段，形状与契约 §6.2 一致。

export interface PlanRequest {
  tool?: string;
  params?: unknown;
  commands?: unknown;
  ref: string;
  baseVersion?: string;
  bytes?: Uint8Array;
}

export interface PlanOutcome {
  plan: unknown;
  options?: unknown;
  summary?: string;
}

export interface ExecRequest {
  ref: string;
  baseVersion: string;
  plan: unknown;
  options?: unknown;
  command?: 'undo' | 'redo';
  dryRun?: boolean;
  bytes?: Uint8Array;
}

export interface ExecResult {
  applied: number;
  rejected?: Array<{ index: number; code: string; message: string }>;
  warnings?: string[];
  notices?: string[];
  audit?: unknown;
  sync?: { kind: 'none' | 'version' | 'manifest' | 'bytes'; blocks?: unknown[]; dataUrl?: string };
}

export interface ExecOutcome {
  bytes?: Uint8Array;
  artifact?: { tempPath: string; size: number; sha256: string };
  modelToken?: string;
  result: ExecResult;
}

export interface EngineCtx {
  readBytes(): Promise<{ bytes: Uint8Array; version: string }>;
  tempDir: string;
  signal?: AbortSignal;
  importEpoch: number;
  budget: { maxBytes: number; deadlineMs: number };
}

export interface ErrorWithCode extends Error {
  code?: string;
}

/** 带契约错误码的异常（宿主原样透出给调用方）。 */
export function coded(code: string, message: string): ErrorWithCode {
  const error = new Error(message) as ErrorWithCode;
  error.code = code;
  return error;
}

export function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
