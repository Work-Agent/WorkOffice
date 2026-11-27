// ─── 执行点 ABI：错误码前缀 ─────────────────────────────────────────
//
// 宿主的传输层只有 5 个码（unknown / denied / invalid / not_found / failed），
// 语义码写在 message 前缀里：`command.stale: …`，页面与助手都按前缀判断。

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
