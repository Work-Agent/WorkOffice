// ─── 命令错误 ─────────────────────────────────────────────────────────
//
// 宿主只把 `code` 归一成 5 个传输码（unknown/denied/invalid/not_found/failed），
// 所以语义码写进 message 前缀：agent 与日志看到的仍是 command.stale-index 这类
// 可处置的码，而不是一句没有出路的 failed。

export class CommandError extends Error {
  constructor(
    readonly code: string,
    message: string,
  ) {
    super(`${code}: ${message}`)
    this.name = 'CommandError'
  }
}

export function invalid(message: string): never {
  throw new CommandError('command.invalid', message)
}

export function notFound(message: string): never {
  throw new CommandError('command.not-found', message)
}

export function stale(message: string): never {
  throw new CommandError('command.stale', message)
}
