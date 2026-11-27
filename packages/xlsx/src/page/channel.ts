/**
 * 页面 ↔ 引擎的唯一通道。
 *
 * 所有 op 都走 `xlsx.page`：只读 op 不要求令牌，写与结算 op 用
 * attach 发放的 token。页面不碰字节，也不拼路径。
 */

export interface CallError {
  code: string
  message: string
}

export class PageError extends Error {
  constructor(
    readonly code: string,
    message: string,
  ) {
    super(message)
    this.name = 'PageError'
  }
}

function bridge() {
  const value = window.workStudio
  if (!value) throw new PageError('unavailable', '宿主 bridge 未注入：请从 WorkOffice 打开该文件')
  return value
}

export class Channel {
  private token = ''
  ref: string

  constructor(
    ref: string,
    readonly pageId: string,
  ) {
    this.ref = ref
  }

  /** 另存为之后把通道指向新文件：后续 op 与保存都落在新 ref 上。 */
  setRef(ref: string): void {
    this.ref = ref
  }

  async call<T>(op: string, input: Record<string, unknown> = {}): Promise<T> {
    const result = await bridge().command<T>('xlsx', 'xlsx', 'page', {
      ref: this.ref,
      pageId: this.pageId,
      ...(this.token ? { token: this.token } : {}),
      op,
      input,
    })
    if (!result.ok) throw new PageError(result.error.code, result.error.message)
    return result.data
  }

  /** 只读 op：输入里的 token 不参与，但保留可读的调用形态。 */
  read<T>(op: string, input: Record<string, unknown> = {}): Promise<T> {
    return this.call<T>(op, input)
  }

  async attach(takeover = false): Promise<void> {
    const result = await this.call<{ token: string }>('attach', {
      pageId: this.pageId,
      ...(takeover ? { takeover: true } : {}),
    })
    this.token = result.token
  }

  async report(patch: Record<string, unknown>): Promise<void> {
    if (!this.token) return
    await this.call('report', { token: this.token, ...patch })
  }

  detach(): void {
    if (!this.token) return
    // 页面关闭时用 notify：不等回执，主进程仍会把 detach 跑完。
    bridge().notify('xlsx', 'xlsx', 'page', {
      ref: this.ref,
      pageId: this.pageId,
      token: this.token,
      op: 'detach',
      input: '{}',
    })
    this.token = ''
  }
}
