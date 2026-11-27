// ─── 页面通道：window.workStudio 的类型化薄封装 ───────────────────────
//
// 地址固定 `pptx/pptx/<action>`；所有错误都是 `code: message` 文本，
// 页面只按前缀区分（command.stale / command.page-unavailable / …）。

import type {
  AttachResult,
  CallResult,
  DomainUpdate,
  LoadResult,
  PageRequest,
  SaveReceipt,
} from './env'

const DOMAIN = 'pptx'

export interface AckPayload {
  receipt?: SaveReceipt
  data?: unknown
  error?: { code: string; message: string; index?: number }
}

async function call<T>(action: string, input?: unknown): Promise<T> {
  const result = (await window.workStudio.command(DOMAIN, DOMAIN, action, input)) as
    | CallResult<T>
    | undefined
  if (!result) throw new Error('failed: 页面桥没有返回')
  if (result.ok !== true) throw new Error(result.error?.message ?? 'failed: 调用失败')
  return result.data
}

export const channel = {
  load: (ref: string) => call<LoadResult>('load', { ref }),

  save: (input: { ref: string; pageId: string; bytes: Uint8Array; baseHash: string }) =>
    call<SaveReceipt>('save', input),

  saveAs: (input: { ref: string; pageId: string; token: string; bytes: Uint8Array }) =>
    call<SaveReceipt>('saveAs', input),

  attach: (input: { ref: string; pageId: string }) => call<AttachResult>('attach', input),

  pending: (id: string) => call<PageRequest>('pending', { id }),

  ack: (input: { id: string; ok: boolean; payload?: AckPayload }) =>
    call<{ ok: boolean }>('ack', {
      ...input,
      ...(input.payload ? { payload: JSON.stringify(input.payload) } : {}),
    }),

  report: (input: {
    ref: string
    pageId: string
    hash?: string
    docRev?: number
    dirty?: boolean
  }) => call<{ ok: boolean }>('report', input),

  detach: (input: { ref: string; pageId: string }) => call<{ ok: boolean }>('detach', input),

  /** 宿主保存对话框：插件只拿到一次性 token。 */
  async pickSaveFile(name: string): Promise<string | null> {
    const result = await window.workStudio.command<
      { cancelled?: boolean; token?: string } | null
    >('shell', 'native', 'saveFile', { name })
    if (result.ok !== true || !result.data || result.data.cancelled) return null
    return result.data.token ?? null
  },

  /** 订阅插件事件（助手请求）。 */
  onRequest(handler: (id: string, type: string) => void): () => void {
    void window.workStudio.subscribe(DOMAIN, 'events')
    return window.workStudio.onUpdate((update: DomainUpdate) => {
      if (update.type !== 'event') return
      if (update.domain !== DOMAIN || update.resource !== 'events') return
      if (update.event !== 'request') return
      const data = update.data as { id?: unknown; type?: unknown } | undefined
      if (typeof data?.id !== 'string') return
      handler(data.id, typeof data.type === 'string' ? data.type : 'apply')
    })
  },

  /** 主题：订阅 shell/appearance，并给出当前值。 */
  async theme(): Promise<'light' | 'dark'> {
    const snapshot = await readSnapshot('shell', 'appearance')
    const value = (snapshot as { theme?: unknown } | null)?.theme
    if (value === 'dark' || value === 'light') return value
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
  },

  onTheme(handler: (theme: 'light' | 'dark') => void): () => void {
    void window.workStudio.subscribe('shell', 'appearance')
    return window.workStudio.onUpdate((update: DomainUpdate) => {
      if (update.domain !== 'shell' || update.resource !== 'appearance') return
      const data = update.data as { theme?: unknown } | undefined
      if (data?.theme === 'dark' || data?.theme === 'light') handler(data.theme)
    })
  },
}

/** 订阅后首个快照；2 秒内拿不到就返回 null（调用方退回系统偏好）。 */
function readSnapshot(domain: string, resource: string): Promise<unknown> {
  return new Promise((resolve) => {
    let done = false
    const finish = (value: unknown): void => {
      if (done) return
      done = true
      clearTimeout(timer)
      off()
      resolve(value)
    }
    const timer = setTimeout(() => finish(null), 2000)
    const off = window.workStudio.onUpdate((update) => {
      if (update.type !== 'snapshot') return
      if (update.domain !== domain || update.resource !== resource) return
      finish(update.data)
    })
    void window.workStudio.subscribe(domain, resource)
  })
}
