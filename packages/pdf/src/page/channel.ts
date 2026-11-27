// ─── 页面通道：window.workStudio 的类型化薄封装 ───────────────────────
//
// 地址固定 `pdf/pdf/<action>`；错误统一按 `code: message` 文本抛出。

import type {
  CallResult,
  ChunkResult,
  DomainUpdate,
  LoadResult,
  SaveFileResult,
} from './env'

const DOMAIN = 'pdf'

async function call<T>(action: string, input?: unknown): Promise<T> {
  const result = (await window.workStudio.command(DOMAIN, DOMAIN, action, input)) as
    | CallResult<T>
    | undefined
  if (!result) throw new Error('failed: 页面桥没有返回')
  if (result.ok !== true) throw new Error(result.error?.message ?? 'failed: 调用失败')
  return result.data
}

export const channel = {
  load: (ref: string, base64 = false) =>
    call<LoadResult>('load', { ref, ...(base64 ? { base64 } : {}) }),

  chunk: (input: {
    ref: string
    start: number
    length: number
    hash?: string
    base64?: boolean
  }) => call<ChunkResult>('chunk', input),

  saveCopy: (input: { ref: string; token: string; bytes: Uint8Array }) =>
    call<{ ok: boolean; size: number }>('saveCopy', input),

  /** 宿主保存对话框：插件只拿到一次性 token。 */
  async pickSaveFile(name: string): Promise<string | null> {
    const result = await window.workStudio.command<SaveFileResult>(
      'shell',
      'native',
      'saveFile',
      { name },
    )
    if (result.ok !== true || !result.data || result.data.cancelled) return null
    return result.data.token ?? null
  },

  /** 订阅同一 ref 的 changed 事件；只提示，不产生写冲突。 */
  watch(ref: string, handler: () => void): () => void {
    void window.workStudio.subscribe('workspace', ref).catch(() => undefined)
    return window.workStudio.onUpdate((update: DomainUpdate) => {
      if (update.type !== 'event') return
      if (update.domain !== 'workspace' || update.resource !== ref) return
      if (update.event !== 'changed') return
      handler()
    })
  },

  /** 主题：订阅 shell/appearance，并给出当前值。 */
  theme(): Promise<'light' | 'dark'> {
    return readSnapshot('shell', 'appearance').then((snapshot) => {
      const value = (snapshot as { theme?: unknown } | null)?.theme
      if (value === 'dark' || value === 'light') return value
      return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
    })
  },

  onTheme(handler: (theme: 'light' | 'dark') => void): () => void {
    void window.workStudio.subscribe('shell', 'appearance').catch(() => undefined)
    return window.workStudio.onUpdate((update: DomainUpdate) => {
      if (update.domain !== 'shell' || update.resource !== 'appearance') return
      if (update.type !== 'snapshot' && update.type !== 'patch') return
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
    void window.workStudio.subscribe(domain, resource).catch(() => finish(null))
  })
}
