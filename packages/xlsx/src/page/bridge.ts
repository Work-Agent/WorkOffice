/**
 * 页面给编辑器用的桥：唯一实现者。
 *
 * 编辑器只认这个对象：流式读（meta/range）、整份字节（回退）、保存/另存、
 * 主题、脏态，全部落到 `xlsx.page` 的 op 上。
 */

import type { Channel } from './channel'

export interface SaveInput {
  bytes?: Uint8Array
  request?: unknown
  baseHash?: string
}

export interface SaveReceipt {
  ref: string
  hash: string
  size: number
  at: number
  applied: number
  notices?: string[]
}

export interface OpenBytes {
  bytes: Uint8Array
  hash: string
  size: number
}

export interface EditorBridge {
  readonly ref: string
  readonly pageId: string
  /** wasm 可用：编辑器走 `meta + range` 流式打开，不读整份字节。 */
  readonly supportsEngineRange: boolean
  /** 打开时的文件哈希；保存后更新，作为下一次写的 baseHash。 */
  hash(): string
  theme(): Promise<'light' | 'dark'>
  onTheme(listener: (theme: 'light' | 'dark') => void): () => void
  /** wasm `open` 的元数据（sheets / styles / definedNames / visuals…）。 */
  meta(): Promise<Record<string, unknown>>
  /** wasm `read_range` 原样结果。 */
  range(sheetId: string, range: unknown): Promise<unknown>
  bytes(): Promise<OpenBytes>
  save(input: SaveInput): Promise<SaveReceipt>
  saveAs(bytes: Uint8Array, saveToken: string): Promise<SaveReceipt>
  /** 宿主保存对话框 → 一次性 `save:` token；取消返回 null。 */
  pickSaveFile(name: string): Promise<{ token: string; name: string } | null>
  /** 崩溃恢复副本：把同一份保存请求写进插件数据目录。 */
  draft(request: unknown): Promise<{ ok: boolean }>
  clearDraft(): Promise<void>
  setDirty(dirty: boolean): void
  toast(message: string): void
}

const READ_CHUNK = 1024 * 1024
/** 超过这个体积走分块保存，避免一次 IPC 拷几十 MB。 */
const SAVE_CHUNK_THRESHOLD = 8 * 1024 * 1024
const SAVE_CHUNK = 4 * 1024 * 1024

let hostTheme: 'light' | 'dark' = 'light'
const themeListeners = new Set<(theme: 'light' | 'dark') => void>()

/** 宿主主题变化（`shell/appearance` 订阅）→ 所有监听者。 */
export function setHostTheme(theme: unknown): void {
  hostTheme = theme === 'dark' ? 'dark' : 'light'
  for (const listener of themeListeners) listener(hostTheme)
}

export function createBridge(
  channel: Channel,
  initial: { hash: string; supportsEngineRange: boolean },
): EditorBridge {
  let currentHash = initial.hash
  let dirtyReported = false
  let dirtyTimer: ReturnType<typeof setTimeout> | null = null

  const reportDirty = (dirty: boolean): void => {
    if (dirty === dirtyReported) return
    dirtyReported = dirty
    void channel.report({ hash: currentHash, dirty, docRev: 0 }).catch(() => undefined)
  }

  const save = async (input: SaveInput): Promise<SaveReceipt> => {
    const baseHash = input.baseHash ?? currentHash
    const bytes = input.bytes
    let receipt: SaveReceipt
    if (bytes && bytes.byteLength > SAVE_CHUNK_THRESHOLD) {
      const begin = await channel.call<{ transferId: string }>('saveBegin', { size: bytes.byteLength })
      let seq = 0
      for (let at = 0; at < bytes.byteLength; at += SAVE_CHUNK, seq += 1) {
        await channel.call('saveChunk', {
          transferId: begin.transferId,
          seq,
          bytes: bytes.subarray(at, Math.min(at + SAVE_CHUNK, bytes.byteLength)),
        })
      }
      receipt = await channel.call<SaveReceipt>('saveCommit', { transferId: begin.transferId, baseHash })
    } else {
      receipt = await channel.call<SaveReceipt>('save', {
        ...(bytes ? { bytes } : {}),
        ...(input.request ? { request: input.request } : {}),
        baseHash,
      })
    }
    currentHash = receipt.hash
    dirtyReported = false
    return receipt
  }

  return {
    ref: channel.ref,
    pageId: channel.pageId,
    supportsEngineRange: initial.supportsEngineRange,
    hash: () => currentHash,

    theme: async () => hostTheme,

    onTheme(listener) {
      themeListeners.add(listener)
      listener(hostTheme)
      return () => themeListeners.delete(listener)
    },

    meta: () => channel.call<Record<string, unknown>>('meta', {}),

    range: (sheetId, range) => channel.call('range', { sheetId, range }),

    async bytes() {
      const opened = await channel.read<{ hash: string; size: number }>('open')
      const parts: Uint8Array[] = []
      let offset = 0
      for (;;) {
        const slice = await channel.read<{ bytes: Uint8Array; next: number; eof: boolean }>('chunk', {
          hash: opened.hash,
          offset,
          length: READ_CHUNK,
        })
        parts.push(slice.bytes)
        offset = slice.next
        if (slice.eof) break
      }
      const bytes = new Uint8Array(offset)
      let at = 0
      for (const part of parts) {
        bytes.set(part, at)
        at += part.byteLength
      }
      currentHash = opened.hash
      return { bytes, hash: opened.hash, size: bytes.byteLength }
    },

    save,

    async saveAs(bytes, saveToken) {
      const receipt = await channel.call<SaveReceipt>('saveAs', {
        bytes,
        saveToken,
        baseHash: currentHash,
      })
      channel.setRef(receipt.ref)
      currentHash = receipt.hash
      dirtyReported = false
      return receipt
    },

    async draft(request) {
      await channel.call('draft', { request })
      return { ok: true }
    },

    async clearDraft() {
      await channel.call('draftClear', {}).catch(() => undefined)
    },

    async pickSaveFile(name) {
      const result = await window.workStudio.command<{
        cancelled: boolean
        token?: string
        name?: string
      }>('shell', 'native', 'saveFile', { name })
      if (!result.ok) throw new Error(result.error.message)
      if (result.data.cancelled || !result.data.token) return null
      return { token: result.data.token, name: result.data.name ?? name }
    },

    setDirty(dirty) {
      // 防抖：编辑期每次事务都上报会把 IPC 打满，脏态只需要最终一致。
      if (dirtyTimer) clearTimeout(dirtyTimer)
      dirtyTimer = setTimeout(() => {
        dirtyTimer = null
        reportDirty(dirty)
      }, 400)
    },

    toast(message) {
      console.info('[xlsx]', message)
    },
  }
}
