// ─── PDF 页面入口（WorkStudio SDK） ───────────────────────────────────
//
// 1. `?ref=` → pdf.load 拿元数据（≤16 MB 同时带字节）
// 2. 大文件用 pdf.chunk 按块拼装（预分配一次）
// 3. 注入 __pdfBoot，启动 PDF.js 查看器
// 4. 页面只读：不调用 save / saveBytes；另存副本走宿主 save token。

import { assembleBytes, MAX_PDF_BYTES } from '../shared/chunk'
import type { PdfBoot } from '../shared/pdf-types'
import { channel } from './channel'
import type { LoadResult } from './env'

const CHUNK_BYTES = 4 * 1024 * 1024

function fail(message: string): void {
  const boot = document.querySelector<HTMLElement>('.ws-boot')
  if (!boot) return
  boot.dataset.error = '1'
  boot.textContent = message
}

function applyTheme(theme: 'light' | 'dark'): void {
  document.documentElement.dataset.theme = theme
  document.documentElement.style.colorScheme = theme
}

function decodeBase64(text: string): Uint8Array {
  const binary = atob(text)
  const out = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i += 1) out[i] = binary.charCodeAt(i)
  return out
}

async function fetchChunks(ref: string, size: number, hash: string): Promise<Uint8Array> {
  return assembleBytes(
    async ({ start, end }) => {
      const length = end - start
      let result = await channel.chunk({ ref, start, length, hash })
      if (result.bytes instanceof Uint8Array) return result.bytes
      if (typeof result.data === 'string') return decodeBase64(result.data)
      result = await channel.chunk({ ref, start, length, hash, base64: true })
      if (typeof result.data === 'string') return decodeBase64(result.data)
      throw new Error('pdf.chunk 没有返回字节')
    },
    { size, chunkBytes: CHUNK_BYTES, maxBytes: MAX_PDF_BYTES },
  )
}

async function loadBytes(ref: string, loaded: LoadResult): Promise<Uint8Array> {
  if (loaded.bytes instanceof Uint8Array) return loaded.bytes
  if (typeof loaded.data === 'string' && loaded.data.length > 0) {
    return decodeBase64(loaded.data)
  }
  // IPC 极端情况下不保留 typed array：小文件用 base64 再取一次。
  if (loaded.inline) {
    const retry = await channel.load(ref, true)
    if (typeof retry.data === 'string' && retry.data.length > 0) {
      return decodeBase64(retry.data)
    }
  }
  return fetchChunks(ref, loaded.size, loaded.hash)
}

async function boot(): Promise<void> {
  const params = new URLSearchParams(location.search)
  const ref = params.get('ref') ?? ''
  const locator = params.get('locator') ?? ''
  if (!ref) {
    fail('缺少 ref：请从工作区打开 PDF 文档')
    return
  }

  try {
    applyTheme(await channel.theme())
    channel.onTheme(applyTheme)

    const loaded = await channel.load(ref)
    const bytes = await loadBytes(ref, loaded)

    const bootData: PdfBoot = {
      ref,
      data: loaded.meta,
      bytes,
      version: loaded.hash,
      locator,
      refetch: async () => {
        const next = await channel.load(ref)
        return loadBytes(ref, next)
      },
      reload: async () => {
        const next = await channel.load(ref)
        return {
          bytes: await loadBytes(ref, next),
          data: next.meta,
          version: next.hash,
        }
      },
      watchExternal: (handler) => channel.watch(ref, handler),
      saveCopy: async (next) => {
        const token = await channel.pickSaveFile(loaded.meta.name)
        if (!token) return false
        await channel.saveCopy({ ref, token, bytes: next })
        return true
      },
    }
    window.__pdfBoot = bootData
    await import('../viewer/main')
  } catch (error) {
    fail(`PDF 打开失败：${error instanceof Error ? error.message : String(error)}`)
  }
}

void boot()
