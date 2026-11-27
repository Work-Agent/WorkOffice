// ─── 单文件守卫与字节工具 ────────────────────────────────────────────
//
// 插件只接受「一个 .pdf 文件」：目录、非 file:// 来源、错误扩展名一律拒绝。
// 上限 128 MB，与旧查看器一致；超过就拒绝，不静默截断。

import { createHash } from 'node:crypto'
import { open, readFile, stat } from 'node:fs/promises'
import { basename, extname } from 'node:path'
import { fileURLToPath } from 'node:url'

export const MAX_PDF_BYTES = 128 * 1024 * 1024
/** 小文件一次 load 就带字节；大文件走 chunk。 */
export const INLINE_BYTES = 16 * 1024 * 1024
export const CHUNK_BYTES = 4 * 1024 * 1024
export const MAX_CHUNK_BYTES = 8 * 1024 * 1024
/** 另存副本上限：与页面现有 SAVE_COPY_LIMIT 对齐。 */
export const MAX_SAVE_BYTES = 64 * 1024 * 1024

export interface PdfFile {
  ref: string
  path: string
  name: string
  size: number
  mtimeMs: number
}

export function fail(code: string, message: string): never {
  throw new Error(`${code}: ${message}`)
}

export function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export function textOf(value: unknown): string {
  return typeof value === 'string' ? value : ''
}

/** `file://` → 绝对路径；其它协议直接拒。 */
export function pathOf(ref: string): string {
  let url: URL
  try {
    url = new URL(ref)
  } catch {
    fail('command.invalid', `ref 必须是 file:// 地址：${ref}`)
  }
  if (url.protocol !== 'file:') {
    fail('command.invalid', `只支持 file:// 的单个文件：${ref}`)
  }
  try {
    return fileURLToPath(url)
  } catch {
    fail('command.invalid', `ref 不是合法的 file:// 地址：${ref}`)
  }
}

export function nameOf(ref: string): string {
  return basename(pathOf(ref)) || 'document.pdf'
}

export async function assertPdfFile(ref: string): Promise<PdfFile> {
  const path = pathOf(ref)
  const info = await stat(path).catch(() => null)
  if (!info) fail('resource.not-found', `文件不存在：${ref}`)
  if (!info.isFile()) fail('command.invalid', `不是单个文件：${ref}`)
  if (extname(path).toLowerCase() !== '.pdf') {
    fail('command.invalid', `只支持 .pdf：${ref}`)
  }
  if (info.size > MAX_PDF_BYTES) {
    fail(
      'command.too-large',
      `PDF ${Math.round(info.size / 1024 / 1024)} MB 超过 ${Math.round(MAX_PDF_BYTES / 1024 / 1024)} MB 上限`,
    )
  }
  return { ref, path, name: basename(path) || 'document.pdf', size: info.size, mtimeMs: info.mtimeMs }
}

/** 命令输入里的 ref：必须存在且是 .pdf；返回原 ref。 */
export function refOf(input: { ref?: unknown }): string {
  const ref = textOf(input.ref)
  if (!ref) fail('command.invalid', '缺少 ref')
  pathOf(ref)
  return ref
}

export async function readBytes(path: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(path))
}

/** 只读文件头部：大文件做元数据探测时避免整份读。 */
export async function readHead(path: string, length: number): Promise<Uint8Array> {
  const handle = await open(path, 'r')
  try {
    const buffer = Buffer.alloc(length)
    const { bytesRead } = await handle.read(buffer, 0, length, 0)
    return new Uint8Array(buffer.subarray(0, bytesRead))
  } finally {
    await handle.close()
  }
}

export function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

/** 页面命令的二进制输入：Uint8Array / ArrayBuffer / base64 文本。 */
export function toBytes(input: Record<string, unknown>): Uint8Array {
  const raw = input['bytes'] ?? input['data']
  if (raw instanceof Uint8Array) return raw
  if (raw instanceof ArrayBuffer) return new Uint8Array(raw)
  if (ArrayBuffer.isView(raw)) {
    return new Uint8Array(raw.buffer, raw.byteOffset, raw.byteLength)
  }
  if (typeof raw === 'string' && raw.length > 0) {
    try {
      return new Uint8Array(Buffer.from(raw, 'base64'))
    } catch {
      fail('command.invalid', 'data 不是合法 base64')
    }
  }
  return new Uint8Array()
}

export function clampInt(value: unknown, fallback: number, min: number, max: number): number {
  const n = typeof value === 'number' ? value : Number(value)
  if (!Number.isFinite(n)) return fallback
  return Math.max(min, Math.min(max, Math.trunc(n)))
}
