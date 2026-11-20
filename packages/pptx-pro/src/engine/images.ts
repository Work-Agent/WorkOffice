// ─── 设计稿路径的图片来源 ─────────────────────────────────────────────
//
// 支持四类来源（与 `genoffice create --spec` 同语义）：
//   - `data:image/…;base64,…`
//   - `http(s)://…`
//   - `file:///…`
//   - 绝对 / 相对路径（相对路径按 deck 所在目录解析）
//
// 与 genoffice 的 `readImageSource` 同语义，但不用 @genoffice/electron-utils
// 的 remote-image（那个模块 import 'undici'，本仓库没有该包）。这里用 Node
// 自带 fetch，自己补上：仅 http/https、逐跳校验目标地址、按流计数的字节上限。
//
// 本地文件是用户主动放进 deck 目录的素材；读取按魔数校验 + 单张字节上限，
// 不是图片的内容不会被塞进 deck。

import { lookup } from 'node:dns/promises'
import { readFile, stat } from 'node:fs/promises'
import { isIP } from 'node:net'
import { dirname, isAbsolute, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

/** 单张图上限；模型可以指任意 URL，先卡死再下载。 */
export const MAX_IMAGE_BYTES = 20 * 1024 * 1024
const FETCH_TIMEOUT_MS = 20_000
const MAX_REDIRECTS = 3
const MAX_DATA_URL_CHARS = Math.ceil((MAX_IMAGE_BYTES * 4) / 3)

export interface FetchedImage {
  bytes: Uint8Array
  ext: string
}

function ascii(bytes: Uint8Array, offset: number, length: number): string {
  let out = ''
  for (let i = 0; i < length; i++) out += String.fromCharCode(bytes[offset + i] ?? 0)
  return out
}

/** 扩展名按魔数判定：Content-Type 会撒谎，把 HTML 错误页当 jpg 存进 deck 是最坏情况。 */
export function imageExtOf(bytes: Uint8Array): string | null {
  if (bytes.length >= 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)
    return 'png'
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'jpg'
  if (bytes.length >= 6 && (ascii(bytes, 0, 4) === 'GIF8')) return 'gif'
  if (bytes.length >= 12 && ascii(bytes, 0, 4) === 'RIFF' && ascii(bytes, 8, 4) === 'WEBP') return 'webp'
  if (bytes.length >= 2 && bytes[0] === 0x42 && bytes[1] === 0x4d) return 'bmp'
  return null
}

/** 私有 / 环回 / 链路本地 / 保留地址；只做基本拦截，防的是"顺手把内网地址写进 spec"。 */
function isBlockedIp(ip: string): boolean {
  const version = isIP(ip)
  if (version === 4) {
    const parts = ip.split('.').map(Number)
    const [a, b] = parts as [number, number]
    if (a === 0 || a === 10 || a === 127) return true
    if (a === 169 && b === 254) return true
    if (a === 172 && b >= 16 && b <= 31) return true
    if (a === 192 && b === 168) return true
    if (a === 100 && b >= 64 && b <= 127) return true
    if (a >= 224) return true
    return false
  }
  if (version === 6) {
    const v = ip.toLowerCase()
    if (v === '::1' || v === '::') return true
    if (v.startsWith('fc') || v.startsWith('fd')) return true
    if (v.startsWith('fe80')) return true
    const mapped = /^::ffff:(\d+\.\d+\.\d+\.\d+)$/.exec(v)
    if (mapped) return isBlockedIp(mapped[1]!)
    return false
  }
  return false
}

async function assertPublicHost(url: URL): Promise<void> {
  const host = url.hostname.replace(/^\[|\]$/g, '')
  if (!host) throw new Error('empty host')
  if (isIP(host)) {
    if (isBlockedIp(host)) throw new Error('private address')
    return
  }
  const addresses = await lookup(host, { all: true })
  if (addresses.length === 0) throw new Error('host does not resolve')
  if (addresses.some((a) => isBlockedIp(a.address))) throw new Error('private address')
}

async function readCapped(resp: Response): Promise<Uint8Array> {
  const declared = Number(resp.headers.get('content-length') ?? '')
  if (Number.isFinite(declared) && declared > MAX_IMAGE_BYTES) throw new Error('image too large')
  if (!resp.body) return new Uint8Array(await resp.arrayBuffer())
  const reader = resp.body.getReader()
  const chunks: Uint8Array[] = []
  let total = 0
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    total += value.byteLength
    if (total > MAX_IMAGE_BYTES) {
      await reader.cancel().catch(() => {})
      throw new Error('image too large')
    }
    chunks.push(value)
  }
  const out = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    out.set(chunk, offset)
    offset += chunk.byteLength
  }
  return out
}

async function fetchHttp(url: string, signal?: AbortSignal): Promise<Uint8Array | null> {
  let current = url
  for (let hop = 0; hop <= MAX_REDIRECTS; hop++) {
    const parsed = new URL(current)
    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') return null
    await assertPublicHost(parsed)
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS)
    const onAbort = () => controller.abort()
    signal?.addEventListener('abort', onAbort, { once: true })
    try {
      const resp = await fetch(current, {
        redirect: 'manual',
        signal: controller.signal,
        headers: { accept: 'image/*' },
      })
      if (resp.status >= 300 && resp.status < 400) {
        const location = resp.headers.get('location')
        if (!location) return null
        current = new URL(location, current).toString()
        continue
      }
      if (!resp.ok) return null
      return await readCapped(resp)
    } catch {
      return null
    } finally {
      clearTimeout(timer)
      signal?.removeEventListener('abort', onAbort)
    }
  }
  return null
}

/** 本地路径 → 图片；读不到 / 不是图片 / 太大都返回 null。 */
async function readLocalImage(path: string): Promise<Uint8Array | null> {
  try {
    const info = await stat(path)
    if (!info.isFile() || info.size === 0 || info.size > MAX_IMAGE_BYTES) return null
    return new Uint8Array(await readFile(path))
  } catch {
    return null
  }
}

/** deck 的所在目录：相对图片路径按它解析（`./assets/hero.png`）。 */
export function baseDirOfRef(ref: string): string | undefined {
  if (!ref) return undefined
  try {
    const url = new URL(ref)
    if (url.protocol === 'file:') return dirname(fileURLToPath(url))
    return undefined
  } catch {
    // 不是 URL：绝对路径就取它的目录，其余交给调用方（相对 cwd）
    return isAbsolute(ref) ? dirname(ref) : undefined
  }
}

function localPathOf(url: string, baseDir?: string): string | null {
  if (url.startsWith('file:')) {
    try {
      return fileURLToPath(url)
    } catch {
      return null
    }
  }
  if (isAbsolute(url)) return url
  return baseDir ? resolve(baseDir, url) : null
}

function fromDataUrl(url: string): FetchedImage | null {
  const match = /^data:image\/([a-z0-9.+-]+);base64,(.*)$/is.exec(url)
  if (!match) return null
  if (match[2]!.length > MAX_DATA_URL_CHARS) return null
  const bytes = new Uint8Array(Buffer.from(match[2]!, 'base64'))
  if (bytes.byteLength > MAX_IMAGE_BYTES) return null
  const ext = imageExtOf(bytes) ?? match[1]!.toLowerCase().replace(/^jpeg$/, 'jpg')
  return { bytes, ext }
}

/**
 * 取一张图；失败返回 null（该页其余元素照常落地，失败列在 imageFailures）。
 * `cache` 由调用方按一次构建共享，同一 URL 多页引用只下载一次。
 */
export async function fetchSpecImage(
  url: string,
  cache: Map<string, Promise<FetchedImage | null>>,
  signal?: AbortSignal,
  baseDir?: string,
): Promise<FetchedImage | null> {
  const key = String(url ?? '')
  const hit = cache.get(key)
  if (hit) return hit

  const task = (async (): Promise<FetchedImage | null> => {
    if (key.startsWith('data:')) return fromDataUrl(key)
    let bytes: Uint8Array | null
    if (/^https?:\/\//i.test(key)) bytes = await fetchHttp(key, signal)
    else {
      const path = localPathOf(key, baseDir)
      bytes = path ? await readLocalImage(path) : null
    }
    if (!bytes) return null
    const ext = imageExtOf(bytes)
    // 只有真的图片字节才进 deck；伪造 Content-Type / HTML 错误页在这里被拦下。
    return ext ? { bytes, ext } : null
  })().catch(() => null)

  cache.set(key, task)
  return task
}
