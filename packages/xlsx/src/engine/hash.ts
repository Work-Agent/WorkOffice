import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'

export function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

export async function sha256File(path: string): Promise<string> {
  return sha256Hex(new Uint8Array(await readFile(path)))
}

/**
 * 文件哈希缓存：`(path,size,mtime)` 命中就不再读盘。
 * `chunk` 每块都要校验基线，没有这个缓存会退化成 O(n²) 读盘。
 */
const cache = new Map<string, { size: number; mtime: number; hash: string }>()

export function rememberHash(path: string, size: number, mtime: number, hash: string): void {
  cache.set(path, { size, mtime, hash })
}

export async function hashOfFile(file: {
  path: string
  size: number
  mtime: number
}): Promise<string> {
  const hit = cache.get(file.path)
  if (hit && hit.size === file.size && hit.mtime === file.mtime) return hit.hash
  const hash = await sha256File(file.path)
  cache.set(file.path, { size: file.size, mtime: file.mtime, hash })
  return hash
}

export function resetHashes(): void {
  cache.clear()
}
