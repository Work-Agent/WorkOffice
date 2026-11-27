// ─── 单文件守卫 ───────────────────────────────────────────────────────
//
// 插件只接受「一个 .xlsx 文件」：目录、非 file:// 来源、错误扩展名一律拒绝。
// 这是 §0 的边界，不是建议 —— 资源适配器、命令、页面通道都从这里过。

import { stat } from 'node:fs/promises'
import { basename, extname } from 'node:path'
import { fileURLToPath } from 'node:url'

import { invalid, notFound } from './errors'

export const XLSX_EXTS = new Set(['.xlsx', '.xlsm', '.xltx'])

/** wasm 底座的上限（实测 150 MB 舒适）；TS 解析低于此值。 */
export const MAX_FILE_BYTES = 150 * 1024 * 1024
/** 页面走 TS 解析时的舒适上限。 */
export const PAGE_PARSE_BYTES = 64 * 1024 * 1024

export type Tier = 'normal' | 'readOnlyOversized'

export interface XlsxFile {
  ref: string
  path: string
  name: string
  size: number
  mtime: number
}

/** `file://` → 绝对路径；其它协议直接拒。 */
export function pathOf(ref: string): string {
  let url: URL
  try {
    url = new URL(ref)
  } catch {
    invalid(`ref 必须是 file:// 地址：${ref}`)
  }
  if (url.protocol !== 'file:') invalid(`只支持 file:// 的单个文件：${ref}`)
  return fileURLToPath(url)
}

export function assertExt(path: string): void {
  if (!XLSX_EXTS.has(extname(path).toLowerCase()))
    invalid(`只支持 xlsx / xlsm / xltx：${basename(path)}`)
}

/** 读取前守卫：必须存在、是普通文件、扩展名正确、不超上限。 */
export async function assertXlsxFile(ref: string): Promise<XlsxFile> {
  const path = pathOf(ref)
  const info = await stat(path).catch(() => null)
  if (!info) notFound(`文件不存在：${ref}`)
  if (!info.isFile()) invalid(`不是单个文件：${ref}`)
  assertExt(path)
  if (info.size > MAX_FILE_BYTES)
    invalid(`工作簿 ${info.size} 字节超过 ${MAX_FILE_BYTES} 上限`)
  return { ref, path, name: basename(path), size: info.size, mtime: info.mtimeMs }
}

export function tierOf(size: number): Tier {
  return size > PAGE_PARSE_BYTES ? 'readOnlyOversized' : 'normal'
}
