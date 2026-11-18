export function dirname(path: string): string {
  const index = path.lastIndexOf('/')
  return index <= 0 ? '.' : path.slice(0, index)
}
export function join(...parts: string[]): string {
  return parts.filter(Boolean).join('/').replace(/\/+/g, '/')
}

/// 上游 xlsx-gateway（f6c2358c）在临时文件命名里用到 basename / extname：
/// 页面 shim 只覆盖 POSIX 风格路径（页面里没有 Windows 盘符语义）。
export function basename(path: string, ext?: string): string {
  const index = path.lastIndexOf('/')
  const name = index < 0 ? path : path.slice(index + 1)
  return ext && name.endsWith(ext) ? name.slice(0, name.length - ext.length) : name
}

export function extname(path: string): string {
  const name = basename(path)
  const dot = name.lastIndexOf('.')
  return dot <= 0 ? '' : name.slice(dot)
}
