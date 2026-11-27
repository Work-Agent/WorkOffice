import type { PluginScope } from '@work-studio/sdk'

let current: PluginScope | null = null

export function setScope(scope: PluginScope): void {
  current = scope
}

export function clearScope(): void {
  current = null
}

export function getScope(): PluginScope {
  if (!current) throw new Error('插件尚未激活')
  return current
}

/** 插件私有临时目录：wasm 实例 / 会话文件的落脚点，跨调用存活。 */
export function tempDir(): string {
  return getScope().tempDir
}
