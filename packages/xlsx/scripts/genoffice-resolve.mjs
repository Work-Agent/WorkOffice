/**
 * GenOffice 源码解析：唯一一份规则，esbuild / vite / tsconfig 三处共用。
 *
 *   @sheets/<path>            → build/sheets/src/renderer/<path>（vendor + src/editor 物化视图）
 *   @xlsx/<path>              → src/<path>（插件自己的页面/共享模块）
 *   @genoffice/<pkg>[/<sub>]  → vendor/genoffice/packages/<pkg>/src/<sub>
 *
 * fork 优先是硬规则：只有「我们确实改过」的文件才进 src/editor，
 * 清单由 scripts/fork-diff.mjs 生成并由 scripts/verify.mjs 校验。
 */

import { existsSync, statSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const root = join(dirname(fileURLToPath(import.meta.url)), '..')
export const vendor = join(root, '../../vendor/genoffice')

const EXTENSIONS = ['', '.ts', '.tsx', '.css', '.json']

function firstFile(base) {
  for (const ext of EXTENSIONS) {
    const candidate = `${base}${ext}`
    try {
      if (existsSync(candidate) && statSync(candidate).isFile()) return candidate
    } catch {
      /* 下一个候选 */
    }
  }
  for (const index of ['index.ts', 'index.tsx']) {
    const candidate = join(base, index)
    try {
      if (existsSync(candidate) && statSync(candidate).isFile()) return candidate
    } catch {
      /* 下一个候选 */
    }
  }
  return null
}

/** `@sheets/App` → fork 优先、upstream 兜底。返回 null 表示不是这个前缀。 */
export function resolveSheets(source) {
  if (!source.startsWith('@sheets/')) return null
  return firstFile(join(root, 'build/sheets/src/renderer', source.slice('@sheets/'.length)))
}

/** `@xlsx/<path>` → src/<path>：fork 文件引用插件自有模块的唯一入口。 */
export function resolveOwn(source) {
  if (!source.startsWith('@xlsx/')) return null
  return firstFile(join(root, 'src', source.slice('@xlsx/'.length)))
}

/** `@genoffice/<pkg>[/<sub>]` → vendor 源码。 */
export function resolveGenoffice(source) {
  if (!source.startsWith('@genoffice/')) return null
  const rest = source.slice('@genoffice/'.length)
  const slash = rest.indexOf('/')
  const pkg = slash < 0 ? rest : rest.slice(0, slash)
  const sub = slash < 0 ? '' : rest.slice(slash + 1)
  return firstFile(join(vendor, 'packages', pkg, 'src', sub))
}

/** 统一入口：先插件自有，再物化视图，最后 vendor。 */
export function resolveSource(source) {
  return resolveOwn(source) ?? resolveSheets(source) ?? resolveGenoffice(source) ?? null
}

/** esbuild 插件。 */
export function esbuildGenoffice() {
  return {
    name: 'genoffice-source',
    setup(build) {
      build.onResolve({ filter: /^@(genoffice|sheets|xlsx)\// }, (args) => {
        const resolved = resolveSource(args.path)
        return resolved ? { path: resolved } : null
      })
    },
  }
}

/** vite 插件：与 esbuild 同一份规则（vite 内部也用它解析 css 子路径）。 */
export function viteGenoffice() {
  return {
    name: 'genoffice-source',
    enforce: 'pre',
    resolveId(source) {
      return resolveSource(source)
    },
  }
}

/** vite 的 alias 兜底（插件没命中时仍能解析裸包）。 */
export function viteAliases() {
  return [
    { find: /^@genoffice\/([^/]+)\/(.*)$/, replacement: join(vendor, 'packages/$1/src/$2') },
    { find: /^@genoffice\/([^/]+)$/, replacement: join(vendor, 'packages/$1/src/index.ts') },
    { find: /^@sheets\/(.*)$/, replacement: join(root, 'build/sheets/src/renderer/$1') },
    { find: /^@xlsx\/(.*)$/, replacement: join(root, 'src/$1') },
  ]
}
