/**
 * 生成「渲染器视图」：vendor/apps/sheets/src 全量 + src/editor 覆盖。
 *
 * 为什么不逐文件回落：renderer 内部全是相对 import（./App、../shared/…），
 * 只覆盖解析 `@sheets/*` 无法让 fork 的 main.tsx 找到未 fork 的 ./App。
 * 所以这里物化一棵完整目录树，fork 只提供「我们改过的文件」。
 *
 *   vendor/genoffice/apps/sheets/src  →  build/sheets/src
 *   src/editor/**                       →  build/sheets/src/renderer/**
 *   editor.manifest.json.removed[]      →  从 build 树里删掉
 */

import { cpSync, existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const manifest = JSON.parse(readFileSync(join(root, 'editor.manifest.json'), 'utf8'))
const vendor = resolve(root, '../../vendor/genoffice', manifest.vendorTree ?? 'apps/sheets/src')
const dest = join(root, 'build/sheets/src')

if (!existsSync(vendor)) {
  console.error(`缺少 vendor 源码：${vendor}\n先运行 pnpm setup 拉取 sparse 路径。`)
  process.exit(1)
}

rmSync(join(root, 'build'), { recursive: true, force: true })
mkdirSync(dest, { recursive: true })
cpSync(vendor, dest, { recursive: true })

const forkRoot = join(root, 'src/editor')
if (existsSync(forkRoot)) cpSync(forkRoot, join(dest, 'renderer'), { recursive: true })
// 渲染层的共享类型（desktop-api / ipc-channels）与它同版本，整体覆盖
const sharedFork = join(root, 'src/editor-shared')
if (existsSync(sharedFork)) cpSync(sharedFork, join(dest, 'shared'), { recursive: true })

for (const rel of manifest.removed ?? []) {
  rmSync(join(dest, 'renderer', rel), { recursive: true, force: true })
}

console.log(`renderer → build/sheets/src（vendor ${manifest.vendorCommit} + ${manifest.forked.length} 个 fork 文件）`)
