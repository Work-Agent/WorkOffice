#!/usr/bin/env node
// 把构建好的插件装进 WorkStudio 的内置插件目录（开发 / 随包构建）或用户安装目录。
//
//   node scripts/install.mjs --target ../WorkStudio/src/plugins
//      → WorkStudio 构建时拷贝到 out/plugins，随应用分发。
//   node scripts/install.mjs --target "~/Library/Application Support/WorkStudio/plugins/node_modules"
//      → 用户安装目录；必须落在 plugins/node_modules 下，PluginService 只扫描这里
//        （或 plugins.md 里显式声明的本地来源）。
//
// 只复制 package.json 与 dist/；源码与 node_modules 不进入宿主。

import { cpSync, existsSync, mkdirSync, rmSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const args = process.argv.slice(2)
const targetArg = args[args.indexOf('--target') + 1]
if (!targetArg) {
  console.error('用法：node scripts/install.mjs --target <plugins-dir>')
  process.exit(1)
}
const target = resolve(targetArg.replace(/^~(?=$|\/|\\)/, homedir()))
const plugins = ['docx', 'xlsx', 'pptx']

mkdirSync(target, { recursive: true })
for (const id of plugins) {
  const pluginSource = join(root, 'packages', id)
  const pluginTarget = join(target, id)
  if (!existsSync(join(pluginSource, 'dist'))) {
    console.error(`缺少 packages/${id}/dist，请先运行 pnpm build`)
    process.exit(1)
  }
  rmSync(pluginTarget, { recursive: true, force: true })
  mkdirSync(pluginTarget, { recursive: true })
  cpSync(join(pluginSource, 'package.json'), join(pluginTarget, 'package.json'))
  cpSync(join(pluginSource, 'dist'), join(pluginTarget, 'dist'), { recursive: true })
  console.log(`installed ${id} → ${pluginTarget}`)
}
