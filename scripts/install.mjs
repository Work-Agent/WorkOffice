#!/usr/bin/env node
// 把构建结果安装到 WorkStudio 的插件目录。
//
//   node scripts/install.mjs
//     → <userData>/plugins/installed/<id>/
//
//   node scripts/install.mjs --target ../WorkStudio/src/plugins
//     → 内置插件目录（仅开发/随应用构建时使用）
//
// 只复制 package.json 与 dist/；源码、node_modules、测试不进入宿主。
// 新机制按目录扫描插件，不需要 npm install。

import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, statSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const pluginIds = ['docx', 'xlsx', 'pptx']

function defaultPluginsRoot() {
  if (process.platform === 'darwin') {
    return join(homedir(), 'Library', 'Application Support', 'WorkStudio', 'plugins')
  }
  if (process.platform === 'win32') {
    return join(process.env.APPDATA ?? join(homedir(), 'AppData', 'Roaming'), 'WorkStudio', 'plugins')
  }
  return join(process.env.XDG_CONFIG_HOME ?? join(homedir(), '.config'), 'WorkStudio', 'plugins')
}

const args = process.argv.slice(2)
const targetArg = args.includes('--target') ? args[args.indexOf('--target') + 1] : undefined
const target = targetArg
  ? resolve(targetArg.replace(/^~(?=$|\/|\\)/, homedir()))
  : join(defaultPluginsRoot(), 'installed')

if (/[/\\]node_modules$/.test(target)) {
  console.warn(
    '警告：plugins/node_modules 已不是 WorkStudio 的插件扫描目录；' +
      '请使用 plugins/installed，或通过设置 → 插件添加本地目录。',
  )
}

function checkManifest(id, pkgRoot) {
  const pkg = JSON.parse(readFileSync(join(pkgRoot, 'package.json'), 'utf8'))
  const ws = pkg.workstudio ?? {}
  if (ws.id !== id) throw new Error(`${id}: workstudio.id 必须为 "${id}"`)
  if (ws.apiVersion !== 1) throw new Error(`${id}: workstudio.apiVersion 必须为 1`)
  if (!pkg.version) throw new Error(`${id}: 缺少 version`)
  const ura = pkg.exports?.['./ura']
  if (!ura) throw new Error(`${id}: exports["./ura"] 缺失`)
  for (const file of [ura, ws.page]) {
    if (!file || !existsSync(join(pkgRoot, file))) throw new Error(`${id}: 文件不存在 ${file}`)
  }
  return pkg
}

mkdirSync(target, { recursive: true })
for (const id of pluginIds) {
  const pluginSource = join(root, 'packages', id)
  const pluginTarget = join(target, id)
  if (!existsSync(join(pluginSource, 'dist'))) {
    console.error(`缺少 packages/${id}/dist，请先运行 pnpm build`)
    process.exit(1)
  }
  const pkg = checkManifest(id, pluginSource)
  rmSync(pluginTarget, { recursive: true, force: true })
  mkdirSync(pluginTarget, { recursive: true })
  cpSync(join(pluginSource, 'package.json'), join(pluginTarget, 'package.json'))
  cpSync(join(pluginSource, 'dist'), join(pluginTarget, 'dist'), { recursive: true })
  console.log(`installed ${id}@${pkg.version} → ${pluginTarget}`)
}
console.log('\n重启 WorkStudio 后生效。')
