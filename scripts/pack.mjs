#!/usr/bin/env node
// 生成 WorkStudio 可安装的插件 zip。
//
//   pnpm pack:plugins
//   node scripts/pack.mjs --base-url https://github.com/org/repo/releases/download/v0.1.0
//
// 输出：
//   release/<id>-<version>.zip
//   release/<id>-<version>.zip.sha256
//   release/registry.md
//
// zip 根目录直接放 package.json + dist/，一个 zip 只包含一个插件。
// 不包含源码、node_modules、测试或构建脚本。

import { createHash } from 'node:crypto'
import { mkdir, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { zipSync } from 'fflate'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const release = join(root, 'release')
const pluginIds = ['docx', 'xlsx', 'pptx']
const baseUrlArg = process.argv.indexOf('--base-url')
const baseUrl = (baseUrlArg >= 0 ? process.argv[baseUrlArg + 1] : 'https://example.com')
  ?.replace(/\/$/, '')
const registryLines = ['# 可安装插件', '']

async function walk(dir, base = dir) {
  const out = []
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const file = join(dir, entry.name)
    if (entry.isDirectory()) out.push(...(await walk(file, base)))
    else if (entry.isFile()) out.push(relative(base, file).replaceAll('\\', '/'))
  }
  return out
}

function minimalManifest(pkg) {
  return {
    name: pkg.name,
    version: pkg.version,
    description: pkg.description,
    author: pkg.author,
    license: pkg.license,
    type: 'module',
    exports: pkg.exports,
    workstudio: pkg.workstudio,
  }
}

await rm(release, { recursive: true, force: true })
await mkdir(release, { recursive: true })

for (const id of pluginIds) {
  const pkgRoot = join(root, 'packages', id)
  const pkg = JSON.parse(await readFile(join(pkgRoot, 'package.json'), 'utf8'))
  const ws = pkg.workstudio ?? {}
  if (ws.id !== id) throw new Error(`${id}: workstudio.id 必须为 "${id}"`)
  if (!pkg.version) throw new Error(`${id}: 缺少 version`)
  if (ws.apiVersion !== 1) throw new Error(`${id}: workstudio.apiVersion 必须为 1`)
  const ura = pkg.exports?.['./ura']
  if (!ura) throw new Error(`${id}: exports["./ura"] 缺失`)
  for (const file of [ura, ws.page]) {
    if (!file) throw new Error(`${id}: 清单缺少入口或页面`)
    const info = await stat(join(pkgRoot, file)).catch(() => null)
    if (!info?.isFile()) throw new Error(`${id}: 文件不存在 ${file}`)
  }

  const entries = {}
  entries['package.json'] = new TextEncoder().encode(
    JSON.stringify(minimalManifest(pkg), null, 2) + '\n',
  )
  const distFiles = await walk(join(pkgRoot, 'dist'))
  for (const file of distFiles) {
    entries[`dist/${file}`] = new Uint8Array(await readFile(join(pkgRoot, 'dist', file)))
  }

  const zip = zipSync(entries, { level: 6 })
  const base = `${id}-${pkg.version}`
  const zipPath = join(release, `${base}.zip`)
  await writeFile(zipPath, zip)
  const sha256 = createHash('sha256').update(zip).digest('hex')
  await writeFile(`${zipPath}.sha256`, sha256 + '\n')
  console.log(`packed ${base}.zip (${(zip.length / 1024 / 1024).toFixed(1)} MB)`)
  console.log(`  sha256 ${sha256}`)
  const registryLine = `- [${ws.label ?? id}](${baseUrl}/${base}.zip#sha256=${sha256})`
  registryLines.push(registryLine)
  console.log(`  registry: ${registryLine}`)
}

await writeFile(join(release, 'registry.md'), registryLines.join('\n') + '\n')
console.log('wrote release/registry.md')
