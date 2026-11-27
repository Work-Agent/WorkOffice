#!/usr/bin/env node
// 生成 WorkStudio 可安装的插件 zip。
//
//   pnpm pack:plugins
//
// 输出：
//   release/<id>-<version>.zip
//   release/<id>-<version>.zip.sha256
//
// zip 根目录直接放 package.json + dist/，一个 zip 只包含一个插件。
// 不包含源码、node_modules、测试或构建脚本。
// WorkStudio 的安装清单由 WorkStudio 管理，不在本仓库生成。

import { createHash } from 'node:crypto'
import { mkdir, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { zipSync } from 'fflate'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const release = join(root, 'release')
/** 默认打全部插件：扫描 packages 下每个 package.json，取声明了 workstudio 的。 */
async function discoverPluginIds() {
  const ids = []
  for (const entry of await readdir(join(root, 'packages'), { withFileTypes: true })) {
    if (!entry.isDirectory()) continue
    const file = join(root, 'packages', entry.name, 'package.json')
    const pkg = await readFile(file, 'utf8').then(JSON.parse).catch(() => null)
    if (pkg?.workstudio) ids.push(entry.name)
  }
  return ids.sort()
}

const pluginIds = await discoverPluginIds()

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

const tagIndex = process.argv.indexOf('--tag')
const tag = tagIndex >= 0 ? process.argv[tagIndex + 1] : undefined
const onlyIndex = process.argv.indexOf('--only')
const onlyArg = onlyIndex >= 0 ? process.argv[onlyIndex + 1] : undefined
if (onlyArg) {
  const wanted = onlyArg.split(',').map((id) => id.trim()).filter(Boolean)
  for (const id of wanted)
    if (!pluginIds.includes(id)) throw new Error(`没有这个插件：packages/${id}`)
  pluginIds.splice(0, pluginIds.length, ...wanted)
}

// 不清空 release：一次只打某个插件时不该删掉别的产物（重名会覆盖）。
await mkdir(release, { recursive: true })

for (const id of pluginIds) {
  const pkgRoot = join(root, 'packages', id)
  const pkg = JSON.parse(await readFile(join(pkgRoot, 'package.json'), 'utf8'))
  const ws = pkg.workstudio ?? {}
  if (ws.id !== id) throw new Error(`${id}: workstudio.id 必须为 "${id}"`)
  if (!pkg.version) throw new Error(`${id}: 缺少 version`)
  const sdkEntry = pkg.exports?.['./workstudio']
  if (sdkEntry) {
    // WorkStudio SDK 新插件
    if (!ws.pages?.content) throw new Error(`${id}: 缺少 workstudio.pages.content`)
    for (const file of [sdkEntry, ws.pages.content, 'commands.json']) {
      const info = await stat(join(pkgRoot, file)).catch(() => null)
      if (!info?.isFile()) throw new Error(`${id}: 文件不存在 ${file}`)
    }
  } else {
    if (ws.apiVersion !== 1 && ws.apiVersion !== 2)
      throw new Error(`${id}: workstudio.apiVersion 必须为 1 或 2`)
    if (ws.apiVersion === 2 && ws.write === 'plan' && !ws.engine?.exec)
      throw new Error(`${id}: write: 'plan' 必须声明 engine.exec`)
    const ura = pkg.exports?.['./ura']
    if (!ura) throw new Error(`${id}: exports["./ura"] 缺失`)
    for (const file of [ura, ws.page]) {
      if (!file) throw new Error(`${id}: 清单缺少入口或页面`)
      const info = await stat(join(pkgRoot, file)).catch(() => null)
      if (!info?.isFile()) throw new Error(`${id}: 文件不存在 ${file}`)
    }
  }

  const entries = {}
  entries['package.json'] = new TextEncoder().encode(
    JSON.stringify(minimalManifest(pkg), null, 2) + '\n',
  )
  const distFiles = await walk(join(pkgRoot, 'dist'))
  for (const file of distFiles) {
    // dist/shared/** 是测试与诊断入口（verify.mjs / node --test 用），
    // 运行时已被 bundle 进 engine / page / ura，不进安装包。
    if (file.startsWith('shared/')) continue
    entries[`dist/${file}`] = new Uint8Array(await readFile(join(pkgRoot, 'dist', file)))
  }
  for (const file of ['commands.json', 'api.json', 'api.d.ts']) {
    const info = await stat(join(pkgRoot, file)).catch(() => null)
    if (info?.isFile()) entries[file] = new Uint8Array(await readFile(join(pkgRoot, file)))
  }
  // 技能目录必须随包：清单声明了 skills，缺目录时宿主会静默跳过该技能。
  for (const entry of ws.skills ?? []) {
    const skill = typeof entry === 'string' ? entry : entry?.dir
    if (!skill) continue
    const skillDir = join(pkgRoot, skill)
    const info = await stat(skillDir).catch(() => null)
    if (!info?.isDirectory()) continue
    for (const file of await walk(skillDir)) {
      entries[`${skill.replaceAll('\\', '/')}/${file}`] = new Uint8Array(
        await readFile(join(skillDir, file)),
      )
    }
  }
  // 执行点产物自证：引擎清单里的 sha256 必须与随包的 wasm 一致，
  // 否则可能把"上次构建的 wasm + 这次构建的清单"混进安装包。
  const engineJson = join(pkgRoot, 'dist/engine/xlsx-engine.json')
  const engineInfo = await stat(engineJson).catch(() => null)
  if (engineInfo?.isFile()) {
    const manifest = JSON.parse(await readFile(engineJson, 'utf8'))
    const wasm = new Uint8Array(await readFile(join(pkgRoot, 'dist/engine/xlsx-engine.wasm')))
    const actual = createHash('sha256').update(wasm).digest('hex')
    if (manifest.sha256 !== actual)
      throw new Error(
        `${id}: wasm 与清单 sha256 不一致（清单 ${String(manifest.sha256).slice(0, 12)}… / 实际 ${actual.slice(0, 12)}…），请重新构建`,
      )
  }
  // 技能目录随插件分发（v2 清单的 skills[].dir 指向它）。
  for (const skill of ws.skills ?? []) {
    const dir = typeof skill === 'string' ? skill : skill?.dir
    if (!dir) continue
    for (const file of await walk(join(pkgRoot, dir))) {
      entries[`${dir}/${file}`] = new Uint8Array(
        await readFile(join(pkgRoot, dir, file)),
      )
    }
  }

  const zip = zipSync(entries, { level: 6 })
  const base = `${id}-${pkg.version}${tag ? `-${tag}` : ''}`
  const zipPath = join(release, `${base}.zip`)
  await writeFile(zipPath, zip)
  const sha256 = createHash('sha256').update(zip).digest('hex')
  await writeFile(`${zipPath}.sha256`, sha256 + '\n')
  console.log(`packed ${base}.zip (${(zip.length / 1024 / 1024).toFixed(1)} MB)`)
  console.log(`  sha256 ${sha256}`)
}
