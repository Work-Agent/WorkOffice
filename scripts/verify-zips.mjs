#!/usr/bin/env node
// ─── 发布前最后一道闸：校验 release/*.zip ──────────────────────────────
//
// 用 node 读 zip（fflate）而不是 shell + unzip，原因：
//   1. `unzip -l "$ZIP" | grep -q ...` 在 `set -o pipefail` 下是真 bug：
//      grep -q 命中即退 → unzip 后续 write 拿到 SIGPIPE(141) → pipefail 让整步失败。
//      列表小的时候还会时红时绿，非常难查。
//   2. sha256 从"格式检查"升级成"内容比对"（以前只 grep 了 64 位十六进制）。
//   3. 少一个系统依赖（unzip），本地/CI 行为一致。
//
// 校验内容：
//   · 每个插件的 zip 与 .sha256 都在，且 sha256 就是该 zip 的哈希
//   · 条目齐全：package.json / dist/engine.js / dist/page/index.html / skills/**
//   · xlsx 额外要求 dist/xlsx-engine.wasm + .json，且清单里的 sha256 与 wasm 字节一致
//   · 包内 package.json：id / version 对得上，且没有 dependencies / devDependencies
//   · 不带 node_modules/、不带 dist/shared/**（那是测试入口，不进安装包）
//
// 用法：node scripts/verify-zips.mjs [--dir release]

import { createHash } from 'node:crypto'
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { unzipSync } from 'fflate'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const argv = process.argv.slice(2)
const dirIndex = argv.indexOf('--dir')
const releaseDir = dirIndex >= 0 ? resolve(argv[dirIndex + 1]) : join(root, 'release')

const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex')
const mb = (bytes) => `${(bytes.byteLength / 1024 / 1024).toFixed(1)} MB`

/** 需要在 zip 里出现的精确条目。 */
const REQUIRED_ENTRIES = ['package.json', 'dist/engine.js', 'dist/page/index.html']
const FORBIDDEN_PREFIXES = ['node_modules/', 'dist/shared/']
const XLSX_ENTRIES = ['dist/xlsx-engine.wasm', 'dist/xlsx-engine.json']

const failures = []

function fail(label, message) {
  failures.push(`${label}: ${message}`)
  console.error(`✗ ${label}: ${message}`)
}

function pluginManifests() {
  const packagesDir = join(root, 'packages')
  return readdirSync(packagesDir, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => {
      const path = join(packagesDir, entry.name, 'package.json')
      if (!existsSync(path)) return null
      const pkg = JSON.parse(readFileSync(path, 'utf8'))
      const id = pkg.workstudio?.id
      return id ? { id, version: pkg.version, pkg } : null
    })
    .filter(Boolean)
    .sort((a, b) => a.id.localeCompare(b.id))
}

function verifyOne({ id, version }) {
  const label = `${id}-${version}.zip`
  const zipPath = join(releaseDir, label)
  const shaPath = `${zipPath}.sha256`

  if (!existsSync(zipPath)) return fail(label, `缺文件（先跑 pnpm pack:plugins）：${zipPath}`)
  if (!existsSync(shaPath)) return fail(label, `缺文件：${shaPath}`)

  const bytes = readFileSync(zipPath)
  const actual = sha256(bytes)

  // sha256 内容比对：以前只校验"是不是 64 位十六进制"，写错了也能过。
  const declared = readFileSync(shaPath, 'utf8').trim().split(/\s+/)[0].toLowerCase()
  if (declared !== actual) fail(label, `sha256 不匹配（清单 ${declared.slice(0, 12)}… / 实际 ${actual.slice(0, 12)}…）`)

  let entries
  try {
    entries = unzipSync(new Uint8Array(bytes))
  } catch (error) {
    return fail(label, `不是合法 zip：${error instanceof Error ? error.message : error}`)
  }

  const names = Object.keys(entries)
  const has = (name) => names.includes(name)

  for (const name of REQUIRED_ENTRIES)
    if (!has(name)) fail(label, `缺条目 ${name}`)
  if (!names.some((name) => name.startsWith('skills/'))) fail(label, '缺 skills/（清单声明了 skills，缺目录宿主会静默跳过）')
  for (const prefix of FORBIDDEN_PREFIXES) {
    const hit = names.find((name) => name.startsWith(prefix))
    if (hit) fail(label, `不该包含 ${hit}`)
  }
  if (id === 'xlsx') {
    for (const name of XLSX_ENTRIES) if (!has(name)) fail(label, `缺条目 ${name}（否则发的是降级包）`)
    if (has('dist/xlsx-engine.wasm') && has('dist/xlsx-engine.json')) {
      const manifest = JSON.parse(new TextDecoder().decode(entries['dist/xlsx-engine.json']))
      const wasmSha = sha256(entries['dist/xlsx-engine.wasm'])
      if (manifest.sha256 !== wasmSha)
        fail(
          label,
          `wasm 与清单 sha256 不一致（清单 ${String(manifest.sha256).slice(0, 12)}… / 实际 ${wasmSha.slice(0, 12)}…）`,
        )
      if (manifest.size !== entries['dist/xlsx-engine.wasm'].byteLength)
        fail(label, `wasm 与清单体积不一致（${manifest.size} vs ${entries['dist/xlsx-engine.wasm'].byteLength}）`)
    }
  }

  if (has('package.json')) {
    const pkg = JSON.parse(new TextDecoder().decode(entries['package.json']))
    if (pkg.workstudio?.id !== id) fail(label, `包内 manifest id 是 ${pkg.workstudio?.id}，期望 ${id}`)
    if (pkg.version !== version) fail(label, `包内版本是 ${pkg.version}，期望 ${version}`)
    if (pkg.dependencies) fail(label, '包内 manifest 不该带 dependencies（运行时依赖必须已 bundle）')
    if (pkg.devDependencies) fail(label, '包内 manifest 不该带 devDependencies')
  }

  const ok = !failures.some((entry) => entry.startsWith(`${label}:`))
  console.log(`${ok ? '✓' : '✗'} ${label.padEnd(24)} ${String(names.length).padStart(4)} 个条目  ${mb(bytes)}`)
}

const plugins = pluginManifests()
if (plugins.length === 0) {
  console.error('✗ 没找到任何带 workstudio.id 的插件')
  process.exit(1)
}
if (!existsSync(releaseDir) || !statSync(releaseDir).isDirectory()) {
  console.error(`✗ release 目录不存在：${releaseDir}（先跑 pnpm pack:plugins）`)
  process.exit(1)
}

for (const plugin of plugins) verifyOne(plugin)

if (failures.length > 0) {
  console.error(`\n${failures.length} 项未通过`)
  process.exit(1)
}
console.log(`\n全部通过：${plugins.length} 个插件 zip`)
