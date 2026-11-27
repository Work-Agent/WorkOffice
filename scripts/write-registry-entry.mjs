#!/usr/bin/env node
/**
 * 把 release/ 下打好的插件 zip 写入 WorkStudio PluginRegistry 的 plugins/<id>.json。
 *
 * 由 release workflow 调用；也可本地预演：
 *   REGISTRY_DIR=../PluginRegistry \
 *   REPO_URL=https://github.com/Work-Agent/WorkOffice \
 *   RELEASE_BASE_URL=https://github.com/Work-Agent/WorkOffice/releases/download/v0.2.0 \
 *   node scripts/write-registry-entry.mjs --only docx,xlsx
 *
 * 需要：
 *   REGISTRY_DIR      PluginRegistry 检出目录（写入 <dir>/plugins/<id>.json）
 *   REPO_URL          本仓库地址（https）
 *   RELEASE_TAG       发布 tag（默认 v<version>；含 v 或不含都行）
 *   RELEASE_BASE_URL  覆盖 zip 下载前缀（可选，默认 <REPO_URL>/releases/download/<tag>）
 *
 * 一个插件一个文件；已有条目只改版本 / 哈希 / 描述，保留人工维护的
 * category / tags / icon（除非清单里的 icon 是首次写入）。
 */

import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const releaseDir = join(root, 'release')

/** 注册表元数据：首次写入时使用，之后保留注册表里的值。 */
const META = {
  docx: {
    category: '文档',
    tags: ['docx', 'word', 'editor', 'agent'],
  },
  xlsx: {
    category: '表格',
    tags: ['xlsx', 'excel', 'workbook', 'agent'],
  },
  pptx: {
    category: '演示',
    tags: ['pptx', 'powerpoint', 'slides', 'agent'],
  },
  pdf: {
    category: '文档',
    tags: ['pdf', 'viewer', 'readonly'],
  },
}

function fail(message) {
  console.error(`::error::${message}`)
  process.exit(1)
}

/** 读环境变量，缺失直接失败（本地预演同样适用）。 */
function env(name) {
  const value = process.env[name]
  if (!value) fail(`缺少环境变量 ${name}`)
  return value
}

const registryDir = resolve(env('REGISTRY_DIR'))
const repo = env('REPO_URL')
if (!/^https:\/\//.test(repo)) fail(`REPO_URL 必须是 https 地址：${repo}`)

function parts(version) {
  return version.split('-')[0].split('.').map((value) => Number(value))
}

function isNewer(a, b) {
  const left = parts(a)
  const right = parts(b)
  for (let index = 0; index < 3; index += 1) {
    if ((left[index] ?? 0) !== (right[index] ?? 0)) {
      return (left[index] ?? 0) > (right[index] ?? 0)
    }
  }
  return false
}

/** 扫描 packages/，取声明了 workstudio 的插件。 */
async function discover() {
  const out = []
  for (const entry of await readdir(join(root, 'packages'), { withFileTypes: true })) {
    if (!entry.isDirectory()) continue
    const pkg = await readFile(join(root, 'packages', entry.name, 'package.json'), 'utf8')
      .then(JSON.parse)
      .catch(() => null)
    if (pkg?.workstudio?.id && pkg.version) out.push({ dir: entry.name, pkg })
  }
  return out.sort((a, b) => a.pkg.workstudio.id.localeCompare(b.pkg.workstudio.id))
}

const onlyIndex = process.argv.indexOf('--only')
const onlyArg = onlyIndex >= 0 ? process.argv[onlyIndex + 1] : undefined
const wanted = onlyArg
  ? onlyArg.split(',').map((value) => value.trim()).filter(Boolean)
  : undefined

const plugins = (await discover()).filter((plugin) =>
  wanted ? wanted.includes(plugin.dir) || wanted.includes(plugin.pkg.workstudio.id) : true,
)
if (plugins.length === 0) fail('没有可写入的插件')

// tag → zip 下载前缀；四种插件共用同一 tag，版本必须一致才能发一个 Release。
const version = plugins[0].pkg.version
for (const plugin of plugins) {
  if (plugin.pkg.version !== version) {
    fail(
      `插件版本不一致：${plugin.pkg.workstudio.id}@${plugin.pkg.version} / ${version}；` +
        '所有插件统一版本后才能发 Release',
    )
  }
}

const tag = process.env.RELEASE_TAG || `v${version}`
const sourceBase = (
  process.env.RELEASE_BASE_URL || `${repo}/releases/download/${tag}`
).replace(/\/+$/, '')

for (const plugin of plugins) {
  const ws = plugin.pkg.workstudio
  const id = ws.id
  const zipName = `${id}-${plugin.pkg.version}.zip`
  const zipPath = join(releaseDir, zipName)

  const sha = await readFile(`${zipPath}.sha256`, 'utf8').catch(() => null)
  if (sha === null) {
    fail(`缺少 ${zipName}.sha256，请先运行 pnpm pack:plugins`)
  }
  const sha256 = sha.trim().toLowerCase()
  if (!/^[a-f0-9]{64}$/.test(sha256)) fail(`${zipName}.sha256 内容非法：${sha}`)

  const entryFile = join(registryDir, 'plugins', `${id}.json`)
  let previous = {}
  try {
    previous = JSON.parse(await readFile(entryFile, 'utf8'))
  } catch {
    // 新插件：文件还不存在。
  }

  if (previous.version && isNewer(previous.version, plugin.pkg.version)) {
    fail(`注册表已有更新版本 ${id}@${previous.version}，拒绝回退到 ${plugin.pkg.version}`)
  }

  const meta = META[id] ?? {}
  const entry = {
    ...previous,
    id,
    name: previous.name ?? ws.label ?? id,
    description: previous.description ?? plugin.pkg.description ?? '',
    version: plugin.pkg.version,
    repo,
    source: `${sourceBase}/${zipName}`,
    sha256,
    icon: previous.icon ?? ws.icon ?? 'file',
    category: previous.category ?? meta.category ?? '办公',
    tags: previous.tags ?? meta.tags ?? [id],
  }

  await mkdir(dirname(entryFile), { recursive: true })
  await writeFile(entryFile, `${JSON.stringify(entry, null, 2)}\n`)
  console.log(`wrote plugins/${id}.json`)
  console.log(`  ${entry.version} ${entry.source}`)
  console.log(`  sha256 ${entry.sha256}`)
}

console.log(`\n${plugins.length} 个插件条目已写入 ${join(registryDir, 'plugins')}`)
