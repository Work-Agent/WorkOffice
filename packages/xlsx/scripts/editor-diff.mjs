/**
 * fork 清单：哪些页面文件是我们改过的，哪些已经和上游一致。
 *
 *   node scripts/editor-diff.mjs           # 重写 editor.manifest.json 的 forked[]
 *   node scripts/editor-diff.mjs --check   # CI：清单与 src/editor 实况不一致就失败
 *
 * 判定按内容（忽略 CRLF）：src/editor 里与上游逐字节相同的文件是「陈旧 fork」，
 * 应当删掉回落上游；清单里缺失 / 多出的条目同样失败。
 */

import { createHash } from 'node:crypto'
import { existsSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const manifestPath = join(root, 'editor.manifest.json')
const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'))
const vendorRoot = resolve(root, '../../vendor/genoffice', manifest.vendorSource ?? 'apps/sheets/src/renderer')
const forkRoot = join(root, 'src/editor')
const check = process.argv.includes('--check')

const hash = (file) =>
  createHash('sha256')
    .update(readFileSync(file).toString('binary').replace(/\r\n/g, '\n'))
    .digest('hex')

function walk(dir) {
  if (!existsSync(dir)) return []
  const out = []
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name)
    if (entry.isDirectory()) out.push(...walk(path))
    else if (entry.isFile()) out.push(path)
  }
  return out
}

const stale = []
const pending = []
const collect = (base, upstreamBase, prefix = '') => {
  for (const file of walk(base)) {
    const rel = prefix + relative(base, file).replaceAll('\\', '/')
    const upstream = join(upstreamBase, relative(base, file))
    if (existsSync(upstream) && hash(upstream) === hash(file)) stale.push(rel)
    else pending.push(rel)
  }
}
collect(forkRoot, vendorRoot)
const sharedFork = join(root, 'src/editor-shared')
if (existsSync(sharedFork)) {
  collect(sharedFork, resolve(root, '../../vendor/genoffice', 'apps/sheets/src/shared'), 'shared/')
}
const listed = new Set(manifest.forked ?? [])
const missing = pending.filter((rel) => !listed.has(rel))
const extra = [...listed].filter((rel) => !pending.includes(rel))
const problems = []
if (stale.length > 0) problems.push(`与上游一致（应删除，回落 vendor）：\n  ${stale.join('\n  ')}`)
if (missing.length > 0) problems.push(`清单缺少：\n  ${missing.join('\n  ')}`)
if (extra.length > 0) problems.push(`清单多出（文件已不存在）：\n  ${extra.join('\n  ')}`)

if (check) {
  if (problems.length > 0) {
    console.error(`fork 清单不一致：\n\n${problems.join('\n\n')}`)
    process.exit(1)
  }
  console.log(`fork 清单一致：${pending.length} 个文件（vendor ${manifest.vendorCommit}）`)
  process.exit(0)
}

const next = {
  vendorCommit: manifest.vendorCommit,
  vendorSource: manifest.vendorSource ?? 'apps/sheets/src/renderer',
  notes: manifest.notes ?? {},
  forked: pending.sort(),
  removed: (manifest.removed ?? []).sort(),
}
writeFileSync(manifestPath, `${JSON.stringify(next, null, 2)}\n`)
console.log(`写回 editor.manifest.json：${pending.length} 个 fork 文件`)
if (stale.length > 0) console.log(`注意：${stale.length} 个文件与上游一致，可从 src/editor 删除`)
