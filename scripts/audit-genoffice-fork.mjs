#!/usr/bin/env node
// 审计 docx 的 renderer fork 与 GenOffice 上游的差异，回答一个问题：
//
//   fork 里哪些文件已经可以改回从 vendor/genoffice 子模块直接取？
//
// 判定分三层（都由本脚本计算，不靠人工印象）：
//   1. 逐文件 diff（先归一化 CRLF；本地工作区是 CRLF，vendor 是 LF）
//   2. “闭包一致”：该文件的相对 import 闭包里每个文件都与上游一致，
//      才可能安全地整体换成 vendor 版本；只差一个文件都不算。
//   3. i18n 键集合：本地被上游覆盖 = 纯旧快照；本地独有键必须做 overlay。
//
// 用法：
//   node scripts/audit-genoffice-fork.mjs
//   node scripts/audit-genoffice-fork.mjs --fork packages/docx/src/editor
//
// 注意：闭包一致 ≠ 现在就值得换。混用 vendor/本地路径会把 upstream 的
// 兄弟模块带进来（见报告 §4），只在重建 fork 时才有意义。

import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const arg = (name, fallback) => {
  const i = process.argv.indexOf(name)
  return i < 0 ? fallback : process.argv[i + 1]
}
const FORK = resolve(root, arg('--fork', 'packages/docx/src/editor'))
const UP = resolve(root, arg('--upstream', 'vendor/genoffice/apps/docs/src/renderer'))

if (!existsSync(FORK) || !existsSync(UP)) {
  console.error(`缺少目录：\n  fork=${FORK}\n  upstream=${UP}\n先跑 pnpm setup 初始化 vendor/genoffice。`)
  process.exit(1)
}

const walk = (dir, base = dir, out = []) => {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name)
    if (statSync(p).isDirectory()) walk(p, base, out)
    else out.push(relative(base, p).replaceAll('\\', '/'))
  }
  return out
}
const read = (p) => readFileSync(p)
const text = (p) => readFileSync(p, 'utf8').replaceAll('\r\n', '\n')
const isText = (p) => !read(p).includes(0)
const isFile = (p) => {
  try {
    return statSync(p).isFile()
  } catch {
    return false
  }
}

const forkFiles = walk(FORK)
const upFiles = new Set(walk(UP))
const same = new Set()
const different = []
for (const f of forkFiles) {
  if (!upFiles.has(f)) continue
  const a = join(FORK, f)
  const b = join(UP, f)
  if (read(a).equals(read(b)) || (isText(a) && text(a) === text(b))) same.add(f)
  else different.push(f)
}
const onlyLocal = forkFiles.filter((f) => !upFiles.has(f))
const onlyUp = [...upFiles].filter((f) => !forkFiles.includes(f))

// —— 闭包一致：沿相对 import 走，任何一环不一致就 blocked ——
const resolveSpec = (from, spec) => {
  if (!spec.startsWith('.')) return null
  const base = resolve(dirname(join(FORK, from)), spec)
  for (const c of [base, `${base}.ts`, `${base}.tsx`, join(base, 'index.ts'), join(base, 'index.tsx')]) {
    if (isFile(c)) return relative(FORK, c).replaceAll('\\', '/')
  }
  return null
}
const depsOf = (f) => {
  const out = []
  for (const m of text(join(FORK, f)).matchAll(/(?:from\s+|import\s*\(\s*)['"]([^'"]+)['"]/g)) {
    const dep = resolveSpec(f, m[1])
    if (dep) out.push(dep)
  }
  return out
}
const memo = new Map()
const closure = (f, stack = new Set()) => {
  if (memo.has(f)) return memo.get(f)
  if (stack.has(f)) return { safe: true, blockedBy: null }
  stack.add(f)
  let blockedBy = null
  if (!same.has(f)) blockedBy = f
  else {
    for (const dep of depsOf(f)) {
      if (!upFiles.has(dep)) {
        blockedBy = `${dep} (only-local)`
        break
      }
      const r = closure(dep, stack)
      if (!r.safe) {
        blockedBy = r.blockedBy ?? dep
        break
      }
    }
  }
  stack.delete(f)
  const r = { safe: blockedBy === null, blockedBy }
  memo.set(f, r)
  return r
}
const closedSafe = []
const blocked = []
for (const f of same) {
  const r = closure(f)
  if (r.safe) closedSafe.push(f)
  else blocked.push([f, r.blockedBy])
}

// —— i18n 键集合对账（只对 per-locale 词典；strings-*.ts 是汇总/兼容层）——
const keys = (p) => new Set([...readFileSync(p, 'utf8').matchAll(/^\s{2}([A-Za-z0-9_]+):/gm)].map((m) => m[1]))
const commonRibbon = resolve(root, 'vendor/genoffice/packages/i18n/src/common-ribbon.ts')
const commonKeys = existsSync(commonRibbon) ? keys(commonRibbon) : new Set()
const i18nDirs = ['i18n/app', 'i18n/ribbon', 'i18n/ai']
const i18nRows = []
for (const dir of i18nDirs) {
  const f = join(FORK, dir, 'zh.ts')
  const u = join(UP, dir, 'zh.ts')
  if (!existsSync(f) || !existsSync(u)) continue
  // ribbon 的公共词已抽到 packages/i18n/common-ribbon.ts，要并进来才是真实可用键集
  const up = dir === 'i18n/ribbon' ? new Set([...keys(u), ...commonKeys]) : keys(u)
  const l = keys(f)
  const missing = [...l].filter((k) => !up.has(k))
  i18nRows.push({ dir, localKeys: l.size, upKeys: up.size, missing })
}

const printList = (title, list, fmt = (x) => x) => {
  console.log(`\n## ${title} (${list.length})`)
  if (!list.length) return console.log('(none)')
  for (const x of list.sort()) console.log(`  ${fmt(x)}`)
}

console.log('docx renderer fork ↔ genoffice apps/docs renderer')
console.log(`  fork     ${relative(root, FORK)} (${forkFiles.length} files)`)
console.log(`  upstream ${relative(root, UP)} (${upFiles.size} files)`)
console.log('\n## 概况')
console.log(`  完全相同（CRLF 归一化后） ${same.size}`)
console.log(`  有差异                      ${different.length}`)
console.log(`  仅 fork 有                  ${onlyLocal.length}`)
console.log(`  仅 upstream 有              ${onlyUp.length}`)
console.log('\n## 可直接取上游的前提：闭包一致')
console.log(`  闭包一致 ${closedSafe.length} / 有差异 ${same.size}（其余 ${blocked.length} 个“文件相同但依赖已改”）`)

printList('闭包一致（可以考虑改为 submodule import）', closedSafe)
printList('文件相同但闭包被本地改动阻断', blocked, ([f, by]) => `${f}  ->  ${by}`)
printList('仅 fork 有（本地新增，必须保留）', onlyLocal)
printList('仅 upstream 有（上游新增，未进 fork）', onlyUp)

console.log('\n## i18n 键集合对账')
for (const row of i18nRows) {
  console.log(`  ${row.dir}/zh.ts  本地 ${row.localKeys} / 上游 ${row.upKeys} / 本地独有 ${row.missing.length}`)
  if (row.missing.length) console.log(`    本地独有键：${row.missing.join(', ')}`)
}
