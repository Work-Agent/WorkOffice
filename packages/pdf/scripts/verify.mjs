#!/usr/bin/env node
// 装载前自检：清单 / 产物 / 静态描述对齐 / 旧面残留。

import { readFile, readdir, stat } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const problems = []
const warnings = []
const fail = (message) => problems.push(message)
const warn = (message) => warnings.push(message)

const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))
const ws = pkg.workstudio ?? {}

// ─── 清单 ───────────────────────────────────────────────────────────

if (ws.id !== 'pdf') fail(`workstudio.id 必须是 pdf（当前 ${ws.id}）`)
if (pkg.name !== '@workstudio/plugin-pdf') {
  fail(`包名应为 @workstudio/plugin-pdf（当前 ${pkg.name}）`)
}
for (const gone of ['apiVersion', 'write', 'content', 'model', 'engine', 'page', 'tools']) {
  if (gone in ws) fail(`清单不应再有 ${gone}`)
}
if (typeof ws.pages?.content !== 'string') fail('缺少 workstudio.pages.content')
if (pkg.exports?.['./workstudio'] !== './dist/engine.js') {
  fail(`exports['./workstudio'] 应为 ./dist/engine.js（当前 ${pkg.exports?.['./workstudio']}）`)
}
const claims = ws.claims ?? []
if (!claims.some((claim) => (claim.extensions ?? []).includes('pdf'))) {
  fail('claims 缺少 pdf')
}
if (!claims.some((claim) => (claim.priority ?? 0) >= 100)) {
  warn('claims.priority 建议 100')
}
for (const skill of ws.skills ?? []) {
  const info = await stat(join(root, skill)).catch(() => null)
  if (!info?.isDirectory()) fail(`skills 目录不存在：${skill}`)
}

// ─── 产物 ───────────────────────────────────────────────────────────

for (const file of [
  'dist/engine.js',
  'dist/text.js',
  'dist/pdf.worker.mjs',
  'dist/page/index.html',
  'dist/page/assets/pdfjs/cmaps',
  'dist/page/assets/pdfjs/standard_fonts',
  'dist/page/assets/pdfjs/wasm',
  'dist/page/assets/pdfjs/iccs',
  'commands.json',
]) {
  const info = await stat(join(root, file)).catch(() => null)
  if (!info) fail(`${file} 缺失（先 npm run build）`)
}

// ─── 静态描述 ↔ 动态命令 ────────────────────────────────────────────

const contribution = (await import(
  `${pathToFileURL(join(root, 'dist/engine.js')).href}?verify=${Date.now()}`
)).workstudio
if (!contribution) fail('engine 没有导出 workstudio')
else {
  const dynamic = contribution.commands ?? {}
  const staticFile = JSON.parse(await readFile(join(root, 'commands.json'), 'utf8'))
  const staticCommands = Array.isArray(staticFile) ? staticFile : (staticFile.commands ?? [])
  const staticNames = new Set(staticCommands.map((command) => command.name))

  for (const name of Object.keys(dynamic)) {
    if (!staticNames.has(name)) fail(`commands.json 缺少 ${name}`)
  }
  for (const name of staticNames) {
    if (!(name in dynamic)) fail(`静态命令没有实现：${name}`)
  }

  const expected = ['pdf.read', 'pdf.load', 'pdf.chunk', 'pdf.saveCopy']
  for (const name of expected) {
    if (!(name in dynamic)) fail(`缺少命令 ${name}`)
  }
  if (Object.keys(dynamic).length !== expected.length) {
    fail(`命令数应为 ${expected.length}（当前 ${Object.keys(dynamic).length}）`)
  }
  if (dynamic['pdf.read']?.expose?.agent !== true) fail('pdf.read 必须暴露给助手')
  for (const name of ['pdf.load', 'pdf.chunk', 'pdf.saveCopy']) {
    if (dynamic[name]?.expose?.agent === true) fail(`${name} 不得暴露给助手`)
  }
  if (!contribution.resources?.pdf) fail('缺少 resources.pdf')
  if (contribution.events === true) fail('只读 PDF 不需要 events 通道')
}

// ─── 旧面残留 ───────────────────────────────────────────────────────

async function walk(dir) {
  const out = []
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name)
    if (entry.isDirectory()) out.push(...(await walk(path)))
    else out.push(path)
  }
  return out
}

const LEGACY = [
  ['window.extension', /window\.extension/],
  ['ura.page', /\bura\.page\b/],
  ['apiVersion', /apiVersion/],
  ['registerCodec', /registerCodec/],
  ['registerPresenter', /registerPresenter/],
  ['registerSidecar', /registerSidecar/],
  ['appBridge', /appBridge/],
  ['sessionId', /sessionId/],
]
for (const file of await walk(join(root, 'src'))) {
  if (/\.(png|jpg|jpeg|gif|webp|svg|css|woff2?|ttf)$/i.test(file)) continue
  const text = await readFile(file, 'utf8').catch(() => '')
  for (const [name, pattern] of LEGACY) {
    if (pattern.test(text)) fail(`${file}: 仍含旧面 ${name}`)
  }
}

// ─── 输出 ───────────────────────────────────────────────────────────

for (const message of warnings) console.warn(`  ! ${message}`)
if (problems.length > 0) {
  for (const message of problems) console.error(`  ✗ ${message}`)
  console.error(`\npdf verify failed: ${problems.length} problem(s)`)
  process.exit(1)
}
console.log('pdf verify: manifest / artifacts / commands / resources / legacy checks passed')
