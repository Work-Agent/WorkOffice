#!/usr/bin/env node
// 装载前自检：清单 / 产物 / 静态描述对齐 / 旧面残留 / 一次往返。
//
// 与 tests/engine.test.mjs 共用 harness：这里多查清单与文件，测试查行为。

import { readFile, readdir, stat } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { applyOnPageBytes, createHarness, loadEngine, slidePathsOf, slideXmlOf } from '../tests/harness.mjs'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const problems = []
const warnings = []
const fail = (message) => problems.push(message)
const warn = (message) => warnings.push(message)

const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))
const ws = pkg.workstudio ?? {}

// ─── 清单 ───────────────────────────────────────────────────────────

if (ws.id !== 'pptx') fail(`workstudio.id 必须是 pptx（当前 ${ws.id}）`)
if (pkg.name !== '@workstudio/plugin-pptx') fail(`包名应为 @workstudio/plugin-pptx（当前 ${pkg.name}）`)
for (const gone of ['apiVersion', 'write', 'content', 'model', 'engine', 'page'])
  if (gone in ws) fail(`清单不应再有 ${gone}`)
if (typeof ws.pages?.content !== 'string') fail('缺少 workstudio.pages.content')
const entry = pkg.exports?.['./workstudio']
if (entry !== './dist/engine.js') fail(`exports['./workstudio'] 应为 ./dist/engine.js（当前 ${entry}）`)
const EXTS = new Set(['pptx', 'pptm', 'potx'])
for (const claim of ws.claims ?? []) {
  for (const ext of claim.extensions ?? []) if (!EXTS.has(ext)) fail(`claims 里有多余扩展名 ${ext}`)
}
for (const ext of EXTS)
  if (!(ws.claims ?? []).some((claim) => (claim.extensions ?? []).includes(ext)))
    fail(`claims 缺少 ${ext}`)
if (!(ws.claims ?? []).some((claim) => (claim.priority ?? 0) >= 100)) warn('claims.priority 建议 100')
for (const skill of ws.skills ?? []) {
  const info = await stat(join(root, skill)).catch(() => null)
  if (!info?.isDirectory()) fail(`skills 目录不存在：${skill}`)
}

// ─── 产物 ───────────────────────────────────────────────────────────

for (const file of ['dist/engine.js', 'dist/seed.pptx', 'dist/page/index.html', 'commands.json']) {
  const info = await stat(join(root, file)).catch(() => null)
  if (!info?.isFile()) fail(`${file} 缺失（先 npm run build）`)
}

// ─── 静态描述 ↔ 动态命令 ────────────────────────────────────────────

const contribution = await loadEngine()
const dynamic = contribution.commands ?? {}
const staticFile = JSON.parse(await readFile(join(root, 'commands.json'), 'utf8'))
const staticCommands = Array.isArray(staticFile) ? staticFile : (staticFile.commands ?? [])
const staticNames = new Set(staticCommands.map((command) => command.name))

for (const name of Object.keys(dynamic))
  if (!staticNames.has(name)) fail(`commands.json 缺少 ${name}`)
for (const name of staticNames)
  if (!(name in dynamic)) fail(`静态命令没有实现：${name}`)
if (Object.keys(dynamic).length !== 13) fail(`命令数应为 13（当前 ${Object.keys(dynamic).length}）`)

const AGENT = ['guide', 'read', 'check', 'create', 'apply']
const PAGE = ['load', 'save', 'saveAs', 'attach', 'pending', 'report', 'detach', 'ack']
for (const action of AGENT) {
  const command = dynamic[`pptx.${action}`]
  if (!command) fail(`缺少助手命令 pptx.${action}`)
  else if (command.expose?.agent !== true) fail(`${action} 必须 expose.agent: true`)
}
for (const action of PAGE) {
  const command = dynamic[`pptx.${action}`]
  if (!command) fail(`缺少页面命令 pptx.${action}`)
  else if (command.expose?.agent === true) fail(`${action} 不得暴露给助手`)
}
if (!contribution.resources?.pptx) fail('缺少 resources.pptx')
if (contribution.events !== true) fail('需要 events: true（页面队列走事件）')

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
  ['agent-request', /agent-request/],
  ['saveBegin', /saveBegin/],
]
for (const file of await walk(join(root, 'src'))) {
  if (/\.(png|jpg|jpeg|gif|webp|svg|css|woff2?|ttf)$/i.test(file)) continue
  const text = await readFile(file, 'utf8').catch(() => '')
  for (const [label, pattern] of LEGACY) {
    if (pattern.test(text)) fail(`${file.replace(root, '.')} 仍有旧面：${label}`)
  }
}
for (const file of ['src/ura.ts', 'src/ura-types.ts', 'src/sdk.ts', 'src/page/agent.ts'])
  if (await stat(join(root, file)).catch(() => null)) fail(`旧实现没有删除：${file}`)

// ─── 一次往返 ───────────────────────────────────────────────────────

const seed = new Uint8Array(await readFile(join(root, 'dist/seed.pptx')))
const harness = createHarness(contribution, seed)
const REF = 'file:///work/sample.pptx'
harness.files.set(REF, seed)

try {
  const guide = await harness.run('guide', { ref: REF })
  for (const group of ['text', 'element', 'insert', 'table', 'slide', 'deck', 'chart'])
    if (!String(guide).includes(group)) fail(`guide 缺少分组 ${group}`)

  const read = await harness.run('read', { ref: REF, what: 'context' })
  if (!String(read).toLowerCase().includes('slide')) fail('read context 没有返回页概览')

  const audit = await harness.run('check', { ref: REF, audit: true })
  if (!/几何审计/.test(String(audit))) fail('check audit 没有返回几何审计')

  const noPage = await harness.run('apply', { ref: REF, commands: [{ op: 'addElement' }] }).then(
    () => null,
    (error) => String(error.message ?? error),
  )
  if (!noPage?.includes('command.page-unavailable')) fail(`无页面 apply 应回 page-unavailable（当前 ${noPage}）`)

  const loaded = await harness.run('load', { ref: REF })
  if (loaded.bytes.byteLength !== seed.byteLength) fail('load 字节长度不符')
  if (!/^[0-9a-f]{64}$/.test(loaded.hash)) fail('load 没有返回 sha256')

  // 新建（文字稿）
  const created = await harness.run('create', {
    ref: 'file:///work/new.pptx',
    slides: JSON.stringify([
      { title: '年度汇报', bullets: ['收入增长', '成本下降'], notes: '讲稿' },
      { title: '第二页', body: '正文' },
    ]),
  })
  const createdBytes = harness.files.get('file:///work/new.pptx')
  if (!createdBytes) fail('create 没有写入目标')
  else if (slidePathsOf(createdBytes).length < 2) fail('新建页数不足 2')
  else if (!slideXmlOf(createdBytes).includes('年度汇报')) fail('新建第一页缺少标题')
  if (!created?.hash) fail('create 没有返回 hash')

  // 新建（设计稿）
  const design = JSON.stringify([
    {
      background: '#0E1A2B',
      elements: [
        { type: 'text', x: 80, y: 200, w: 900, h: 140, paragraphs: [{ runs: [{ text: 'Verify deck', sizePt: 54, bold: true, color: '#FFFFFF' }] }] },
        { type: 'shape', shape: 'rect', x: 80, y: 400, w: 300, h: 8, fill: '#FFB000' },
      ],
    },
  ])
  await harness.run('create', { ref: 'file:///work/design.pptx', pages: design })
  const designBytes = harness.files.get('file:///work/design.pptx')
  if (!designBytes) fail('设计稿 create 没有写入目标')
  else {
    const xml = slideXmlOf(designBytes)
    if (!xml.includes('Verify deck')) fail('设计稿第一页缺少文本')
    if (!/<a:prstGeom prst="rect"/.test(xml)) fail('设计稿缺少形状元素')
    if (!xml.includes('<p:bg>')) fail('设计稿背景没有被提升成原生背景')
  }

  // 页面通道：attach → apply（入队）→ 假页面执行 → save → ack
  const attached = await harness.run('attach', { ref: REF, pageId: 'page-1' })
  if (!attached?.hash) fail('attach 没有返回 hash')
  const { id, promise } = await harness.startApply({
    ref: REF,
    commands: [
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'textbox',
        offset: { x: 1000000, y: 1000000, cx: 3000000, cy: 1000000 },
        paragraphs: [{ runs: [{ text: 'Hello pptx' }] }],
      },
    ],
  })
  const request = await harness.pending(id)
  if (request.type !== 'apply' || request.commands.length !== 1) fail('pending 正文不对')
  const edited = await applyOnPageBytes(harness.files.get(REF), request.commands)
  if (!edited.ok) fail(`页面事务失败：${edited.error}`)
  else {
    const receipt = await harness.run('save', {
      ref: REF,
      pageId: 'page-1',
      baseHash: attached.hash,
      bytes: edited.bytes,
    })
    if (!receipt?.hash) fail('save 没有返回回执')
    if (!slideXmlOf(harness.files.get(REF)).includes('Hello pptx')) fail('保存后缺少新文本')
    await harness.ack(id, true, { receipt })
    const settled = await promise
    if (!settled?.hash) fail('apply 没有等待到回执')
  }

  // CAS：旧 baseHash 必须被拒
  const stale = await harness.run('save', {
    ref: REF,
    pageId: 'page-1',
    baseHash: attached.hash,
    bytes: harness.files.get(REF),
  }).then(
    () => null,
    (error) => String(error.message ?? error),
  )
  if (!stale?.includes('command.stale')) fail(`旧 baseHash 应被拒（当前 ${stale}）`)

  // 包校验：非 zip 字节
  const bad = await harness.run('save', {
    ref: REF,
    pageId: 'page-1',
    baseHash: (await harness.run('load', { ref: REF })).hash,
    bytes: new Uint8Array([1, 2, 3, 4]),
  }).then(
    () => null,
    (error) => String(error.message ?? error),
  )
  if (!bad?.includes('command.invalid')) fail(`坏包应被拒（当前 ${bad}）`)

  // 释放租约后写必须失败
  await harness.run('detach', { ref: REF, pageId: 'page-1' })
  const detached = await harness.run('apply', { ref: REF, commands: [{ op: 'addElement' }] }).then(
    () => null,
    (error) => String(error.message ?? error),
  )
  if (!detached?.includes('command.page-unavailable')) fail('detach 后 apply 应被拒')
} finally {
  harness.stop()
}

for (const message of warnings) console.warn(`warn  ${message}`)
if (problems.length > 0) {
  for (const message of problems) console.error(`error ${message}`)
  console.error(`\nverify failed: ${problems.length} 个问题`)
  process.exit(1)
}
console.log(
  `verify ok · ${pkg.name}@${pkg.version} · commands=${Object.keys(dynamic).length} · ` +
    `agent=${AGENT.length} · page=${PAGE.length}`,
)
