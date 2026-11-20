#!/usr/bin/env node
// 装载前自检：清单 / 产物 / 工具声明 / 执行点 ABI / 一次往返。

import { readFile, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { strFromU8, unzipSync } from 'fflate'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const problems = []
const warnings = []
const fail = (message) => problems.push(message)
const warn = (message) => warnings.push(message)

const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))
const ws = pkg.workstudio ?? {}

if (ws.id !== 'pptx-pro') fail(`workstudio.id 必须是 pptx-pro（当前 ${ws.id}）`)
if (ws.apiVersion !== 2) fail(`apiVersion 必须是 2（当前 ${ws.apiVersion}）`)
if (ws.write !== 'plan') fail(`write 必须是 plan（当前 ${ws.write}）`)
if (ws.content !== 'pulled') fail(`content 必须是 pulled（当前 ${ws.content}）`)
if (ws.model !== 'resident') fail(`model 必须是 resident（当前 ${ws.model}）`)
if (!ws.engine?.exec) fail('write: plan 必须声明 engine.exec')
if (Boolean(ws.engine?.plan) !== Boolean(ws.engine?.exec))
  fail('engine.plan 与 engine.exec 必须同时出现')

const ura = pkg.exports?.['./ura']
if (!ura) fail('exports["./ura"] 缺失')
for (const [name, file] of Object.entries({ ura, ...(ws.engine ?? {}) })) {
  if (typeof file !== 'string') {
    fail(`${name}: 入口不是字符串`)
    continue
  }
  const info = await stat(join(root, file)).catch(() => null)
  if (!info?.isFile()) fail(`${name}: 文件不存在 ${file}`)
}
for (const file of ['dist/seed.pptx', 'dist/page/index.html']) {
  const info = await stat(join(root, file)).catch(() => null)
  if (!info?.isFile()) fail(`${file} 缺失`)
}
for (const skill of ws.skills ?? []) {
  if (!skill?.name || !skill?.dir) {
    fail('skills[] 缺少 name / dir')
    continue
  }
  const info = await stat(join(root, skill.dir)).catch(() => null)
  if (!info?.isDirectory()) fail(`skills: 目录不存在 ${skill.dir}`)
}

const load = (file) => import(pathToFileURL(join(root, file)).href)
const { tools } = await load(ws.engine.tools)
if (!Array.isArray(tools) || tools.length === 0) fail('engine.tools 没有导出 tools 数组')
const NAME = /^[a-z][a-z0-9._-]*$/i
const seen = new Set()
for (const decl of tools ?? []) {
  const at = decl?.name ?? '(匿名)'
  if (!decl?.name || !NAME.test(decl.name)) fail(`工具名非法：${at}`)
  if (!decl?.label) fail(`工具缺少 label：${at}`)
  if (!decl?.description) fail(`工具缺少 description：${at}`)
  if (!Array.isArray(decl?.kinds) || decl.kinds.length === 0) fail(`工具缺少 kinds：${at}`)
  if (!decl?.parameters || typeof decl.parameters !== 'object') fail(`工具缺少 parameters：${at}`)
  if (typeof decl?.readOnly !== 'boolean') fail(`工具缺少 readOnly：${at}`)
  if (seen.has(decl?.name)) fail(`工具名重复：${at}`)
  seen.add(decl?.name)
}
for (const name of ['pptx_guide', 'pptx_read', 'pptx_check', 'pptx_create', 'pptx_apply'])
  if (!seen.has(name)) fail(`工具面缺少 ${name}`)
if (seen.size !== 5) fail(`工具面应只有 5 条声明（实际 ${seen.size}）`)
if ((tools ?? []).filter((d) => d.readOnly).length !== 3) fail('工具面应恰好 3 个只读工具')
if ((tools ?? []).filter((d) => !d.readOnly).length !== 2) fail('工具面应恰好 2 个写工具')

const planMod = await load(ws.engine.plan)
const execMod = await load(ws.engine.exec)
const readMod = await load(ws.engine.read)
const runtimeMod = await load(ws.engine.runtime)
for (const [name, mod, exports] of [
  ['plan', planMod, ['plan']],
  ['exec', execMod, ['exec', 'commitModel', 'abortModel']],
  ['read', readMod, ['project', 'query']],
  ['runtime', runtimeMod, ['runtime']],
]) {
  for (const key of exports)
    if (typeof mod[key] !== 'function') fail(`执行点 ${name} 缺少导出 ${key}`)
}

const report = await runtimeMod.runtime({
  readBytes: async () => ({ bytes: new Uint8Array(), version: '' }),
  tempDir: '',
  importEpoch: 0,
  budget: { maxBytes: 1, deadlineMs: 1 },
})
if (report?.model !== 'resident') fail(`runtime.model 应为 resident（当前 ${report?.model}）`)
if (report?.features?.undo !== 'page') fail('runtime.features.undo 应为 page')
if (report?.features?.tokens?.inPage !== 'docRev')
  fail('runtime.features.tokens.inPage 应为 docRev')
if (report?.features?.tokens?.baselineSync !== 'version')
  fail('runtime.features.tokens.baselineSync 应为 version')

const seed = new Uint8Array(await readFile(join(root, 'dist/seed.pptx')))
const REF = 'verify://fixture.pptx'
function ctxFor(bytes) {
  return {
    readBytes: async () => ({ bytes, version: 'v1' }),
    tempDir: tmpdir(),
    importEpoch: 0,
    budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 },
  }
}
const ctx = ctxFor(seed)

const guide = await planMod.plan(
  { tool: 'pptx_guide', params: { ref: REF }, ref: REF, bytes: seed },
  ctx,
)
for (const group of ['text', 'element', 'insert', 'table', 'slide', 'deck', 'chart'])
  if (!guide?.summary?.includes(group)) fail(`pptx_guide 缺少分组 ${group}`)
if (!/ops=\d+/.test(guide?.summary ?? '')) fail('pptx_guide 没有返回 ops 数量')

const read = await planMod.plan(
  { tool: 'pptx_read', params: { ref: REF, what: 'context' }, ref: REF, bytes: seed },
  ctx,
)
if (!read?.summary?.includes('slides=')) fail('pptx_read 没有返回页概览')

const applied = await planMod.plan(
  {
    tool: 'pptx_apply',
    params: {
      ref: REF,
      commands: [
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: 1000000, y: 1000000, cx: 3000000, cy: 1000000 },
          paragraphs: [{ runs: [{ text: 'Hello pptx-pro' }] }],
        },
      ],
    },
    ref: REF,
    baseVersion: 'v1',
    bytes: seed,
  },
  ctx,
)
if (applied?.plan?.kind !== 'txn') fail(`pptx_apply 没有产出 txn 计划：${applied?.plan?.kind}`)
const produced = await execMod.exec(
  { ref: REF, baseVersion: 'v1', plan: applied.plan, bytes: seed },
  ctx,
)
if (!produced?.artifact?.tempPath) fail('exec 没有产出 artifact')
else {
  const bytes = new Uint8Array(await readFile(produced.artifact.tempPath))
  const xml = strFromU8(
    unzipSync(bytes, { filter: (f) => f.name === 'ppt/slides/slide1.xml' })['ppt/slides/slide1.xml'],
  )
  if (!xml.includes('Hello pptx-pro')) fail('写回结果缺少新文本')
  await execMod.commitModel(produced.modelToken)
}

const blocked = await planMod.plan(
  {
    tool: 'pptx_apply',
    params: { ref: REF, commands: [{ op: 'addPicture', target: { slide: 0 } }] },
    ref: REF,
    baseVersion: 'v1',
    bytes: seed,
  },
  ctx,
)
if (blocked?.plan?.kind !== 'reject') fail('无页面 addPicture 应产出 reject 计划')
else if (blocked.plan.rejected?.[0]?.code !== 'command.page-unavailable')
  fail(`无页面 addPicture 应回 command.page-unavailable（当前 ${blocked.plan.rejected?.[0]?.code}）`)

const created = await planMod.plan(
  {
    tool: 'pptx_create',
    params: {
      ref: 'verify://new.pptx',
      slides: [
        { title: '年度汇报', bullets: ['收入增长', '成本下降'], notes: '讲稿' },
        { title: '第二页', body: '正文' },
      ],
    },
    ref: 'verify://new.pptx',
    missing: true,
  },
  ctxFor(new Uint8Array()),
)
if (created?.plan?.kind !== 'new') fail(`pptx_create 没有产出 new 计划：${created?.plan?.kind}`)
else {
  const createdOut = await execMod.exec(
    {
      ref: 'verify://new.pptx',
      baseVersion: '',
      plan: created.plan,
      missing: true,
    },
    ctxFor(new Uint8Array()),
  )
  if (!createdOut?.artifact?.tempPath) fail('新建没有产出 artifact')
  else {
    const bytes = new Uint8Array(await readFile(createdOut.artifact.tempPath))
    const files = unzipSync(bytes)
    const slides = Object.keys(files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    if (slides.length < 2) fail(`新建页数不足：${slides.length}`)
    const xml = strFromU8(files['ppt/slides/slide1.xml'])
    if (!xml.includes('年度汇报')) fail('新建第一页缺少标题')
  }
}

// 设计稿路径：pages spec → build → pptx_check audit
const DESIGN = [
  {
    background: '#0E1A2B',
    elements: [
      {
        type: 'text',
        x: 80,
        y: 200,
        w: 900,
        h: 140,
        paragraphs: [
          { runs: [{ text: 'Verify deck', sizePt: 54, bold: true, color: '#FFFFFF' }] },
        ],
      },
      { type: 'shape', shape: 'rect', x: 80, y: 400, w: 300, h: 8, fill: '#FFB000' },
    ],
  },
]
const designPlan = await planMod.plan(
  { tool: 'pptx_create', params: { pages: DESIGN }, ref: 'verify://design.pptx', missing: true },
  ctxFor(new Uint8Array()),
)
if (designPlan?.plan?.kind !== 'build')
  fail(`pptx_create pages 应产出 build 计划（当前 ${designPlan?.plan?.kind}）`)
else {
  const designOut = await execMod.exec(
    { ref: 'verify://design.pptx', baseVersion: '', plan: designPlan.plan, missing: true },
    ctxFor(new Uint8Array()),
  )
  if (!designOut?.artifact?.tempPath) fail('设计稿新建没有产出 artifact')
  else {
    const bytes = new Uint8Array(await readFile(designOut.artifact.tempPath))
    const files = unzipSync(bytes)
    const xml = strFromU8(files['ppt/slides/slide1.xml'])
    if (!xml.includes('Verify deck')) fail('设计稿第一页缺少文本')
    if (!/<a:prstGeom prst="rect"/.test(xml)) fail('设计稿缺少形状元素')
    if (!xml.includes('<p:bg>')) fail('设计稿背景未被提升成原生背景')
    const audit = await planMod.plan(
      { tool: 'pptx_check', params: { audit: true }, ref: 'verify://design.pptx', bytes },
      ctxFor(bytes),
    )
    if (!/几何审计/.test(audit?.summary ?? '')) fail('pptx_check audit 没有返回几何审计')
  }
}

const outlineCheck = await planMod.plan(
  {
    tool: 'pptx_check',
    params: {
      outline: {
        core_hook: 'x',
        pages: [
          {
            title: 'A',
            type: 'content',
            layout: 'three_column_cards',
            brief: 'a'.repeat(60),
            image_queries: [],
          },
          {
            title: 'B',
            type: 'content',
            layout: 'three_column_cards',
            brief: 'b'.repeat(60),
            image_queries: [],
          },
        ],
      },
    },
    ref: 'verify://new.pptx',
    missing: true,
  },
  ctxFor(new Uint8Array()),
)
if (!/repeats the previous page/.test(outlineCheck?.summary ?? ''))
  fail('pptx_check 没有报连续同版式')

for (const message of warnings) console.warn(`warn  ${message}`)
if (problems.length > 0) {
  for (const message of problems) console.error(`error ${message}`)
  console.error(`\nverify failed: ${problems.length} 个问题`)
  process.exit(1)
}
console.log(
  `verify ok · ${pkg.name}@${pkg.version} · tools=${seen.size} · ` +
    `read=${(tools ?? []).filter((d) => d.readOnly).length} · ` +
    `write=${(tools ?? []).filter((d) => !d.readOnly).length}`,
)
