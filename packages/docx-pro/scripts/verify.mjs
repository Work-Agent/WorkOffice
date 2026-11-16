#!/usr/bin/env node
// 装载前自检：清单 / 声明位 / 产物 / 工具声明 / 执行点 ABI / 一次往返。
//
// 规则与宿主 `src/shared/plugin-contract.ts` 的 parseDeclaration 与
// toolDeclIssues 保持一致——这里不 import 宿主源码，只镜像判定，避免两份
// 定义漂移时静默通过。

import { readFile, stat } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { strFromU8, unzipSync, zipSync } from 'fflate'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const problems = []
const warnings = []

const fail = (message) => problems.push(message)
const warn = (message) => warnings.push(message)

const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))
const ws = pkg.workstudio ?? {}

// ─── 清单 ────────────────────────────────────────────────────────────
if (ws.id !== 'docx-pro') fail(`workstudio.id 必须是 docx-pro（当前 ${ws.id}）`)
if (ws.apiVersion !== 2) fail(`apiVersion 必须是 2（当前 ${ws.apiVersion}）`)
if (!['bytes', 'plan'].includes(ws.write)) fail(`write 非法：${ws.write}`)
if (ws.write === 'plan' && !ws.engine?.exec) fail('write: plan 必须声明 engine.exec')
if (ws.write === 'bytes' && ws.engine?.exec) fail('write: bytes 不得声明 engine.exec')
if (Boolean(ws.engine?.plan) !== Boolean(ws.engine?.exec))
  fail('engine.plan 与 engine.exec 必须同时出现')
if (ws.content === 'pulled' && !ws.page) fail("content: 'pulled' 必须声明 page")
if (ws.model === 'resident' && ws.write !== 'plan') fail("model: 'resident' 需要 write: 'plan'")

const ura = pkg.exports?.['./ura']
if (!ura) fail('exports["./ura"] 缺失')

for (const [name, file] of Object.entries({ ura, page: ws.page, ...(ws.engine ?? {}) })) {
  if (typeof file !== 'string') {
    fail(`${name}: 入口不是字符串`)
    continue
  }
  const info = await stat(join(root, file)).catch(() => null)
  if (!info?.isFile()) fail(`${name}: 文件不存在 ${file}`)
}
{
  const seed = join(root, 'dist/seed.docx')
  const info = await stat(seed).catch(() => null)
  if (!info?.isFile()) fail('dist/seed.docx 缺失（新建模板）')
}

for (const skill of ws.skills ?? []) {
  if (!skill?.name || !skill?.dir) {
    fail('skills[] 缺少 name / dir')
    continue
  }
  const info = await stat(join(root, skill.dir)).catch(() => null)
  if (!info?.isDirectory()) fail(`skills: 目录不存在 ${skill.dir}`)
}

// ─── 工具声明（宿主 toolDeclIssues 的镜像） ────────────────────────────
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
  if (!decl?.readOnly && !ws.write) fail(`写工具出现在只读插件里：${at}`)
}
const buildInfo = JSON.parse(
  await readFile(join(root, 'dist/build.json'), 'utf8').catch(() => '{}'),
)
const compact = buildInfo?.tools !== 'legacy'
if (compact) {
  for (const name of ['docx_guide', 'docx_read', 'docx_create', 'docx_apply'])
    if (!seen.has(name)) fail(`四工具面缺少 ${name}`)
  if (seen.size !== 4) fail(`四工具面应只有 4 条声明（实际 ${seen.size}）`)
  if ((tools ?? []).filter((d) => d.readOnly).length !== 2) fail('四工具面应恰好 2 个只读工具')
  if ((tools ?? []).filter((d) => !d.readOnly).length !== 2) fail('四工具面应恰好 2 个写工具')
} else {
  if (seen.size < 35) fail(`legacy 工具面应 ≥35 条（实际 ${seen.size}）`)
}

// ─── 执行点 ABI ──────────────────────────────────────────────────────
const planMod = await load(ws.engine.plan)
const execMod = await load(ws.engine.exec)
const readMod = await load(ws.engine.read)
const runtimeMod = await load(ws.engine.runtime)
for (const [name, mod, exports] of [
  ['plan', planMod, ['plan']],
  ['exec', execMod, ['exec']],
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
if (report?.write !== 'plan') fail(`runtime.write 应为 plan（当前 ${report?.write}）`)
if (report?.features?.undo !== 'page') fail('runtime.features.undo 应为 page')
if (report?.features?.tokens?.inPage !== 'docRev+epoch')
  fail('runtime.features.tokens.inPage 应为 docRev+epoch')

// ─── 一次真实往返：块命令 → 计划 → 字节 ───────────────────────────────
const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>标题</w:t></w:r></w:p>
<w:p><w:r><w:t>正文</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body></w:document>`
const fixture = zipSync({
  '[Content_Types].xml': new TextEncoder().encode('<Types/>'),
  'word/document.xml': new TextEncoder().encode(DOCUMENT),
})
const ctx = {
  readBytes: async () => ({ bytes: fixture, version: 'v1' }),
  tempDir: '',
  importEpoch: 0,
  budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 10_000 },
}
const planned = await planMod.plan(
  {
    tool: 'docx_apply',
    params: { ref: 'verify://fixture.docx', commands: [{ op: 'setText', index: 0, text: '新标题' }] },
    ref: 'verify://fixture.docx',
    baseVersion: 'v1',
    bytes: fixture,
  },
  ctx,
)
if (planned?.plan?.kind !== 'save') fail(`docx_apply 没有产出 save 计划：${planned?.plan?.kind}`)
const produced = await execMod.exec(
  { ref: 'verify://fixture.docx', baseVersion: 'v1', plan: planned.plan, bytes: fixture },
  ctx,
)
if (!(produced?.bytes instanceof Uint8Array)) fail('exec 没有产出字节')
else {
  const xml = strFromU8(unzipSync(produced.bytes, { filter: (f) => f.name === 'word/document.xml' })['word/document.xml'])
  if (!xml.includes('新标题')) fail('写回结果缺少新文本')
}

const guide = await planMod.plan(
  { tool: 'docx_guide', params: { ref: 'verify://fixture.docx' }, ref: 'verify://fixture.docx', bytes: fixture },
  ctx,
)
if (!guide?.summary?.includes('insertParagraph')) fail('docx_guide 没有返回块命令目录')
for (const group of ['blocks', 'editor', 'actions'])
  if (!guide?.summary?.includes(group)) fail(`docx_guide 缺少分组 ${group}`)

const read = await planMod.plan(
  { tool: 'docx_read', params: { ref: 'verify://fixture.docx', what: 'context' }, ref: 'verify://fixture.docx', bytes: fixture },
  ctx,
)
if (!read?.summary?.includes('blocks=')) fail('docx_read 没有返回块投影')

const blocked = await planMod
  .plan(
    {
      tool: 'docx_apply',
      params: { ref: 'verify://fixture.docx', commands: [{ op: 'setFont', target: { blockIndexes: [0] } }] },
      ref: 'verify://fixture.docx',
      bytes: fixture,
    },
    ctx,
  )
  .then(() => null)
  .catch((error) => error)
if (blocked?.code !== 'command.page-unavailable')
  fail(`无页面写应回 command.page-unavailable（当前 ${blocked?.code ?? '未报错'}）`)

// ─── 结果 ────────────────────────────────────────────────────────────
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
