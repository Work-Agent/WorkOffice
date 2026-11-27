// 执行点契约：声明 / 目录 / 计划 / 字节 / 只读。
//
// 默认构建是**四工具**面（docx_guide / docx_read / docx_create / docx_apply）；
// `--legacy-tools` 构建会导出 35 条逐工具声明，这里按 dist/build.json 分支断言。
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { strFromU8, unzipSync, zipSync } from 'fflate'
import { buildDocx } from './fixtures.mjs'

const { tools } = await import('../dist/engine/tools.mjs')
const planMod = await import('../dist/engine/plan.mjs')
const execMod = await import('../dist/engine/exec.mjs')
const runtimeMod = await import('../dist/engine/runtime.mjs')
const readMod = await import('../dist/engine/read.mjs')

const BUILD = JSON.parse(readFileSync(new URL('../dist/build.json', import.meta.url), 'utf8'))
const COMPACT = BUILD.tools !== 'legacy'

const BUDGET = { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 }
const REF = 'verify://fixture.docx'

function ctxFor(bytes) {
  return { readBytes: async () => ({ bytes, version: 'v1' }), tempDir: '', importEpoch: 0, budget: BUDGET }
}

function xmlOf(bytes) {
  return strFromU8(
    unzipSync(bytes, { filter: (f) => f.name === 'word/document.xml' })['word/document.xml'],
  )
}

const guide = (params = {}) =>
  planMod.plan({ tool: 'docx_guide', params: { ref: REF, ...params }, ref: REF, bytes: buildDocx() }, ctxFor(buildDocx()))

test('工具声明满足宿主 §6.1 形状（四工具面）', () => {
  for (const decl of tools) {
    assert.match(decl.name, /^[a-z][a-z0-9._-]*$/i, decl.name)
    assert.ok(decl.label, decl.name)
    assert.ok(decl.description, decl.name)
    assert.ok(Array.isArray(decl.kinds) && decl.kinds.length > 0, decl.name)
    assert.equal(typeof decl.readOnly, 'boolean', decl.name)
    assert.ok(decl.parameters && typeof decl.parameters === 'object', decl.name)
  }
  const names = tools.map((t) => t.name)
  assert.equal(new Set(names).size, names.length, '工具名必须唯一')
  if (COMPACT) {
    assert.deepEqual(names, ['docx_guide', 'docx_read', 'docx_create', 'docx_apply'])
    assert.equal(tools.filter((t) => t.readOnly).length, 2)
    assert.equal(tools.find((t) => t.name === 'docx_apply').requiresBaseVersion, true)
    assert.equal(tools.find((t) => t.name === 'docx_create').requiresBaseVersion, false)
    // 声明体积：四工具面必须远小于逐工具面（这是这次收敛的目的）
    const size = tools.reduce(
      (sum, t) => sum + JSON.stringify({ n: t.name, d: t.description, p: t.parameters }).length,
      0,
    )
    assert.ok(size < 6 * 1024, `四工具声明应 < 6 KB，实际 ${size}`)
  } else {
    assert.ok(tools.length >= 35, `legacy 构建应有 35 条，实际 ${tools.length}`)
  }
})

test('docx_guide：三级发现（分组 / 组内 schema / 单条）', async () => {
  const top = await guide()
  assert.match(top.summary, /blocks/)
  assert.match(top.summary, /editor/)
  assert.match(top.summary, /actions/)
  assert.match(top.summary, /ops=\d+/)

  const blocks = await guide({ group: 'blocks' })
  assert.match(blocks.summary, /insertParagraph/)
  assert.match(blocks.summary, /schema:/)

  const editorGroup = await guide({ group: 'editor' })
  assert.match(editorGroup.summary, /setFont/)

  const actions = await guide({ group: 'actions' })
  assert.match(actions.summary, /insert_content/)

  const one = await guide({ op: 'setText' })
  assert.match(one.summary, /setText/)
  assert.match(one.summary, /example:/)

  const editorOne = await guide({ op: 'setFont' })
  assert.match(editorOne.summary, /signature:/)

  const actionOne = await guide({ op: 'insert_content' })
  assert.match(actionOne.summary, /parameters:/)

  const unknown = await guide({ op: 'nope_not_an_op' })
  assert.match(unknown.summary, /未知 op/)
})

test('docx_apply：blocks 组 → 计划 → 字节（未改块保真）', async () => {
  const bytes = buildDocx()
  const planned = await planMod.plan(
    {
      tool: 'docx_apply',
      params: { ref: REF, commands: [{ op: 'setText', index: 1, text: '改写后的段落' }] },
      ref: REF,
      baseVersion: 'v1',
      bytes,
    },
    ctxFor(bytes),
  )
  assert.equal(planned.plan.kind, 'save')

  const outcome = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan: planned.plan, bytes },
    ctxFor(bytes),
  )
  assert.ok(outcome.bytes instanceof Uint8Array)
  assert.equal(outcome.result.applied, 1)
  assert.match(xmlOf(outcome.bytes), /改写后的段落/)

  const before = unzipSync(bytes)
  const after = unzipSync(outcome.bytes)
  for (const name of Object.keys(before)) {
    if (name === 'word/document.xml') continue
    assert.deepEqual(Buffer.from(after[name]), Buffer.from(before[name]), name)
  }
})

test('docx_apply：块命令批（插入 + 改格式）', async () => {
  const bytes = buildDocx()
  const planned = await planMod.plan(
    {
      tool: 'docx_apply',
      params: {
        ref: REF,
        commands: [
          { op: 'insertParagraph', after: 0, text: '新插入的段落' },
          { op: 'setHeadingLevel', index: 0, level: 2 },
        ],
      },
      ref: REF,
      baseVersion: 'v1',
      bytes,
    },
    ctxFor(bytes),
  )
  assert.equal(planned.plan.kind, 'save')
  const outcome = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan: planned.plan, bytes },
    ctxFor(bytes),
  )
  const xml = xmlOf(outcome.bytes)
  assert.match(xml, /新插入的段落/)
  assert.match(xml, /Heading2/)
})

test('docx_apply：未知 op → 拒绝（带 index 与 code）', async () => {
  const bytes = buildDocx()
  const planned = await planMod.plan(
    {
      tool: 'docx_apply',
      params: { ref: REF, commands: [{ op: 'setText', index: 0, text: 'x' }, { op: 'nope' }] },
      ref: REF,
      baseVersion: 'v1',
      bytes,
    },
    ctxFor(bytes),
  )
  assert.equal(planned.plan.kind, 'reject')
  assert.equal(planned.plan.rejected[0].index, 1)
  assert.equal(planned.plan.rejected[0].code, 'command.not-found')
  const outcome = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan: planned.plan, bytes },
    ctxFor(bytes),
  )
  assert.equal(outcome.bytes, undefined, '被拒批不产字节')
})

test('docx_apply：editor / actions 组在无页面时 → page-unavailable', async () => {
  const bytes = buildDocx()
  for (const op of ['setFont', 'insert_content']) {
    await assert.rejects(
      () =>
        planMod.plan(
          { tool: 'docx_apply', params: { ref: REF, commands: [{ op, target: { blockIndexes: [0] } }] }, ref: REF, bytes },
          ctxFor(bytes),
        ),
      (error) => error.code === 'command.page-unavailable',
      op,
    )
  }
})

test('docx_read：七种 what 都返回文本', async () => {
  const bytes = buildDocx()
  for (const what of ['context', 'blocks', 'comments', 'notes', 'revisions', 'styles', 'audit']) {
    const outcome = await planMod.plan(
      {
        tool: 'docx_read',
        params: { ref: REF, what, startBlockIndex: 0, endBlockIndex: 2 },
        ref: REF,
        bytes,
      },
      ctxFor(bytes),
    )
    assert.equal(typeof outcome.summary, 'string', what)
    assert.ok(outcome.summary.length > 0, what)
  }
  await assert.rejects(
    () =>
      planMod.plan(
        { tool: 'docx_read', params: { ref: REF, what: 'nope' }, ref: REF, bytes },
        ctxFor(bytes),
      ),
    (error) => error.code === 'command.invalid',
  )
})

test('旧工具名仍可用（迁移期）：apply_ops 无页面 → page-unavailable', async () => {
  await assert.rejects(
    () =>
      planMod.plan(
        { tool: 'apply_ops', params: { ref: REF, ops: [] }, ref: REF, bytes: buildDocx() },
        ctxFor(buildDocx()),
      ),
    (error) => error.code === 'command.page-unavailable',
  )
})

test('bytes 计划：校验包结构后原样输出', async () => {
  const bytes = buildDocx()
  const ok = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan: { kind: 'bytes', bytes } },
    ctxFor(bytes),
  )
  assert.deepEqual(Buffer.from(ok.bytes), Buffer.from(bytes))
  assert.equal(ok.result.sync.kind, 'version')

  await assert.rejects(
    () =>
      execMod.exec(
        { ref: REF, baseVersion: 'v1', plan: { kind: 'bytes', bytes: new Uint8Array([1, 2, 3]) } },
        ctxFor(bytes),
      ),
    (error) => error.code === 'command.invalid',
  )
})

test('docx_create：Markdown → 新建 DOCX', async () => {
  const ref = 'verify://new-report.docx'
  const planned = await planMod.plan(
    {
      tool: 'docx_create',
      params: {
        content:
          '# 年度报告\n\n正文 **加粗** 和 *斜体*\n\n- A\n- B\n\n| 列1 | 列2 |\n| --- | --- |\n| a | b |',
      },
      ref,
      missing: true,
    },
    ctxFor(new Uint8Array()),
  )
  assert.equal(planned.plan.kind, 'new')

  const out = await execMod.exec(
    {
      ref,
      baseVersion: '',
      plan: planned.plan,
      bytes: new Uint8Array(),
      missing: true,
    },
    ctxFor(new Uint8Array()),
  )
  assert.ok(out.bytes && out.bytes.byteLength > 0)
  const xml = xmlOf(out.bytes)
  assert.match(xml, /年度报告/)
  assert.match(xml, /加粗/)
  assert.match(xml, /w:tbl/)
  assert.match(xml, /Heading1/)
  assert.equal(out.result.audit.source, 'create')

  await assert.rejects(
    () =>
      planMod.plan(
        { tool: 'docx_create', params: { content: '# x' }, ref, missing: false },
        ctxFor(buildDocx()),
      ),
    (error) => error.code === 'command.invalid',
  )
})

test('read 模块的 project / query ABI', async () => {
  const bytes = buildDocx()
  const projected = await readMod.project({ ref: REF, bytes }, ctxFor(bytes))
  assert.match(projected, /blocks=/)
  const queried = await readMod.query({ ref: REF, params: { kind: 'headings' }, bytes }, ctxFor(bytes))
  assert.ok(Array.isArray(queried.rows))
})

test('runtime 报告如实声明能力', async () => {
  const report = await runtimeMod.runtime(ctxFor(buildDocx()))
  assert.equal(report.write, 'plan')
  assert.equal(report.execution, 'isolated')
  assert.equal(report.features.undo, 'page')
  assert.equal(report.features.isolation, 'batch')
  assert.equal(report.features.tokens.inPage, 'docRev+epoch')
  assert.ok(report.limits.maxCommitBytes > 0)
})

test('未改条目的压缩流逐字节不变', async () => {
  const bytes = zipSync({
    '[Content_Types].xml': new TextEncoder().encode('<Types/>'),
    'word/document.xml': new TextEncoder().encode(
      '<w:document><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>',
    ),
    'word/media/image1.png': new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]),
  })
  const planned = await planMod.plan(
    {
      tool: 'docx_apply',
      params: { ref: REF, commands: [{ op: 'setText', index: 0, text: 'Z' }] },
      ref: REF,
      baseVersion: 'v1',
      bytes,
    },
    ctxFor(bytes),
  )
  const outcome = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan: planned.plan, bytes },
    ctxFor(bytes),
  )
  assert.match(xmlOf(outcome.bytes), />Z</)
  assert.match(xmlOf(outcome.bytes), />B</)
})
