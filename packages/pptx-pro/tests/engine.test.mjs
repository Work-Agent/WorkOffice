import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { strFromU8, unzipSync } from 'fflate'

const { tools } = await import('../dist/engine/tools.mjs')
const planMod = await import('../dist/engine/plan.mjs')
const execMod = await import('../dist/engine/exec.mjs')
const readMod = await import('../dist/engine/read.mjs')
const runtimeMod = await import('../dist/engine/runtime.mjs')

const SEED = new Uint8Array(await readFile(new URL('../dist/seed.pptx', import.meta.url)))
const BUDGET = { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 }
const REF = 'verify://engine.pptx'

function ctxFor(bytes, version = 'v1') {
  return {
    readBytes: async () => ({ bytes, version }),
    tempDir: tmpdir(),
    importEpoch: 0,
    budget: BUDGET,
  }
}

const ctx = ctxFor(SEED)

function slideXml(bytes, name = 'ppt/slides/slide1.xml') {
  return strFromU8(unzipSync(bytes, { filter: (f) => f.name === name })[name])
}

test('工具声明：4 个工具 / 2 读 2 写', () => {
  const names = tools.map((tool) => tool.name)
  assert.deepEqual(names, ['pptx_guide', 'pptx_read', 'pptx_create', 'pptx_apply'])
  assert.equal(tools.filter((tool) => tool.readOnly).length, 2)
  assert.equal(tools.filter((tool) => !tool.readOnly).length, 2)
  assert.equal(tools.find((tool) => tool.name === 'pptx_apply').requiresBaseVersion, true)
  assert.equal(tools.find((tool) => tool.name === 'pptx_create').requiresBaseVersion, false)
})

test('runtime：resident + engine undo + docRev/version', async () => {
  const report = await runtimeMod.runtime(ctx)
  assert.equal(report.model, 'resident')
  assert.equal(report.features.undo, 'page')
  assert.equal(report.features.tokens.inPage, 'docRev')
  assert.equal(report.features.tokens.baselineSync, 'version')
  assert.equal(report.limits.maxFileBytes, 64 * 1024 * 1024)
  assert.equal(report.limits.maxCommitBytes, 64 * 1024 * 1024)
})

test('guide：七组与 ops 数量', async () => {
  const top = await planMod.plan({ tool: 'pptx_guide', params: { ref: REF }, ref: REF, bytes: SEED }, ctx)
  for (const group of ['text', 'element', 'insert', 'table', 'slide', 'deck', 'chart'])
    assert.match(top.summary, new RegExp(group))
  assert.match(top.summary, /ops=\d+/)
  const chart = await planMod.plan(
    { tool: 'pptx_guide', params: { ref: REF, group: 'chart' }, ref: REF, bytes: SEED },
    ctx,
  )
  assert.match(chart.summary, /editChart/)
  const one = await planMod.plan(
    { tool: 'pptx_guide', params: { ref: REF, op: 'setText' }, ref: REF, bytes: SEED },
    ctx,
  )
  assert.match(one.summary, /signature:/)
  assert.match(one.summary, /example:/)
})

test('read：context / slide / audit', async () => {
  const context = await planMod.plan(
    { tool: 'pptx_read', params: { ref: REF, what: 'context' }, ref: REF, bytes: SEED },
    ctx,
  )
  assert.match(context.summary, /slides=1/)
  const slide = await planMod.plan(
    { tool: 'pptx_read', params: { ref: REF, what: 'slide', slide: 0 }, ref: REF, bytes: SEED },
    ctx,
  )
  assert.match(slide.summary, /slide 0 s_1/)
  const audit = await planMod.plan(
    { tool: 'pptx_read', params: { ref: REF, what: 'audit' }, ref: REF, bytes: SEED },
    ctx,
  )
  assert.match(audit.summary, /audit: slides=1/)
})

test('pptx_apply：txn → artifact → commitModel → 按 hash 重键', async () => {
  const planned = await planMod.plan(
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
      bytes: SEED,
    },
    ctx,
  )
  assert.equal(planned.plan.kind, 'txn')
  const outcome = await execMod.exec(
    { ref: REF, baseVersion: 'v1', plan: planned.plan, bytes: SEED },
    ctx,
  )
  assert.ok(outcome.artifact?.tempPath)
  assert.ok(outcome.modelToken)
  const bytes = new Uint8Array(await readFile(outcome.artifact.tempPath))
  assert.match(slideXml(bytes), /Hello pptx-pro/)
  await execMod.commitModel(outcome.modelToken)

  // 下一次请求带新 version + 新字节：应通过 lastArtifactSha256 认领缓存，
  // 而不是重新 openPptx；读取应看到刚才的文本。
  const reread = await planMod.plan(
    { tool: 'pptx_read', params: { ref: REF, what: 'slide', slide: 0 }, ref: REF, baseVersion: 'v2', bytes },
    ctxFor(bytes, 'v2'),
  )
  assert.match(reread.summary, /Hello pptx-pro/)
})

test('pptx_apply：abortModel 恢复模型', async () => {
  const ref = 'verify://abort.pptx'
  const planned = await planMod.plan(
    {
      tool: 'pptx_apply',
      params: {
        ref,
        commands: [
          {
            op: 'addElement',
            target: { slide: 0 },
            kind: 'textbox',
            offset: { x: 100000, y: 100000, cx: 1000000, cy: 500000 },
            paragraphs: [{ runs: [{ text: 'ghost' }] }],
          },
        ],
      },
      ref,
      baseVersion: 'v1',
      bytes: SEED,
    },
    ctxFor(SEED),
  )
  const outcome = await execMod.exec(
    { ref, baseVersion: 'v1', plan: planned.plan, bytes: SEED },
    ctxFor(SEED),
  )
  await execMod.abortModel(outcome.modelToken)
  const reread = await planMod.plan(
    { tool: 'pptx_read', params: { ref, what: 'slide', slide: 0 }, ref, baseVersion: 'v1', bytes: SEED },
    ctxFor(SEED),
  )
  assert.doesNotMatch(reread.summary, /ghost/)
})

test('pptx_apply：dryRun 不产 artifact', async () => {
  const planned = await planMod.plan(
    {
      tool: 'pptx_apply',
      params: {
        ref: 'verify://dry.pptx',
        commands: [{ op: 'setText', target: { slide: 0, el: 'e_1' }, paragraphs: [] }],
        dryRun: true,
      },
      ref: 'verify://dry.pptx',
      baseVersion: 'v1',
      bytes: SEED,
    },
    ctxFor(SEED),
  )
  // 空 paragraphs 可能被上游拒绝；无论 plan 是 txn 还是 reject，都不应该有 artifact。
  if (planned.plan.kind === 'txn') {
    const outcome = await execMod.exec(
      { ref: 'verify://dry.pptx', baseVersion: 'v1', plan: planned.plan, bytes: SEED, dryRun: true },
      ctxFor(SEED),
    )
    assert.equal(outcome.artifact, undefined)
  }
})

test('页面 op：无页面回 command.page-unavailable', async () => {
  const blocked = await planMod.plan(
    {
      tool: 'pptx_apply',
      params: { ref: REF, commands: [{ op: 'addPicture', target: { slide: 0 } }] },
      ref: REF,
      baseVersion: 'v1',
      bytes: SEED,
    },
    ctx,
  )
  assert.equal(blocked.plan.kind, 'reject')
  assert.equal(blocked.plan.rejected[0].code, 'command.page-unavailable')
})

test('未知 op：reject command.not-found', async () => {
  const blocked = await planMod.plan(
    {
      tool: 'pptx_apply',
      params: { ref: REF, commands: [{ op: 'nope_not_an_op' }] },
      ref: REF,
      baseVersion: 'v1',
      bytes: SEED,
    },
    ctx,
  )
  assert.equal(blocked.plan.kind, 'reject')
  assert.equal(blocked.plan.rejected[0].code, 'command.not-found')
})

test('pptx_create：setup/content 两阶段 + 多页', async () => {
  const planned = await planMod.plan(
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
  assert.equal(planned.plan.kind, 'new')
  const outcome = await execMod.exec(
    { ref: 'verify://new.pptx', baseVersion: '', plan: planned.plan, missing: true },
    ctxFor(new Uint8Array()),
  )
  assert.ok(outcome.artifact?.tempPath)
  const bytes = new Uint8Array(await readFile(outcome.artifact.tempPath))
  const files = unzipSync(bytes)
  const slides = Object.keys(files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
  assert.ok(slides.length >= 2)
  assert.match(strFromU8(files['ppt/slides/slide1.xml']), /年度汇报/)
})

test('txn.batches：先建页、再编辑新页（页面保存形态）', async () => {
  const ref = 'verify://batches.pptx'
  const outcome = await execMod.exec(
    {
      ref,
      baseVersion: 'v1',
      bytes: SEED,
      plan: {
        kind: 'txn',
        txn: {
          batches: [
            [{ op: 'addBlankSlide', target: { slide: 0 } }],
            [
              {
                op: 'addElement',
                target: { slide: 1 },
                kind: 'textbox',
                offset: { x: 100000, y: 100000, cx: 1000000, cy: 500000 },
                paragraphs: [{ runs: [{ text: 'second slide' }] }],
              },
            ],
          ],
        },
      },
    },
    ctxFor(SEED),
  )
  assert.ok(outcome.artifact?.tempPath)
  const bytes = new Uint8Array(await readFile(outcome.artifact.tempPath))
  const files = unzipSync(bytes)
  const slides = Object.keys(files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
  assert.ok(slides.length >= 2)
  assert.match(strFromU8(files['ppt/slides/slide2.xml']), /second slide/)
  await execMod.abortModel(outcome.modelToken)
})

test('engine undo / redo：快照恢复 + artifact', async () => {
  const ref = 'verify://undo.pptx'
  const planned = await planMod.plan(
    {
      tool: 'pptx_apply',
      params: {
        ref,
        commands: [
          {
            op: 'addElement',
            target: { slide: 0 },
            kind: 'textbox',
            offset: { x: 100000, y: 100000, cx: 1000000, cy: 500000 },
            paragraphs: [{ runs: [{ text: 'to be undone' }] }],
          },
        ],
      },
      ref,
      baseVersion: 'v1',
      bytes: SEED,
    },
    ctxFor(SEED),
  )
  const applied = await execMod.exec(
    { ref, baseVersion: 'v1', plan: planned.plan, bytes: SEED },
    ctxFor(SEED),
  )
  await execMod.commitModel(applied.modelToken)
  const undone = await execMod.exec(
    { ref, baseVersion: 'v1', command: 'undo', plan: { kind: 'noop' }, bytes: SEED },
    ctxFor(SEED),
  )
  assert.ok(undone.artifact?.tempPath)
  const bytes = new Uint8Array(await readFile(undone.artifact.tempPath))
  assert.doesNotMatch(slideXml(bytes), /to be undone/)
  await execMod.commitModel(undone.modelToken)
})

test('pageSave：写入页面已生成的 slide XML，不重放 ops', async () => {
  const ref = 'verify://pagesave.pptx'
  const planned = await planMod.plan(
    {
      tool: 'pptx_apply',
      params: {
        ref,
        commands: [
          {
            op: 'addElement',
            target: { slide: 0 },
            kind: 'textbox',
            offset: { x: 100000, y: 100000, cx: 1000000, cy: 500000 },
            paragraphs: [{ runs: [{ text: 'page xml' }] }],
          },
        ],
      },
      ref,
      baseVersion: 'v1',
      bytes: SEED,
    },
    ctxFor(SEED),
  )
  const applied = await execMod.exec(
    { ref, baseVersion: 'v1', plan: planned.plan, bytes: SEED },
    ctxFor(SEED),
  )
  const generated = new Uint8Array(await readFile(applied.artifact.tempPath))
  const xml = slideXml(generated)

  const saved = await execMod.exec(
    {
      ref,
      baseVersion: 'v1',
      bytes: SEED,
      plan: {
        kind: 'pageSave',
        slides: [{ path: 'ppt/slides/slide1.xml', xml }],
        parts: [],
        deletedParts: [],
        changedCount: 1,
      },
    },
    ctxFor(SEED),
  )
  assert.ok(saved.artifact?.tempPath)
  assert.match(slideXml(new Uint8Array(await readFile(saved.artifact.tempPath))), /page xml/)
})
