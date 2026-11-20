import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
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

test('工具声明：5 个工具 / 3 读 2 写', () => {
  const names = tools.map((tool) => tool.name)
  assert.deepEqual(names, ['pptx_guide', 'pptx_read', 'pptx_check', 'pptx_create', 'pptx_apply'])
  assert.equal(tools.filter((tool) => tool.readOnly).length, 3)
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

test('pptx_create：3 页 + 显式 16:9（回归：批内坐标越界 / 尺寸非幂等）', async () => {
  const spec = {
    size: '16:9',
    slides: [{ title: '一' }, { title: '二' }, { title: '三' }],
  }
  const planned = await planMod.plan(
    { tool: 'pptx_create', params: spec, ref: 'verify://three.pptx', missing: true },
    ctxFor(new Uint8Array()),
  )
  assert.equal(planned.plan.kind, 'new')
  const outcome = await execMod.exec(
    { ref: 'verify://three.pptx', baseVersion: '', plan: planned.plan, missing: true },
    ctxFor(new Uint8Array()),
  )
  assert.ok(outcome.artifact?.tempPath)
  const bytes = new Uint8Array(await readFile(outcome.artifact.tempPath))
  const files = unzipSync(bytes)
  const slides = Object.keys(files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
  assert.equal(slides.length, 3)
  // 页序必须与 spec 一致（空白页统一插在第 0 页之后）
  const outline = await planMod.plan(
    { tool: 'pptx_read', params: { what: 'context' }, ref: 'verify://three.pptx', bytes },
    ctxFor(bytes),
  )
  assert.match(outline.summary, /\[0\] s_\d+ "一"/)
  assert.match(outline.summary, /\[1\] s_\d+ "二"/)
  assert.match(outline.summary, /\[2\] s_\d+ "三"/)
})

test('read：元素 box 是真实坐标（回归：transform.offset）', async () => {
  const planned = await planMod.plan(
    {
      tool: 'pptx_create',
      params: { slides: [{ title: '坐标' }, { title: '第二页' }] },
      ref: 'verify://box.pptx',
      missing: true,
    },
    ctxFor(new Uint8Array()),
  )
  const outcome = await execMod.exec(
    { ref: 'verify://box.pptx', baseVersion: '', plan: planned.plan, missing: true },
    ctxFor(new Uint8Array()),
  )
  const bytes = new Uint8Array(await readFile(outcome.artifact.tempPath))
  const slide = await planMod.plan(
    { tool: 'pptx_read', params: { what: 'slide', slide: 1 }, ref: 'verify://box.pptx', bytes },
    ctxFor(bytes),
  )
  // 标题框 x=8%、y=6%、cx=84%、cy=18%（EMU → in）
  assert.match(slide.summary, /box=1\.07,0\.45,11\.2,1\.35in/)
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

// ─── 设计稿路径（pages spec） ─────────────────────────────────────────

const DESIGN_PAGES = [
  {
    background: '#0E1A2B',
    elements: [
      {
        type: 'text',
        x: 80,
        y: 200,
        w: 900,
        h: 140,
        valign: 'middle',
        paragraphs: [
          {
            align: 'left',
            runs: [{ text: 'The 40% squeeze', sizePt: 54, bold: true, color: '#FFFFFF' }],
          },
        ],
      },
      { type: 'shape', shape: 'rect', x: 80, y: 400, w: 300, h: 8, fill: '#FFB000' },
    ],
  },
  {
    background: '#FFFFFF',
    elements: [
      {
        type: 'text',
        x: 80,
        y: 60,
        w: 900,
        h: 80,
        paragraphs: [
          { runs: [{ text: 'Where the margin went', sizePt: 36, bold: true, color: '#0E1A2B' }] },
        ],
      },
      {
        type: 'text',
        x: 80,
        y: 180,
        w: 420,
        h: 30,
        paragraphs: [
          {
            runs: [
              {
                text: 'Fuel +31% / wages +18% / spot rates -22% — 2024 vs 2022, three sources',
                sizePt: 16,
                color: '#333333',
              },
            ],
          },
        ],
      },
    ],
  },
]

async function createDesignDeck(pages, ref = 'verify://design.pptx') {
  const planned = await planMod.plan(
    { tool: 'pptx_create', params: { pages }, ref, missing: true },
    ctxFor(new Uint8Array()),
  )
  assert.equal(planned.plan.kind, 'build')
  const outcome = await execMod.exec(
    { ref, baseVersion: '', plan: planned.plan, missing: true },
    ctxFor(new Uint8Array()),
  )
  const bytes = new Uint8Array(await readFile(outcome.artifact.tempPath))
  return { planned, outcome, bytes }
}

test('pptx_create pages：设计稿 → 形状 + 文本 + 背景提升', async () => {
  const { planned, bytes } = await createDesignDeck(DESIGN_PAGES)
  assert.match(planned.summary, /pages=2/)
  const files = unzipSync(bytes)
  const slide1 = strFromU8(files['ppt/slides/slide1.xml'])
  const slide2 = strFromU8(files['ppt/slides/slide2.xml'])
  assert.match(slide1, /The 40% squeeze/)
  assert.match(slide1, /<a:prstGeom prst="rect"/)
  assert.match(slide1, /<p:bg>/) // 背景矩形被提升成原生背景
  assert.match(slide2, /Where the margin went/)
  const outline = await planMod.plan(
    { tool: 'pptx_read', params: { what: 'context' }, ref: 'x', bytes },
    ctxFor(bytes),
  )
  assert.match(outline.summary, /slides=2/)
})

test('pptx_create pages：文本框按真实字宽长高（模型估小也溢不出来）', async () => {
  const { bytes } = await createDesignDeck(DESIGN_PAGES, 'verify://grow.pptx')
  const slide = await planMod.plan(
    { tool: 'pptx_read', params: { what: 'slide', slide: 1 }, ref: 'x', bytes },
    ctxFor(bytes),
  )
  // 原 spec 的正文框是 30px（=0.3125in），一行 16pt 的长句在里面放不下 → 必须被长高
  const boxes = [...slide.summary.matchAll(/box=([\d.]+),([\d.]+),([\d.]+),([\d.]+)in/g)]
  const short = Math.min(...boxes.map((m) => Number(m[4])))
  assert.ok(short > 0.3, `最短的框应仍是最初的标题框高度，实际 ${short}`)
  assert.ok(
    boxes.some((m) => Number(m[4]) > 0.3125),
    `应有框被长高到 0.3125in 以上：${slide.summary}`,
  )
})

test('pptx_create pages：越界元素被 clamp / 空元素被丢弃并回注 warnings', async () => {
  const pages = [
    {
      elements: [
        { type: 'text', x: 1200, y: 600, w: 400, h: 400, paragraphs: [{ runs: [{ text: 'edge' }] }] },
        { type: 'text', x: 10, y: 10, w: 100, h: 40, paragraphs: [] },
      ],
    },
  ]
  const planned = await planMod.plan(
    { tool: 'pptx_create', params: { pages }, ref: 'verify://clamp.pptx', missing: true },
    ctxFor(new Uint8Array()),
  )
  assert.equal(planned.plan.kind, 'build')
  assert.match(planned.summary, /警告 — element 1/)
  const outcome = await execMod.exec(
    { ref: 'verify://clamp.pptx', baseVersion: '', plan: planned.plan, missing: true },
    ctxFor(new Uint8Array()),
  )
  const bytes = new Uint8Array(await readFile(outcome.artifact.tempPath))
  const slide = await planMod.plan(
    { tool: 'pptx_read', params: { what: 'slide', slide: 0 }, ref: 'x', bytes },
    ctxFor(bytes),
  )
  // 越界的 1200,600,400,400 被 clamp 进 1280×720：x+w 不超过 13.33in，y+h 不超过 7.5in
  const boxes = [...slide.summary.matchAll(/box=([\d.]+),([\d.]+),([\d.]+),([\d.]+)in/g)]
  for (const m of boxes) {
    assert.ok(Number(m[1]) + Number(m[3]) <= 13.34, `x 越界未 clamp：${m[0]}`)
    assert.ok(Number(m[2]) + Number(m[4]) <= 7.51, `y 越界未 clamp：${m[0]}`)
  }
})

test('pptx_check：outline 校验（未知版式 / 连续重复 / 缺 core_hook）', async () => {
  const outline = {
    core_hook: '',
    pages: [
      { title: 'A', type: 'content', layout: 'three_column_cards', brief: 'x'.repeat(60), image_queries: [] },
      { title: 'B', type: 'content', layout: 'three_column_cards', brief: 'y'.repeat(60), image_queries: [] },
      { title: 'C', type: 'content', layout: 'no_such_layout', brief: 'z'.repeat(60), image_queries: [] },
    ],
  }
  const result = await planMod.plan(
    { tool: 'pptx_check', params: { outline }, ref: 'verify://new.pptx', missing: true },
    ctxFor(new Uint8Array()),
  )
  assert.match(result.summary, /error deck: "core_hook"/)
  assert.match(result.summary, /https?|repeats the previous page|is not a content variant/)
})

test('pptx_check：合规 outline 通过', async () => {
  const outline = {
    core_hook: 'Margins fell 40% while volume grew',
    pages: [
      { title: 'Cover', type: 'cover', layout: 'cover_typography_hero', brief: 'a'.repeat(60), image_queries: [] },
      { title: 'One', type: 'content', layout: 'three_column_cards', brief: 'b'.repeat(60), image_queries: [] },
      { title: 'Two', type: 'data', layout: 'kpi_cards_row', brief: 'c'.repeat(60), image_queries: [] },
      { title: 'End', type: 'closing', layout: 'closing_thank_you', brief: 'd'.repeat(60), image_queries: [] },
    ],
  }
  const result = await planMod.plan(
    { tool: 'pptx_check', params: { outline }, ref: 'verify://new.pptx', missing: true },
    ctxFor(new Uint8Array()),
  )
  assert.match(result.summary, /outline 检查通过/)
})

test('pptx_check：pages 对账 outline + style 调色板', async () => {
  const outline = {
    core_hook: 'x',
    pages: [
      { title: 'The 40% squeeze', type: 'cover', layout: 'cover_typography_hero', brief: 'a'.repeat(60), image_queries: [] },
      { title: 'Where the margin went', type: 'content', layout: 'three_column_cards', brief: 'b'.repeat(60), image_queries: [] },
    ],
  }
  const style = '# Style\n- background: #0E1A2B\n- ink: #FFFFFF\n- accent: #FFB000\n'
  const result = await planMod.plan(
    {
      tool: 'pptx_check',
      params: { pages: DESIGN_PAGES, outline, style },
      ref: 'verify://new.pptx',
      missing: true,
    },
    ctxFor(new Uint8Array()),
  )
  assert.match(result.summary, /pages 检查：2 页/)
  // 第二页的 #FFFFFF 背景与 #333333 正文字色不在 style.md 里（纯白放行，深灰报出）
  assert.match(result.summary, /#333333/)
  // 标题在页面文本里 → 不该报 "does not appear"
  assert.doesNotMatch(result.summary, /does not appear/)
})

test('pptx_check audit：越界元素带可套用的 setTransform', async () => {
  const { bytes } = await createDesignDeck(DESIGN_PAGES, 'verify://audit.pptx')
  const applied = await planMod.plan(
    {
      tool: 'pptx_apply',
      params: {
        commands: [
          {
            op: 'addElement',
            target: { slide: 0 },
            kind: 'rect',
            offset: { x: 12000000, y: 100000, cx: 2000000, cy: 1000000 },
            fill: '#FF0000',
          },
        ],
      },
      ref: 'verify://audit.pptx',
      baseVersion: 'v1',
      bytes,
    },
    ctxFor(bytes),
  )
  assert.equal(applied.plan.kind, 'txn')
  const out = await execMod.exec(
    { ref: 'verify://audit.pptx', baseVersion: 'v1', plan: applied.plan, bytes },
    ctxFor(bytes),
  )
  const next = new Uint8Array(await readFile(out.artifact.tempPath))
  const audit = await planMod.plan(
    { tool: 'pptx_check', params: { audit: true }, ref: 'verify://audit.pptx', bytes: next },
    ctxFor(next),
  )
  assert.match(audit.summary, /几何审计/)
  assert.match(audit.summary, /out_of_bounds/)
  assert.match(audit.summary, /"op":"setTransform"/)
})

test('pptx_check skeleton：抽出标题框骨架', async () => {
  const pages = [
    { background: '#0E1A2B', elements: [{ type: 'text', x: 80, y: 120, w: 900, h: 120, paragraphs: [{ runs: [{ text: 'Cover', sizePt: 54, color: '#FFFFFF' }] }] }] },
    { background: '#FFFFFF', elements: [{ type: 'text', x: 80, y: 60, w: 900, h: 80, paragraphs: [{ runs: [{ text: 'Body', sizePt: 36, color: '#0E1A2B' }] }] }] },
    { background: '#FFFFFF', elements: [{ type: 'text', x: 80, y: 260, w: 900, h: 80, paragraphs: [{ runs: [{ text: 'Thanks', sizePt: 36, color: '#0E1A2B' }] }] }] },
  ]
  const { bytes } = await createDesignDeck(pages, 'verify://skeleton.pptx')
  const result = await planMod.plan(
    { tool: 'pptx_check', params: { skeleton: true }, ref: 'verify://skeleton.pptx', bytes },
    ctxFor(bytes),
  )
  assert.match(result.summary, /版面骨架/)
  assert.match(result.summary, /title box/)
})

test('pptx_check audit：文字溢出 → 报溢出并给完整 setTransform', async () => {
  const pages = [
    {
      elements: [
        {
          type: 'text',
          x: 80,
          y: 80,
          w: 400,
          h: 60,
          paragraphs: [{ runs: [{ text: 'Tiny box', sizePt: 12 }] }],
        },
      ],
    },
  ]
  const { bytes } = await createDesignDeck(pages, 'verify://overflow.pptx')
  const slide = await planMod.plan(
    { tool: 'pptx_read', params: { what: 'slide', slide: 0 }, ref: 'x', bytes },
    ctxFor(bytes),
  )
  const el = /\[0\] (e_\w+)/.exec(slide.summary)?.[1]
  assert.ok(el, slide.summary)
  const applied = await planMod.plan(
    {
      tool: 'pptx_apply',
      params: {
        commands: [
          {
            op: 'setText',
            target: { slide: 0, el },
            paragraphs: Array.from({ length: 8 }, () => ({
              runs: [{ text: '这是一段很长的正文，用来撑爆这个只有一个标题高度的文本框。', fontSize: 28 }],
            })),
          },
        ],
      },
      ref: 'verify://overflow.pptx',
      baseVersion: 'v1',
      bytes,
    },
    ctxFor(bytes),
  )
  const out = await execMod.exec(
    { ref: 'verify://overflow.pptx', baseVersion: 'v1', plan: applied.plan, bytes },
    ctxFor(bytes),
  )
  const next = new Uint8Array(await readFile(out.artifact.tempPath))
  const audit = await planMod.plan(
    { tool: 'pptx_check', params: { audit: true }, ref: 'verify://overflow.pptx', bytes: next },
    ctxFor(next),
  )
  assert.match(audit.summary, /text_overflow/)
  assert.match(audit.summary, /"op":"setTransform"/)
  assert.match(audit.summary, /"box":\{"x":\d+,"y":\d+,"cx":\d+,"cy":\d+\}/)
})

// ─── 高层版式（layout + 内容）+ 风格预设 + 本地图片 ─────────────────────

// 1×1 PNG（透明），用来验证图片真的被写进 media/
const PNG_1PX = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  'base64',
)

const HIGH_LEVEL_PAGES = [
  { type: 'cover', title: 'The 40% squeeze', subtitle: 'Regional logistics review · 2024', kicker: 'ACME' },
  {
    type: 'content',
    layout: 'three_column_cards',
    title: 'Where the margin went',
    items: [
      { value: '+31%', title: '燃油', note: '2024 vs 2022' },
      { value: '+18%', title: '人力', note: '2024 vs 2022' },
      { value: '-22%', title: '即期运价', note: '2024 vs 2022' },
    ],
    footnote: '来源：公司年报 2024',
  },
  { type: 'closing', title: '谢谢', subtitle: 'questions@acme.com' },
]

test('pptx_create：高层版式（layout + 内容）→ 真的摆出卡片', async () => {
  const planned = await planMod.plan(
    { tool: 'pptx_create', params: { pages: HIGH_LEVEL_PAGES, style: 'consulting-navy' }, ref: 'verify://hl.pptx', missing: true },
    ctxFor(new Uint8Array()),
  )
  assert.equal(planned.plan.kind, 'build')
  const outcome = await execMod.exec(
    { ref: 'verify://hl.pptx', baseVersion: '', plan: planned.plan, missing: true },
    ctxFor(new Uint8Array()),
  )
  const bytes = new Uint8Array(await readFile(outcome.artifact.tempPath))
  const files = unzipSync(bytes)
  const slide1 = strFromU8(files['ppt/slides/slide1.xml'])
  // 封面：深色底 + 大标题
  assert.match(slide1, /The 40% squeeze/)
  assert.match(slide1, /<p:bg>/) // 背景提升
  assert.match(slide1, /0E1A2B/i)
  const slide2 = strFromU8(files['ppt/slides/slide2.xml'])
  // 三张卡片 = 至少 4 个 roundRect（3 卡 + 顶部装饰线）+ 三个数字
  const rounds = (slide2.match(/prst="roundRect"/g) ?? []).length
  assert.ok(rounds >= 3, `应至少 3 个圆角卡片，实际 ${rounds}`)
  assert.match(slide2, /\+31%/)
  assert.match(slide2, /-22%/)
  // 预设的强调色 / 卡片底色真实落盘
  assert.match(slide2, /FFB000/i)
  assert.match(slide2, /F4F6F8/i)
  // 元素数：标题 + 线 + 3×(卡+数字+标题+说明) + 来源
  const slide = await planMod.plan(
    { tool: 'pptx_read', params: { what: 'slide', slide: 1 }, ref: 'x', bytes },
    ctxFor(bytes),
  )
  assert.match(slide.summary, /elements:/)
  const count = Number(/elements:\s*([\s\S]*)$/.exec(slide.summary)?.[1]?.match(/^\s*\[\d+\]/gm)?.length ?? 0)
  assert.ok(count >= 12, `三种卡片页应有 12 个以上元素，实际 ${count}`)
})

test('pptx_create：换预设只换皮，版式不变', async () => {
  const build = async (style) => {
    const planned = await planMod.plan(
      { tool: 'pptx_create', params: { pages: HIGH_LEVEL_PAGES, style }, ref: `verify://${style}.pptx`, missing: true },
      ctxFor(new Uint8Array()),
    )
    const outcome = await execMod.exec(
      { ref: `verify://${style}.pptx`, baseVersion: '', plan: planned.plan, missing: true },
      ctxFor(new Uint8Array()),
    )
    return new Uint8Array(await readFile(outcome.artifact.tempPath))
  }
  const navy = unzipSync(await build('consulting-navy'))
  const tech = unzipSync(await build('tech-dark'))
  const navy2 = strFromU8(navy['ppt/slides/slide2.xml'])
  const tech2 = strFromU8(tech['ppt/slides/slide2.xml'])
  assert.match(navy2, /FFB000/i)
  assert.match(tech2, /3DDC97/i)
  // 版式几何一致：卡片数相同
  const rounds = (xml) => (xml.match(/prst="roundRect"/g) ?? []).length
  assert.equal(rounds(navy2), rounds(tech2))
})

test('pptx_create：不写 layout 时按内容自动挑版式', async () => {
  const pages = [
    { title: 'Cover' },
    { title: 'Three things', items: [{ title: 'a' }, { title: 'b' }, { title: 'c' }] },
    { title: 'One number', items: [{ value: '92%', title: '命中率' }] },
    { title: 'Bye' },
  ]
  const planned = await planMod.plan(
    { tool: 'pptx_create', params: { pages }, ref: 'verify://auto.pptx', missing: true },
    ctxFor(new Uint8Array()),
  )
  assert.equal(planned.plan.kind, 'build')
  // 首页 cover / 末页 closing / 三栏 / 大数字都能落地
  const outcome = await execMod.exec(
    { ref: 'verify://auto.pptx', baseVersion: '', plan: planned.plan, missing: true },
    ctxFor(new Uint8Array()),
  )
  const files = unzipSync(new Uint8Array(await readFile(outcome.artifact.tempPath)))
  const slide2 = strFromU8(files['ppt/slides/slide2.xml'])
  const slide3 = strFromU8(files['ppt/slides/slide3.xml'])
  assert.ok((slide2.match(/prst="roundRect"/g) ?? []).length >= 3)
  assert.match(slide3, /92%/)
})

test('pptx_create：未知版式退回自动版式并回一条提示', async () => {
  const planned = await planMod.plan(
    {
      tool: 'pptx_create',
      params: { pages: [{ title: 'Mystery', layout: 'no_such_layout', body: ['a'] }] },
      ref: 'verify://unknown.pptx',
      missing: true,
    },
    ctxFor(new Uint8Array()),
  )
  assert.equal(planned.plan.kind, 'build')
  assert.match(planned.summary, /未知版式/)
})

test('pptx_check styles：列出风格预设', async () => {
  const result = await planMod.plan(
    { tool: 'pptx_check', params: { styles: true }, ref: 'verify://new.pptx', missing: true },
    ctxFor(new Uint8Array()),
  )
  for (const id of ['consulting-navy', 'tech-dark', 'editorial-warm', 'minimal-mono'])
    assert.match(result.summary, new RegExp(id))
})

test('pptx_check pages：高层页面按预设调色板检查（不该误报）', async () => {
  const outline = {
    core_hook: 'Margins fell 40%',
    pages: HIGH_LEVEL_PAGES.map((p) => ({
      title: p.title,
      type: p.type,
      layout: p.layout ?? (p.type === 'cover' ? 'cover_typography_hero' : 'closing_thank_you'),
      brief: 'a'.repeat(60),
      image_queries: [],
    })),
  }
  const result = await planMod.plan(
    {
      tool: 'pptx_check',
      params: { pages: HIGH_LEVEL_PAGES, outline, style: 'consulting-navy' },
      ref: 'verify://new.pptx',
      missing: true,
    },
    ctxFor(new Uint8Array()),
  )
  assert.match(result.summary, /pages 检查：3 页/)
  assert.match(result.summary, /风格 咨询深蓝/)
  assert.doesNotMatch(result.summary, /不在预设的调色板里/)
})

test('pptx_create：本地图片（file:// 与相对 deck 目录）真的进 media/', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'pptx-pro-img-'))
  const deckPath = join(dir, 'deck.pptx')
  await writeFile(join(dir, 'hero.png'), PNG_1PX)
  await writeFile(join(dir, 'not-an-image.txt'), 'this is not a png')
  const ref = pathToFileURL(deckPath).href
  const pages = [
    {
      type: 'cover',
      layout: 'cover_split_image',
      title: 'With a local photo',
      image: './hero.png',
    },
  ]
  const planned = await planMod.plan(
    { tool: 'pptx_create', params: { pages }, ref, missing: true },
    ctxFor(new Uint8Array()),
  )
  assert.equal(planned.plan.kind, 'build')
  const outcome = await execMod.exec(
    { ref, baseVersion: '', plan: planned.plan, missing: true },
    ctxFor(new Uint8Array()),
  )
  const files = unzipSync(new Uint8Array(await readFile(outcome.artifact.tempPath)))
  const media = Object.keys(files).filter((n) => n.startsWith('ppt/media/') && !n.endsWith('/'))
  assert.equal(media.length, 1, `应写入一张图片，实际 ${media.join(',')}`)
  assert.ok(media[0].endsWith('.png'))
})

test('pptx_create：取不到图片时只降级该图，不阻断整份 deck', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'pptx-pro-img2-'))
  const ref = pathToFileURL(join(dir, 'deck.pptx')).href
  await writeFile(join(dir, 'notes.txt'), 'not an image')
  const pages = [
    { type: 'cover', layout: 'cover_split_image', title: 'No photo', image: './notes.txt' },
  ]
  const planned = await planMod.plan(
    { tool: 'pptx_create', params: { pages }, ref, missing: true },
    ctxFor(new Uint8Array()),
  )
  const outcome = await execMod.exec(
    { ref, baseVersion: '', plan: planned.plan, missing: true },
    ctxFor(new Uint8Array()),
  )
  const bytes = new Uint8Array(await readFile(outcome.artifact.tempPath))
  const files = unzipSync(bytes)
  assert.equal(Object.keys(files).filter((n) => n.startsWith('ppt/media/') && !n.endsWith('/')).length, 0)
  assert.ok(
    (outcome.result.warnings ?? []).some((w) => /图片取不到/.test(w)),
    JSON.stringify(outcome.result.warnings),
  )
})

test('版式库：17 种版式全部几何审计通过（含图片版式）', async () => {
  const dataUrl = `data:image/png;base64,${PNG_1PX.toString('base64')}`
  const pages = [
    { type: 'cover', layout: 'cover_typography_hero', kicker: 'K', title: 'Hero cover', subtitle: 'sub' },
    { type: 'cover', layout: 'cover_dark_minimal', title: 'Dark minimal', subtitle: 'sub' },
    { type: 'cover', layout: 'cover_split_color', title: 'Split colour', items: [{ value: '40%' }] },
    { type: 'cover', layout: 'cover_magazine', title: 'Magazine', subtitle: 'sub', kicker: 'ISSUE 01' },
    { type: 'cover', layout: 'cover_split_image', title: 'Split image', image: dataUrl, imageCaption: 'caption' },
    { type: 'cover', layout: 'cover_full_image_overlay', title: 'Full image', image: dataUrl },
    {
      type: 'content',
      layout: 'three_column_cards',
      title: 'Cards',
      items: [
        { value: '+31%', title: 'a', note: 'note a' },
        { value: '+18%', title: 'b', note: 'note b' },
        { value: '-22%', title: 'c', note: 'note c' },
      ],
    },
    {
      type: 'data',
      layout: 'kpi_cards_row',
      title: 'KPIs',
      items: [
        { value: '92%', title: 'a', note: 'n' },
        { value: '12%', title: 'b', note: 'n' },
        { value: '4.1d', title: 'c', note: 'n' },
      ],
      body: ['one', 'two'],
    },
    {
      type: 'data',
      layout: 'hero_big_number',
      title: 'Hero',
      items: [{ value: '+38%', title: 'growth', note: 'vs 2024' }],
      body: ['one', 'two'],
    },
    {
      type: 'content',
      layout: 'two_column_comparison',
      title: 'Compare',
      items: [
        { title: 'Before', note: 'a\nb\nc' },
        { title: 'After', note: 'a\nb\nc', value: '+38%' },
      ],
    },
    {
      type: 'data',
      layout: 'two_by_two_grid',
      title: 'Grid',
      items: [{ title: 'a' }, { title: 'b' }, { title: 'c' }, { title: 'd' }],
    },
    {
      type: 'content',
      layout: 'timeline_horizontal',
      title: 'Timeline',
      items: [
        { value: 'Q1', title: 'a', note: 'n' },
        { value: 'Q2', title: 'b', note: 'n' },
        { value: 'Q3', title: 'c', note: 'n' },
        { value: 'Q4', title: 'd', note: 'n' },
      ],
    },
    {
      type: 'content',
      layout: 'left_text_right_image',
      title: 'Text and image',
      body: ['one', 'two'],
      image: dataUrl,
    },
    { type: 'content', layout: 'full_image_text_overlay', title: 'Overlay', body: ['one'], image: dataUrl },
    { type: 'content', layout: 'title_and_bullets', title: 'Bullets', body: ['one', 'two', 'three'] },
    { type: 'closing', layout: 'closing_thank_you', title: '谢谢', subtitle: 'mail@acme.com' },
    {
      type: 'closing',
      layout: 'closing_cta',
      title: 'Let us go',
      subtitle: 'next step',
      items: [{ value: '安排复盘' }],
      footnote: 'owner · 2025-01-15',
    },
  ]
  for (const style of ['consulting-navy', 'tech-dark', 'editorial-warm', 'minimal-mono']) {
    const planned = await planMod.plan(
      { tool: 'pptx_create', params: { pages, style }, ref: `verify://layouts-${style}.pptx`, missing: true },
      ctxFor(new Uint8Array()),
    )
    assert.equal(planned.plan.kind, 'build', planned.summary)
    const outcome = await execMod.exec(
      { ref: `verify://layouts-${style}.pptx`, baseVersion: '', plan: planned.plan, missing: true },
      ctxFor(new Uint8Array()),
    )
    const bytes = new Uint8Array(await readFile(outcome.artifact.tempPath))
    const audit = await planMod.plan(
      { tool: 'pptx_check', params: { audit: true }, ref: 'x', bytes },
      ctxFor(bytes),
    )
    assert.match(audit.summary, /几何审计通过/, `[${style}] ${audit.summary}`)
    const files = unzipSync(bytes)
    assert.equal(
      Object.keys(files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length,
      pages.length,
    )
  }
})
