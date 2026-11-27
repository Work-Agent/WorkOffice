import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { createHarness, loadEngine, applyOnPageBytes, slidePathsOf, slideXmlOf } from './harness.mjs'

const seed = new Uint8Array(await readFile(new URL('../dist/seed.pptx', import.meta.url)))
const contribution = await loadEngine()
const REF = 'file:///work/sample.pptx'

function harnessWithSeed() {
  const harness = createHarness(contribution, seed)
  harness.files.set(REF, seed)
  return harness
}

test('guide 列出七个分组', async () => {
  const harness = harnessWithSeed()
  try {
    const guide = await harness.run('guide', { ref: REF })
    for (const group of ['text', 'element', 'insert', 'table', 'slide', 'deck', 'chart'])
      assert.ok(guide.includes(group), `缺少分组 ${group}`)
  } finally {
    harness.stop()
  }
})

test('read/check 不需要页面', async () => {
  const harness = harnessWithSeed()
  try {
    const read = await harness.run('read', { ref: REF, what: 'context' })
    assert.match(String(read).toLowerCase(), /slide/)
    const audit = await harness.run('check', { ref: REF, audit: true })
    assert.match(String(audit), /几何审计/)
  } finally {
    harness.stop()
  }
})

test('无页面 apply 直接拒绝', async () => {
  const harness = harnessWithSeed()
  try {
    await assert.rejects(
      () => harness.run('apply', { ref: REF, commands: [{ op: 'addElement' }] }),
      /command\.page-unavailable/,
    )
  } finally {
    harness.stop()
  }
})

test('非演示文稿扩展名拒绝', async () => {
  const harness = harnessWithSeed()
  try {
    await assert.rejects(() => harness.run('load', { ref: 'file:///work/a.docx' }), /resource\.not-found/)
  } finally {
    harness.stop()
  }
})

test('create 文字稿', async () => {
  const harness = harnessWithSeed()
  try {
    await harness.run('create', {
      ref: 'file:///work/new.pptx',
      slides: JSON.stringify([{ title: '年度汇报', bullets: ['收入增长'] }, { title: '第二页' }]),
    })
    const bytes = harness.files.get('file:///work/new.pptx')
    assert.ok(bytes && bytes.byteLength > 0)
    assert.equal(slidePathsOf(bytes).length >= 2, true)
    assert.match(slideXmlOf(bytes), /年度汇报/)
  } finally {
    harness.stop()
  }
})

test('create 设计稿', async () => {
  const harness = harnessWithSeed()
  try {
    await harness.run('create', {
      ref: 'file:///work/design.pptx',
      pages: JSON.stringify([
        {
          background: '#0E1A2B',
          elements: [
            {
              type: 'text',
              x: 80,
              y: 200,
              w: 900,
              h: 140,
              paragraphs: [{ runs: [{ text: 'Verify deck', sizePt: 54, bold: true }] }],
            },
          ],
        },
      ]),
    })
    const xml = slideXmlOf(harness.files.get('file:///work/design.pptx'))
    assert.match(xml, /Verify deck/)
    assert.match(xml, /<p:bg>/)
  } finally {
    harness.stop()
  }
})

test('create 目标已存在时拒绝', async () => {
  const harness = harnessWithSeed()
  try {
    await assert.rejects(
      () => harness.run('create', { ref: REF, slides: JSON.stringify([{ title: 'x' }]) }),
      /resource\.exists/,
    )
  } finally {
    harness.stop()
  }
})

test('页面队列：attach → apply → save → ack', async () => {
  const harness = harnessWithSeed()
  try {
    const attached = await harness.run('attach', { ref: REF, pageId: 'p1' })
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
    assert.equal(request.type, 'apply')

    const edited = await applyOnPageBytes(harness.files.get(REF), request.commands)
    assert.equal(edited.ok, true, edited.error)

    const receipt = await harness.run('save', {
      ref: REF,
      pageId: 'p1',
      baseHash: attached.hash,
      bytes: edited.bytes,
    })
    assert.match(receipt.hash, /^[0-9a-f]{64}$/)
    assert.match(slideXmlOf(harness.files.get(REF)), /Hello pptx/)

    await harness.ack(id, true, { receipt })
    const settled = await promise
    assert.equal(settled.hash, receipt.hash)
  } finally {
    harness.stop()
  }
})

test('保存：旧 baseHash 报 stale，坏包报 invalid', async () => {
  const harness = harnessWithSeed()
  try {
    const attached = await harness.run('attach', { ref: REF, pageId: 'p1' })
    const edited = await applyOnPageBytes(seed, [
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'textbox',
        offset: { x: 1000000, y: 1000000, cx: 3000000, cy: 1000000 },
        paragraphs: [{ runs: [{ text: 'once' }] }],
      },
    ])
    assert.equal(edited.ok, true, edited.error)
    await harness.run('save', { ref: REF, pageId: 'p1', baseHash: attached.hash, bytes: edited.bytes })

    await assert.rejects(
      () => harness.run('save', { ref: REF, pageId: 'p1', baseHash: attached.hash, bytes: edited.bytes }),
      /command\.stale/,
    )
    const hash = (await harness.run('load', { ref: REF })).hash
    await assert.rejects(
      () => harness.run('save', { ref: REF, pageId: 'p1', baseHash: hash, bytes: new Uint8Array([1, 2, 3]) }),
      /command\.invalid/,
    )
  } finally {
    harness.stop()
  }
})

test('助手不能调页面命令', async () => {
  const harness = harnessWithSeed()
  try {
    for (const action of ['load', 'save', 'attach', 'ack', 'report', 'detach', 'pending']) {
      assert.notEqual(contribution.commands[`pptx.${action}`]?.expose?.agent, true, action)
    }
    for (const action of ['guide', 'read', 'check', 'create', 'apply']) {
      assert.equal(contribution.commands[`pptx.${action}`]?.expose?.agent, true, action)
    }
  } finally {
    harness.stop()
  }
})

test('detach 后写被拒', async () => {
  const harness = harnessWithSeed()
  try {
    await harness.run('attach', { ref: REF, pageId: 'p1' })
    await harness.run('detach', { ref: REF, pageId: 'p1' })
    await assert.rejects(
      () => harness.run('apply', { ref: REF, commands: [{ op: 'addElement' }] }),
      /command\.page-unavailable/,
    )
  } finally {
    harness.stop()
  }
})
