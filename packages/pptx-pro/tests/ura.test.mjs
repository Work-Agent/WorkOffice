import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'

const seen = { codecs: [], presenters: [], sidecars: [], stores: [] }
globalThis.__workstudioSdk = {
  registerTransport() {},
  registerCodec(codec) { seen.codecs.push(codec) },
  registerPresenter(presenter) { seen.presenters.push(presenter) },
  registerSidecar(sidecar) { seen.sidecars.push(sidecar) },
  registerStore(store) { seen.stores.push(store) },
  registerCollection() {},
  registerViewAction() {},
}

await import('../dist/ura.js?test=' + Date.now())

const SEED = new Uint8Array(await readFile(new URL('../dist/seed.pptx', import.meta.url)))
const raw = { bytes: SEED, version: 'v1', size: SEED.byteLength, type: 'file' }

test('codec：pptx-pro / patchType bytes / 拒绝 encode', async () => {
  assert.equal(seen.codecs.length, 1)
  const codec = seen.codecs[0]
  assert.equal(codec.id, 'pptx-pro')
  assert.equal(codec.patchType, 'bytes')
  assert.ok(codec.extensions.includes('pptx'))
  const decoded = await codec.decode(raw, 'file:///w/demo.pptx')
  assert.match(decoded.anchorMap, /第 1 页/)
  assert.equal(decoded.outline.length, 1)
  assert.throws(() => codec.encode(decoded, { type: 'replace-all', content: SEED }), /command\.forbidden/)
})

test('presenter：只放元数据，不放 bytes', () => {
  const presenter = seen.presenters[0]
  assert.equal(presenter.surface, 'pptx-pro')
  const decoded = seen.codecs[0].decode(raw, 'file:///w/demo.pptx')
  const presentation = presenter.present(decoded, raw, 'file:///w/demo.pptx')
  assert.equal(presentation.type, 'page')
  assert.equal(presentation.data.kind, 'pptx')
  assert.equal(presentation.data.bytes, undefined)
  assert.equal(presentation.data.slideCount, 1)
  assert.equal(presentation.data.runtime.model, 'resident')
})

test('sidecar / store：幻灯片文本与查询', async () => {
  const sidecar = seen.sidecars[0]
  const decoded = seen.codecs[0].decode(raw, 'file:///w/demo.pptx')
  const text = sidecar.sidecar(decoded)
  assert.equal(text.type, 'text')
  assert.match(text.content, /第 1 页/)
  const store = seen.stores[0]
  const result = store.query(decoded, { where: {}, limit: 10 })
  assert.equal(result.total, 1)
  assert.equal(result.rows[0].index, 1)
})
