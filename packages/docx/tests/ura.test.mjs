// 注册面：codec / sidecar / store / presenter。
//
// 宿主装载插件时先注入 `globalThis.__workstudioSdk`，再由注册入口登记能力。
// 这里用同一个注入面跑一遍，断言：
//   - codec 能解包并产出块投影（decode 在主进程，必须便宜）
//   - encode 必须拒绝（字节只由 exec 产生）
//   - presenter.data 不含整份内容（content: 'pulled'）
//   - sidecar 给出可读文本、store 能按块查询

import { test } from 'node:test'
import assert from 'node:assert/strict'
import { pathToFileURL } from 'node:url'
import { resolve } from 'node:path'
import { buildDocx } from './fixtures.mjs'

const MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

const registered = {}
globalThis.__workstudioSdk = {
  registerCodec: (codec) => (registered.codec = codec),
  registerSidecar: (provider) => (registered.sidecar = provider),
  registerStore: (store) => (registered.store = store),
  registerPresenter: (presenter) => (registered.presenter = presenter),
}

await import(pathToFileURL(resolve('dist/ura.js')).href)

const bytes = buildDocx()
const raw = { bytes, version: 'v1', mime: MIME }
const ref = 'file:///w/%E5%B9%B4%E5%BA%A6%E6%8A%A5%E5%91%8A.docx'

test('注册了四个面，且 codec 身份正确', () => {
  for (const key of ['codec', 'sidecar', 'store', 'presenter'])
    assert.ok(registered[key], `缺少 ${key}`)
  assert.equal(registered.codec.id, 'docx-pro')
  assert.deepEqual(registered.codec.extensions, ['docx', 'docm', 'dotx'])
  assert.equal(registered.codec.keys[0].mime, MIME)
  assert.equal(registered.codec.patchType, 'bytes')
})

test('decode：块投影 / 大纲 / 文件名', () => {
  const decoded = registered.codec.decode(raw, ref)
  assert.equal(decoded.doc.name, '年度报告.docx')
  assert.equal(decoded.doc.blockCount, 4)
  assert.equal(decoded.doc.tableCount, 1)
  assert.match(decoded.anchorMap, /年度报告/)
  assert.ok(decoded.outline.some((item) => item.level === 1))
  assert.ok(Array.isArray(decoded.doc.blocks))
  assert.equal(decoded.doc.blocks[0].index, 0)
  assert.equal(decoded.doc.blocks[0].heading, 1)
})

test('encode 必须拒绝（字节只由 exec 产生）', () => {
  assert.throws(() => registered.codec.encode({}, { type: 'replace-all', bytes }), /command\.forbidden/)
})

test('presenter.data 不含整份内容（content: pulled）', () => {
  const decoded = registered.codec.decode(raw, ref)
  const presented = registered.presenter.present(decoded, raw, ref)
  assert.equal(presented.type, 'page')
  assert.equal(presented.surface, 'docx-pro')
  const data = presented.data
  assert.equal(data.bytes, undefined)
  assert.equal(data.size, bytes.byteLength)
  assert.equal(data.version, 'v1')
  assert.ok(!JSON.stringify(data).includes('PK'), 'data 里不应出现 zip 内容')
})

test('sidecar 给可读投影；store 能按块查询', () => {
  const decoded = registered.codec.decode(raw, ref)
  const sidecar = registered.sidecar.sidecar(decoded)
  assert.equal(sidecar.type, 'text')
  assert.match(sidecar.content, /blocks=/)
  const all = registered.store.query(decoded, { limit: 10 })
  assert.ok(all.rows.length > 0)
  const tables = registered.store.query(decoded, { where: { type: 'tbl' } })
  assert.equal(tables.rows.length, 1)
  const headings = registered.store.query(decoded, { where: { heading: 1 } })
  assert.equal(headings.rows.length, 1)
})
