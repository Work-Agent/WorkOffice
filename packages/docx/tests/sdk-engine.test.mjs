// WorkStudio SDK 引擎契约：guide / read / load / 无页面拒绝 / 页面队列 / save。
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFile, unlink, writeFile } from 'node:fs/promises'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { buildDocx } from './fixtures.mjs'

const { workstudio } = await import('../dist/engine.js')
const file = new URL('../.test-sdk.docx', import.meta.url)
const ref = pathToFileURL(fileURLToPath(file)).href
await writeFile(file, buildDocx())

let emit = () => {}
const scope = { pluginId: 'docx-pro', events: { emit: (name, data) => emit(name, data) }, log: console }
await workstudio.lifecycle.activate(scope)
const ctx = {
  fs: {
    read: async (target) => new Uint8Array(await readFile(fileURLToPath(target))),
    write: async (target, bytes) => writeFile(fileURLToPath(target), bytes),
  },
  changed: () => {},
  log: console,
}

test('guide / read / load', async () => {
  const guide = await workstudio.commands['docx-pro.guide'].run({ ref }, ctx)
  assert.match(guide, /blocks/)
  const read = await workstudio.commands['docx-pro.read'].run({ ref, what: 'context' }, ctx)
  assert.match(read, /blocks=/)
  const loaded = await workstudio.commands['docx-pro.load'].run({ ref }, ctx)
  assert.ok(loaded.bytes instanceof Uint8Array)
  assert.equal(typeof loaded.hash, 'string')
})

test('create：只创建空白 seed，目标存在拒绝', async () => {
  const target = new URL('../.test-created.docx', import.meta.url)
  const targetRef = pathToFileURL(fileURLToPath(target)).href
  const created = await workstudio.commands['docx-pro.create'].run({ ref: targetRef }, ctx)
  assert.equal(created.created, true)
  const loaded = await workstudio.commands['docx-pro.load'].run({ ref: targetRef }, ctx)
  assert.ok(loaded.bytes instanceof Uint8Array)
  await assert.rejects(
    () => workstudio.commands['docx-pro.create'].run({ ref: targetRef }, ctx),
    /resource\.exists/,
  )
  await unlink(target).catch(() => {})
})

test('无页面 apply 直接拒绝', async () => {
  await assert.rejects(
    () => workstudio.commands['docx-pro.apply'].run({ ref, commands: '[]' }, ctx),
    /command\.page-unavailable/,
  )
})

test('页面队列：apply → pending → ack', async () => {
  await workstudio.commands['docx-pro.attach'].run({ ref, pageId: 'p1' }, ctx)
  const eventP = new Promise((resolve) => {
    emit = (name, data) => name === 'request' && resolve(data)
  })
  const pending = workstudio.commands['docx-pro.apply'].run(
    { ref, commands: '[{"op":"setText","index":1,"text":"新"}]' },
    ctx,
  )
  const event = await eventP
  assert.equal(event.type, 'apply')
  const request = await workstudio.commands['docx-pro.pending'].run({ id: event.id }, ctx)
  assert.equal(request.type, 'apply')
  assert.equal(request.commands.length, 1)
  const receipt = { ref, hash: 'h', size: 1, at: 1, applied: 1 }
  await workstudio.commands['docx-pro.ack'].run(
    { id: event.id, ok: true, payload: JSON.stringify({ receipt }) },
    ctx,
  )
  assert.equal((await pending).hash, 'h')
})

test('save：typed array 往返 + hash CAS', async () => {
  const loaded = await workstudio.commands['docx-pro.load'].run({ ref }, ctx)
  const saved = await workstudio.commands['docx-pro.save'].run(
    { ref, pageId: 'p1', bytes: loaded.bytes, baseHash: loaded.hash, type: 'page' },
    ctx,
  )
  assert.equal(saved.hash, loaded.hash)
  await assert.rejects(
    () =>
      workstudio.commands['docx-pro.save'].run(
        { ref, pageId: 'p1', bytes: loaded.bytes, baseHash: 'stale', type: 'page' },
        ctx,
      ),
    /command\.stale/,
  )
})

test('cleanup', async () => {
  await workstudio.lifecycle.deactivate(scope)
  await unlink(file).catch(() => {})
})
