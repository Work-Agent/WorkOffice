import assert from 'node:assert/strict'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import test from 'node:test'
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'

const engine = await import(new URL(`../dist/engine.js?test=${Date.now()}`, import.meta.url))

async function makePdf(dir) {
  const doc = await PDFDocument.create()
  doc.setTitle('Demo PDF')
  doc.setAuthor('Tester')
  const font = await doc.embedFont(StandardFonts.Helvetica)
  for (const text of ['Hello WorkStudio PDF text', 'Second page revenue report']) {
    const page = doc.addPage([595, 842])
    page.drawText(text, { x: 60, y: 760, size: 18, font, color: rgb(0, 0, 0) })
  }
  const bytes = await doc.save()
  const path = join(dir, 'demo.pdf')
  await writeFile(path, bytes)
  return { path, ref: pathToFileURL(path).href, bytes }
}

function contextFor() {
  return {
    fs: {
      read: async (ref) => new Uint8Array(await readFile(new URL(ref))),
      write: async () => undefined,
    },
    access: {},
    settings: { get: () => undefined, watch: () => () => undefined },
    credentials: {},
    changed: () => undefined,
    log: { info: () => undefined, warn: () => undefined, error: () => undefined },
  }
}

function activate(saveTarget) {
  const scope = {
    pluginId: 'pdf',
    dataDir: tmpdir(),
    tempDir: tmpdir(),
    settings: { get: (_k, fallback) => fallback, set: () => undefined, watch: () => () => undefined },
    secrets: { availability: () => ({ available: false }), get: async () => null, set: async () => undefined },
    credentials: {},
    events: { emit: () => undefined },
    files: {
      resolveInput: () => {
        throw new Error('not used')
      },
      resolveSave: (token) => {
        if (token !== 'save:test') throw new Error('bad token')
        return saveTarget
      },
    },
    log: { info: () => undefined, warn: () => undefined, error: () => undefined },
  }
  engine.workstudio.lifecycle?.activate?.(scope)
}

test('engine exposes the read-only SDK contract', () => {
  const commands = engine.workstudio.commands ?? {}
  assert.deepEqual(Object.keys(commands).sort(), [
    'pdf.chunk',
    'pdf.load',
    'pdf.read',
    'pdf.saveCopy',
  ])
  assert.equal(commands['pdf.read'].expose?.agent, true)
  assert.notEqual(commands['pdf.load'].expose?.agent, true)
  assert.ok(engine.workstudio.resources?.pdf)
  assert.notEqual(engine.workstudio.events, true)
})

test('load / chunk / read / saveCopy round-trip', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'pdf-plugin-'))
  try {
    const { path, ref, bytes } = await makePdf(dir)
    const ctx = contextFor()
    const commands = engine.workstudio.commands

    const loaded = await commands['pdf.load'].run({ ref }, ctx)
    assert.equal(loaded.name, 'demo.pdf')
    assert.equal(loaded.size, bytes.byteLength)
    assert.equal(loaded.inline, true)
    assert.ok(loaded.bytes instanceof Uint8Array)
    assert.equal(loaded.meta.pageCount, 2)
    assert.equal(loaded.meta.title, 'Demo PDF')

    const chunk = await commands['pdf.chunk'].run(
      { ref, start: 0, length: 16, hash: loaded.hash },
      ctx,
    )
    assert.equal(chunk.hash, loaded.hash)
    assert.deepEqual([...chunk.bytes], [...bytes.subarray(0, 16)])

    const context = await commands['pdf.read'].run({ ref, what: 'context' }, ctx)
    assert.match(context, /pages=2/)
    assert.match(context, /title: Demo PDF/)

    const text = await commands['pdf.read'].run({ ref, what: 'text', page: 1 }, ctx)
    assert.match(text, /Hello WorkStudio PDF text/)

    const found = await commands['pdf.read'].run(
      { ref, what: 'find', query: 'revenue' },
      ctx,
    )
    assert.match(found, /p\.2:/)

    const outline = await commands['pdf.read'].run({ ref, what: 'outline' }, ctx)
    assert.equal(typeof outline, 'string')

    const pages = await commands['pdf.read'].run({ ref, what: 'pages' }, ctx)
    assert.match(pages, /p\.1: 595×842/)

    const target = join(dir, 'copy.pdf')
    activate(target)
    const saved = await commands['pdf.saveCopy'].run(
      { ref, token: 'save:test', bytes: new Uint8Array(bytes) },
      ctx,
    )
    assert.equal(saved.ok, true)
    assert.deepEqual(new Uint8Array(await readFile(target)), new Uint8Array(bytes))

    // resources.pdf：read 元数据、grep 页文本、open 透传 locator。
    const resource = engine.workstudio.resources.pdf
    const items = await resource.read(ref)
    assert.equal(items[0].locator, 'pdf:doc')
    assert.equal(items[0].pageCount, 2)
    assert.equal(items[0].title, 'Demo PDF')

    const hits = await resource.grep(ref, 'revenue')
    assert.equal(hits.length, 1)
    assert.equal(hits[0].locator, 'pdf:page:2')

    const opened = await resource.open(ref, 'pdf:page:2')
    assert.deepEqual(opened, { ref, locator: 'pdf:page:2' })

    // 非 PDF 拒绝。
    const bad = join(dir, 'bad.pdf')
    await writeFile(bad, 'not a pdf')
    await assert.rejects(
      commands['pdf.load'].run({ ref: pathToFileURL(bad).href }, ctx),
      /不是有效的 PDF/,
    )
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
})
