import assert from 'node:assert/strict'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'

const text = await import(new URL(`../dist/text.js?test=${Date.now()}`, import.meta.url))

test('text worker entry extracts page text and outline placeholders', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'pdf-text-'))
  try {
    const doc = await PDFDocument.create()
    doc.setTitle('Text Demo')
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const page = doc.addPage([300, 400])
    page.drawText('Alpha Beta Gamma', { x: 20, y: 360, size: 14, font, color: rgb(0, 0, 0) })
    const path = join(dir, 'text.pdf')
    await writeFile(path, await doc.save())

    const result = await text.handleTask({
      kind: 'text',
      path,
      pages: 'all',
      maxPages: 10,
      maxChars: 10_000,
      perPageChars: 10_000,
    })
    assert.equal(result.pageCount, 1)
    assert.match(result.pages[0].text, /Alpha Beta Gamma/)

    const pages = await text.handleTask({ kind: 'pages', path, limit: 10 })
    assert.equal(pages.pageCount, 1)
    assert.equal(pages.pages[0].page, 1)

    const outline = await text.handleTask({ kind: 'outline', path })
    assert.deepEqual(outline, [])
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
})
