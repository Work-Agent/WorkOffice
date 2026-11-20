import assert from 'node:assert/strict'
import test from 'node:test'
import { hasPdfHeader, probePdfMeta } from '../dist/shared/pdf-meta.mjs'

const encoder = new TextEncoder()

/** 生成一个最小可解析 PDF（经典 xref 表）。 */
function buildPdf({ pages = 2, title = 'Hello', author = 'World', encrypt = false } = {}) {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    `<< /Type /Pages /Count ${pages} /Kids [3 0 R] >>`,
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>',
    `<< /Title (${title}) /Author (${author}) >>`,
    ...(encrypt ? ['<< /Filter /Standard /V 2 /R 3 /P -3904 >>'] : []),
  ]
  let body = '%PDF-1.7\n%\xE2\xE3\xCF\xD3\n'
  const offsets = []
  objects.forEach((object, index) => {
    offsets.push(encoder.encode(body).length)
    body += `${index + 1} 0 obj\n${object}\nendobj\n`
  })
  const xrefAt = encoder.encode(body).length
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  for (const offset of offsets) body += `${String(offset).padStart(10, '0')} 00000 n \n`
  body += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R /Info 4 0 R${encrypt ? ' /Encrypt 5 0 R' : ''} >>\n`
  body += `startxref\n${xrefAt}\n%%EOF\n`
  return encoder.encode(body)
}

test('probePdfMeta: header + classic xref + info + page count', () => {
  const bytes = buildPdf()
  assert.equal(hasPdfHeader(bytes), true)
  const meta = probePdfMeta(bytes)
  assert.equal(meta.headerVersion, '1.7')
  assert.equal(meta.xref, 'table')
  assert.equal(meta.encrypted, false)
  assert.equal(meta.pageCount, 2)
  assert.equal(meta.title, 'Hello')
  assert.equal(meta.author, 'World')
  assert.equal(meta.warning, undefined)
})

test('probePdfMeta: encrypted flag comes from trailer', () => {
  const meta = probePdfMeta(buildPdf({ encrypt: true }))
  assert.equal(meta.encrypted, true)
  assert.equal(meta.pageCount, 2)
})

test('probePdfMeta: non-PDF bytes are rejected by the header check', () => {
  const bytes = encoder.encode('hello world')
  assert.equal(hasPdfHeader(bytes), false)
  assert.equal(probePdfMeta(bytes).warning, 'not-pdf')
})

test('probePdfMeta: damaged xref degrades instead of throwing', () => {
  const bytes = encoder.encode('%PDF-1.4\n1 0 obj\n<< /Type /Catalog >>\nendobj\n')
  const meta = probePdfMeta(bytes)
  assert.equal(meta.headerVersion, '1.4')
  assert.equal(meta.xref, 'missing')
  assert.equal(meta.warning, 'xref-missing')
})

test('probePdfMeta: xref stream dict is read without inflating streams', () => {
  const head = '%PDF-1.5\n'
  const offset = encoder.encode(head).length
  const body =
    `${offset} 0 obj\n<< /Type /XRef /Size 9 /Root 1 0 R /Encrypt 3 0 R /W [1 2 1] >>\n` +
    'stream\nbinary\nendstream\nendobj\n'
  const tail = `startxref\n${offset}\n%%EOF\n`
  const meta = probePdfMeta(encoder.encode(head + body + tail))
  assert.equal(meta.xref, 'stream')
  assert.equal(meta.size, 9)
  assert.equal(meta.encrypted, true)
})
