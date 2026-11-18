import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'

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

await import('../dist/ura.js?test=1')
const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const bytes = new Uint8Array(readFileSync(new URL('../dist/seed.xlsx', import.meta.url)))

test('codec 注册与 encode 拒绝', () => {
  assert.equal(seen.codecs.length, 1)
  const codec = seen.codecs[0]
  assert.equal(codec.id, 'xlsx-pro')
  assert.ok(codec.keys.some((key) => key.mime === MIME))
  assert.throws(() => codec.encode({}, { type: 'replace-all', content: bytes }), /command\.forbidden/)
})

test('presenter / sidecar / store 投影', async () => {
  const codec = seen.codecs[0]
  const decoded = codec.decode({ bytes, version: 'v1', mime: MIME, size: bytes.length, type: 'file' }, 'file:///w/a.xlsx')
  const presenter = seen.presenters[0]
  const presentation = await presenter.present(decoded, { bytes, version: 'v1' }, 'file:///w/a.xlsx')
  assert.equal(presentation.surface, 'xlsx-pro')
  assert.equal(presentation.data.bytes, undefined)
  const sidecar = seen.sidecars[0]
  const text = sidecar.sidecar(decoded)
  assert.equal(text.type, 'text')
  const store = seen.stores[0]
  const result = store.query(decoded, { where: { sheet: 'Sheet1' }, limit: 5 })
  assert.ok(Array.isArray(result.rows))
})

test('带命名空间前缀的 XLSX 也能解析（有些生产工具会写 x:workbook）', async () => {
  const { zipSync, strToU8 } = await import('fflate')
  const ct = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`
  const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
  const wb = `<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:sheets><x:sheet name="数据" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>`
  const wbr = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../worksheets/sheet1.xml"/></Relationships>`
  const sh = `<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>你好</x:t></x:is></x:c></x:row></x:sheetData></x:worksheet>`
  const prefixed = zipSync({
    '[Content_Types].xml': strToU8(ct),
    '_rels/.rels': strToU8(rels),
    'xl/workbook.xml': strToU8(wb),
    'xl/_rels/workbook.xml.rels': strToU8(wbr),
    'xl/worksheets/sheet1.xml': strToU8(sh),
  })
  const codec = seen.codecs[0]
  const decoded = codec.decode({ bytes: prefixed, version: 'v1', mime: MIME, size: prefixed.length, type: 'file' }, 'file:///w/prefixed.xlsx')
  assert.equal(decoded.doc.sheets.length, 1)
  assert.equal(decoded.doc.sheets[0].name, '数据')
  assert.equal(decoded.doc.sheets[0].cells.A1.value, '你好')
})
