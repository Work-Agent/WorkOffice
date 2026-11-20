import assert from 'node:assert/strict'
import test from 'node:test'
import { pathToFileURL } from 'node:url'
import { resolve } from 'node:path'

const encoder = new TextEncoder()

function buildPdf(title = 'Demo') {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Count 1 /Kids [3 0 R] >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>',
    `<< /Title (${title}) /Author (Tester) >>`,
  ]
  let body = '%PDF-1.7\n'
  const offsets = []
  objects.forEach((object, index) => {
    offsets.push(encoder.encode(body).length)
    body += `${index + 1} 0 obj\n${object}\nendobj\n`
  })
  const xrefAt = encoder.encode(body).length
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  for (const offset of offsets) body += `${String(offset).padStart(10, '0')} 00000 n \n`
  body += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R /Info 4 0 R >>\nstartxref\n${xrefAt}\n%%EOF\n`
  return encoder.encode(body)
}

async function loadPlugin() {
  const seen = {}
  globalThis.__workstudioSdk = {
    registerTransport() {},
    registerStore() {},
    registerCollection() {},
    registerViewAction() {},
    registerCodec(codec) {
      seen.codec = codec
    },
    registerPresenter(presenter) {
      seen.presenter = presenter
    },
    registerSidecar(provider) {
      seen.sidecar = provider
    },
  }
  const entry = resolve('dist/ura.js')
  await import(`${pathToFileURL(entry).href}?test=${Date.now()}`)
  return seen
}

test('codec / presenter / sidecar keep the read-only contract', async () => {
  const { codec, presenter, sidecar } = await loadPlugin()
  assert.equal(codec.id, 'pdf')
  assert.equal(codec.patchType, 'bytes')
  assert.deepEqual(codec.extensions, ['pdf'])

  const bytes = buildPdf()
  const raw = { version: 'v1', type: 'file', size: bytes.byteLength, bytes }
  const decoded = await codec.decode(raw, 'file:///tmp/Demo.pdf')
  assert.equal(decoded.doc.doc.pageCount, 1)
  assert.equal(decoded.doc.doc.title, 'Demo')

  const presentation = presenter.present(decoded, raw, 'file:///tmp/Demo.pdf')
  assert.equal(presentation.type, 'page')
  assert.equal(presentation.surface, 'pdf-viewer')
  assert.equal(presentation.data.kind, 'pdf')
  assert.equal(presentation.data.size, bytes.byteLength)
  const serialized = JSON.stringify(presentation.data)
  assert.equal(serialized.includes('base64'), false)
  assert.equal('bytes' in presentation.data, false)

  // 只读插件拒绝 encode。
  assert.throws(() => codec.encode(decoded, {}), /command\.forbidden/)

  // sidecar 只拼元数据：给它一个一读 bytes 就抛错的 raw，也必须成功。
  const trapRaw = new Proxy(raw, {
    get(target, key, receiver) {
      if (key === 'bytes') throw new Error('sidecar must not read raw.bytes')
      return Reflect.get(target, key, receiver)
    },
  })
  const text = sidecar.sidecar(decoded, trapRaw, 'file:///tmp/Demo.pdf')
  assert.equal(text.type, 'text')
  assert.match(text.content, /pages=1/)
  assert.match(text.content, /title: Demo/)
})

test('codec rejects non-PDF bytes', async () => {
  const { codec } = await loadPlugin()
  const bytes = encoder.encode('not a pdf')
  await assert.rejects(
    codec.decode(
      { version: 'v1', type: 'file', size: bytes.byteLength, bytes },
      'file:///tmp/x.pdf',
    ),
    /不是有效的 PDF/,
  )
})
