import assert from 'node:assert/strict'
import test from 'node:test'
import { assembleBytes, MAX_PDF_BYTES } from '../dist/shared/chunk.mjs'

const encoder = new TextEncoder()

function readerFor(bytes, chunkSize = 4) {
  const calls = []
  const read = async ({ start, end }) => {
    calls.push([start, end])
    return bytes.subarray(start, Math.min(bytes.length, end, start + chunkSize))
  }
  return { read, calls }
}

test('assembleBytes preallocates once and fills chunk by chunk', async () => {
  const source = encoder.encode('abcdefghij')
  const { read, calls } = readerFor(source)
  const bytes = await assembleBytes(read, { size: source.length, chunkBytes: 4 })
  assert.equal(new TextDecoder().decode(bytes), 'abcdefghij')
  assert.equal(calls[0][0], 0)
  assert.equal(bytes.byteLength, source.length)
})

test('assembleBytes rejects a short read instead of silently truncating', async () => {
  const source = encoder.encode('1234567')
  const read = async ({ start, end }) => source.subarray(start, Math.min(end, source.length))
  await assert.rejects(
    assembleBytes(read, { size: 10, chunkBytes: 4 }),
    /读取不完整/,
  )
})

test('assembleBytes refuses oversized documents', async () => {
  await assert.rejects(
    assembleBytes(async () => new Uint8Array(0), { size: MAX_PDF_BYTES + 1 }),
    /PDF 过大/,
  )
})

test('assembleBytes supports unknown size by stopping on an empty chunk', async () => {
  const source = encoder.encode('abcdef')
  let offset = 0
  const read = async ({ end }) => {
    const chunk = source.subarray(offset, Math.min(source.length, end))
    offset += chunk.length
    return chunk
  }
  const bytes = await assembleBytes(read, { size: 0, chunkBytes: 4 })
  assert.equal(new TextDecoder().decode(bytes), 'abcdef')
})
