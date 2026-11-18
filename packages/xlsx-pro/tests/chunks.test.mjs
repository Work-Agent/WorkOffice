import { test } from 'node:test'
import assert from 'node:assert/strict'

const { assembleChunks, hasZipMagic, ChunkError } = await import('../dist/shared/chunks.mjs')

const bytesOf = (...parts) => parts.map((part) => new Uint8Array(part))

test('分块拼装：长度与声明一致即通过', () => {
  const all = assembleChunks(bytesOf([1, 2], [3, 4, 5]), 5)
  assert.deepEqual([...all], [1, 2, 3, 4, 5])
})

test('分块拼装：单块直接返回（不复制）', () => {
  const only = new Uint8Array([1, 2, 3])
  assert.equal(assembleChunks([only], 3), only)
})

test('分块拼装：声明大小未知时按块数兜底，空块不追加', () => {
  const all = assembleChunks(bytesOf([1], [2]), 0)
  assert.deepEqual([...all], [1, 2])
  assert.throws(
    () => assembleChunks(Array.from({ length: 5 }, () => new Uint8Array([1])), 0, { maxChunks: 2 }),
    (error) => error instanceof ChunkError && error.code === 'command.too-large',
  )
})

test('分块拼装：截断必须报错，不得当成完整字节', () => {
  assert.throws(
    () => assembleChunks(bytesOf([1, 2]), 10),
    (error) =>
      error instanceof ChunkError &&
      error.code === 'command.invalid' &&
      /不完整/.test(error.message),
  )
})

test('分块拼装：超出声明大小同样报错', () => {
  assert.throws(
    () => assembleChunks(bytesOf([1, 2, 3]), 2),
    (error) => error instanceof ChunkError && /超出声明大小/.test(error.message),
  )
})

test('分块拼装：空集合与超上限', () => {
  assert.throws(() => assembleChunks([], 0), /没有返回内容/)
  assert.throws(
    () => assembleChunks(bytesOf([1, 2, 3]), 3, { maxBytes: 2 }),
    (error) => error instanceof ChunkError && error.code === 'command.too-large',
  )
})

test('ZIP 魔数：本地文件头 / 空包通过，其它拒绝', () => {
  assert.equal(hasZipMagic(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x14])), true)
  assert.equal(hasZipMagic(new Uint8Array([0x50, 0x4b, 0x05, 0x06])), true)
  assert.equal(hasZipMagic(new Uint8Array([0x50, 0x4b, 0x09, 0x09])), false)
  assert.equal(hasZipMagic(new Uint8Array([0x50, 0x4b])), false)
  assert.equal(hasZipMagic(new Uint8Array([0x25, 0x50, 0x44, 0x46])), false)
})
