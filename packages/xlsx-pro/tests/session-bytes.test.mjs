// 会话字节取用：分块拼装 + 完整性校验。页面现在默认不读整份文件（P5），
// 这条路径只在"老宿主"和"保存回退"时走，所以更要把截断挡在这里。

import { test } from 'node:test'
import assert from 'node:assert/strict'

const { fetchSessionBytes } = await import('../dist/shared/session-bytes.mjs')

/** 造一个"看起来像 zip"的字节流：PK\x03\x04 + 填充。 */
function zipLike(length) {
  const bytes = new Uint8Array(length)
  bytes.set([0x50, 0x4b, 0x03, 0x04], 0)
  for (let index = 4; index < length; index += 1) bytes[index] = index % 251
  return bytes
}

function fetcherFor(bytes, { chunk = 4 } = {}) {
  return async (_id, { start, end }) => bytes.subarray(start, Math.min(end, start + chunk, bytes.length))
}

test('按声明大小拼装（多块）', async () => {
  const bytes = zipLike(20)
  const got = await fetchSessionBytes(fetcherFor(bytes), 'session-1', 20)
  assert.deepEqual([...got], [...bytes])
})

test('大小未知：以空块终止', async () => {
  const bytes = zipLike(13)
  const got = await fetchSessionBytes(fetcherFor(bytes), 'session-1', 0)
  assert.equal(got.byteLength, 13)
})

test('截断必须报错（不能把半份字节交给解析器）', async () => {
  const bytes = zipLike(6)
  await assert.rejects(
    fetchSessionBytes(fetcherFor(bytes), 'session-1', 20),
    /分块读取不完整/,
  )
})

test('size 只是提示：保存后文件变大（多读）不得报错', async () => {
  // 真实故障：打开时 17819 字节，保存后 17821 → 旧实现按"长度必须相等"直接抛
  const bytes = zipLike(17821)
  const got = await fetchSessionBytes(fetcherFor(bytes, { chunk: 1 << 20 }), 'session-1', 17819)
  assert.equal(got.byteLength, 17821)
  assert.deepEqual([...got], [...bytes])
})

test('size 未知（0）：以空块终止，不做长度断言', async () => {
  const bytes = zipLike(40)
  const got = await fetchSessionBytes(fetcherFor(bytes, { chunk: 7 }), 'session-1', 0)
  assert.equal(got.byteLength, 40)
})

test('声明超过页面上限 → command.too-large', async () => {
  await assert.rejects(
    fetchSessionBytes(fetcherFor(zipLike(4)), 'session-1', 200 * 1024 * 1024),
    (error) => error.code === 'command.too-large',
  )
})

test('非 ZIP 内容被挡在打开阶段', async () => {
  const notZip = new Uint8Array([1, 2, 3, 4, 5, 6])
  await assert.rejects(fetchSessionBytes(fetcherFor(notZip), 'session-1', 6), /不是有效的 XLSX 包/)
})

test('空会话', async () => {
  await assert.rejects(fetchSessionBytes(async () => new Uint8Array(0), 'session-1', 0), /没有返回内容/)
})
