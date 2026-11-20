import assert from 'node:assert/strict'
import test from 'node:test'
import {
  foldCase,
  matchContext,
  searchInIndex,
} from '../dist/shared/search.mjs'

function entry(text, items = []) {
  return { text, lower: foldCase(text), items }
}

function item(text, start, x, y, w, h) {
  return { start, end: start + text.length, x, y, w, h }
}

test('searchInIndex is case-insensitive by default', () => {
  const index = [entry('Hello World', [item('Hello World', 0, 0, 0, 100, 10)])]
  assert.equal(searchInIndex(index, 'hello').length, 1)
  assert.equal(searchInIndex(index, 'hello', { matchCase: true }).length, 0)
  assert.equal(searchInIndex(index, 'Hello', { matchCase: true }).length, 1)
})

test('whole-word matching rejects longer Latin words but allows CJK', () => {
  const latin = [entry('cat catalog', [item('cat catalog', 0, 0, 0, 100, 10)])]
  assert.equal(searchInIndex(latin, 'cat', { wholeWord: true }).length, 1)
  const cjk = [entry('中文搜索测试', [item('中文搜索测试', 0, 0, 0, 100, 10)])]
  assert.equal(searchInIndex(cjk, '搜索', { wholeWord: true }).length, 1)
})

test('match spanning several text items produces several rects', () => {
  const text = 'abcdef'
  const index = [entry(text, [item('abc', 0, 0, 0, 30, 10), item('def', 3, 30, 0, 30, 10)])]
  const matches = searchInIndex(index, 'cde')
  assert.equal(matches.length, 1)
  assert.equal(matches[0].rects.length, 2)
})

test('empty query returns no matches and results are capped by context format', () => {
  assert.deepEqual(searchInIndex([entry('abc')], ' '), [])
  assert.equal(matchContext(entry('abcdef'), 'cd'), 'abcdef')
})
