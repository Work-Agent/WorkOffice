import assert from 'node:assert/strict'
import test from 'node:test'
import {
  BASE_DPI,
  MAX_PRINT_PIXELS,
  TARGET_DPI,
  normalizePrintPages,
  printScaleForAreas,
} from '../dist/shared/print.mjs'

test('printScaleForAreas uses the 200 DPI target within budget', () => {
  assert.equal(printScaleForAreas([612 * 792]), TARGET_DPI / 72)
})

test('printScaleForAreas scales down long documents, never below 150 DPI', () => {
  const areas = Array.from({ length: 60 }, () => 612 * 792)
  const scale = printScaleForAreas(areas)
  assert.ok(scale < TARGET_DPI / 72)
  assert.ok(scale >= BASE_DPI / 72)
  assert.ok(areas.reduce((sum, area) => sum + area, 0) * scale * scale <= MAX_PRINT_PIXELS * 1.000001)
  // 超出预算太多时守住 150 DPI 下限（宁可超预算也不低于基线质量）。
  const huge = Array.from({ length: 400 }, () => 612 * 792)
  assert.equal(printScaleForAreas(huge), BASE_DPI / 72)
})

test('printScaleForAreas ignores invalid areas', () => {
  assert.equal(printScaleForAreas([0, Number.NaN, -1]), TARGET_DPI / 72)
})

test('normalizePrintPages expands all pages and rejects empty selections', () => {
  assert.deepEqual(normalizePrintPages(undefined, 3), [1, 2, 3])
  assert.deepEqual(normalizePrintPages([3, 1, 3], 3), [1, 3])
  assert.throws(() => normalizePrintPages([9], 3), /没有可打印的页码/)
})
