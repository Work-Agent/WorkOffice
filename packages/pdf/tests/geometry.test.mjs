import assert from 'node:assert/strict'
import test from 'node:test'
import {
  geomDispSize,
  pdfRectToCss,
  pdfToView,
  quadToRect,
  viewToPdf,
  withViewRotation,
} from '../dist/shared/geometry.mjs'

const base = { pw: 600, ph: 800, rot: 0, x0: 0, y0: 0, userUnit: 1 }

test('geomDispSize swaps width/height at 90/270', () => {
  assert.deepEqual(geomDispSize(base), { width: 600, height: 800 })
  assert.deepEqual(geomDispSize({ ...base, rot: 90 }), { width: 800, height: 600 })
  assert.deepEqual(geomDispSize({ ...base, rot: 180 }), { width: 600, height: 800 })
  assert.deepEqual(geomDispSize({ ...base, rot: 270 }), { width: 800, height: 600 })
})

test('viewToPdf / pdfToView round-trip at every rotation', () => {
  for (const rot of [0, 90, 180, 270]) {
    const geom = { ...base, rot }
    for (const [x, y] of [
      [0, 0],
      [10, 20],
      [599, 799],
    ]) {
      const [vx, vy] = pdfToView(geom, x, y)
      const [rx, ry] = viewToPdf(geom, vx, vy)
      assert.ok(Math.abs(rx - x) < 1e-6, `rot ${rot} x ${rx} != ${x}`)
      assert.ok(Math.abs(ry - y) < 1e-6, `rot ${rot} y ${ry} != ${y}`)
    }
  }
})

test('cropBox offset is preserved', () => {
  const geom = { ...base, x0: 10, y0: 20 }
  assert.deepEqual(pdfToView(geom, 10, 20), [0, 800])
  assert.deepEqual(viewToPdf(geom, 0, 0), [10, 820])
})

test('userUnit scales layout coordinates only', () => {
  const geom = { ...base, pw: 1200, ph: 1600, userUnit: 2 }
  assert.deepEqual(pdfToView(geom, 0, 0), [0, 1600])
  assert.deepEqual(pdfToView(geom, 600, 800), [1200, 0])
  assert.deepEqual(viewToPdf(geom, 0, 1600), [0, 0])
})

test('withViewRotation accumulates page /Rotate and view delta', () => {
  assert.equal(withViewRotation({ ...base, rot: 90 }, 180).rot, 270)
  assert.equal(withViewRotation({ ...base, rot: 270 }, 180).rot, 90)
})

test('pdfRectToCss maps a PDF rect into a scaled CSS box', () => {
  const geom = { ...base, rot: 90 }
  const box = pdfRectToCss(geom, [0, 0, 100, 50], 2)
  assert.deepEqual(box, { left: 0, top: 0, width: 100, height: 200 })
})

test('quadToRect returns the bounding box', () => {
  assert.deepEqual(quadToRect([10, 20, 30, 20, 30, 40, 10, 40]), [10, 20, 30, 40])
})
