import assert from 'node:assert/strict'
import test from 'node:test'
import { RENDER_POLICY } from '../dist/shared/render-policy.mjs'

test('read-only render policy keeps eval, system fonts and XFA off', () => {
  assert.equal(RENDER_POLICY.isEvalSupported, false)
  assert.equal(RENDER_POLICY.useSystemFonts, false)
  assert.equal(RENDER_POLICY.enableXfa, false)
  // 只读页面画 /AP 外观，不挂可交互表单 widget。
  assert.equal(RENDER_POLICY.annotationMode, 'ENABLE')
})
