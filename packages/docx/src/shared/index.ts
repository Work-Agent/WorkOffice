// ─── docx-pro 共享实现（页面 / 执行点同源） ─────────────────────────────
//
// 同源是契约 §14.1 的验收项之一：同一份源码分别打进
// `dist/engine/*.mjs`（执行点）与 `dist/page/assets/*.js`（页面），
// 不在两侧各写一遍实现。

export * from './types'
export * from './xml'
export * from './xml-ops'
export * from './document'
export * from './catalog'
export * from './project'
export * from './sha256'
