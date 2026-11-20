// 只读渲染策略的单一事实源（测试会静态断言这些值）。
//
// `isEvalSupported` 在 PDF.js 6 已移除：内嵌 JavaScript 始终不执行，
// 因此这里只记录语义，不再作为 getDocument 参数传入。

export const RENDER_POLICY = {
  isEvalSupported: false,
  useSystemFonts: false,
  enableXfa: false,
  annotationMode: 'ENABLE',
} as const
