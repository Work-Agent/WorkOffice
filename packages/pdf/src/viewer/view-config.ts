// CMaps / standard fonts / wasm 由 scripts/build.mjs 复制到 page 产物的 pdfjs/。
export const ASSET_BASE = new URL('pdfjs/', document.baseURI).href

export const ZOOM_STEPS = [
  0.1, 0.25, 0.5, 0.67, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4, 5, 6, 8,
]
export const MIN_SCALE = ZOOM_STEPS[0]!
export const MAX_SCALE = ZOOM_STEPS[ZOOM_STEPS.length - 1]!
export const PAGE_GAP = 16
export const SCROLL_PAD = 24
/** 单页位图预算：800% 下 Letter 页否则接近 125M 像素。 */
export const MAX_PAGE_RENDER_PIXELS = 48_000_000
export const THUMB_WIDTH = 140
export const PRINT_ALL = 'all'

export type FitMode = 'width' | 'page' | null
