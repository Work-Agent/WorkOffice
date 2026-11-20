// ─── 打印参数（纯函数，供页面与测试共用） ────────────────────────────
//
// 一期是页面栅格打印：150–200 DPI，总像素预算 150M，逐页渲染、canvas 即用即弃。
// 原文件向量打印需要宿主提供 print-to-pdf / print-file，列为上游缺口。

export const BASE_DPI = 150
export const TARGET_DPI = 200
export const MAX_PRINT_PIXELS = 150_000_000
export const PRINT_DIALOG_TIMEOUT_MS = 300_000

/** 选中页面积（scale=1 的 PDF point²）→ 打印栅格 scale。 */
export function printScaleForAreas(areas: number[]): number {
  const target = TARGET_DPI / 72
  const total = areas.reduce((sum, area) => sum + (Number.isFinite(area) && area > 0 ? area : 0), 0)
  if (!(total > 0) || !Number.isFinite(total)) return target
  const atTarget = total * target * target
  if (!Number.isFinite(atTarget) || atTarget <= 0) return target
  if (atTarget <= MAX_PRINT_PIXELS) return target
  return Math.max(BASE_DPI / 72, target * Math.sqrt(MAX_PRINT_PIXELS / atTarget))
}

/** 空页列表（或全部越界）必须拒绝，否则会打印一份空白文档。 */
export function normalizePrintPages(pages: number[] | undefined, numPages: number): number[] {
  if (!pages || pages.length === 0) {
    return Array.from({ length: numPages }, (_value, index) => index + 1)
  }
  const targets = [...new Set(pages)]
    .filter((n) => Number.isInteger(n) && n >= 1 && n <= numPages)
    .sort((a, b) => a - b)
  if (targets.length === 0) {
    throw new Error(`没有可打印的页码（文档共 ${numPages} 页）`)
  }
  return targets
}
