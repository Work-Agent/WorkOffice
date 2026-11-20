// ─── 页面几何（与 GenOffice annotations.ts 同构的只读子集） ───────────
//
// 渲染基准盒是 CropBox（缺省 = MediaBox）。`/UserUnit` 只影响布局坐标，
// PDF API 坐标不乘；`/Rotate` 与视图旋转相加，不是覆盖。

import type { PDFPageProxy } from 'pdfjs-dist'
import type { PageGeom } from '../shared/pdf-types'

export function normRot(deg: number): number {
  return ((deg % 360) + 360) % 360
}

/** 从 PDF.js 页面读取几何；失败页由调用方沿用上一页几何。 */
export function geomOfPage(page: PDFPageProxy): PageGeom {
  const box = page.view
  const unit = page.userUnit ?? 1
  return {
    pw: Math.max(0, (box[2]! - box[0]!) * unit),
    ph: Math.max(0, (box[3]! - box[1]!) * unit),
    rot: normRot(page.rotate ?? 0),
    x0: box[0]!,
    y0: box[1]!,
    userUnit: unit,
    // PDF.js 只公开 view（CropBox，缺省即 MediaBox）。
    mediaBox: box as [number, number, number, number],
    cropBox: box as [number, number, number, number],
  }
}

/** 页面几何 + 视图旋转 = 显示几何。 */
export function withViewRotation(geom: PageGeom, delta: number): PageGeom {
  const rot = normRot(geom.rot + delta)
  return rot === geom.rot ? geom : { ...geom, rot }
}

export function sameGeom(a: PageGeom | undefined, b: PageGeom | undefined): boolean {
  if (a === b) return true
  if (!a || !b) return false
  return (
    a.pw === b.pw &&
    a.ph === b.ph &&
    a.rot === b.rot &&
    (a.x0 ?? 0) === (b.x0 ?? 0) &&
    (a.y0 ?? 0) === (b.y0 ?? 0) &&
    (a.userUnit ?? 1) === (b.userUnit ?? 1)
  )
}

/** 显示尺寸（scale=1；宽高随旋转交换）。 */
export function geomDispSize(geom: PageGeom): { width: number; height: number } {
  return normRot(geom.rot) % 180 === 0
    ? { width: geom.pw, height: geom.ph }
    : { width: geom.ph, height: geom.pw }
}

/** 显示坐标（左上原点、y 向下、scale=1）→ PDF user space（y 向上）。 */
export function viewToPdf(geom: PageGeom, vx: number, vy: number): [number, number] {
  const x0 = geom.x0 ?? 0
  const y0 = geom.y0 ?? 0
  const unit = geom.userUnit ?? 1
  const pw = geom.pw / unit
  const ph = geom.ph / unit
  vx /= unit
  vy /= unit
  switch (normRot(geom.rot)) {
    case 90:
      return [x0 + vy, y0 + vx]
    case 180:
      return [x0 + pw - vx, y0 + vy]
    case 270:
      return [x0 + pw - vy, y0 + ph - vx]
    default:
      return [x0 + vx, y0 + ph - vy]
  }
}

/** PDF user space → 显示坐标（左上原点、y 向下、scale=1）。 */
export function pdfToView(geom: PageGeom, x: number, y: number): [number, number] {
  x -= geom.x0 ?? 0
  y -= geom.y0 ?? 0
  const unit = geom.userUnit ?? 1
  const pw = geom.pw / unit
  const ph = geom.ph / unit
  let result: [number, number]
  switch (normRot(geom.rot)) {
    case 90:
      result = [y, x]
      break
    case 180:
      result = [pw - x, y]
      break
    case 270:
      result = [ph - y, pw - x]
      break
    default:
      result = [x, ph - y]
  }
  return [result[0] * unit, result[1] * unit]
}

/** PDF-space rect [x1,y1,x2,y2] → 显示像素框（乘 scale）。 */
export function pdfRectToCss(
  geom: PageGeom,
  rect: readonly [number, number, number, number],
  scale: number,
): { left: number; top: number; width: number; height: number } {
  const [ax, ay] = pdfToView(geom, rect[0], rect[1])
  const [bx, by] = pdfToView(geom, rect[2], rect[3])
  return {
    left: Math.min(ax, bx) * scale,
    top: Math.min(ay, by) * scale,
    width: Math.abs(bx - ax) * scale,
    height: Math.abs(by - ay) * scale,
  }
}

/** quad（四角任意顺序）→ PDF-space 包围盒。 */
export function quadToRect(q: readonly number[]): [number, number, number, number] {
  const xs = [q[0]!, q[2]!, q[4]!, q[6]!]
  const ys = [q[1]!, q[3]!, q[5]!, q[7]!]
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
}
