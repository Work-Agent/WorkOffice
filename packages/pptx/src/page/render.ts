// ─── RenderSlide → SVG（只读） ───────────────────────────────────────
//
// GenOffice 的 pptx-render 输出数据化绘制列表；这里用 SVG 画出来，
// 不引入 Konva/React，页面逻辑归 WorkStudio。

import type {
  ChartRenderNode,
  GlyphRun,
  PictureRenderNode,
  RenderFill,
  RenderNode,
  RenderSlide,
  RenderStroke,
  RenderTextLayout,
  ShapeRenderNode,
  TableRenderNode,
} from '@genoffice/pptx-render'

const NS = 'http://www.w3.org/2000/svg'

interface Ctx {
  svg: SVGSVGElement
  defs: SVGDefsElement
  seq: number
}

function el<K extends keyof SVGElementTagNameMap>(
  name: K,
  attrs: Record<string, string | number> = {},
): SVGElementTagNameMap[K] {
  const node = document.createElementNS(NS, name)
  for (const [key, value] of Object.entries(attrs)) {
    node.setAttribute(key, String(value))
  }
  return node
}

function color(value: string | undefined): string {
  if (!value || value === 'none') return 'transparent'
  if (/^#?[0-9a-f]{8}$/i.test(value)) {
    const raw = value.replace(/^#/, '')
    const r = parseInt(raw.slice(0, 2), 16)
    const g = parseInt(raw.slice(2, 4), 16)
    const b = parseInt(raw.slice(4, 6), 16)
    const a = parseInt(raw.slice(6, 8), 16) / 255
    return `rgba(${r},${g},${b},${a.toFixed(3)})`
  }
  return value.startsWith('#') || value.startsWith('rgb') ? value : `#${value}`
}

function fillOf(ctx: Ctx, fill: RenderFill | undefined, w: number, h: number): string {
  if (!fill || fill.kind === 'none') return 'none'
  if (fill.kind === 'solid') return color(fill.color)
  if (fill.kind === 'gradient') {
    const id = `g${++ctx.seq}`
    const gradient =
      fill.radial || fill.path
        ? el('radialGradient', { id, cx: '50%', cy: '50%', r: '65%' })
        : el('linearGradient', { id, x1: '0%', y1: '0%', x2: '100%', y2: '0%' })
    if (!fill.radial && !fill.path) {
      const angle = ((fill.angleDeg - 90) * Math.PI) / 180
      gradient.setAttribute('x1', `${50 - Math.cos(angle) * 50}%`)
      gradient.setAttribute('y1', `${50 - Math.sin(angle) * 50}%`)
      gradient.setAttribute('x2', `${50 + Math.cos(angle) * 50}%`)
      gradient.setAttribute('y2', `${50 + Math.sin(angle) * 50}%`)
    }
    for (const stop of fill.stops) {
      gradient.append(
        el('stop', {
          offset: `${Math.max(0, Math.min(1, stop.pos)) * 100}%`,
          'stop-color': color(stop.color),
        }),
      )
    }
    ctx.defs.append(gradient)
    return `url(#${id})`
  }
  if (fill.kind === 'image' && fill.dataUrl) {
    const id = `p${++ctx.seq}`
    const pattern = el('pattern', {
      id,
      patternUnits: 'userSpaceOnUse',
      width: Math.max(1, w),
      height: Math.max(1, h),
    })
    pattern.append(
      el('image', {
        href: fill.dataUrl,
        width: Math.max(1, w),
        height: Math.max(1, h),
        preserveAspectRatio: 'none',
      }),
    )
    ctx.defs.append(pattern)
    return `url(#${id})`
  }
  if (fill.kind === 'pattern') return color(fill.fg)
  return 'none'
}

function strokeOf(stroke: RenderStroke | undefined): Record<string, string | number> {
  if (!stroke || stroke.widthPx <= 0) return {}
  return {
    stroke: color(stroke.color),
    'stroke-width': stroke.widthPx,
    ...(stroke.dash?.length ? { 'stroke-dasharray': stroke.dash.join(' ') } : {}),
    ...(stroke.cap ? { 'stroke-linecap': stroke.cap } : {}),
    ...(stroke.join ? { 'stroke-linejoin': stroke.join } : {}),
  }
}

function textStyle(run: GlyphRun): Record<string, string | number> {
  return {
    fill: color(run.color),
    'font-family': run.fontFamily || 'Arial, sans-serif',
    'font-size': Math.max(1, run.fontSizePx),
    'font-weight': run.bold ? 700 : 400,
    'font-style': run.italic ? 'italic' : 'normal',
    ...(run.underline ? { 'text-decoration': 'underline' } : {}),
    ...(run.strike ? { 'text-decoration': 'line-through' } : {}),
    ...(run.letterSpacingPx ? { 'letter-spacing': run.letterSpacingPx } : {}),
  }
}

function renderText(
  parent: SVGElement,
  text: RenderTextLayout | undefined,
  ox: number,
  oy: number,
): void {
  if (!text) return
  for (const line of text.lines) {
    for (const run of line.runs) {
      if (run.highlight) {
        parent.append(
          el('rect', {
            x: ox + (text.insets.l ?? 0) + run.x,
            y: oy + (text.insets.t ?? 0) + line.top,
            width: Math.max(0, run.widthPx),
            height: Math.max(0, line.height),
            fill: color(run.highlight),
          }),
        )
      }
      const node = el('text', {
        x: ox + (text.insets.l ?? 0) + run.x,
        y: oy + (text.insets.t ?? 0) + run.baselineY,
        ...textStyle(run),
        ...(run.rotate90
          ? {
              transform: `rotate(90 ${ox + (text.insets.l ?? 0) + run.x} ${oy + (text.insets.t ?? 0) + run.baselineY})`,
            }
          : {}),
      })
      node.textContent = run.text
      parent.append(node)
    }
  }
}

function renderShape(ctx: Ctx, parent: SVGElement, shape: ShapeRenderNode): void {
  const { box } = shape
  const group = el('g', {
    'data-source-id': shape.sourceId,
    'data-node-type': shape.type,
    ...(box.rotationDeg
      ? { transform: `rotate(${box.rotationDeg} ${box.centerX} ${box.centerY})` }
      : {}),
  })
  parent.append(group)
  const fill = fillOf(ctx, shape.fill, box.w, box.h)
  const stroke = strokeOf(shape.stroke)

  if (shape.line) {
    const points = shape.line.points
    const pairs: string[] = []
    for (let i = 0; i + 1 < points.length; i += 2) {
      pairs.push(`${box.x + points[i]!},${box.y + points[i + 1]!}`)
    }
    if (shape.line.bezier?.length) {
      const d = [`M ${box.x + points[0]!} ${box.y + points[1]!}`]
      const b = shape.line.bezier
      for (let i = 0; i + 5 < b.length; i += 6) {
        d.push(
          `C ${box.x + b[i]!} ${box.y + b[i + 1]!} ${box.x + b[i + 2]!} ${box.y + b[i + 3]!} ${box.x + b[i + 4]!} ${box.y + b[i + 5]!}`,
        )
      }
      group.append(el('path', { d: d.join(' '), fill: 'none', ...stroke }))
    } else if (pairs.length) {
      group.append(el('polyline', { points: pairs.join(' '), fill: 'none', ...stroke }))
    }
  } else if (shape.pathData) {
    group.append(
      el('path', {
        d: shape.pathData,
        transform: `translate(${box.x} ${box.y})`,
        fill,
        ...stroke,
      }),
    )
  } else if (shape.polygonPoints?.length) {
    const points: string[] = []
    for (let i = 0; i + 1 < shape.polygonPoints.length; i += 2) {
      points.push(
        `${box.x + shape.polygonPoints[i]!},${box.y + shape.polygonPoints[i + 1]!}`,
      )
    }
    group.append(el('polygon', { points: points.join(' '), fill, ...stroke }))
  } else {
    group.append(
      el('rect', {
        x: box.x,
        y: box.y,
        width: Math.max(0, box.w),
        height: Math.max(0, box.h),
        rx: shape.cornerRadiusPx ?? 0,
        fill,
        ...stroke,
      }),
    )
  }
  renderText(group, shape.text, box.x, box.y)
  if (shape.type === 'text' && !shape.fill) {
    // text box has no shape fill by default
    const first = group.firstElementChild
    if (first && first.tagName === 'rect' && !shape.stroke) first.setAttribute('fill', 'none')
  }
}

function renderPicture(parent: SVGElement, picture: PictureRenderNode): void {
  const { box } = picture
  const group = el('g', {
    ...(box.rotationDeg
      ? { transform: `rotate(${box.rotationDeg} ${box.centerX} ${box.centerY})` }
      : {}),
  })
  parent.append(group)
  if (picture.dataUrl) {
    group.append(
      el('image', {
        href: picture.dataUrl,
        x: box.x,
        y: box.y,
        width: Math.max(0, box.w),
        height: Math.max(0, box.h),
        preserveAspectRatio: 'none',
        ...(picture.opacity !== undefined ? { opacity: picture.opacity } : {}),
      }),
    )
  } else {
    group.append(
      el('rect', {
        x: box.x,
        y: box.y,
        width: Math.max(0, box.w),
        height: Math.max(0, box.h),
        fill: '#f1f3f5',
      }),
    )
  }
  const stroke = strokeOf(picture.stroke)
  if (Object.keys(stroke).length) {
    group.append(
      el('rect', {
        x: box.x,
        y: box.y,
        width: Math.max(0, box.w),
        height: Math.max(0, box.h),
        fill: 'none',
        ...stroke,
      }),
    )
  }
}

function renderTable(ctx: Ctx, parent: SVGElement, table: TableRenderNode): void {
  const { box } = table
  const group = el('g', {
    ...(box.rotationDeg
      ? { transform: `rotate(${box.rotationDeg} ${box.centerX} ${box.centerY})` }
      : {}),
  })
  parent.append(group)
  if (table.bgFill) {
    group.append(
      el('rect', {
        x: box.x,
        y: box.y,
        width: box.w,
        height: box.h,
        fill: fillOf(ctx, table.bgFill, box.w, box.h),
      }),
    )
  }
  for (const cell of table.cells) {
    const x = box.x + cell.x
    const y = box.y + cell.y
    group.append(
      el('rect', {
        x,
        y,
        width: cell.w,
        height: cell.h,
        fill: fillOf(ctx, cell.fill, cell.w, cell.h),
      }),
    )
    for (const border of [cell.borders?.l, cell.borders?.r, cell.borders?.t, cell.borders?.b]) {
      if (!border) continue
      const stroke = strokeOf(border)
      if (!Object.keys(stroke).length) continue
      const coords =
        border === cell.borders?.l
          ? [x, y, x, y + cell.h]
          : border === cell.borders?.r
            ? [x + cell.w, y, x + cell.w, y + cell.h]
            : border === cell.borders?.t
              ? [x, y, x + cell.w, y]
              : [x, y + cell.h, x + cell.w, y + cell.h]
      group.append(
        el('line', {
          x1: coords[0]!,
          y1: coords[1]!,
          x2: coords[2]!,
          y2: coords[3]!,
          ...stroke,
        }),
      )
    }
    renderText(group, cell.text, x, y)
  }
}

function polar(cx: number, cy: number, r: number, deg: number): [number, number] {
  const rad = ((deg - 90) * Math.PI) / 180
  return [cx + r * Math.cos(rad), cy + r * Math.sin(rad)]
}

function wedgePath(
  cx: number,
  cy: number,
  outer: number,
  inner: number,
  start: number,
  sweep: number,
): string {
  const [x0, y0] = polar(cx, cy, outer, start)
  const [x1, y1] = polar(cx, cy, outer, start + sweep)
  const [x2, y2] = polar(cx, cy, inner, start + sweep)
  const [x3, y3] = polar(cx, cy, inner, start)
  const large = Math.abs(sweep) > 180 ? 1 : 0
  return `M ${x0} ${y0} A ${outer} ${outer} 0 ${large} 1 ${x1} ${y1} L ${x2} ${y2} A ${inner} ${inner} 0 ${large} 0 ${x3} ${y3} Z`
}

function renderChart(ctx: Ctx, parent: SVGElement, chart: ChartRenderNode): void {
  const { box } = chart
  const group = el('g', { transform: `translate(${box.x} ${box.y})` })
  parent.append(group)
  if (chart.bgFill) {
    group.append(
      el('rect', { x: 0, y: 0, width: box.w, height: box.h, fill: fillOf(ctx, chart.bgFill, box.w, box.h) }),
    )
  }
  if (chart.plotRect) {
    const p = chart.plotRect
    group.append(
      el('rect', {
        x: p.x,
        y: p.y,
        width: p.w,
        height: p.h,
        fill: p.fill ? fillOf(ctx, p.fill, p.w, p.h) : 'none',
        ...(p.borderColor
          ? { stroke: color(p.borderColor), 'stroke-width': p.borderWidthPx ?? 1 }
          : {}),
      }),
    )
  }
  for (const line of chart.gridLines) {
    group.append(
      el('line', {
        x1: line.x1,
        y1: line.y1,
        x2: line.x2,
        y2: line.y2,
        stroke: color(line.color),
        'stroke-width': line.widthPx ?? 1,
        ...(line.dash?.length ? { 'stroke-dasharray': line.dash.join(' ') } : {}),
      }),
    )
  }
  for (const line of chart.axisLines) {
    group.append(
      el('line', {
        x1: line.x1,
        y1: line.y1,
        x2: line.x2,
        y2: line.y2,
        stroke: color(line.color),
        'stroke-width': line.widthPx,
      }),
    )
  }
  for (const bar of chart.bars) {
    group.append(
      el('rect', {
        x: bar.x,
        y: bar.y,
        width: bar.w,
        height: bar.h,
        fill: bar.fill ? fillOf(ctx, bar.fill, bar.w, bar.h) : color(bar.color),
      }),
    )
  }
  for (const line of chart.polylines) {
    const pairs: string[] = []
    for (let i = 0; i + 1 < line.points.length; i += 2) {
      pairs.push(`${line.points[i]},${line.points[i + 1]}`)
    }
    group.append(
      el('polyline', {
        points: pairs.join(' '),
        fill: line.closed ? (line.fill ? color(line.fill) : 'none') : 'none',
        stroke: color(line.color),
        'stroke-width': line.widthPx,
        ...(line.dash?.length ? { 'stroke-dasharray': line.dash.join(' ') } : {}),
      }),
    )
  }
  for (const path of chart.paths ?? []) {
    group.append(
      el('path', {
        d: path.d,
        fill: color(path.fill),
        ...(path.stroke
          ? { stroke: color(path.stroke), 'stroke-width': path.strokeWidthPx ?? 1 }
          : {}),
      }),
    )
  }
  for (const wedge of chart.wedges ?? []) {
    group.append(
      el('path', {
        d: wedgePath(
          wedge.cx,
          wedge.cy,
          wedge.outerR,
          wedge.innerR,
          wedge.startDeg,
          wedge.sweepDeg,
        ),
        fill: wedge.noFill ? 'none' : color(wedge.color),
        ...(wedge.noFill && wedge.stroke
          ? { stroke: color(wedge.stroke), 'stroke-width': wedge.strokeWidthPx ?? 1 }
          : {}),
      }),
    )
  }
  for (const marker of chart.markers) {
    group.append(
      el('circle', {
        cx: marker.x,
        cy: marker.y,
        r: marker.r,
        fill: color(marker.color),
      }),
    )
  }
  for (const swatch of chart.swatches) {
    group.append(
      el('rect', {
        x: swatch.x,
        y: swatch.y,
        width: swatch.w,
        height: swatch.h,
        fill: color(swatch.color),
      }),
    )
  }
  for (const label of chart.labels) {
    const node = el('text', {
      x: label.x,
      y: label.y + label.fontSizePx * 0.8,
      fill: color(label.color),
      'font-family': label.fontFamily || 'Arial, sans-serif',
      'font-size': label.fontSizePx,
      'font-weight': label.bold ? 700 : 400,
      'font-style': label.italic ? 'italic' : 'normal',
      ...(label.rotationDeg
        ? { transform: `rotate(${label.rotationDeg} ${label.x} ${label.y})` }
        : {}),
    })
    node.textContent = label.text
    group.append(node)
  }
}

function renderNode(ctx: Ctx, parent: SVGElement, node: RenderNode): void {
  if (node.type === 'shape' || node.type === 'text') {
    renderShape(ctx, parent, node)
  } else if (node.type === 'picture') {
    renderPicture(parent, node)
  } else if (node.type === 'group') {
    const group = el('g', {
      transform: `translate(${node.box.x} ${node.box.y})`,
    })
    parent.append(group)
    for (const child of node.children) renderNode(ctx, group, child)
  } else if (node.type === 'table') {
    renderTable(ctx, parent, node)
  } else if (node.type === 'chart') {
    renderChart(ctx, parent, node)
  } else if (node.type === 'placeholder-chip') {
    const group = el('g', { transform: `translate(${node.box.x} ${node.box.y})` })
    parent.append(group)
    group.append(
      el('rect', {
        x: 0,
        y: 0,
        width: node.box.w,
        height: node.box.h,
        fill: '#f4f4f5',
        stroke: '#d4d4d8',
      }),
    )
    const text = el('text', {
      x: 12,
      y: 24,
      fill: '#71717a',
      'font-family': 'Arial, sans-serif',
      'font-size': 13,
    })
    text.textContent = node.label
    group.append(text)
  }
}

export function renderSlideToSvg(slide: RenderSlide): SVGSVGElement {
  const svg = el('svg', {
    xmlns: NS,
    viewBox: `0 0 ${slide.widthPx} ${slide.heightPx}`,
    preserveAspectRatio: 'xMidYMid meet',
    class: 'slide-svg',
  })
  const defs = el('defs')
  svg.append(defs)
  const ctx: Ctx = { svg, defs, seq: 0 }
  svg.append(
    el('rect', {
      x: 0,
      y: 0,
      width: slide.widthPx,
      height: slide.heightPx,
      fill: fillOf(ctx, slide.background, slide.widthPx, slide.heightPx),
    }),
  )
  for (const node of slide.nodes) renderNode(ctx, svg, node)
  return svg
}
