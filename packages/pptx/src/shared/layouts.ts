// ─── 版式库：内容 + 预设 → 1280×720 的页面元素 ─────────────────────────
//
// 模型只写"这一页是什么版式 + 里面有什么内容"，坐标由这里算。
// 理由和 genoffice 的 layout script 一样：坐标不该经过 LLM token 往返，
// 手算出来的间距一定不齐，版式也就没有层次。
//
// 每个版式都只依赖 StylePreset 的 token（色/字号/圆角），换预设就换皮。

import type {
  SpecElement,
  SpecImage,
  SpecParagraph,
  SpecShape,
  SpecText,
} from '@genoffice/pipelines/slides/page-spec'
import type { StylePreset } from './styles'

export const CANVAS_W = 1280
export const CANVAS_H = 720
/** 左右安全边距 */
const M = 80
const W = CANVAS_W - M * 2
const GAP = 32

export interface DesignItem {
  /** 卡片标题 */
  title?: string
  /** 大数字 / 关键值 */
  value?: string
  /** 说明文字（可多段，用 \n 分） */
  note?: string
  /** 纯符号列表项 */
  bullet?: string
}

export interface DesignPage {
  /** 版式名；给了 layout 就忽略 elements（元素由引擎摆） */
  layout?: string
  type?: string
  title?: string
  /** 眉标：章节号 / 分类 / 日期 */
  kicker?: string
  subtitle?: string
  items?: DesignItem[]
  body?: string | string[]
  /** 来源行，放页面底部 */
  footnote?: string
  /** 图片 URL / 本地路径 / data: URL（版式里有图位时用） */
  image?: string
  imageCaption?: string
  /** 页级背景覆盖 */
  background?: string
  /** 页级预设覆盖 */
  style?: string
  /** 低层直摆（和 layout 二选一） */
  elements?: SpecElement[]
}

// ─── 元素构造小工具 ──────────────────────────────────────────────────

function box(x: number, y: number, w: number, h: number) {
  return { x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(h) }
}

function text(
  x: number,
  y: number,
  w: number,
  h: number,
  paragraphs: SpecParagraph[],
  valign?: 'top' | 'middle' | 'bottom',
): SpecText {
  return { type: 'text', ...box(x, y, w, h), paragraphs, ...(valign ? { valign } : {}) }
}

function para(
  content: string,
  o: {
    sizePt: number
    bold?: boolean
    color: string
    font?: string
    align?: 'left' | 'center' | 'right'
    lineSpacingPct?: number
    spaceAfterPt?: number
    bullet?: boolean
  },
): SpecParagraph {
  return {
    runs: [
      {
        text: content,
        sizePt: o.sizePt,
        color: o.color,
        ...(o.bold ? { bold: true } : {}),
        ...(o.font ? { font: o.font } : {}),
      },
    ],
    ...(o.align ? { align: o.align } : {}),
    ...(o.lineSpacingPct ? { lineSpacingPct: o.lineSpacingPct } : {}),
    ...(o.spaceAfterPt !== undefined ? { spaceAfterPt: o.spaceAfterPt } : {}),
    ...(o.bullet ? { bullet: true } : {}),
  }
}

function rect(
  x: number,
  y: number,
  w: number,
  h: number,
  fill: string,
  o: { radius?: number; stroke?: string; widthPt?: number } = {},
): SpecShape {
  return {
    type: 'shape',
    shape: (o.radius ?? 0) > 0 ? 'roundRect' : 'rect',
    ...box(x, y, w, h),
    fill,
    ...(o.stroke ? { stroke: { color: o.stroke, widthPt: o.widthPt ?? 1 } } : {}),
  }
}

function picture(x: number, y: number, w: number, h: number, url: string, caption?: string): SpecElement[] {
  const out: SpecElement[] = [{ type: 'image', ...box(x, y, w, h), url } as SpecImage]
  if (caption)
    out.push(
      text(x, y + h + 12, w, 28, [para(caption, { sizePt: 12, color: '#8A93A3' })]),
    )
  return out
}

function rule(x: number, y: number, w: number, preset: StylePreset): SpecElement[] {
  if (preset.decor === 'none') return []
  if (preset.decor === 'block') return [rect(x, y, Math.min(w, 96), 8, preset.palette.accent)]
  return [rect(x, y, Math.min(w, 220), 3, preset.palette.accent)]
}

function bodyParagraphs(page: DesignPage, preset: StylePreset, color: string): SpecParagraph[] {
  const p = preset.palette
  const lines = Array.isArray(page.body) ? page.body : page.body ? page.body.split('\n') : []
  const out = lines
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line, i, all) =>
      para(line, {
        sizePt: preset.scale.body,
        color,
        lineSpacingPct: 130,
        bullet: true,
        ...(i === all.length - 1 ? {} : { spaceAfterPt: Math.round(preset.scale.body * 0.9) }),
      }),
    )
  for (const item of page.items ?? []) {
    if (item.value || item.title) continue
    if (item.note)
      out.push(para(item.note, { sizePt: preset.scale.body, color, lineSpacingPct: 130, bullet: true }))
  }
  return out.length > 0 ? out : [para(' ', { sizePt: preset.scale.body, color: p.inkOnLight })]
}

/** 卡片：底 + 数字 + 标题 + 说明。 */
function card(
  x: number,
  y: number,
  w: number,
  h: number,
  item: DesignItem,
  preset: StylePreset,
): SpecElement[] {
  const p = preset.palette
  const pad = 28
  const out: SpecElement[] = [
    rect(x, y, w, h, p.card, { radius: preset.radius, stroke: p.border, widthPt: 1 }),
  ]
  let cursor = y + pad
  if (item.value) {
    out.push(
      text(x + pad, cursor, w - pad * 2, Math.round(preset.scale.number * 1.5), [
        para(item.value, { sizePt: preset.scale.number, bold: true, color: p.accent }),
      ]),
    )
    cursor += Math.round(preset.scale.number * 1.5) + 8
  }
  if (item.title) {
    out.push(
      text(x + pad, cursor, w - pad * 2, Math.round(preset.scale.h2 * 1.7), [
        para(item.title, { sizePt: preset.scale.h2, bold: true, color: p.inkOnLight }),
      ]),
    )
    cursor += Math.round(preset.scale.h2 * 1.7) + 8
  }
  if (item.note) {
    out.push(
      text(x + pad, cursor, w - pad * 2, Math.max(48, y + h - pad - cursor), [
        para(item.note, { sizePt: preset.scale.caption + 2, color: p.mutedOnLight, lineSpacingPct: 130 }),
      ]),
    )
  }
  return out
}

/** 标题区：眉标 + 页标题 + 强调线。 */
function heading(page: DesignPage, preset: StylePreset, y = 64): SpecElement[] {
  const p = preset.palette
  const out: SpecElement[] = []
  let cursor = y
  if (page.kicker) {
    out.push(text(M, cursor, W, 28, [para(page.kicker.toUpperCase(), { sizePt: preset.scale.caption + 2, bold: true, color: p.accent })]))
    cursor += 34
  }
  if (page.title) {
    out.push(
      text(M, cursor, W, Math.round(preset.scale.h1 * 1.4), [
        para(page.title, { sizePt: preset.scale.h1, bold: true, color: p.inkOnLight, font: preset.font.title }),
      ]),
    )
    cursor += Math.round(preset.scale.h1 * 1.4) + 10
  }
  out.push(...rule(M, cursor, W, preset))
  return out
}

function footnote(page: DesignPage, preset: StylePreset): SpecElement[] {
  if (!page.footnote) return []
  return [
    text(M, CANVAS_H - 56, W, 28, [
      para(page.footnote, { sizePt: preset.scale.caption, color: preset.palette.mutedOnLight }),
    ]),
  ]
}

// ─── 封面 ────────────────────────────────────────────────────────────

function coverTypographyHero(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const out: SpecElement[] = []
  if (page.kicker)
    out.push(text(M, 120, W, 30, [para(page.kicker.toUpperCase(), { sizePt: preset.scale.caption + 3, bold: true, color: p.accent })]))
  out.push(
    text(M, 180, W, 260, [
      para(page.title ?? '', {
        sizePt: preset.scale.cover,
        bold: true,
        color: p.ink,
        font: preset.font.title,
        lineSpacingPct: 108,
      }),
    ], 'middle'),
  )
  out.push(...rule(M, 462, W, preset))
  if (page.subtitle)
    out.push(
      text(M, 494, Math.min(W, 900), 80, [
        para(page.subtitle, { sizePt: preset.scale.h2, color: p.muted, lineSpacingPct: 130 }),
      ]),
    )
  out.push(...footnote(page, preset))
  return out
}

function coverDarkMinimal(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const out: SpecElement[] = [rect(M, 140, 72, 10, p.accent)]
  out.push(
    text(M, 260, Math.min(W, 1000), 240, [
      para(page.title ?? '', {
        sizePt: preset.scale.cover,
        bold: true,
        color: p.ink,
        font: preset.font.title,
        lineSpacingPct: 104,
      }),
    ], 'bottom'),
  )
  if (page.subtitle)
    out.push(text(M, 526, Math.min(W, 900), 40, [para(page.subtitle, { sizePt: preset.scale.h2, color: p.muted })]))
  if (page.kicker)
    out.push(text(M, 580, W, 28, [para(page.kicker, { sizePt: preset.scale.caption + 2, color: p.muted })]))
  return out
}

function coverSplitColor(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const splitW = Math.round(CANVAS_W * 0.42)
  const out: SpecElement[] = [
    rect(CANVAS_W - splitW, 0, splitW, CANVAS_H, p.accent),
    text(M, 200, CANVAS_W - splitW - M * 1.6, 220, [
      para(page.title ?? '', {
        sizePt: preset.scale.cover - 8,
        bold: true,
        color: p.ink,
        font: preset.font.title,
        lineSpacingPct: 110,
      }),
    ], 'middle'),
  ]
  if (page.subtitle)
    out.push(
      text(M, 452, CANVAS_W - splitW - M * 1.6, 90, [
        para(page.subtitle, { sizePt: preset.scale.h2, color: p.muted, lineSpacingPct: 130 }),
      ]),
    )
  if (page.kicker)
    out.push(text(M, 150, 600, 30, [para(page.kicker.toUpperCase(), { sizePt: preset.scale.caption + 2, bold: true, color: p.accent })]))
  // 色块里放一个反白的关键词，让 40% 的色块有内容
  const key = page.items?.[0]?.value ?? page.items?.[0]?.title
  if (key)
    out.push(
      text(CANVAS_W - splitW + 40, 300, splitW - 80, 140, [
        para(key, { sizePt: preset.scale.number, bold: true, color: p.onAccent, align: 'left' }),
      ], 'middle'),
    )
  return out
}

function coverMagazine(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const out: SpecElement[] = []
  if (page.kicker)
    out.push(text(M, 90, W, 28, [para(page.kicker.toUpperCase(), { sizePt: preset.scale.caption + 2, bold: true, color: p.accent })]))
  out.push(...rule(M, 132, W, preset))
  out.push(
    text(M, 170, W, 300, [
      para(page.title ?? '', {
        sizePt: preset.scale.cover,
        bold: false,
        color: p.ink,
        font: preset.font.title,
        lineSpacingPct: 112,
      }),
    ], 'middle'),
  )
  if (page.subtitle)
    out.push(
      text(M, 500, Math.min(W, 820), 90, [
        para(page.subtitle, { sizePt: preset.scale.h2, color: p.mutedOnLight, lineSpacingPct: 130 }),
      ]),
    )
  out.push(...footnote(page, preset))
  return out
}

function coverSplitImage(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const leftW = Math.round(CANVAS_W * 0.52)
  const out: SpecElement[] = []
  if (page.image) out.push(...picture(leftW, 0, CANVAS_W - leftW, CANVAS_H, page.image, page.imageCaption))
  else out.push(rect(leftW, 0, CANVAS_W - leftW, CANVAS_H, p.accent))
  out.push(
    text(M, 200, leftW - M * 1.8, 220, [
      para(page.title ?? '', {
        sizePt: preset.scale.cover - 10,
        bold: true,
        color: p.ink,
        font: preset.font.title,
        lineSpacingPct: 110,
      }),
    ], 'middle'),
  )
  if (page.subtitle)
    out.push(
      text(M, 452, leftW - M * 1.8, 100, [
        para(page.subtitle, { sizePt: preset.scale.h2, color: p.muted, lineSpacingPct: 130 }),
      ]),
    )
  return out
}

function coverFullImageOverlay(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  if (!page.image) return coverTypographyHero(page, preset)
  return [
    ...picture(0, 0, CANVAS_W, CANVAS_H, page.image, page.imageCaption),
    rect(0, 0, CANVAS_W, CANVAS_H, `${p.bg}D9`),
    text(M, 300, Math.min(W, 960), 220, [
      para(page.title ?? '', {
        sizePt: preset.scale.cover,
        bold: true,
        color: p.ink,
        font: preset.font.title,
        lineSpacingPct: 108,
      }),
    ], 'bottom'),
    ...(page.subtitle
      ? [text(M, 540, Math.min(W, 820), 60, [para(page.subtitle, { sizePt: preset.scale.h2, color: p.ink })])]
      : []),
  ]
}

// ─── 内容页 ──────────────────────────────────────────────────────────

function threeColumnCards(page: DesignPage, preset: StylePreset): SpecElement[] {
  const items = (page.items ?? []).slice(0, 4)
  if (items.length === 0) return heroBigNumber(page, preset)
  const n = items.length
  const w = Math.round((W - GAP * (n - 1)) / n)
  const y = 220
  const h = 300
  const out: SpecElement[] = heading(page, preset)
  items.forEach((item, i) => out.push(...card(M + i * (w + GAP), y, w, h, item, preset)))
  out.push(...footnote(page, preset))
  return out
}

function kpiCardsRow(page: DesignPage, preset: StylePreset): SpecElement[] {
  const items = (page.items ?? []).slice(0, 4)
  if (items.length === 0) return heroBigNumber(page, preset)
  const n = items.length
  const w = Math.round((W - GAP * (n - 1)) / n)
  const out: SpecElement[] = heading(page, preset)
  const y = 236
  const h = 204
  const valueH = Math.round(preset.scale.number * 1.5)
  items.forEach((item, i) => {
    const x = M + i * (w + GAP)
    const p = preset.palette
    out.push(rect(x, y, w, h, p.card, { radius: preset.radius, stroke: p.border, widthPt: 1 }))
    out.push(
      text(x + 24, y + 24, w - 48, valueH, [
        para(item.value ?? item.title ?? '', {
          sizePt: preset.scale.number,
          bold: true,
          color: p.accent,
        }),
      ]),
    )
    out.push(
      text(x + 24, y + 24 + valueH + 4, w - 48, 34, [
        para(item.title ?? '', { sizePt: preset.scale.h2 - 2, color: p.inkOnLight, bold: true }),
      ]),
    )
    if (item.note)
      out.push(
        text(x + 24, y + h - 38, w - 48, 26, [
          para(item.note, { sizePt: preset.scale.caption, color: p.mutedOnLight }),
        ]),
      )
  })
  if (page.body) {
    out.push(
      text(M, 470, W, 140, bodyParagraphs(page, preset, preset.palette.inkOnLight)),
    )
  }
  out.push(...footnote(page, preset))
  return out
}

function heroBigNumber(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const items = page.items ?? []
  const head = items[0]
  const rest = items.slice(1)
  const out: SpecElement[] = heading(page, preset)
  const bigY = 250
  out.push(
    text(M, bigY, 620, 190, [
      para(head?.value ?? head?.title ?? '', {
        sizePt: Math.max(preset.scale.number, 72),
        bold: true,
        color: p.accent,
        font: preset.font.title,
      }),
    ], 'middle'),
  )
  if (head?.title)
    out.push(text(M, bigY + 190, 620, 70, [para(head.title, { sizePt: preset.scale.h2 + 2, color: p.inkOnLight })]))
  if (head?.note)
    out.push(text(M, bigY + 260, 620, 90, [para(head.note, { sizePt: preset.scale.body - 2, color: p.mutedOnLight, lineSpacingPct: 130 })]))
  const list = page.body ?? rest.map((i) => i.note ?? i.title ?? '').filter(Boolean)
  const listPage: DesignPage = { body: list, items: rest.filter((i) => i.bullet) }
  if (list.length > 0 || listPage.items?.length)
    out.push(
      text(760, bigY - 10, CANVAS_W - M - 760, 340, bodyParagraphs(listPage, preset, p.inkOnLight)),
    )
  out.push(...footnote(page, preset))
  return out
}

function twoColumnComparison(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const items = (page.items ?? []).slice(0, 2)
  const out: SpecElement[] = heading(page, preset)
  const w = Math.round((W - GAP) / 2)
  const y = 210
  const h = 340
  items.forEach((item, i) => {
    const x = M + i * (w + GAP)
    const fill = i === 0 ? p.card : p.accentSoft
    const border = i === 0 ? p.border : p.accentSoft
    out.push(rect(x, y, w, h, fill, { radius: preset.radius, stroke: border, widthPt: 1 }))
    // 关键值做成右上角徽标：卡片下半部分整块留给论据，长文案长高也撞不到它。
    const badgeW = item.value ? 150 : 0
    out.push(
      text(x + 32, y + 24, w - 64 - badgeW, 44, [
        para(item.title ?? '', { sizePt: preset.scale.h2 + 2, bold: true, color: p.inkOnLight }),
      ]),
    )
    if (item.value)
      out.push(
        text(x + w - 32 - badgeW, y + 12, badgeW, 64, [
          para(item.value, {
            sizePt: preset.scale.number - 8,
            bold: true,
            color: p.accent,
            align: 'right',
          }),
        ], 'middle'),
      )
    const lines = (item.note ?? '').split('\n').filter((l) => l.trim())
    if (lines.length > 0)
      out.push(
        text(
          x + 32,
          y + 92,
          w - 64,
          h - 92 - 24,
          lines.map((line, k) =>
            para(line, {
              sizePt: preset.scale.body - 1,
              color: p.mutedOnLight,
              bullet: true,
              lineSpacingPct: 135,
              ...(k === lines.length - 1 ? {} : { spaceAfterPt: 10 }),
            }),
          ),
        ),
      )
  })
  out.push(...footnote(page, preset))
  return out
}

function twoByTwoGrid(page: DesignPage, preset: StylePreset): SpecElement[] {
  const items = (page.items ?? []).slice(0, 4)
  const out: SpecElement[] = heading(page, preset)
  const w = Math.round((W - GAP) / 2)
  const h = 190
  items.forEach((item, i) => {
    const x = M + (i % 2) * (w + GAP)
    const y = 200 + Math.floor(i / 2) * (h + 24)
    out.push(...card(x, y, w, h, item, preset))
  })
  out.push(...footnote(page, preset))
  return out
}

function timelineHorizontal(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const items = (page.items ?? []).slice(0, 5)
  const out: SpecElement[] = heading(page, preset)
  const y = 330
  const n = Math.max(items.length, 1)
  // 首末节点的文字左右各占 110px：行程区要内缩，否则最后一列会被画布裁掉。
  const half = 110
  const usable = W - half * 2
  out.push(rect(M + half, y, usable, 2, p.border))
  const step = n === 1 ? 0 : usable / (n - 1)
  items.forEach((item, i) => {
    const cx = Math.round(M + half + i * step)
    out.push(rect(cx - 9, y - 8, 18, 18, p.accent, { radius: 9 }))
    if (item.value)
      out.push(
        text(cx - 100, y - 78, 200, 56, [
          para(item.value, { sizePt: preset.scale.number - 10, bold: true, color: p.accent, align: 'center' }),
        ]),
      )
    if (item.title)
      out.push(
        text(cx - 110, y + 30, 220, 40, [
          para(item.title, { sizePt: preset.scale.h2 - 2, bold: true, color: p.inkOnLight, align: 'center' }),
        ]),
      )
    if (item.note)
      out.push(
        text(cx - 110, y + 74, 220, 100, [
          para(item.note, { sizePt: preset.scale.caption + 1, color: p.mutedOnLight, align: 'center', lineSpacingPct: 130 }),
        ]),
      )
  })
  if (page.body)
    out.push(text(M, 520, W, 100, bodyParagraphs(page, preset, p.inkOnLight)))
  out.push(...footnote(page, preset))
  return out
}

function titleAndBullets(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const out: SpecElement[] = heading(page, preset)
  out.push(text(M, 220, Math.min(W, 980), 400, bodyParagraphs(page, preset, p.inkOnLight)))
  out.push(...footnote(page, preset))
  return out
}

function leftTextRightImage(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const out: SpecElement[] = heading(page, preset)
  const imgW = 540
  const textW = W - imgW - GAP
  const y = 210
  if (page.image) out.push(...picture(CANVAS_W - M - imgW, y, imgW, 420, page.image, page.imageCaption))
  if (page.body)
    out.push(text(M, y + 8, textW, 380, bodyParagraphs(page, preset, p.inkOnLight)))
  const item = page.items?.[0]
  if (item?.value)
    out.push(text(M, 520, textW, 60, [para(item.value, { sizePt: preset.scale.number, bold: true, color: p.accent })]))
  out.push(...footnote(page, preset))
  return out
}

function fullImageTextOverlay(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  if (!page.image) return leftTextRightImage(page, preset)
  const out: SpecElement[] = [...picture(0, 0, CANVAS_W, CANVAS_H, page.image, page.imageCaption)]
  out.push(rect(0, CANVAS_H - 320, CANVAS_W, 320, `${p.bg}E6`))
  out.push(
    text(M, CANVAS_H - 290, Math.min(W, 900), 120, [
      para(page.title ?? '', {
        sizePt: preset.scale.h1 + 6,
        bold: true,
        color: p.ink,
        font: preset.font.title,
        lineSpacingPct: 112,
      }),
    ]),
  )
  if (page.body || page.subtitle)
    out.push(
      text(
        M,
        CANVAS_H - 160,
        Math.min(W, 900),
        90,
        page.body
          ? bodyParagraphs({ body: page.body }, preset, p.ink)
          : [para(page.subtitle ?? '', { sizePt: preset.scale.h2, color: p.muted })],
      ),
    )
  return out
}

// ─── 结尾 ────────────────────────────────────────────────────────────

function closingThankYou(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const out: SpecElement[] = [
    text(0, 260, CANVAS_W, 160, [
      para(page.title ?? '谢谢', {
        sizePt: preset.scale.cover,
        bold: true,
        color: p.ink,
        font: preset.font.title,
        align: 'center',
      }),
    ], 'middle'),
  ]
  if (page.subtitle)
    out.push(
      text(0, 440, CANVAS_W, 60, [para(page.subtitle, { sizePt: preset.scale.h2, color: p.muted, align: 'center' })]),
    )
  if (page.footnote)
    out.push(text(0, 540, CANVAS_W, 40, [para(page.footnote, { sizePt: preset.scale.caption, color: p.muted, align: 'center' })]))
  return out
}

function closingCta(page: DesignPage, preset: StylePreset): SpecElement[] {
  const p = preset.palette
  const out: SpecElement[] = []
  if (page.title)
    out.push(
      text(M, 240, Math.min(W, 900), 200, [
        para(page.title, {
          sizePt: preset.scale.cover - 8,
          bold: true,
          color: p.ink,
          font: preset.font.title,
          lineSpacingPct: 110,
        }),
      ], 'middle'),
    )
  if (page.subtitle)
    out.push(text(M, 460, Math.min(W, 820), 80, [para(page.subtitle, { sizePt: preset.scale.h2, color: p.muted, lineSpacingPct: 130 })]))
  const cta = page.items?.[0]?.value ?? page.items?.[0]?.title
  if (cta) {
    const width = Math.max(220, Math.min(420, cta.length * (preset.scale.h2 + 6) + 96))
    out.push(rect(M, 570, width, 64, p.accent, { radius: preset.radius }))
    out.push(
      text(M, 570, width, 64, [
        para(cta, { sizePt: preset.scale.h2, bold: true, color: p.onAccent, align: 'center' }),
      ], 'middle'),
    )
  }
  if (page.footnote)
    out.push(text(M + (cta ? Math.max(220, Math.min(420, cta.length * (preset.scale.h2 + 6) + 96)) + 24 : 0), 570, 520, 64, [
      para(page.footnote, { sizePt: preset.scale.caption, color: p.muted }),
    ], 'middle'))
  return out
}

// ─── 注册表 ──────────────────────────────────────────────────────────

type Builder = (page: DesignPage, preset: StylePreset) => SpecElement[]

export const LAYOUTS: Readonly<Record<string, Builder>> = {
  cover_typography_hero: coverTypographyHero,
  cover_dark_minimal: coverDarkMinimal,
  cover_split_color: coverSplitColor,
  cover_magazine: coverMagazine,
  cover_split_image: coverSplitImage,
  cover_full_image_overlay: coverFullImageOverlay,
  three_column_cards: threeColumnCards,
  kpi_cards_row: kpiCardsRow,
  hero_big_number: heroBigNumber,
  two_column_comparison: twoColumnComparison,
  two_by_two_grid: twoByTwoGrid,
  timeline_horizontal: timelineHorizontal,
  title_and_bullets: titleAndBullets,
  left_text_right_image: leftTextRightImage,
  full_image_text_overlay: fullImageTextOverlay,
  closing_thank_you: closingThankYou,
  closing_cta: closingCta,
}

export const LAYOUT_NAMES = Object.keys(LAYOUTS)

/** 需要图片的版式：没给 image 时回退到不带图的版式，不画灰块。 */
const NEEDS_IMAGE = new Set([
  'cover_split_image',
  'cover_full_image_overlay',
  'left_text_right_image',
  'full_image_text_overlay',
])

const FALLBACK: Readonly<Record<string, string>> = {
  cover_split_image: 'cover_split_color',
  cover_full_image_overlay: 'cover_typography_hero',
  left_text_right_image: 'hero_big_number',
  full_image_text_overlay: 'hero_big_number',
}

/** 按版式把内容摆成元素；未知版式 → null（交给调用方报错）。 */
export function buildLayoutElements(
  page: DesignPage,
  preset: StylePreset,
): { elements: SpecElement[]; layout: string; fallback?: string } | null {
  const requested = page.layout ?? ''
  let name = requested
  let fallback: string | undefined
  if (NEEDS_IMAGE.has(name) && !page.image) {
    name = FALLBACK[name] ?? 'hero_big_number'
    fallback = name
  }
  const builder = LAYOUTS[name]
  if (!builder) return null
  const elements = builder(page, preset)
  return {
    elements,
    layout: name,
    ...(fallback ? { fallback: requested } : {}),
  }
}
