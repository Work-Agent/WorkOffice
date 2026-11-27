// ─── 设计稿展开：高层页面（版式 + 内容）→ 低层 spec（px 元素） ──────────
//
// 模型写"这一页是什么版式、里面有什么"，引擎按预设的 token 把元素摆出来。
// 这一步做完，页面就是普通的 `{background, elements}` spec，后面走同一条
// parseDeckSpec + buildDeckPptx，校验和构建都不需要知道"高层"这回事。

import type { SpecElement } from '@genoffice/pipelines/slides/page-spec'
import { buildLayoutElements, LAYOUT_NAMES, type DesignItem, type DesignPage } from './layouts'
import {
  DEFAULT_PRESET,
  findPreset,
  presetColors,
  STYLE_PRESETS,
  type StylePreset,
} from './styles'

export interface StyleResolution {
  preset: StylePreset | null
  /** 自定义 style.md 文本（只用来做调色板检查） */
  markdown?: string
  notes: string[]
}

const COLOR_KEYS = [
  'bg',
  'bgContent',
  'ink',
  'inkOnLight',
  'muted',
  'mutedOnLight',
  'accent',
  'accentSoft',
  'card',
  'border',
  'onAccent',
] as const

function presetFromColors(name: string, colors: Record<string, unknown>): StylePreset | null {
  const found = findPreset(name) ?? findPreset(DEFAULT_PRESET)
  if (!found) return null
  const palette = { ...found.palette }
  for (const key of COLOR_KEYS) {
    const value = colors[key]
    if (typeof value === 'string' && /^#?[0-9a-f]{6}([0-9a-f]{2})?$/i.test(value.trim())) {
      const hex = value.trim()
      palette[key] = hex.startsWith('#') ? hex.toUpperCase() : `#${hex.toUpperCase()}`
    }
  }
  return { ...found, id: `${found.id}+custom`, label: `${found.label}（自定义色）`, palette }
}

/**
 * 解析 `style` 参数：
 *   - 预设 id / 中文名 → 用预设；
 *   - 对象 `{preset}` / `{colors}` → 预设或预设+自定义色；
 *   - 其他字符串 → 当作 style.md 文本（只做调色板检查）；
 *   - 不给 → 有高层页面就用默认预设，否则不设风格。
 */
export function resolveStyle(styleInput: unknown, hasHighLevel: boolean): StyleResolution {
  const notes: string[] = []
  if (styleInput === undefined || styleInput === null || styleInput === '') {
    return { preset: hasHighLevel ? findPreset(DEFAULT_PRESET) : null, notes }
  }
  if (typeof styleInput === 'object') {
    const obj = styleInput as Record<string, unknown>
    const name = typeof obj['preset'] === 'string' ? obj['preset'] : typeof obj['name'] === 'string' ? obj['name'] : DEFAULT_PRESET
    const colors = obj['colors']
    const preset =
      colors && typeof colors === 'object'
        ? presetFromColors(name, colors as Record<string, unknown>)
        : findPreset(name)
    if (!preset) notes.push(`未知预设「${name}」，可用：${STYLE_PRESETS.map((p) => p.id).join(' / ')}`)
    return { preset: preset ?? (hasHighLevel ? findPreset(DEFAULT_PRESET) : null), notes }
  }
  const preset = findPreset(styleInput)
  if (preset) return { preset, notes }
  const text = String(styleInput)
  if (hasHighLevel)
    notes.push(
      '自定义 style.md 不会改版式配色：高层页面按默认预设上色。要自定义颜色请传 { colors: { accent: "#…", bg: "#…" } }，或改用低层 elements 自己写颜色。',
    )
  return { preset: hasHighLevel ? findPreset(DEFAULT_PRESET) : null, markdown: text, notes }
}

function autoLayout(page: DesignPage, index: number, total: number): string {
  const type = page.type ?? ''
  const items = page.items ?? []
  if (type === 'cover' || (index === 0 && !type)) return 'cover_typography_hero'
  if (type === 'closing' || (total >= 3 && index === total - 1 && !type))
    return 'closing_thank_you'
  if (type === 'data') {
    if (items.length >= 3) return 'kpi_cards_row'
    if (items.length === 2) return 'two_by_two_grid'
    return 'hero_big_number'
  }
  if (items.length >= 3) return 'three_column_cards'
  if (items.length === 2) return 'two_column_comparison'
  if (items.length === 1 && (items[0]?.value ?? '').length > 0) return 'hero_big_number'
  if (page.image && page.body) return 'left_text_right_image'
  return 'title_and_bullets'
}

export interface ExpandedDeck {
  /** 低层页面 spec（`{title,type,layout,background,elements}`） */
  pages: Array<Record<string, unknown>>
  /** 展开时用的预设（调色板检查用） */
  preset: StylePreset | null
  /** style.md 文本（没给预设时做调色板检查） */
  markdown?: string
  /** 有页面走了高层版式 */
  highLevel: boolean
  notes: string[]
}

/** 高层页面 → 低层 spec。低层页面（自带 elements）原样透传。 */
export function expandDeck(rawPages: unknown[], styleInput: unknown): ExpandedDeck {
  const pagesIn = rawPages as DesignPage[]
  const highLevel = pagesIn.some((page) => page && typeof page === 'object' && !Array.isArray(page.elements))
  const style = resolveStyle(styleInput, highLevel)
  const notes = [...style.notes]
  const pages: Array<Record<string, unknown>> = []

  for (const [index, rawPage] of pagesIn.entries()) {
    const page = (rawPage && typeof rawPage === 'object' ? rawPage : {}) as DesignPage
    if (Array.isArray(page.elements)) {
      pages.push({ ...page, elements: page.elements })
      continue
    }
    const requested = page.layout ?? autoLayout(page, index, pagesIn.length)
    const preset = (page.style ? findPreset(page.style) : null) ?? style.preset
    if (!preset) {
      notes.push(`p${index}: 没有可用预设，无法展开版式 ${requested}`)
      pages.push({ ...page, elements: [] })
      continue
    }
    const built = buildLayoutElements({ ...page, layout: requested }, preset)
    if (!built) {
      const fallbackLayout = autoLayout(page, index, pagesIn.length)
      notes.push(
        `p${index}: 未知版式「${requested}」，按「${fallbackLayout}」摆。可用：${LAYOUT_NAMES.join(' ')}`,
      )
      const retry = buildLayoutElements({ ...page, layout: fallbackLayout }, preset)
      pages.push({
        ...page,
        title: page.title,
        type: page.type,
        layout: requested,
        background: page.background ?? backgroundFor(fallbackLayout, preset),
        elements: retry?.elements ?? [],
      })
      continue
    }
    if (built.fallback)
      notes.push(
        `p${index}: 版式「${built.fallback}」需要图片但没给 image，已改用「${built.layout}」`,
      )
    pages.push({
      ...page,
      title: page.title,
      type: page.type,
      layout: requested,
      background: page.background ?? backgroundFor(built.layout, preset),
      elements: built.elements,
    })
  }

  return {
    pages,
    preset: style.preset,
    ...(style.markdown !== undefined ? { markdown: style.markdown } : {}),
    highLevel,
    notes,
  }
}

function backgroundFor(layout: string, preset: StylePreset): string {
  const dark =
    layout.startsWith('cover_') || layout.startsWith('closing_') || layout === 'full_image_text_overlay'
  // 极简黑白 / 杂志暖橙这两套的"深色页"其实是浅底，按预设自己的 bg 走就行。
  return dark ? preset.palette.bg : preset.palette.bgContent
}

/** 展开后用于调色板检查的颜色集合。 */
export function paletteOf(deck: ExpandedDeck): Set<string> | null {
  if (deck.preset) return new Set(presetColors(deck.preset))
  return null
}

export type { DesignItem, DesignPage, SpecElement }
