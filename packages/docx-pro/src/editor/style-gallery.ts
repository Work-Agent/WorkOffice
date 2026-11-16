import type { Node as PmNode } from '@tiptap/pm/model'
import type { CSSProperties } from 'react'
import type { DocDefaults, StyleInfo } from '@genoffice/docx-engine'
import type { StringKey, TFunc } from './i18n/locale'
import { cssFontFamily } from './line-metrics'

export type StyleMap = ReadonlyMap<string, StyleInfo>

/** What the gallery lists when no document styles are known (tests, no open document) */
export const FALLBACK_STYLES: StyleMap = new Map<string, StyleInfo>(
  [
    { styleId: 'Normal', name: 'Normal', type: 'paragraph', isDefault: true, qFormat: true },
    {
      styleId: 'Heading1',
      name: 'heading 1',
      type: 'paragraph',
      headingLevel: 1,
      qFormat: true,
      uiPriority: 9,
    },
    {
      styleId: 'Heading2',
      name: 'heading 2',
      type: 'paragraph',
      headingLevel: 2,
      qFormat: true,
      uiPriority: 9,
    },
    {
      styleId: 'Heading3',
      name: 'heading 3',
      type: 'paragraph',
      headingLevel: 3,
      qFormat: true,
      uiPriority: 9,
    },
  ].map((s) => [s.styleId, s as StyleInfo]),
)

const byPriorityThenName = (a: StyleInfo, b: StyleInfo): number =>
  (a.uiPriority ?? 0) - (b.uiPriority ?? 0) ||
  a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' })

const listable = (info: StyleInfo): boolean =>
  (info.type === 'paragraph' || info.type === 'character') && !info.linkedCharShell

interface CanonicalStyleSpec {
  id: string
  name: RegExp
  level?: number
}

/** Minimum standard paragraph styles every .docx must expose. */
const BASELINE_QUICK_STYLES: CanonicalStyleSpec[] = [
  { id: 'Normal', name: /^normal$/i },
  { id: 'Heading1', name: /^heading\s*1$/i, level: 1 },
  { id: 'Heading2', name: /^heading\s*2$/i, level: 2 },
  { id: 'Heading3', name: /^heading\s*3$/i, level: 3 },
  { id: 'Heading4', name: /^heading\s*4$/i, level: 4 },
]

const isBaselineSemantic = (info: StyleInfo): boolean =>
  info.styleId === 'Normal' ||
  (info.type === 'paragraph' &&
    ((info.headingLevel !== undefined && info.headingLevel >= 1 && info.headingLevel <= 4) ||
      /^heading\s*[1-4]$/i.test(info.name)))

/**
 * The five styles that must exist in every document, in a fixed order.
 * Everything else in the gallery comes from the document itself.
 */
export function canonicalStyleEntries(styles: StyleMap): StyleInfo[] {
  const out: StyleInfo[] = []
  const seen = new Set<string>()
  const push = (info: StyleInfo | undefined): void => {
    if (!info || seen.has(info.styleId)) return
    seen.add(info.styleId)
    out.push(info)
  }

  for (const spec of BASELINE_QUICK_STYLES) {
    let info: StyleInfo | undefined
    if (spec.level !== undefined) {
      const id = headingStyleId(styles, spec.level)
      if (id) info = styles.get(id)
    } else {
      info = styles.get(spec.id)
      if (!info)
        info = [...styles.values()].find(
          (candidate) => candidate.type === 'paragraph' && candidate.isDefault,
        )
    }
    push(info)
  }
  return out
}

export function quickStyleEntries(styles: StyleMap, used: ReadonlySet<string>): StyleInfo[] {
  const baseline = canonicalStyleEntries(styles)
  const baseIds = new Set(baseline.map((info) => info.styleId))
  const extras = [...styles.values()].filter(
    (info) =>
      listable(info) &&
      !baseIds.has(info.styleId) &&
      !isBaselineSemantic(info) &&
      (info.qFormat ||
        (info.type === 'paragraph' && info.isDefault) ||
        (info.semiHidden && info.unhideWhenUsed && used.has(info.styleId))),
  )
  extras.sort(byPriorityThenName)
  return [...baseline, ...extras]
}

/** Styles pane "All styles": every paragraph and character style, same order */
export function allStyleEntries(styles: StyleMap): StyleInfo[] {
  return [...styles.values()].filter(listable).sort(byPriorityThenName)
}

/** The paragraph style Word applies to unstyled paragraphs (w:default), by id */
export function defaultParagraphStyleId(styles: StyleMap | undefined): string | undefined {
  if (!styles) return undefined
  for (const info of styles.values())
    if (info.type === 'paragraph' && info.isDefault) return info.styleId
  return undefined
}

/** The document's own "heading N" style (its id varies: Heading1, "1", localized ids) */
export function headingStyleId(styles: StyleMap | undefined, level: number): string | undefined {
  if (!styles) return undefined
  let fallback: string | undefined
  for (const info of styles.values()) {
    if (info.type !== 'paragraph' || info.headingLevel !== level || info.headingLevelInherited)
      continue
    if (/^heading\s*[1-9]$/i.test(info.name) || info.styleId === `Heading${level}`)
      return info.styleId
    fallback ??= info.styleId
  }
  return fallback
}

/**
 * Style ids the document body uses, for the gallery's unhide-when-used rule.
 * Only the candidates that can change the gallery are looked for (semiHidden
 * quick styles), and runs are scanned only when one of them is a character style.
 */
export function collectUsedStyleIds(doc: PmNode | null, styles: StyleMap | undefined): Set<string> {
  const used = new Set<string>()
  if (!doc || !styles) return used
  const candidates = new Set<string>()
  let needRuns = false
  for (const info of styles.values()) {
    if (info.qFormat && info.semiHidden && info.unhideWhenUsed && listable(info)) {
      candidates.add(info.styleId)
      if (info.type === 'character') needRuns = true
    }
  }
  if (candidates.size === 0) return used
  const headingIds = new Map<number, string | undefined>()
  doc.descendants((node) => {
    if (!node.isTextblock) return true
    let id = typeof node.attrs.styleId === 'string' ? node.attrs.styleId : null
    if (!id && node.type.name === 'docHeading') {
      const level = Number(node.attrs.level) || 1
      if (!headingIds.has(level)) headingIds.set(level, headingStyleId(styles, level))
      id = headingIds.get(level) ?? null
    }
    if (id && candidates.has(id)) used.add(id)
    if (needRuns) {
      node.forEach((child) => {
        for (const m of child.marks) {
          const sid = m.type.name === 'docTextStyle' ? m.attrs.styleId : null
          if (typeof sid === 'string' && candidates.has(sid)) used.add(sid)
        }
      })
    }
    return false
  })
  return used
}

/** gallery / pane highlight: the character style at the caret wins, else the paragraph's style */
export function activeStyleKey(
  fs: { charStyleId: string | null; paraStyleId: string | null; headingLevel: number | null },
  styles: StyleMap | undefined,
): string {
  if (fs.charStyleId) return `char:${fs.charStyleId}`
  const paraId =
    fs.paraStyleId ??
    (fs.headingLevel !== null
      ? (headingStyleId(styles, fs.headingLevel) ?? `Heading${fs.headingLevel}`)
      : (defaultParagraphStyleId(styles) ?? 'Normal'))
  return `para:${paraId}`
}

export const styleKeyOf = (info: StyleInfo): string =>
  `${info.type === 'character' ? 'char' : 'para'}:${info.styleId}`

/** Word's built-in style names, localized like Word does in its gallery */
const BUILT_IN_LABELS: Record<string, StringKey> = {
  normal: 'ribbonStyleNormal',
  'no spacing': 'ribbonStyleNoSpacing',
  'heading 1': 'ribbonStyleHeading1',
  'heading 2': 'ribbonStyleHeading2',
  'heading 3': 'ribbonStyleHeading3',
  title: 'ribbonStyleTitle',
  subtitle: 'ribbonStyleSubtitle',
  'subtle emphasis': 'ribbonStyleSubtleEmphasis',
  emphasis: 'ribbonStyleEmphasis',
  'intense emphasis': 'ribbonStyleIntenseEmphasis',
  strong: 'ribbonStyleStrong',
  quote: 'ribbonStyleQuote',
  'intense quote': 'ribbonStyleIntenseQuote',
  'subtle reference': 'ribbonStyleSubtleReference',
  'intense reference': 'ribbonStyleIntenseReference',
  'book title': 'ribbonStyleBookTitle',
  'list paragraph': 'ribbonStyleListParagraph',
}

export function styleLabel(info: Pick<StyleInfo, 'name'>, t: TFunc): string {
  const name = info.name.trim()
  const key = BUILT_IN_LABELS[name.toLowerCase()]
  if (key) return t(key)
  const heading = /^heading\s*([1-9])$/i.exec(name)
  if (heading) return t('ribbonStyleHeadingN', { n: heading[1] })
  return name.charAt(0).toUpperCase() + name.slice(1)
}

/** card preview: the style's resolved face, weight, slant, color and caps at gallery scale */
export function stylePreviewCss(info: StyleInfo, docDefaults?: DocDefaults): CSSProperties {
  const d = info.display ?? {}
  const css: CSSProperties = {}
  const face = d.fontAscii ?? d.font
  if (face) css.fontFamily = cssFontFamily(face)
  if (info.type === 'paragraph') {
    const pt = (d.sizeHalfPoints ?? docDefaults?.sizeHalfPoints ?? 22) / 2
    css.fontSize = `${Math.min(20, Math.max(11, Math.round((pt * 15) / 11)))}px`
  }
  if (d.bold ?? docDefaults?.bold) css.fontWeight = 700
  if (d.italic ?? docDefaults?.italic) css.fontStyle = 'italic'
  const deco = [d.underline ? 'underline' : '', d.strike ? 'line-through' : ''].filter(Boolean)
  if (deco.length) css.textDecoration = deco.join(' ')
  if (d.caps === 'small') css.fontVariant = 'small-caps'
  else if (d.caps === 'all') css.textTransform = 'uppercase'
  const color = d.color ?? docDefaults?.color
  if (color && color !== 'auto') css.color = `#${color}`
  return css
}
