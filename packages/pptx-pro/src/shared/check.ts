// ─── 设计稿检查（页面与执行点共用） ───────────────────────────────────
//
// pptx_check 的全部逻辑都在这里，四个投影：
//   1. outline：核心 hook / 每页 type+layout / 占位符 / 版式重复；
//   2. pages：逐页解析页面 spec，并和 outline 条目、style.md 调色板对账；
//   3. audit：几何审计（越界 / 文字溢出 / 重叠 / 图片变形），每条带 setTransform；
//   4. skeleton：从现有 deck 抽出可复用的版面骨架。
//
// 只依赖 `OpenedPptx` + pptx-render / pipelines 的纯函数，所以同一份代码能跑在
// 两个地方：执行点（磁盘字节）与页面（活的 OpenedPptx，含未保存修改）。

import type { OpenedPptx } from '@genoffice/pptx-engine'
import { elementDurableId, slideDurableId } from '@genoffice/pptx-engine/identity'
import { auditSlideFindings, type AuditFinding } from '@genoffice/pipelines/slides/layout-audit'
import { parseOutline, type OutlineIssue } from '@genoffice/pipelines/slides/outline'
import { parsePageSpecObject } from '@genoffice/pipelines/slides/page-spec'
import {
  checkPageAgainstOutline,
  offPaletteColors,
  stylePalette,
} from '@genoffice/pipelines/slides/stage-check'
import {
  buildRenderSlide,
  EMU_PER_PX_96,
  imageSizeFromBytes,
  type RenderSlide,
} from '@genoffice/pptx-render'
import { coded } from '../engine/abi'
import { expandDeck } from './design'
import { presetColors, presetsGuide } from './styles'
import {
  extractLayoutSkeleton,
  formatSkeletonForPrompt,
  type LayoutSkeleton,
  type SkeletonRole,
} from './skeleton'

/** 设计稿画布；与 `genoffice create --spec` 一致。 */
export const SPEC_CANVAS_W = 1280
export const SPEC_CANVAS_H = 720

const MAX_REPORT_LINES = 200

export interface CheckParams {
  outline?: unknown
  pages?: unknown
  /** 预设 id / 中文名 / `{preset}` / `{colors}` / style.md 文本 */
  style?: unknown
  /** 列出可用风格预设 */
  styles?: unknown
  audit?: unknown
  skeleton?: unknown
  slide?: unknown
}

/** 接受数组、`{pages:[…]}` 或单页对象；与 `genoffice create --spec` 同语义。 */
export function asPagesInput(raw: unknown): { pages: unknown[] } | null {
  if (Array.isArray(raw)) return { pages: raw }
  if (raw && typeof raw === 'object' && Array.isArray((raw as { pages?: unknown }).pages))
    return raw as { pages: unknown[] }
  if (raw && typeof raw === 'object' && 'elements' in (raw as object)) return { pages: [raw] }
  return null
}

function clip(lines: string[]): string[] {
  return lines.length <= MAX_REPORT_LINES
    ? lines
    : [...lines.slice(0, MAX_REPORT_LINES), `…（还有 ${lines.length - MAX_REPORT_LINES} 行）`]
}

function outlineOf(raw: unknown) {
  const parsed = parseOutline(typeof raw === 'string' ? raw : JSON.stringify(raw))
  if (!parsed.ok) throw coded('command.invalid', `outline 解析失败：${parsed.error}`)
  return parsed
}

function formatOutlineIssues(issues: OutlineIssue[], pages: number): string[] {
  if (issues.length === 0) return [`outline 检查通过：${pages} 页，无错误 / 警告。`]
  const errors = issues.filter((i) => i.level === 'error').length
  const lines = [`outline 检查：${pages} 页 · ${errors} error / ${issues.length - errors} warning`]
  for (const issue of issues) {
    const where = issue.page === undefined ? 'deck' : `p${issue.page}`
    lines.push(`${issue.level === 'error' ? 'error' : 'warning'} ${where}: ${issue.message}`)
  }
  if (errors > 0) lines.push('error 必须改完再写页面；warning 是建议。')
  return lines
}

/**
 * 逐页解析页面 spec，再和 outline 条目、调色板对账。
 * 高层页面（版式 + 内容）先按同一套展开逻辑摆成低层 spec，再来检查——
 * 检查的是"真正会落盘的那一页"。
 */
function pagesBlock(raw: unknown, outlineRaw: unknown, styleInput: unknown): string[] {
  const input = asPagesInput(raw)
  if (!input) throw coded('command.invalid', 'pages 必须是页面数组、{pages:[…]} 或单个页面对象')
  if (input.pages.length === 0) throw coded('command.invalid', 'pages 是空的')

  const deck = expandDeck(input.pages, styleInput)
  const rawPages = deck.pages
  const outline = outlineRaw === undefined ? null : outlineOf(outlineRaw)
  const entries = outline?.outline.pages ?? null
  const palette = deck.preset
    ? new Set(presetColors(deck.preset))
    : deck.markdown
      ? stylePalette(deck.markdown)
      : null

  const lines: string[] = [
    `pages 检查：${rawPages.length} 页${entries ? ` · outline ${entries.length} 页` : ''}${
      deck.preset ? ` · 风格 ${deck.preset.label}` : palette && palette.size > 0 ? ' · style 调色板' : ''
    }`,
    ...deck.notes,
  ]
  if (entries && entries.length !== rawPages.length)
    lines.push(`error deck: ${rawPages.length} 份页面 spec 对 ${entries.length} 条 outline —— 页数必须一致`)

  for (const [index, rawPage] of rawPages.entries()) {
    const parsed = parsePageSpecObject(rawPage, SPEC_CANVAS_W, SPEC_CANVAS_H, { localImages: true })
    if (!parsed.ok) {
      lines.push(`p${index} error: 丢弃 — ${parsed.error}`)
      continue
    }
    const issues = [...parsed.warnings]
    if (entries) {
      const entry = entries[index]
      if (!entry) lines.push(`p${index} error: outline 里没有这一页`)
      else
        for (const issue of checkPageAgainstOutline(rawPage, parsed.spec, entry, index))
          issues.push(`${issue.level === 'error' ? 'error' : 'warning'}: ${issue.message}`)
    }
    if (palette && palette.size > 0) {
      const off = offPaletteColors(parsed.spec, palette)
      if (off.length > 0)
        issues.push(
          `warning: 这些颜色不在${deck.preset ? '预设' : ' style.md '}的调色板里 → ${off.join(' ')}`,
        )
    }
    if (issues.length === 0) lines.push(`p${index} ok`)
    else for (const issue of issues) lines.push(`p${index} ${issue}`)
  }
  return lines
}

// ─── 几何审计 ────────────────────────────────────────────────────────

export interface SlideAudit {
  /** 0-based 页序 */
  slide: number
  /** durable slide id（`s_*`），与 pptx_apply 的 target.slide 一致 */
  id: string
  findings: AuditFinding[]
}

/** 按原生像素尺寸排版整份 deck（骨架抽取 / 审计共用）。 */
export function renderSlides(opened: OpenedPptx): RenderSlide[] {
  const fitWidthPx = opened.deck.size.cx / EMU_PER_PX_96
  return opened.deck.slides.map((slide, index) =>
    buildRenderSlide(slide, opened.deck.size, { fitWidthPx, slideNo: index + 1 }),
  )
}

export function auditOpened(opened: OpenedPptx, only?: number): SlideAudit[] {
  const out: SlideAudit[] = []
  const sizes = new Map<string, { w: number; h: number } | null>()
  for (const [index, slide] of opened.deck.slides.entries()) {
    if (only !== undefined && only !== index) continue
    const ids = new Map(slide.elements.map((el) => [el.id, elementDurableId(el) ?? el.id]))
    const pictures = new Map<string, string>()
    for (const el of slide.elements) {
      if (el.type === 'picture' && !el.tile && el.mediaRef) pictures.set(el.id, el.mediaRef)
    }
    const pictureSize = (sourceId: string) => {
      const ref = pictures.get(sourceId)
      if (!ref) return undefined
      let cached = sizes.get(ref)
      if (cached === undefined) {
        const bytes = opened.archive.readBytes(ref)
        const size = bytes ? imageSizeFromBytes(bytes) : undefined
        cached = size ? { w: size.w, h: size.h } : null
        sizes.set(ref, cached)
      }
      return cached ?? undefined
    }
    const rendered = buildRenderSlide(slide, opened.deck.size, {
      fitWidthPx: opened.deck.size.cx / EMU_PER_PX_96,
      slideNo: index + 1,
    })
    const slideId = slideDurableId(slide)
    const findings = auditSlideFindings(rendered, (id) => ids.get(id) ?? id, { pictureSize }).map(
      (finding) =>
        finding.suggest
          ? {
              ...finding,
              suggest: {
                ...finding.suggest,
                target: { slide: slideId, el: finding.suggest.target.el },
              },
            }
          : finding,
    )
    out.push({ slide: index, id: slideId, findings })
  }
  return out
}

export function formatAuditReport(report: SlideAudit[], totalPages: number): string {
  const count = report.reduce((sum, page) => sum + page.findings.length, 0)
  const lines = [
    count === 0
      ? `几何审计通过：${report.length}/${totalPages} 页，无越界 / 文字溢出 / 重叠 / 图片变形。`
      : `几何审计：${report.length}/${totalPages} 页，${count} 个问题。`,
  ]
  let hasOp = false
  for (const page of report) {
    for (const finding of page.findings) {
      lines.push(
        `[${page.slide}] ${page.id} ${finding.el} ${finding.level} ${finding.code}: ${finding.message}`,
      )
      if (finding.suggest) {
        hasOp = true
        lines.push(
          `    → ${JSON.stringify({
            op: 'setTransform',
            target: finding.suggest.target,
            box: finding.suggest.box,
            ...(finding.suggest.rotDeg !== undefined ? { rotDeg: finding.suggest.rotDeg } : {}),
          })}`,
        )
      } else if (finding.code === 'text_overflow' && finding.overflowPx) {
        // 文字溢出没有固定几何解：给一条“只加高”的完整 setTransform（box 四个数都要给）。
        const emu = (px: number) => Math.round(px * 9525)
        const b = finding.box
        lines.push(
          `    → 可以加高：${JSON.stringify({
            op: 'setTransform',
            target: { slide: page.id, el: finding.el },
            box: { x: emu(b.x), y: emu(b.y), cx: emu(b.w), cy: emu(b.h + finding.overflowPx) },
          })}（或减小字号 / 删文案）`,
        )
      } else if (finding.code === 'text_overflow_width' && finding.overflowPx) {
        lines.push('    → 把框加宽或减小字号（wrap 已开，宽了就不再溢出）')
      }
    }
  }
  if (hasOp) lines.push('修法：把 → 后面的 op 原样放进 pptx_apply（可一次提交多条）。')
  return lines.join('\n')
}

// ─── 版面骨架 ────────────────────────────────────────────────────────

function skeletonOf(opened: OpenedPptx): LayoutSkeleton | null {
  return extractLayoutSkeleton(renderSlides(opened))
}

/**
 * 版面骨架：从当前 deck 抽出来的"骨架"（背景 / 标题框 / 品牌图槽 / 装饰形状）。
 * 抄进 style.md 后下一次按它生成，标题和 logo 就不会逐页漂移。
 */
export function skeletonText(opened: OpenedPptx): string {
  const skeleton = skeletonOf(opened)
  if (!skeleton) return '这个 deck 没有可复用的版面骨架（页数太少或每页都是自由版式）。'
  const roles: SkeletonRole[] = ['cover', 'content', 'closing']
  const blocks = roles
    .map((role) => formatSkeletonForPrompt(skeleton, role))
    .filter((text) => text.length > 0)
  const header =
    `版面骨架（canvas ${skeleton.canvas.w}×${skeleton.canvas.h} px）——把它抄进 style.md，` +
    '下一次按这些固定盒位生成，标题/品牌图/装饰就不会逐页漂移。'
  return [header, ...blocks].join('\n\n')
}

// ─── 统一入口 ────────────────────────────────────────────────────────

/**
 * `loadOpened` 只在需要看文档（audit / skeleton）时才调用：只检查 outline 或
 * pages 时不必打开文件。
 */
export async function runCheck(
  params: CheckParams,
  loadOpened: () => Promise<OpenedPptx>,
): Promise<string> {
  const hasOutline = params.outline !== undefined
  const hasPages = params.pages !== undefined
  const wantsSkeleton = params.skeleton === true
  const wantsStyles = params.styles === true
  // 没给 outline/pages 就是想审现有文件；给了就默认只做对账（audit 显式要才跑）。
  const wantsAudit =
    params.audit === true || (!hasOutline && !hasPages && !wantsSkeleton && !wantsStyles)

  const blocks: string[] = []
  if (wantsStyles) blocks.push(presetsGuide())
  if (hasOutline) {
    const outline = outlineOf(params.outline)
    blocks.push(formatOutlineIssues(outline.issues, outline.outline.pages.length).join('\n'))
  }
  if (hasPages) blocks.push(pagesBlock(params.pages, params.outline, params.style).join('\n'))

  if (wantsAudit || wantsSkeleton) {
    const opened = await loadOpened()
    if (wantsAudit) {
      const only =
        typeof params.slide === 'number' && Number.isInteger(params.slide) ? params.slide : undefined
      const total = only === undefined ? opened.deck.slides.length : 1
      blocks.push(formatAuditReport(auditOpened(opened, only), total))
    }
    if (wantsSkeleton) blocks.push(skeletonText(opened))
  }

  if (blocks.length === 0)
    throw coded('command.invalid', '没有可检查的内容：给 outline / pages，或省略它们来审计该文件')
  return clip(blocks.join('\n\n').split('\n')).join('\n')
}
