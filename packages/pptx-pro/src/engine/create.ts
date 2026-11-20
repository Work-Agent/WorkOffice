// ─── DeckSpec → 新建事务 ops ──────────────────────────────────────────
//
// 新建使用 seed.pptx（16:9 一页空白），再用 addElement / addBlankSlide 生成
// 标题、正文、项目符号、备注与背景。P0 不依赖母版占位符。

import type { OpenedPptx } from '@genoffice/pptx-engine'
import type { DeckSpec, SlideSpec } from '../shared/spec'
import type { PptxOp } from '../shared/ops'

const TITLE_PT = 40
const BODY_PT = 20

function paragraphsOf(spec: SlideSpec): Array<Record<string, unknown>> {
  const paragraphs: Array<Record<string, unknown>> = []
  const body = Array.isArray(spec.body) ? spec.body : spec.body ? [spec.body] : []
  for (const text of body) {
    if (!text.trim()) continue
    paragraphs.push({
      runs: [{ text, fontSize: BODY_PT }],
    })
  }
  for (const text of spec.bullets ?? []) {
    if (!text.trim()) continue
    paragraphs.push({
      runs: [{ text, fontSize: BODY_PT }],
      bullet: 'char',
      bulletChar: '•',
    })
  }
  return paragraphs
}

export interface CreateOps {
  /** 先执行的建页/尺寸/主题；下游内容必须在同一批之外执行（runTxn 先整体校验）。 */
  setup: PptxOp[]
  /** 再执行的内容：背景/标题/正文/备注。 */
  content: PptxOp[]
}

export function buildCreateOps(opened: OpenedPptx, spec: DeckSpec): CreateOps {
  const setup: PptxOp[] = []
  const content: PptxOp[] = []
  const size = opened.deck.size

  // setSlideSize 对“已经是这个尺寸”返回 false（op 会当成失败）；
  // 新建时 16:9 正好是 seed 的尺寸，所以先比对再发。
  const pushSize = (cx: number, cy: number): void => {
    if (size.cx === cx && size.cy === cy) return
    setup.push({ op: 'setSlideSize', cx, cy })
  }
  if (spec.size === '4:3') pushSize(9144000, 6858000)
  else if (spec.size === '16:9') pushSize(12192000, 6858000)
  else if (spec.size && typeof spec.size === 'object') pushSize(spec.size.cx, spec.size.cy)

  if (spec.theme) {
    setup.push({
      op: 'applyTheme',
      name: spec.theme.name,
      colors: spec.theme.colors,
      ...(spec.theme.majorFont ? { majorFont: spec.theme.majorFont } : {}),
      ...(spec.theme.minorFont ? { minorFont: spec.theme.minorFont } : {}),
    })
  }

  const slides = spec.slides.length > 0 ? spec.slides : [{}]
  for (const [index, slide] of slides.entries()) {
    // 批内坐标要按“批开始时的单页 seed”校验：用 index-1 会在第 3 页起越界。
    // 新插入的空白页顺序无所谓（content 随后按最终索引写），统一插入到第 0 页之后。
    if (index > 0) setup.push({ op: 'addBlankSlide', target: { slide: 0 } })

    if (slide.background)
      content.push({
        op: 'setBackground',
        target: { slide: index },
        kind: 'solid',
        color: slide.background,
      })

    if (slide.title) {
      content.push({
        op: 'addElement',
        target: { slide: index },
        kind: 'textbox',
        offset: {
          x: Math.round(size.cx * 0.08),
          y: Math.round(size.cy * 0.06),
          cx: Math.round(size.cx * 0.84),
          cy: Math.round(size.cy * 0.18),
        },
        paragraphs: [{ runs: [{ text: slide.title, fontSize: TITLE_PT, bold: true }] }],
        bodyPr: { autoFit: 'shrink' },
      })
    }

    const paragraphs = paragraphsOf(slide)
    if (paragraphs.length > 0) {
      content.push({
        op: 'addElement',
        target: { slide: index },
        kind: 'textbox',
        offset: {
          x: Math.round(size.cx * 0.08),
          y: Math.round(size.cy * 0.28),
          cx: Math.round(size.cx * 0.84),
          cy: Math.round(size.cy * 0.58),
        },
        paragraphs,
        bodyPr: { autoFit: 'shrink' },
      })
    }

    if (slide.notes) content.push({ op: 'setNotes', target: { slide: index }, text: slide.notes })
  }

  return { setup, content }
}
