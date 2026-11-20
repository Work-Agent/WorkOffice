// ─── engine.build：设计稿路径（页面 spec → pptx 字节） ────────────────
//
// 与 create.ts 的文字稿路径并列：
//   - 文字稿：`slides` / `content` → seed.pptx + addElement（标题 + 正文两段）
//   - 设计稿：`pages`（1280×720 px 的元素数组）→ buildDeckPptx（形状/文本/图片）
//
// 设计稿路径复用 vendor 的 pipelines：解析、校验、构建、文本按真实字体度量
// 长高、图片居中裁剪、背景提升，全部在那边实现；这里只负责把宿主的能力
// （图片下载、字体度量、错误翻译）接上。

import {
  buildDeckPptx,
  parseDeckSpec,
  MAX_DECK_PAGES,
  type DeckPageIssue,
  type DeckSpec,
} from '@genoffice/pipelines/slides/deck-spec'
import type { BuildPageDeps } from '@genoffice/pipelines/slides/page-spec'
import { HeuristicMetrics, imageSizeFromBytes } from '@genoffice/pptx-render'
import { coded, messageOf } from './abi'
import { fetchSpecImage, type FetchedImage } from './images'
import { asPagesInput } from '../shared/check'
import { expandDeck } from '../shared/design'

/** 字体度量：CLI 侧同样用启发式；真实字体索引只有页面宿主有。 */
const METRICS = new HeuristicMetrics()

export interface DesignDeck {
  spec: DeckSpec
  issues: DeckPageIssue[]
  /** 展开时的提示（未知版式 / 需要图但没给图 / 自定义 style.md 不影响配色） */
  notes: string[]
}

export interface BuiltDesignDeck {
  bytes: Uint8Array
  pages: number
  issues: DeckPageIssue[]
  imageFailures: Array<{ page: number; url: string }>
  imageWarnings: Array<{ page: number; message: string }>
}

/** 解析 + 逐页校验；失败直接抛 `command.invalid`（计划阶段就能回绝）。 */
export function parseDesignDeck(rawPages: unknown, styleInput?: unknown): DesignDeck {
  const input = asPagesInput(rawPages)
  if (!input) throw coded('command.invalid', 'pages 必须是页面数组、{pages:[…]} 或单个页面对象')
  // 高层页面（版式 + 内容）先摆成低层 spec；低层页面原样透传。
  const expanded = expandDeck(input.pages, styleInput)
  let parsed: ReturnType<typeof parseDeckSpec>
  try {
    parsed = parseDeckSpec(JSON.stringify({ pages: expanded.pages }), { localImages: true })
  } catch (error) {
    throw coded('command.invalid', `页面 spec 解析失败：${messageOf(error)}`)
  }
  if (!parsed.ok) throw coded('command.invalid', parsed.error)
  if (parsed.spec.pages.length > MAX_DECK_PAGES)
    throw coded('command.too-large', `页数 ${parsed.spec.pages.length} 超过上限 ${MAX_DECK_PAGES}`)
  return { spec: parsed.spec, issues: parsed.issues, notes: expanded.notes }
}

/** 构建字节；图片下载失败不阻断整份 deck（按页记录）。 */
export async function buildDesignDeck(
  deck: DesignDeck,
  opts: { signal?: AbortSignal; baseDir?: string } = {},
): Promise<BuiltDesignDeck> {
  const cache = new Map<string, Promise<FetchedImage | null>>()
  const deps: BuildPageDeps = {
    fetchImage: (url) => fetchSpecImage(url, cache, opts.signal, opts.baseDir),
    imageDims: (bytes) => {
      const size = imageSizeFromBytes(bytes)
      return size ? { width: size.w, height: size.h } : null
    },
    fontMetrics: METRICS,
  }
  const built = await buildDeckPptx(deck.spec, deps)
  return {
    bytes: built.bytes,
    pages: deck.spec.pages.length,
    issues: deck.issues,
    imageFailures: built.imageFailures,
    imageWarnings: built.imageWarnings,
  }
}

/** 校验结果的紧凑文本：计划阶段 / dryRun 都回给模型看。 */
export function formatPageIssues(issues: DeckPageIssue[]): string[] {
  const lines: string[] = []
  for (const issue of issues) {
    if (issue.error) lines.push(`page ${issue.page}: 丢弃 — ${issue.error}`)
    for (const warning of issue.warnings ?? []) lines.push(`page ${issue.page}: 警告 — ${warning}`)
  }
  return lines
}
