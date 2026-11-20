// ─── engine.read：字节侧只读投影（执行点） ───────────────────────────
//
// 投影本身在 shared/read.ts（页面与执行点共用，语义必须一致）；
// 这里只负责：按 baseline 字节打开模型、给出 project/query。

import type { OpenedPptx } from '@genoffice/pptx-engine'
import { deckSummary } from '../shared/deck'
import { contextOf } from '../shared/read'
import { coded, type EngineCtx, type PlanRequest } from './abi'
import { getOrOpen, sha256Hex } from './state'

export {
  audit,
  contextOf,
  find,
  readCharts,
  readComments,
  readMany,
  readNotes,
  readSlide,
  readSlides,
  readTables,
  readTheme,
} from '../shared/read'

export async function loadOpened(
  req: Pick<PlanRequest, 'ref' | 'baseVersion' | 'bytes'>,
  ctx: EngineCtx,
): Promise<OpenedPptx> {
  const bytes = req.bytes ?? (await ctx.readBytes()).bytes
  if (!bytes || bytes.byteLength === 0) throw coded('command.invalid', 'PPTX 内容为空')
  const version = req.baseVersion ?? sha256Hex(bytes)
  const entry = await getOrOpen(req.ref, version, bytes)
  return entry.opened
}

export async function project(req: PlanRequest, ctx: EngineCtx): Promise<string> {
  const opened = await loadOpened(req, ctx)
  const params = (req.params ?? {}) as Record<string, unknown>
  return contextOf(opened, {
    ...(typeof params['name'] === 'string' ? { name: params['name'] } : {}),
    ...(req.baseVersion ? { version: req.baseVersion } : {}),
  })
}

export async function query(req: PlanRequest, ctx: EngineCtx): Promise<unknown> {
  const opened = await loadOpened(req, ctx)
  const summary = deckSummary(opened)
  return {
    total: summary.slides.length,
    rows: summary.slides,
  }
}

export { summarizeSlide, titleOfSlide } from '../shared/deck'
