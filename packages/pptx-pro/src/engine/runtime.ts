// ─── engine.runtime：探活 → 能力报告 ─────────────────────────────────

import { runtimeReport, type RuntimeReportLike } from '../shared/runtime'
import type { EngineCtx } from './abi'

export async function runtime(ctx: EngineCtx): Promise<RuntimeReportLike> {
  return {
    ...runtimeReport(),
    probe: { at: Date.now(), ok: Boolean(ctx) },
  }
}

export const REPORT = runtimeReport()
