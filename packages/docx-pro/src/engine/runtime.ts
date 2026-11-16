// ─── engine.runtime：探活 → 能力报告（契约 §10） ───────────────────────
//
// 三处同源（页面 / Agent / UI）的事实来源。这里只回报本插件真实具备的能力：
//   - write: 'plan'（页面交整份字节，执行点交块级切片）
//   - undo: 'page'（Ctrl+Z = 编辑器模型历史；无页面路径没有 undo）
//   - isolation: 'batch'（一条被拒 = 整批不落盘）
//   - baselineSync: 'manifest'（块级切片提交会重排 docxIndex）

import { MAX_COMMIT_BYTES, MAX_FILE_BYTES } from '../shared/runtime'
import type { EngineCtx } from './abi'
import type { RuntimeReport } from '../tools/types'

export const REPORT: RuntimeReport = {
  write: 'plan',
  exec: 'js',
  render: 'js',
  execution: 'isolated',
  features: {
    edit: true,
    recalc: false,
    bigObjects: false,
    legacyImport: false,
    exportPdf: false,
    imagePreview: true,
    isolation: 'batch',
    undo: 'page',
    requiresBaseVersion: true,
    tokens: {
      commit: 'version',
      inPage: 'docRev+epoch',
      baselineSync: 'manifest',
    },
  },
  limits: {
    maxFileBytes: MAX_FILE_BYTES,
    maxCommitBytes: MAX_COMMIT_BYTES,
  },
}

export async function runtime(ctx: EngineCtx): Promise<RuntimeReport> {
  return {
    ...REPORT,
    probe: { at: Date.now(), ok: Boolean(ctx) },
  }
}
