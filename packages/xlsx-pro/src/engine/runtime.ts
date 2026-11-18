// ─── engine.runtime：能力报告 ──────────────────────────────────────────

import type { EngineCtx } from './abi'
import type { RuntimeReport } from '../tools/types'
import { probeWasm } from './wasm'

/** TS 回退底座的上限（全量解析 + 整包重压缩）。 */
export const MAX_FILE_BYTES = 64 * 1024 * 1024
/** wasm 底座的上限：实测 150 MB 舒适、300 MB 边缘，这里只承诺前者。 */
export const MAX_WASM_FILE_BYTES = 150 * 1024 * 1024
export const MAX_COMMIT_BYTES = 16 * 1024 * 1024

export const REPORT: RuntimeReport = {
  write: 'plan',
  exec: 'js',
  render: 'js',
  execution: 'isolated',
  features: {
    edit: true,
    recalc: false,
    bigObjects: true,
    legacyImport: false,
    exportPdf: false,
    imagePreview: true,
    isolation: 'batch',
    undo: 'page',
    requiresBaseVersion: true,
    tokens: {
      commit: 'version',
      // A1 在值编辑下稳定，只有结构改动让它失效（设计稿 §6.3）。
      // epoch 是提交代数，只做诊断，不参与 op 判定。
      inPage: 'docRev+structureEpoch',
      baselineSync: 'manifest',
    },
  },
  limits: {
    maxFileBytes: MAX_FILE_BYTES,
    maxCommitBytes: MAX_COMMIT_BYTES,
  },
}

export async function runtime(ctx: EngineCtx): Promise<RuntimeReport> {
  // 探活失败不抛、只降级（N6）：wasm 不可用时 exec 走 TS 回退路径。
  // 探活用稳定的会话目录：与 exec / read 落在同一个 wasm 实例上。
  const wasm = await probeWasm(ctx.sessionDir ?? ctx.tempDir)
  const degraded = wasm.ok
    ? // 单线程 reactor 没有 reader 线程，Cancel 只能命中下一个请求
      ['cancel-unavailable']
    : ['engine-init-failed', 'cancel-unavailable']
  return {
    ...REPORT,
    exec: wasm.ok ? 'wasm' : 'js',
    features: { ...REPORT.features, recalc: wasm.ok },
    limits: { ...REPORT.limits, maxFileBytes: wasm.ok ? MAX_WASM_FILE_BYTES : MAX_FILE_BYTES },
    degraded,
    probe: {
      at: Date.now(),
      ok: Boolean(ctx),
      engine: wasm.ok ? 'wasm' : 'js',
      ...(wasm.reason === undefined ? {} : { reason: wasm.reason }),
      ...(wasm.protocolVersion === undefined ? {} : { protocolVersion: wasm.protocolVersion }),
    },
  }
}
