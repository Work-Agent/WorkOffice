// ─── 计划 / 执行点共用的能力报告与预算 ───────────────────────────────
//
// 当前宿主：
//   - ENGINE_BUDGET.plan/exec.maxBytes = 64 MB
//   - authority.readArtifact 会整份读入内存并受 maxCommitBytes 约束
// 因此 P0 的实际上限是 64 MB，而不是 200 MB。

export const MAX_FILE_BYTES = 64 * 1024 * 1024
export const MAX_COMMIT_BYTES = 64 * 1024 * 1024
export const MAX_CACHE_MODELS = 2
export const MAX_CACHE_BYTES = 512 * 1024 * 1024

export const TOKENS = {
  commit: 'version',
  inPage: 'docRev',
  baselineSync: 'version',
} as const

export interface RuntimeReportLike {
  write: 'none' | 'bytes' | 'plan'
  model?: 'none' | 'resident'
  exec?: 'js' | 'wasm' | 'native' | 'host-render'
  render?: 'js' | 'wasm' | 'host-render' | 'none'
  execution: 'inline' | 'isolated'
  features: {
    edit: boolean
    recalc: boolean
    bigObjects: boolean
    legacyImport: boolean
    exportPdf: boolean
    imagePreview: boolean
    isolation: 'batch' | 'per-op'
    undo: 'none' | 'engine' | 'page'
    requiresBaseVersion: boolean
    tokens?: {
      commit: 'version'
      inPage: 'docRev' | 'docRev+epoch' | 'docRev+structureEpoch'
      baselineSync: 'none' | 'version' | 'manifest' | 'bytes'
    }
  }
  limits: {
    maxFileBytes: number
    maxCommitBytes: number
    cacheModels?: number
    cacheBytes?: number
  }
  tier?: 'normal' | 'lite' | 'readOnly' | 'refused'
  degraded?: Array<
    | 'engine-init-failed'
    | 'font-missing'
    | 'print-unavailable'
    | 'probe-pending'
    | 'probe-timeout'
  >
  probe?: { at: number; ok: boolean }
}

export function runtimeReport(): RuntimeReportLike {
  return {
    write: 'plan',
    model: 'resident',
    exec: 'js',
    render: 'js',
    execution: 'isolated',
    features: {
      edit: true,
      recalc: false,
      bigObjects: true,
      legacyImport: false,
      exportPdf: true,
      imagePreview: true,
      isolation: 'batch',
      undo: 'page',
      requiresBaseVersion: true,
      tokens: { ...TOKENS },
    },
    limits: {
      maxFileBytes: MAX_FILE_BYTES,
      maxCommitBytes: MAX_COMMIT_BYTES,
      cacheModels: MAX_CACHE_MODELS,
      cacheBytes: MAX_CACHE_BYTES,
    },
    tier: 'normal',
    degraded: [],
    probe: { at: Date.now(), ok: true },
  }
}
