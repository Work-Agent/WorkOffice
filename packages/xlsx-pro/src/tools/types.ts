// ─── 宿主契约的本地子集（与 docx-pro / WorkStudio plugin-contract 对齐） ──

export interface ToolDecl {
  name: string
  label: string
  description: string
  kinds: string[]
  readOnly: boolean
  requiresBaseVersion?: boolean
  tokens?: 'docRev' | 'docRev+epoch' | 'docRev+structureEpoch'
  parameters: unknown
  plan?: string
  maxItems?: number
}

export interface RuntimeReport {
  write: 'none' | 'bytes' | 'plan'
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
  degraded?: string[]
  probe?: {
    at: number
    ok: boolean
    /** 当前执行底座：wasm 引擎或 TS 回退 */
    engine?: 'wasm' | 'js'
    /** 降级原因（engine='js' 时）：module-missing / wasi-unavailable / init-failed / abi-mismatch */
    reason?: string
    protocolVersion?: number
  }
  tier?: 'normal' | 'lite' | 'readOnly' | 'refused'
}
