// ─── 宿主契约的本地子集 ────────────────────────────────────────────────
//
// 宿主契约的唯一事实源是 WorkStudio `src/shared/plugin-contract.ts`；插件产
// 物在独立进程里求值，不能 import 宿主源码。这里只声明本插件用到的字段，
// 形状与契约 §6.1 / §10 一致。多声明不会出错，少声明会在装载期被拒。

export interface ToolDecl {
  name: string
  label: string
  description: string
  /** 归属校验：页面调用与资源匹配都用它（MIME / 扩展名 / `*`）。 */
  kinds: string[]
  readOnly: boolean
  /** 写盘基线：**无页面**写必填；有页面时基线由页面会话持有（§7.5 / §8.3）。 */
  requiresBaseVersion?: boolean
  /** 页面内新鲜度令牌的声明（取值由插件决定，宿主只透传，见 §10）。 */
  tokens?: 'docRev' | 'docRev+epoch' | 'docRev+structureEpoch'
  /** JSON Schema；宿主转成 agent 的工具参数类型。 */
  parameters: unknown
  /** 执行点里的 plan 名；缺省 = name。 */
  plan?: string
  /** 批上限；宿主强制执行。 */
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
  probe?: { at: number; ok: boolean }
  tier?: 'normal' | 'lite' | 'readOnly' | 'refused'
}
