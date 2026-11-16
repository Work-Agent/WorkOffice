// ─── 计划 / 执行点共用的能力报告 ───────────────────────────────────────

export const MAX_FILE_BYTES = 64 * 1024 * 1024;
export const MAX_COMMIT_BYTES = 16 * 1024 * 1024;

export const TOKENS = {
  commit: 'version',
  inPage: 'docRev+epoch',
  baselineSync: 'manifest',
} as const;

/** 与宿主 `RuntimeReport`（WorkStudio src/shared/plugin-contract.ts）结构一致。 */
export interface RuntimeReportLike {
  write: 'none' | 'bytes' | 'plan';
  exec?: 'js' | 'wasm' | 'native' | 'host-render';
  render?: 'js' | 'wasm' | 'host-render' | 'none';
  execution: 'inline' | 'isolated';
  features: {
    edit: boolean;
    recalc: boolean;
    bigObjects: boolean;
    legacyImport: boolean;
    exportPdf: boolean;
    imagePreview: boolean;
    isolation: 'batch' | 'per-op';
    undo: 'none' | 'engine' | 'page';
    requiresBaseVersion: boolean;
    tokens?: {
      commit: 'version';
      inPage: 'docRev' | 'docRev+epoch' | 'docRev+structureEpoch';
      baselineSync: 'none' | 'version' | 'manifest' | 'bytes';
    };
  };
  limits: {
    maxFileBytes: number;
    maxCommitBytes: number;
    cacheModels?: number;
    cacheBytes?: number;
  };
  tier?: 'normal' | 'lite' | 'readOnly' | 'refused';
  degraded: Array<
    'engine-init-failed' | 'font-missing' | 'print-unavailable' | 'probe-pending' | 'probe-timeout'
  >;
  probe: { at: number; ok: boolean };
}

/**
 * docx 的能力报告（契约 §4.3 / §4.2.1）：
 * - `model: 'none'` 指**执行点**无常驻模型；页面仍有读侧模型（parsed / 块模型）。
 * - `undo: 'page'`：Ctrl+Z 是页面模型历史；无页面路径不提供 docx undo。
 * - `baselineSync: 'manifest'`：提交后必须 rebase `docxIndex`（I15）。
 */
export function runtimeReport(): RuntimeReportLike {
  return {
    write: 'plan',
    exec: 'js',
    render: 'js',
    execution: 'isolated',
    features: {
      edit: true,
      recalc: false,
      bigObjects: false,
      legacyImport: false,
      exportPdf: true,
      imagePreview: false,
      isolation: 'batch',
      undo: 'page',
      requiresBaseVersion: true,
      tokens: { ...TOKENS },
    },
    limits: {
      maxFileBytes: MAX_FILE_BYTES,
      maxCommitBytes: MAX_COMMIT_BYTES,
    },
    tier: 'normal',
    degraded: [],
    probe: { at: Date.now(), ok: true },
  };
}
