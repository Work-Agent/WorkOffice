// ─── 页面保存走哪条路径（设计稿 §9 P2） ───────────────────────────────
//
//   plan  `plan { kind:'save' }` ：请求（几十 KB）交给执行点，字节由 exec 产
//   bytes `plan { kind:'bytes' }`：页面自己调 gateway + JSZip，整份字节上行
//
// `bytes` 不是历史包袱，是三种情况的唯一正确路径：
//   1. 含内联二进制（单元格图片 / 视觉对象的 dataUrl）——请求会变成几 MB；
//   2. save-as / recovery —— 目标不是当前 ref，计划必须落在当前基线上；
//   3. 计划体积超过预算（超大批量编辑）。
//
// 判据必须是**纯函数**：页面侧没有「试一次失败再回退」的余地，
// 半途换路径会让同一批编辑被应用两次。

import type { XlsxSaveRequest } from '@xlsx/shared/save-request'

/** 计划 JSON 的默认体积上限：超过就走字节路径。 */
export const DEFAULT_MAX_PLAN_BYTES = 2 * 1024 * 1024

export interface SavePlanDecision {
  readonly usePlan: boolean
  readonly reason: 'ok' | 'save-as' | 'inline-media' | 'too-large'
  readonly planBytes: number
}

export interface SavePlanOptions {
  /** App 的保存模式；只有 `save` 落在当前 ref 上。 */
  readonly mode?: 'save' | 'save-as' | 'recovery' | string
  readonly maxPlanBytes?: number
}

export function planJsonBytes(plan: XlsxSaveRequest): number {
  try {
    return JSON.stringify(plan).length
  } catch {
    return Number.POSITIVE_INFINITY
  }
}

export function shouldUseSavePlan(
  plan: XlsxSaveRequest,
  options: SavePlanOptions = {},
): SavePlanDecision {
  const planBytes = planJsonBytes(plan)
  if ((options.mode ?? 'save') !== 'save')
    return { usePlan: false, reason: 'save-as', planBytes }
  if (hasInlineMedia(plan)) return { usePlan: false, reason: 'inline-media', planBytes }
  const maxPlanBytes = options.maxPlanBytes ?? DEFAULT_MAX_PLAN_BYTES
  if (!(planBytes <= maxPlanBytes)) return { usePlan: false, reason: 'too-large', planBytes }
  return { usePlan: true, reason: 'ok', planBytes }
}

/**
 * 图片 / 媒体以 data URL 内联进请求时，计划体积不再可预测（一张粘贴图几 MB）。
 * 判据放在请求上而不是编辑日志上：只要最终请求里出现内联载荷就走字节路径。
 */
function hasInlineMedia(plan: XlsxSaveRequest): boolean {
  if ((plan.cellImageAdditions?.length ?? 0) > 0) return true
  for (const addition of plan.visualAdditions ?? []) {
    const record = addition as unknown as Record<string, unknown>
    if (typeof record['dataUrl'] === 'string' || record['image'] !== undefined) return true
  }
  return false
}
