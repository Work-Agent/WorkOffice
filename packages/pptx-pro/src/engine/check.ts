// ─── engine.check：pptx_check 的实现（只读） ──────────────────────────
//
// 逻辑在 shared/check.ts（页面与执行点同源）；这里只负责按需打开基线字节。

import { openPptx, type OpenedPptx } from '@genoffice/pptx-engine'
import { coded, messageOf } from './abi'
import { runCheck, type CheckParams } from '../shared/check'

async function openRef(ref: string, bytes: Uint8Array | undefined): Promise<OpenedPptx> {
  if (!bytes) throw coded('resource.not-found', `文件不存在：${ref}`)
  try {
    return await openPptx(bytes)
  } catch (error) {
    throw coded('command.invalid', `PPTX 解析失败：${messageOf(error)}`)
  }
}

export async function check(
  ref: string,
  params: CheckParams,
  bytes: Uint8Array | undefined,
): Promise<string> {
  return runCheck(params, () => openRef(ref, bytes))
}

export type { CheckParams }
