// ─── 引擎侧只读入口：字节 → OpenedPptx → 投影 ────────────────────────
//
// 投影本身在 shared/read.ts（页面与引擎共用，语义必须一致）；
// 这里只按 ref 的字节打开模型，供 read / check 复用。

import { openPptx, type OpenedPptx } from '@genoffice/pptx-engine'
import { coded, messageOf } from './abi'

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

export { summarizeSlide, titleOfSlide } from '../shared/deck'

export async function openBytes(bytes: Uint8Array, ref: string): Promise<OpenedPptx> {
  if (!bytes || bytes.byteLength === 0) throw coded('resource.not-found', `文件不存在：${ref}`)
  try {
    return await openPptx(bytes)
  } catch (error) {
    throw coded('command.invalid', `PPTX 解析失败：${messageOf(error)}`)
  }
}
