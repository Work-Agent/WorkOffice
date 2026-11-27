// ─── resources.xlsx：单个文件的资源适配器 ─────────────────────────────
//
// 只回答「这一个文件是什么」。ls / find / grep 恒为空：没有目录、没有集合，
// 内容检索走 xlsx.read（range / find），这样 agent 拿到的永远是
// 带 locator 的单文件投影，不会再出现第二套文件枚举。

import type { ResourceAdapter } from '@work-studio/sdk'

import { hashOfFile } from './hash'
import { assertXlsxFile, tierOf } from './ref'

export const xlsxResource: ResourceAdapter = {
  ls: async () => [],
  find: async () => [],
  grep: async () => [],
  read: async (ref) => {
    const file = await assertXlsxFile(ref)
    return [
      {
        locator: 'xlsx:file',
        ref,
        name: file.name,
        size: file.size,
        mtime: file.mtime,
        hash: await hashOfFile(file),
        tier: tierOf(file.size),
      },
    ]
  },
  open: async (ref, locator) => ({ ref, ...(locator ? { locator } : {}) }),
}
