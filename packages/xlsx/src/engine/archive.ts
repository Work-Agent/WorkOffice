// ─── wasm `save_archive`：只写 touched part ───────────────────────────
//
// 未改条目的压缩流逐字节原样拷贝（I1）。JSZip 那类实现会重压缩整包：
// 解压内容一样，但压缩流全变了 —— 大文件上既是性能问题也是漂移来源。
// 失败时调用方回退到 JSZip。

import { randomUUID } from 'node:crypto'
import { readFile, rm } from 'node:fs/promises'
import { join } from 'node:path'

import { CommandError } from './errors'
import { driftedEntries } from '../shared/zip-raw'
import type { SavePlan } from './gateway'
import type { WasmEngine } from './wasm'

export async function saveViaWasm(
  engine: WasmEngine,
  baseline: Uint8Array,
  plan: SavePlan,
): Promise<Uint8Array> {
  const runId = randomUUID().slice(0, 8)
  const written: string[] = []
  const tempName = (suffix: string): string => {
    const name = `save-${runId}-${suffix}`
    written.push(name)
    return name
  }

  try {
    const sourceName = tempName('base.xlsx')
    await engine.writeTemp(sourceName, baseline)

    const replacements: Array<{ name: string; contentPath: string }> = []
    let index = 0
    for (const [name, content] of plan.replaced) {
      const fileName = tempName(`r${index++}.entry`)
      replacements.push({ name, contentPath: await engine.writeTemp(fileName, content) })
    }
    const additions: Array<{ name: string; contentPath: string }> = []
    for (const [name, content] of plan.added) {
      const fileName = tempName(`a${index++}.entry`)
      additions.push({ name, contentPath: await engine.writeTemp(fileName, content) })
    }
    for (const [name, content] of plan.addedBinary) {
      const fileName = tempName(`b${index++}.entry`)
      additions.push({ name, contentPath: await engine.writeTemp(fileName, content) })
    }

    const targetName = tempName('out.xlsx')
    await engine.invoke({
      version: 1,
      requestId: `save-${runId}`,
      command: 'save_archive',
      sourcePath: engine.guestPath(sourceName),
      targetPath: engine.guestPath(targetName),
      replacements,
      removals: [...plan.removedEntries],
      additions,
      streamPatches: [],
    })

    const bytes = new Uint8Array(await readFile(join(engine.tempDir, targetName)))
    // 被删条目本来就该从包里消失：算进「涉及条目」，否则 I1 检查会把删除当成漂移。
    const drifted = driftedEntries(baseline, bytes, [
      ...plan.touchedEntries,
      ...plan.removedEntries,
    ])
    if (drifted.length > 0)
      throw new CommandError(
        'command.invalid',
        `保存改动了未涉及的 ZIP 条目：${drifted.slice(0, 5).join(', ')}`,
      )
    return bytes
  } catch (error) {
    if (error instanceof CommandError) throw error
    throw new CommandError(
      'command.invalid',
      `wasm save_archive 失败：${error instanceof Error ? error.message : String(error)}`,
    )
  } finally {
    await Promise.all(
      written.map((name) => rm(join(engine.tempDir, name), { force: true }).catch(() => undefined)),
    )
  }
}
