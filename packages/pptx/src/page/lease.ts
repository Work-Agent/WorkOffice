// ─── 页面队列：events → pending → 执行 → ack ──────────────────────────
//
// 引擎只入队（pptx.apply）；正文由页面取回，执行与落盘都在页面里：
//   runTxn（一个事务、一个 undo 步）→ savePptx → pptx.save → report → ack
//
// 只读请求（read / check）同样走 pending：有页面时读的是**活的**模型
// （含未保存修改），这样助手看到的和用户屏幕上的一致。

import { channel } from './channel'
import type { PageRequest } from './env'
import type { SlidesHost } from './host'

const HEARTBEAT_MS = 15_000

export interface QueueOptions {
  ref: string
  pageId: string
  host: SlidesHost
}

export interface PageQueue {
  /** 处理一条请求（attach 返回的 pending 也用它）。 */
  handle(id: string): Promise<void>
  stop(): void
}

export function installPageQueue(options: QueueOptions): PageQueue {
  const { ref, pageId, host } = options
  let stopped = false

  const beat = setInterval(() => {
    void channel
      .report({ ref, pageId, hash: host.getHash(), docRev: host.getDocRev() })
      .catch(() => undefined)
  }, HEARTBEAT_MS)

  const off = channel.onRequest((id) => {
    void handle(id)
  })

  async function handle(id: string): Promise<void> {
    if (stopped) return
    let request: PageRequest
    try {
      request = await channel.pending(id)
    } catch {
      // 引擎侧已超时（30s）或页面被接管：这条请求不再归属本页面。
      return
    }
    try {
      if (request.type === 'read') {
        const text =
          request.what === 'check'
            ? await host.check(request.params)
            : request.what === 'guide'
              ? host.guide(request.params)
              : host.read(request.params)
        await ackOk(id, { data: text })
        return
      }

      const result = await host.applyBatch(request.commands, request.dryRun === true)
      if (!result.ok) {
        const failure = result.failures?.[0]
        await ackFail(id, {
          code: 'command.invalid',
          message: failure?.error ?? result.error ?? '页面执行失败',
          ...(typeof failure?.index === 'number' ? { index: failure.index } : {}),
        })
        return
      }
      if (request.dryRun) {
        await ackOk(id, { data: { dryRun: true, applied: 0, plan: result.plan ?? null } })
        return
      }

      const saved = await host.save()
      if (!saved.ok) {
        await ackFail(id, {
          code: saved.error?.startsWith('command.stale')
            ? 'command.stale'
            : 'command.failed',
          message: saved.error ?? '保存失败',
        })
        return
      }
      await ackOk(id, {
        receipt: {
          ref,
          hash: host.getHash(),
          size: host.getBytes().byteLength,
          at: Date.now(),
          applied: result.applied ?? request.commands.length,
        },
      })
    } catch (error) {
      await ackFail(id, {
        code: 'command.failed',
        message: error instanceof Error ? error.message : String(error),
      })
    }
  }

  async function ackOk(id: string, payload: Record<string, unknown>): Promise<void> {
    await channel.ack({ id, ok: true, payload }).catch(() => undefined)
  }

  async function ackFail(
    id: string,
    error: { code: string; message: string; index?: number },
  ): Promise<void> {
    await channel.ack({ id, ok: false, payload: { error } }).catch(() => undefined)
  }

  return {
    handle,
    stop(): void {
      if (stopped) return
      stopped = true
      clearInterval(beat)
      off()
      void channel.detach({ ref, pageId }).catch(() => undefined)
    },
  }
}
