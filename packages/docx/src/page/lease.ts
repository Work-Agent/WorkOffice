// ─── 页面队列：WorkStudio events + pending + ack ───────────────────────
//
// 旧宿主的 agent-request 页面代理已不存在。这里用插件自己的事件通道
// 复刻同一语义：引擎 emit request → 页面取正文 → 执行 → ack。
// Agent 的写永远经过页面 PM 事务，最后仍由 docx.save 落盘。

import type { Editor } from '@tiptap/core'
import type { FileActionContext } from '../editor/file-actions'
import { buildDocBytes, save } from '../editor/file-actions'
import { readMany } from '../engine/read'
import { applyCommandBatch } from './command-batch'
import type { DocxHost } from './host'
import type { PageRequest, Receipt } from './env'

interface QueueOptions {
  ref: string
  pageId: string
  host: DocxHost
  locator?: string
}

interface AckPayload {
  data?: unknown
  receipt?: Receipt
  error?: { code: string; message: string }
}

function bridge() {
  return window.workStudio
}

function editorOf(): Editor | null {
  return (window.__docx?.getEditor() as Editor | null | undefined) ?? null
}

function ctxOf(): FileActionContext | null {
  return (window.__docx?.getCtx() as FileActionContext | null | undefined) ?? null
}

function coded(error: unknown): { code: string; message: string } {
  const message = error instanceof Error ? error.message : String(error)
  const match = /^([a-z][a-z0-9.-]*):\s*(.*)$/i.exec(message)
  return match
    ? { code: match[1]!, message: match[2] || message }
    : { code: 'command.invalid', message }
}

async function ack(id: string, payload: AckPayload): Promise<void> {
  await bridge().command('docx', 'docx', 'ack', {
    id,
    ok: !payload.error,
    payload: JSON.stringify(payload),
  })
}

export function installPageQueue(options: QueueOptions) {
  const { ref, pageId, host } = options
  // 必须先订阅本插件 events，否则 receipt / request 不会投递到页面。
  void bridge().subscribe('docx', 'events').catch(() => undefined)
  let docRev = 0
  let attached: Editor | null = null
  let reportTimer: ReturnType<typeof setTimeout> | null = null
  let closed = false

  const report = (): void => {
    if (reportTimer) return
    reportTimer = setTimeout(() => {
      reportTimer = null
      void bridge()
        .command('docx', 'docx', 'report', {
          ref,
          pageId,
          docRev,
          hash: host.getHash(),
        })
        .catch(() => undefined)
    }, 200)
  }

  const attachEditor = (editor: Editor): void => {
    if (attached === editor) return
    attached = editor
    editor.on('transaction', () => {
      docRev += 1
      report()
    })
    if (options.locator) {
      const index = Number(/^docx:block:(\d+)$/.exec(options.locator)?.[1] ?? -1)
      const node = index >= 0 ? editor.state.doc.child(index) : null
      if (node) {
        let pos = 0
        for (let i = 0; i < index; i += 1) pos += editor.state.doc.child(i).nodeSize
        editor.commands.setTextSelection(pos + 1)
      }
    }
  }

  const poll = setInterval(() => {
    const editor = editorOf()
    if (editor) attachEditor(editor)
  }, 200)

  // 心跳：更新 pageId 的 lastSeen，供诊断；租约本身不按时间失效。
  const heartbeat = setInterval(() => {
    if (!closed) report()
  }, 15_000)

  const handleRead = async (request: Extract<PageRequest, { type: 'read' }>): Promise<AckPayload> => {
    const ctx = ctxOf()
    if (!ctx?.doc) throw new Error('command.page-unavailable: 文档尚未加载完成')
    const bytes = await buildDocBytes(ctx)
    if (!bytes) throw new Error('command.page-unavailable: 无法序列化当前文档')
    // 先同步租约，避免 Agent 刚 read 完的 docRev 在 apply 时被判 stale-index。
    await bridge().command('docx', 'docx', 'report', {
      ref,
      pageId,
      docRev,
      hash: host.getHash(),
    }).catch(() => undefined)
    return {
      data: {
        text: readMany(bytes, {
          ...request.params,
          what: request.what,
          name: ctx.doc.fileName,
        }),
        hash: host.getHash(),
        docRev,
      },
    }
  }

  const handleApply = async (request: Extract<PageRequest, { type: 'apply' }>): Promise<AckPayload> => {
    const editor = editorOf()
    const ctx = ctxOf()
    if (!editor || !ctx?.doc) throw new Error('command.page-unavailable: 编辑器尚未就绪')
    const snapshot = editor.getJSON()
    let applied = 0
    try {
      const outcome = await applyCommandBatch(editor, request.commands, ctx)
      applied = outcome.applied
      if (request.dryRun) {
        editor.commands.setContent(snapshot)
        return { data: { applied, dryRun: true } }
      }
      if (!outcome.mutated) return { receipt: host.lastReceipt() ?? { applied: 0 } as Receipt }
      host.setOrigin('agent')
      const ok = await save(ctx, false, true)
      if (!ok) throw new Error(`command.invalid: ${host.lastError() ?? '保存失败'}`)
      return { receipt: host.lastReceipt() ?? { applied } as Receipt }
    } catch (error) {
      try {
        editor.commands.setContent(snapshot)
      } catch {
        /* 回滚失败则保留页面当前状态 */
      }
      throw error
    } finally {
      host.setOrigin('page')
    }
  }

  const handle = async (request: PageRequest): Promise<void> => {
    try {
      const payload = request.type === 'read'
        ? await handleRead(request)
        : await handleApply(request)
      await ack(request.id, payload)
    } catch (error) {
      await ack(request.id, { error: coded(error) })
    }
  }

  const handlePending = async (requests: PageRequest[]): Promise<void> => {
    for (const request of requests) await handle(request)
  }

  const off = bridge().onUpdate((update) => {
    if (closed) return
    if (update.domain !== 'docx' || update.resource !== 'events') return
    if (update.type !== 'event' || update.event !== 'request') return
    const id = (update.data as { id?: unknown } | null)?.id
    if (typeof id !== 'string') return
    void bridge()
      .command<PageRequest | null>('docx', 'docx', 'pending', { id })
      .then((result) => {
        if (result.ok && result.data) return handle(result.data)
      })
      .catch(() => undefined)
  })

  const onVisible = (): void => {
    if (!closed && !document.hidden) report()
  }
  document.addEventListener('visibilitychange', onVisible)
  window.addEventListener('focus', onVisible)

  window.addEventListener('pagehide', () => {
    if (closed) return
    closed = true
    clearInterval(poll)
    clearInterval(heartbeat)
    if (reportTimer) clearTimeout(reportTimer)
    off()
    document.removeEventListener('visibilitychange', onVisible)
    window.removeEventListener('focus', onVisible)
    void bridge().command('docx', 'docx', 'detach', { ref, pageId }).catch(() => undefined)
  })

  return {
    handlePending,
    touch: report,
    dispose(): void {
      if (closed) return
      closed = true
      clearInterval(poll)
      clearInterval(heartbeat)
      if (reportTimer) clearTimeout(reportTimer)
      off()
      document.removeEventListener('visibilitychange', onVisible)
      window.removeEventListener('focus', onVisible)
    },
  }
}

export type PageQueue = ReturnType<typeof installPageQueue>
