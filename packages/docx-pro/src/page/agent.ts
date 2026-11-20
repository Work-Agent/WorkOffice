// ─── 页面代理：agent-request → 编辑器（契约 §7.5） ─────────────────────
//
// 有页面时，页面的编辑会话是该 ref 的**唯一**编辑模型：
//   - context：返回 live 投影 + 新鲜度令牌（docRev / epoch）。
//   - apply：把 agent 命令应用在同一个编辑器实例上，然后走编辑器自己的保存
//     管线（`save()`）→ 序列化 → 宿主 `commit`（origin: 'agent'）。
//
// 走 `save()` 而不是直接 `serialize()` 是有意的：`saveOnce` 会在写完后重新
// 解析字节、刷新 `parsed` / 各类 dirty 标记。跳过它就等于让编辑器继续拿旧
// 基线做下一次保存（新图片、编号、页眉页脚都会错位）。

import type { Editor } from '@tiptap/core'
import { executeTool } from '../editor/ai/tools'
import { buildDocBytes, save, type FileActionContext } from '../editor/file-actions'
import { buildDocContext } from '../editor/ai/protocol'
import { applyBlockCommands, BlockCommandError } from './block-commands'
import { applyCommandBatch, numIdsFor } from './command-batch'
import { readMany } from '../engine/read'
import type { AgentContextRequest, DocxProBridge, Receipt } from './env'
import type { DocxHost } from './host'

export interface AgentBridgeDeps {
  sessionId: string
  host: DocxHost
}

type AnyRecord = Record<string, unknown>

function coded(code: string, message: string): Error & { code: string } {
  const error = new Error(message) as Error & { code: string }
  error.code = code
  return error
}

function bridgeOf(): DocxProBridge {
  const bridge = window.__docxPro
  if (!bridge) throw coded('command.page-unavailable', '编辑器尚未就绪')
  return bridge
}

function editorOf(): Editor {
  const editor = bridgeOf().getEditor() as Editor | null
  if (!editor) throw coded('command.page-unavailable', '编辑器尚未就绪')
  return editor
}

/**
 * 一次 agent 请求的命令形态（宿主在 host/plugins/host.ts 里生成）：
 * `commands: [{ tool: 'apply_ops', params: { ops: [...] } }]`
 */
interface AgentCommand {
  tool?: unknown
  params?: unknown
}

export function installAgentBridge(deps: AgentBridgeDeps): () => void {
  let docRev = 0
  let attached: Editor | null = null
  const attach = (editor: Editor): void => {
    if (attached === editor) return
    attached = editor
    editor.on('transaction', () => {
      docRev += 1
    })
  }
  const poll = setInterval(() => {
    const editor = window.__docxPro?.getEditor() as Editor | null | undefined
    if (editor) attach(editor)
  }, 200)

  const unsubscribe = window.extension.bus.on((raw) => {
    if (!raw || raw.type !== 'agent-request') return
    const request = raw as unknown as AgentContextRequest
    void handle(deps, request, () => docRev).then(
      (result) =>
        void window.extension.ura.page.agentResult(deps.sessionId, {
          type: 'agent-result',
          id: request.id,
          ok: true,
          ...result,
        }),
      (error: unknown) =>
        void window.extension.ura.page.agentResult(deps.sessionId, {
          type: 'agent-result',
          id: request.id,
          ok: false,
          error: {
            code:
              (error as { code?: string }).code ??
              (error instanceof BlockCommandError ? error.code : 'command.invalid'),
            message: error instanceof Error ? error.message : String(error),
          },
        }),
    )
  })

  return () => {
    clearInterval(poll)
    unsubscribe()
  }
}

async function handle(
  deps: AgentBridgeDeps,
  request: AgentContextRequest,
  docRevOf: () => number,
): Promise<{ context?: unknown; receipt?: Receipt }> {
  const editor = editorOf()
  if (request.kind === 'context') return { context: contextOf(editor, deps, docRevOf) }

  // 只读工具进页面时读 live 文档（含未保存修改），不回退磁盘。
  if (request.kind === 'read') {
    const tool = String(request.tool ?? '')
    const params = (request.read ?? {}) as AnyRecord
    // 命令目录不依赖文档内容：交给执行点生成。
    if (
      tool === 'docx_guide' ||
      tool === 'docx.guide' ||
      params['group'] !== undefined ||
      params['op'] !== undefined
    )
      throw coded('command.unsupported', `只读命令目录走执行点：${tool || 'docx_guide'}`)
    // 旧宿主没传 tool：无 what 时无法区分 docx_read（默认 context）与无参 guide。
    if (!tool && params['what'] === undefined)
      throw coded('command.unsupported', '旧宿主未提供工具名，无法区分读投影与命令目录')

    const ctx = bridgeOf().getCtx() as FileActionContext
    if (!ctx?.doc) throw coded('command.page-unavailable', '文档尚未加载完成')
    const bytes = await buildDocBytes(ctx)
    if (!bytes) throw coded('command.page-unavailable', '读取时无法序列化当前文档')
    return {
      context: readMany(bytes, {
        name: ctx.doc.fileName,
        version: deps.host.getVersion(),
        epoch: deps.host.getEpoch(),
        ...params,
      }),
    }
  }

  // 新鲜度：位置型坐标只在同一份基线上有效（契约 §8.3）。
  if (request.epoch !== undefined && request.epoch !== deps.host.getEpoch())
    throw coded(
      'command.stale-index',
      `页面基线已变（epoch ${request.epoch} → ${deps.host.getEpoch()}）；重新读取块索引后再发命令`,
    )

  const commands = Array.isArray(request.commands) ? (request.commands as AgentCommand[]) : []
  if (commands.length === 0) throw coded('command.invalid', '缺少 commands')

  const ctx = bridgeOf().getCtx() as FileActionContext
  if (!ctx?.doc) throw coded('command.page-unavailable', '文档尚未加载完成')

  let mutated = false
  let applied = 0
  for (const command of commands) {
    const tool = String(command.tool ?? '')
    const params = (command.params ?? {}) as AnyRecord
    if (tool === 'docx_apply' || tool === 'docx.apply') {
      const outcome = await applyCommandBatch(editor, params['commands'], ctx)
      applied += outcome.applied
      mutated ||= outcome.mutated
      continue
    }
    if (tool === 'docx_edit' || tool === 'docx.edit') {
      const outcome = applyBlockCommands(editor, params['commands'])
      applied += outcome.applied
      mutated ||= outcome.mutated
      continue
    }
    const outcome = await executeTool(
      editor,
      { id: 'docx-pro', name: tool, input: params } as never,
      numIdsFor(ctx),
    )
    if (outcome.isError)
      throw coded('command.invalid', outcome.output || `${tool} 失败`)
    applied += outcome.mutated ? 1 : 0
    mutated ||= outcome.mutated
  }

  if (!mutated) return { receipt: deps.host.lastReceipt() ?? undefined }

  if (request.dryRun) return { receipt: deps.host.lastReceipt() ?? undefined }

  // 走编辑器自己的保存管线：序列化 → commit(origin: agent) → 重解析基线。
  deps.host.setOrigin('agent')
  let ok = false
  try {
    ok = await save(ctx, false, true)
  } finally {
    deps.host.setOrigin('page')
  }
  if (!ok) {
    throw coded(
      'command.invalid',
      deps.host.lastError() ?? '保存失败：文档未落盘',
    )
  }
  const receipt = deps.host.lastReceipt()
  if (!receipt) throw coded('command.invalid', '提交未返回回执')
  return { receipt }
}

function contextOf(
  editor: Editor,
  deps: AgentBridgeDeps,
  docRevOf: () => number,
): AnyRecord {
  const doc = editor.state.doc
  const blocks: Array<{ index: number; type: string; text: string }> = []
  doc.forEach((node, _offset, index) => {
    blocks.push({
      index,
      type: node.type.name,
      text: node.textContent.replace(/\s+/g, ' ').slice(0, 160),
    })
  })
  const { from, to, empty } = editor.state.selection
  const selection = empty
    ? null
    : { from, to, text: doc.textBetween(from, to, '\n').slice(0, 2000) }
  return {
    docRev: String(docRevOf()),
    epoch: deps.host.getEpoch(),
    version: deps.host.getVersion(),
    blocks: blocks.slice(0, 400),
    blockCount: doc.childCount,
    selection,
    // 与编辑器自带的上下文构建同源（标题、表格、批注、页眉页脚、分节）
    text: safeContext(editor),
  }
}

function safeContext(editor: Editor): string {
  try {
    return buildDocContext(editor)
  } catch {
    return ''
  }
}
