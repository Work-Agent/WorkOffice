// ─── 页面代理：agent-request → 同一个页面模型 / 同一个 undo 栈 ────────
//
// 人类与 LLM 共用 page.opened：agent 命令也走 runTxn，然后 pageSave。
// 一个 agent 请求 = 一个页面 undo 步；执行点不再二次 apply。

import type { AgentContextRequest, PptxProHost, Receipt } from './env'

type AnyRecord = Record<string, unknown>

function coded(code: string, message: string): Error & { code: string } {
  const error = new Error(message) as Error & { code: string }
  error.code = code
  return error
}

async function flushActiveEdit(): Promise<void> {
  const active = document.activeElement as HTMLElement | null
  if (!active?.isContentEditable) return
  active.blur()
  // 等 blur 触发的 TextEditOverlay 提交落到模型
  for (let i = 0; i < 20; i++) {
    await new Promise((resolve) => setTimeout(resolve, 25))
    if (!(document.activeElement as HTMLElement | null)?.isContentEditable) break
  }
}

async function waitForHost(): Promise<PptxProHost> {
  const deadline = Date.now() + 30_000
  while (Date.now() < deadline) {
    const host = window.__pptxProHost
    if (host) return host
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  throw coded('command.page-unavailable', '演示文稿页面尚未就绪')
}

function formatFailures(result: {
  failures?: Array<{ index: number; error: string }>
  error?: string
}): string {
  if (result.error) return result.error
  if (result.failures?.length)
    return result.failures.map((f) => `[${f.index}] ${f.error}`).join('\n')
  return '应用失败'
}

export function installAgentBridge(deps: { sessionId: string }): () => void {
  const unsubscribe = window.extension.bus.on((raw) => {
    if (!raw || raw.type !== 'agent-request') return
    const request = raw as unknown as AgentContextRequest
    void handle(deps, request).then(
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
            code: (error as { code?: string }).code ?? 'command.invalid',
            message: error instanceof Error ? error.message : String(error),
          },
        }),
    )
  })
  return unsubscribe
}

async function handle(
  deps: { sessionId: string },
  request: AgentContextRequest,
): Promise<{ context?: unknown; receipt?: Receipt }> {
  const host = await waitForHost()

  if (request.kind === 'context') {
    return { context: host.context() }
  }

  const commands = Array.isArray(request.commands) ? (request.commands as AnyRecord[]) : []
  if (commands.length === 0) throw coded('command.invalid', '缺少 commands')

  const batches: AnyRecord[][] = []
  let dryRun = request.dryRun === true
  let requestedDocRev: string | undefined
  for (const [index, command] of commands.entries()) {
    const tool = String(command['tool'] ?? '')
    if (tool !== 'pptx_apply' && tool !== 'pptx.apply')
      throw coded(
        'command.page-unavailable',
        `页面代理不支持工具 ${tool}；写操作请用 pptx_apply`,
      )
    const params = (command['params'] ?? {}) as AnyRecord
    if (typeof params['docRev'] === 'string') requestedDocRev = params['docRev']
    if (params['dryRun'] === true) dryRun = true
    const batch = Array.isArray(params['commands'])
      ? params['commands']
      : Array.isArray(params['ops'])
        ? params['ops']
        : []
    if (batch.length === 0) throw coded('command.invalid', `commands[${index}] 缺少 commands`)
    for (const op of batch)
      if (!op || typeof op !== 'object') throw coded('command.invalid', 'op 不是对象')
    batches.push(batch as AnyRecord[])
  }

  if (requestedDocRev !== undefined && requestedDocRev !== String(host.getDocRev()))
    throw coded(
      'command.stale-index',
      `页面模型已变化（docRev ${requestedDocRev} → ${host.getDocRev()}）；重新读取后再发命令`,
    )

  await flushActiveEdit()
  const result = await host.applyAgentTxn(batches, dryRun)
  if (!result.ok) throw coded('command.invalid', formatFailures(result))
  if (dryRun) {
    const receipt = host.getReceipt()
    return { receipt: receipt ?? undefined }
  }
  if (!result.receipt) throw coded('command.invalid', '提交未返回回执')
  return { receipt: result.receipt }
}
