// ─── xlsx-pro 页面入口 ─────────────────────────────────────────────────
//
// 1. 宿主能力自检（apiVersion 2：bytes / commit / agentResult）
// 2. page.get 取元数据；**有 `invoke` 时不读整份字节**（P5）：
//    页面用引擎的元数据 + 范围读取，字节只在回退路径按需取
// 3. 通过 window.__xlsxProBoot 注入编辑器启动数据
// 4. 启动本包自己的 Sheets 编辑器副本（src/editor）
// 5. 安装 agent 桥（页面是编辑权威；保存仍走 commit）

import { installAgentBridge } from './agent'
import type { PageSessionInfo } from './env'
import type { XlsxPageData } from '../editor/workstudio-host'
import { fetchSessionBytes } from './session-bytes'

function fail(message: string): void {
  const boot = document.querySelector<HTMLElement>('.ws-boot')
  if (!boot) return
  boot.dataset.error = '1'
  boot.textContent = message
}

function assertHost(): void {
  const page = (window.extension as { ura?: { page?: Record<string, unknown> } })?.ura?.page
  const missing = ['bytes', 'commit', 'agentResult'].filter(
    (name) => typeof page?.[name] !== 'function',
  )
  if (missing.length > 0)
    throw new Error(
      `宿主版本过低：缺少 ura.page.${missing.join(' / ')}。\n` +
        'xlsx-pro 需要 apiVersion 2 的 WorkStudio（0.3.x+，含计划型写与页面代理）。\n' +
        '在这个宿主上请改用 apiVersion 1 的 xlsx 插件。',
    )
}

/**
 * WorkStudio 的 Ctrl+S / 关闭前保存走视图总线：
 * `{ type: 'action', id: 'save' }`（见 view-runtime.prepare）。
 * 页面自己接住它，调用同一个 `__xlsxPro.save()`，并回报 save-result。
 */
function installSaveAction(): () => void {
  let pending = false
  return window.extension.bus.on((message) => {
    if (message.type !== 'action' || message.id !== 'save') return
    if (pending) return
    pending = true
    void (async () => {
      try {
        const deadline = Date.now() + 10_000
        while (!window.__xlsxPro && Date.now() < deadline) {
          await new Promise((resolve) => setTimeout(resolve, 50))
        }
        const bridge = window.__xlsxPro
        if (!bridge) {
          window.extension.bus.send({
            type: 'save-result',
            ok: false,
            error: '工作簿页面尚未就绪',
          })
          return
        }
        const result = await bridge.save()
        window.extension.bus.send({
          type: 'save-result',
          ok: result?.ok !== false,
          ...(result?.error ? { error: result.error } : {}),
        })
      } catch (error) {
        window.extension.bus.send({
          type: 'save-result',
          ok: false,
          error: error instanceof Error ? error.message : String(error),
        })
      } finally {
        pending = false
      }
    })()
  })
}

async function boot(): Promise<void> {
  const sessionId = new URLSearchParams(location.search).get('session')
  if (!sessionId) {
    fail('缺少页面会话：请从 WorkStudio 资源面板打开 XLSX 工作簿')
    return
  }

  try {
    assertHost()
    const info = await window.extension.ura.page.get<PageSessionInfo>(sessionId)
    const meta = info.data as Omit<XlsxPageData, 'bytes'>
    const size = Number(info.size ?? meta.size ?? 0)
    const invokeAvailable = typeof window.extension?.ura?.page?.invoke === 'function'
    // P5：有 invoke 就不在打开时读整份字节 —— 页面用引擎的元数据 + 范围读取，
    // 只有回退路径（老宿主 / bytes 计划）才按需调用 `ensureBytes()`。
    const bytes = invokeAvailable
      ? undefined
      : await fetchSessionBytes(
          (id, range) => window.extension.ura.page.bytes(id, range),
          sessionId,
          size,
        )
    const data: XlsxPageData = {
      ...meta,
      version: meta.version ?? info.version,
      size: bytes ? bytes.byteLength : size,
    }
    window.__xlsxProBoot = { sessionId, data, epoch: info.epoch ?? 0, ...(bytes ? { bytes } : {}) }

    await import('../editor/main')
    installAgentBridge({ sessionId })
    installSaveAction()
    document.querySelector<HTMLElement>('.ws-boot')?.remove()
  } catch (error) {
    fail(`XLSX 打开失败：${error instanceof Error ? error.message : String(error)}`)
  }
}

void boot()
