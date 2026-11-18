// ─── PPTX 页面入口：GenOffice Slides renderer + WorkStudio host ─────

import './shims/buffer-global'

import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlDir, htmlLang, type Lang } from '@genoffice/i18n'
import { installScreenTips } from '@genoffice/ui'
import { App } from './App'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import { installWorkStudioSlidesHost, readHostTheme } from './workstudio-host'

import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/color-picker.css'
import '@genoffice/ui/dropdown.css'
import '@genoffice/ui/ribbon-collapse.css'
import '@genoffice/ui/markdown.css'
import './styles.css'

const root = document.getElementById('root')
if (!root) throw new Error('Missing application root.')

installScreenTips()

function fail(message: string): void {
  if (!root) return
  root.innerHTML = `<div style="padding:24px;color:#b91c1c;white-space:pre-wrap">${message}</div>`
}

function applyTheme(theme: 'light' | 'dark'): void {
  document.documentElement.dataset.theme = theme
  document.documentElement.style.colorScheme = theme
}

async function bootstrap(): Promise<void> {
  const boot = window.__pptxProBoot
  const sessionId =
    boot?.sessionId ?? new URLSearchParams(location.search).get('session')
  if (!sessionId) {
    fail('缺少页面会话：请从 WorkStudio 资源面板打开 PPTX 演示文稿')
    return
  }
  if (!boot?.bytes) {
    fail('PPTX 页面启动数据缺失：请使用 apiVersion 2 的 WorkStudio 页面通道')
    return
  }

  let theme: 'light' | 'dark' = 'light'
  try {
    theme = await readHostTheme()
    applyTheme(theme)
  } catch {
    /* 主题失败不阻断打开 */
  }

  const lang: Lang = 'zh'
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  document.documentElement.dir = htmlDir(lang)

  await installWorkStudioSlidesHost({
    sessionId,
    data: boot.data,
    bytes: boot.bytes,
    version: boot.version,
    epoch: boot.epoch,
    onReceipt: (receipt) => {
      window.extension.bus.send({
        type: 'pptx-receipt',
        version: receipt.version,
        epoch: receipt.epoch,
        applied: receipt.applied,
      })
    },
  })

  window.extension.bus.on((message) => {
    if (message.type === 'theme') {
      applyTheme(message.theme === 'dark' ? 'dark' : 'light')
    }
  })

  createRoot(root as HTMLElement).render(
    <React.StrictMode>
      <LocaleProvider initial={lang}>
        <App />
      </LocaleProvider>
    </React.StrictMode>,
  )
  window.extension.bus.loaded()
  window.extension.bus.send({ type: 'dirty', dirty: false })
}

void bootstrap()
