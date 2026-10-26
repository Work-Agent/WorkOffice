// ─── PPTX 页面入口：GenOffice Slides renderer + WorkStudio host ─────

import './shims/buffer-global'

import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlDir, htmlLang, type Lang } from '@genoffice/i18n'
import { installScreenTips } from '@genoffice/ui'
import { App } from './App'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import { installWorkStudioSlidesHost } from './workstudio-host'

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
  const sessionId = new URLSearchParams(location.search).get('session')
  if (!sessionId) {
    fail('缺少页面会话：请从 WorkStudio 资源面板打开 PPTX 演示文稿')
    return
  }

  let info: { id: string; data: any } | null = null
  let theme: 'light' | 'dark' = 'light'
  try {
    info = await window.extension.ura.page.get<{
      id: string
      data: any
    }>(sessionId)
    const ui = await window.extension.invoke<{ theme?: string }>('app:ui.get')
    theme = ui?.theme === 'dark' ? 'dark' : 'light'
    applyTheme(theme)
  } catch (error) {
    fail(
      `PPTX 打开失败：${error instanceof Error ? error.message : String(error)}`,
    )
    return
  }
  if (!info?.data?.bytes) {
    fail('PPTX 页面数据缺少字节内容')
    return
  }

  const lang: Lang = 'zh'
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  document.documentElement.dir = htmlDir(lang)

  await installWorkStudioSlidesHost({ sessionId, data: info.data })

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
