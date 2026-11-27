// ─── GenOffice Slides 页面挂载 ─────────────────────────────────────────
//
// 只负责：语言 / 主题 / 字体 / React 挂载。
// 打开（load）、保存、助手请求都在 src/page/main.ts 里接线，
// 这样 editor 目录不出现任何宿主 I/O。

import './shims/buffer-global'

import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlDir, htmlLang, type Lang } from '@genoffice/i18n'
import { installScreenTips } from '@genoffice/ui'
import { App } from './App'
import { LocaleProvider, setModuleLang } from './i18n/locale'

import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/color-picker.css'
import '@genoffice/ui/dropdown.css'
import '@genoffice/ui/ribbon-collapse.css'
import '@genoffice/ui/markdown.css'
import './styles.css'

export function mountEditor(): void {
  const root = document.getElementById('root')
  if (!root) throw new Error('Missing application root.')

  const boot = window.__pptxProBoot
  const lang: Lang = 'zh'
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  document.documentElement.dir = htmlDir(lang)

  installScreenTips()

  createRoot(root).render(
    <React.StrictMode>
      <LocaleProvider initial={lang}>
        <App />
      </LocaleProvider>
    </React.StrictMode>,
  )

  if (!boot) console.warn('[pptx] boot 数据缺失：页面没有经过 pptx.load')
}
