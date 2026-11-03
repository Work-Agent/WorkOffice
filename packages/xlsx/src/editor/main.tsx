// ─── XLSX 页面入口：GenOffice Sheets renderer + WorkStudio host ─────

import './shims/buffer-global'

import ReactDOM from 'react-dom/client'
import { htmlDir, htmlLang, type Lang } from '@genoffice/i18n'
import { installScreenTips } from '@genoffice/ui'

import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/color-picker.css'
import '@genoffice/ui/dropdown.css'
import '@genoffice/ui/ribbon-collapse.css'
import '@genoffice/ui/markdown.css'
import '@univerjs/preset-sheets-core/lib/index.css'

import { App } from './App'
import {
  installCanvasFontFallback,
  registerTextDingbatsFace,
  untilCellFontsLoaded,
} from './cell-font-fallback'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import { installWorkStudioHost, readHostTheme, type XlsxPageData } from './workstudio-host'
import './styles.css'

const root = document.getElementById('root')
if (!root) throw new Error('Missing application root.')

installScreenTips()
installCanvasFontFallback()

function fail(message: string): void {
  if (!root) return
  root.innerHTML = `<div style="padding:24px;color:#b91c1c;white-space:pre-wrap">${message}</div>`
}

function applyTheme(theme: 'light' | 'dark'): void {
  document.documentElement.dataset.theme = theme
  document.documentElement.style.colorScheme = theme
}

async function loadCellFonts(): Promise<void> {
  const loads: Promise<unknown>[] = [registerTextDingbatsFace()]
  for (const variant of ['', 'bold ', 'italic ', 'italic bold ']) {
    for (const family of ['Calibri', 'Aptos', "'Aptos Narrow'", 'Carlito']) {
      loads.push(document.fonts?.load?.(`${variant}16px ${family}`)?.catch(() => {}) ?? [])
    }
  }
  await untilCellFontsLoaded(Promise.all(loads))
}

async function bootstrap(): Promise<void> {
  const sessionId = new URLSearchParams(location.search).get('session')
  if (!sessionId) {
    fail('缺少页面会话：请从 WorkStudio 资源面板打开 XLSX 工作簿')
    return
  }

  let data: XlsxPageData | null = null
  let theme: 'light' | 'dark' = 'light'
  try {
    const info = await window.extension.ura.page.get<{
      id: string
      data: XlsxPageData
    }>(sessionId)
    data = info.data
    theme = await readHostTheme()
    applyTheme(theme)
  } catch (error) {
    fail(
      `XLSX 打开失败：${error instanceof Error ? error.message : String(error)}`,
    )
    return
  }
  if (!data) {
    fail('XLSX 页面数据为空')
    return
  }

  const lang: Lang = 'zh'
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  document.documentElement.dir = htmlDir(lang)

  await installWorkStudioHost({
    sessionId,
    data,
    onDirty: (dirty) =>
      window.extension.bus.send({ type: 'dirty', dirty }),
    onToast: (message) => console.info('[xlsx]', message),
  })

  // 主题变化：WorkStudio bus → 页面 UI（Univer 自己由 App 处理）
  window.extension.bus.on((message) => {
    if (message.type === 'theme') {
      applyTheme(message.theme === 'dark' ? 'dark' : 'light')
    }
  })

  await loadCellFonts()
  ReactDOM.createRoot(root as HTMLElement).render(
    <LocaleProvider initial={lang}>
      <App queuedWorkbookAtBoot={false} />
    </LocaleProvider>,
  )
  window.extension.bus.loaded()
  window.extension.bus.send({ type: 'dirty', dirty: false })
}

void bootstrap()
