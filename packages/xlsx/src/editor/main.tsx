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
  const boot = window.__xlsxBoot
  const sessionId = boot?.sessionId ?? new URLSearchParams(location.search).get('ref')
  if (!sessionId) {
    fail('缺少 ?ref=：请从工作区打开一个 .xlsx 文件')
    return
  }
  if (!window.__xlsxPlus) {
    fail('页面桥未就绪：请从 WorkStudio 打开该文件')
    return
  }

  let data: XlsxPageData | null = null
  try {
    data = boot?.data ?? null
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
  applyTheme(await readHostTheme())

  const lang: Lang = 'zh'
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  document.documentElement.dir = htmlDir(lang)

  await installWorkStudioHost({
    sessionId,
    data,
    bytes: window.__xlsxBoot?.bytes,
    epoch: window.__xlsxBoot?.epoch ?? 0,
    onDirty: (dirty) => window.__xlsxPlus?.setDirty(dirty),
    onToast: (message) => window.__xlsxPlus?.toast(message),
    onReceipt: () => undefined,
  })

  // 主题变化：宿主 `shell/appearance` → 页面 UI（Univer 自己由 App 处理）
  window.__xlsxPlus?.onTheme(applyTheme)

  await loadCellFonts()
  ReactDOM.createRoot(root as HTMLElement).render(
    <LocaleProvider initial={lang}>
      <App queuedWorkbookAtBoot={false} />
    </LocaleProvider>,
  )
  window.__xlsxPlus?.setDirty(false)
}

void bootstrap()
