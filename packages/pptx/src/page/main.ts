// ─── PPTX 只读页面 ──────────────────────────────────────────────────
//
// codec 给原始字节；页面用 GenOffice 的 pptx-engine 解析、
// pptx-render 生成绘制列表，再由我们的 SVG 渲染器画出来。

import { Buffer } from 'buffer'
import { openPptx } from '@genoffice/pptx-engine'
import { buildRenderSlide, type RenderSlide } from '@genoffice/pptx-render'
import type { PageSessionInfo } from './env'
import { renderSlideToSvg } from './render'

interface PptxPageData {
  bytes: string
  name: string
  ref: string
  version: string
  slideCount: number
}

// pptx-engine 的 OOXML 辅助代码使用 Buffer；只读页面用浏览器 polyfill。
;(globalThis as unknown as { Buffer: typeof Buffer }).Buffer = Buffer

const root = document.getElementById('root')!

function fail(message: string): void {
  root.innerHTML = ''
  const boot = document.createElement('div')
  boot.className = 'boot'
  boot.dataset.error = '1'
  boot.textContent = message
  root.append(boot)
}

function applyTheme(theme: 'light' | 'dark'): void {
  document.documentElement.dataset.theme = theme
  document.documentElement.style.colorScheme = theme
}

async function currentTheme(): Promise<'light' | 'dark'> {
  try {
    const ui = await window.extension.invoke<{ theme?: 'light' | 'dark' }>(
      'app:ui.get',
    )
    if (ui?.theme === 'light' || ui?.theme === 'dark') return ui.theme
  } catch {
    /* 系统主题兜底 */
  }
  return window.matchMedia('(prefers-color-scheme: dark)').matches
    ? 'dark'
    : 'light'
}

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

const MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/emf',
  wmf: 'image/wmf',
}

function mimeOf(path: string): string {
  const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
  return MIME[ext] ?? 'application/octet-stream'
}

function mount(data: PptxPageData, slides: RenderSlide[]): void {
  root.innerHTML = ''
  const app = document.createElement('div')
  app.className = 'app'

  const toolbar = document.createElement('div')
  toolbar.className = 'toolbar'
  const name = document.createElement('div')
  name.className = 'file-name'
  name.textContent = data.name
  name.title = data.ref
  const badge = document.createElement('span')
  badge.className = 'badge'
  badge.textContent = '只读预览'
  const spacer = document.createElement('span')
  spacer.className = 'spacer'

  const nav = document.createElement('div')
  nav.className = 'nav'
  const prev = document.createElement('button')
  prev.textContent = '‹'
  prev.title = '上一页'
  const next = document.createElement('button')
  next.textContent = '›'
  next.title = '下一页'
  const counter = document.createElement('span')
  counter.className = 'counter'
  nav.append(prev, counter, next)
  toolbar.append(name, badge, spacer, nav)

  const body = document.createElement('div')
  body.className = 'body'
  const thumbs = document.createElement('div')
  thumbs.className = 'thumbs'
  const stageWrap = document.createElement('div')
  stageWrap.className = 'stage-wrap'
  const stage = document.createElement('div')
  stage.className = 'stage'
  stageWrap.append(stage)
  body.append(thumbs, stageWrap)

  app.append(toolbar, body)
  root.append(app)

  let current = 0
  const thumbButtons: HTMLButtonElement[] = []
  for (let i = 0; i < slides.length; i++) {
    const button = document.createElement('button')
    button.className = 'thumb'
    button.textContent = String(i + 1)
    button.title = `第 ${i + 1} 页`
    button.addEventListener('click', () => select(i))
    thumbButtons.push(button)
    thumbs.append(button)
  }

  const select = (index: number): void => {
    if (index < 0 || index >= slides.length) return
    current = index
    const slide = slides[index]!
    stage.replaceChildren(renderSlideToSvg(slide))
    counter.textContent = `${index + 1} / ${slides.length}`
    prev.disabled = index === 0
    next.disabled = index === slides.length - 1
    for (let i = 0; i < thumbButtons.length; i++) {
      thumbButtons[i]!.classList.toggle('on', i === index)
    }
  }

  prev.addEventListener('click', () => select(current - 1))
  next.addEventListener('click', () => select(current + 1))

  window.addEventListener('keydown', (event) => {
    if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
      event.preventDefault()
      select(current - 1)
    } else if (event.key === 'ArrowRight' || event.key === 'PageDown') {
      event.preventDefault()
      select(current + 1)
    }
  })

  select(0)
}

async function boot(): Promise<void> {
  const sessionId = new URLSearchParams(location.search).get('session')
  if (!sessionId) {
    fail('缺少页面会话：请从 WorkStudio 资源面板打开 PPTX 演示文稿')
    return
  }
  try {
    applyTheme(await currentTheme())
    const info = await window.extension.ura.page.get<PageSessionInfo>(sessionId)
    const data = info.data as PptxPageData | null
    if (!data?.bytes) throw new Error('PPTX 页面数据缺少字节内容')

    const bytes = base64ToBytes(data.bytes)
    const opened = await openPptx(bytes)
    const media = (ref: string): string | undefined => {
      const path = ref.replace(/^\/+/, '')
      const file = opened.archive.readBytes(path)
      if (!file) return undefined
      return `data:${mimeOf(path)};base64,${btoa(
        Array.from(file, (b) => String.fromCharCode(b)).join(''),
      )}`
    }
    const slides = opened.deck.slides.map((slide, index) =>
      buildRenderSlide(slide, opened.deck.size, {
        fitWidthPx: 960,
        media,
        slideNo: index + 1,
      }),
    )
    if (slides.length === 0) throw new Error('演示文稿没有可读幻灯片')
    mount(data, slides)
    window.extension.bus.on((message) => {
      if (message.type === 'theme') {
        applyTheme(message.theme === 'dark' ? 'dark' : 'light')
      }
    })
    window.extension.bus.loaded()
    window.extension.bus.send({ type: 'dirty', dirty: false })
  } catch (error) {
    fail(
      `PPTX 打开失败：${error instanceof Error ? error.message : String(error)}`,
    )
  }
}

void boot()
