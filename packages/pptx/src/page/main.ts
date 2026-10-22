// ─── PPTX 只读页面 ──────────────────────────────────────────────────
//
// codec 给原始字节；页面用 GenOffice 的 pptx-engine 解析、
// pptx-render 生成绘制列表，再由我们的 SVG 渲染器画出来。

import { Buffer } from 'buffer'
import { openPptx, savePptx, type OpenedPptx } from '@genoffice/pptx-engine'
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

function buildSlides(opened: OpenedPptx): RenderSlide[] {
  const media = (ref: string): string | undefined => {
    const path = ref.replace(/^\/+/, '')
    const file = opened.archive.readBytes(path)
    if (!file) return undefined
    return `data:${mimeOf(path)};base64,${btoa(
      Array.from(file, (b) => String.fromCharCode(b)).join(''),
    )}`
  }
  return opened.deck.slides.map((slide, index) =>
    buildRenderSlide(slide, opened.deck.size, {
      fitWidthPx: 960,
      media,
      slideNo: index + 1,
    }),
  )
}

interface TextElementLike {
  id?: string
  type?: string
  dirty?: boolean
  text?: {
    paragraphs: Array<{ runs?: Array<{ text?: string; [key: string]: unknown }>; [key: string]: unknown }>
    [key: string]: unknown
  }
  children?: TextElementLike[]
}

function findElement(elements: TextElementLike[], id: string): TextElementLike | null {
  for (const element of elements) {
    if (element.id === id) return element
    if (element.type === 'group' && element.children) {
      const found = findElement(element.children, id)
      if (found) return found
    }
  }
  return null
}

function elementText(element: TextElementLike): string {
  if (!element.text) return ''
  return element.text.paragraphs
    .map((paragraph) => (paragraph.runs ?? []).map((run) => run.text ?? '').join(''))
    .join('\n')
}

function setElementText(element: TextElementLike, value: string): boolean {
  if (!element.text) return false
  const template = element.text.paragraphs[0]
  const templateRun = template?.runs?.[0] ?? { text: '' }
  element.text.paragraphs = value.split('\n').map((line) => ({
    ...(template ?? {}),
    runs: [{ ...templateRun, text: line }],
  }))
  element.dirty = true
  return true
}

function mount(
  data: PptxPageData,
  sessionId: string,
  initial: OpenedPptx,
): void {
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
  badge.textContent = '双击文本编辑'
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

  let doc = initial
  let slides = buildSlides(doc)
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
    stage.replaceChildren(renderSlideToSvg(slides[index]!))
    counter.textContent = `${index + 1} / ${slides.length}`
    prev.disabled = index === 0
    next.disabled = index === slides.length - 1
    for (let i = 0; i < thumbButtons.length; i++) {
      thumbButtons[i]!.classList.toggle('on', i === index)
    }
  }

  const openTextEditor = (element: TextElementLike): void => {
    const overlay = document.createElement('div')
    overlay.className = 'text-editor'
    const panel = document.createElement('div')
    panel.className = 'text-editor-panel'
    const title = document.createElement('div')
    title.className = 'text-editor-title'
    title.textContent = '编辑文本'
    const textarea = document.createElement('textarea')
    textarea.value = elementText(element)
    const actions = document.createElement('div')
    actions.className = 'text-editor-actions'
    const cancel = document.createElement('button')
    cancel.textContent = '取消'
    const apply = document.createElement('button')
    apply.className = 'primary'
    apply.textContent = '应用并保存'
    const status = document.createElement('span')
    status.className = 'text-editor-status'
    actions.append(status, cancel, apply)
    panel.append(title, textarea, actions)
    overlay.append(panel)
    document.body.append(overlay)
    textarea.focus()
    textarea.select()

    const close = (): void => overlay.remove()
    cancel.addEventListener('click', close)
    overlay.addEventListener('click', (event) => {
      if (event.target === overlay) close()
    })
    apply.addEventListener('click', () => {
      if (!setElementText(element, textarea.value)) {
        status.textContent = '该元素没有文本模型'
        return
      }
      apply.disabled = true
      status.textContent = '保存中…'
      void (async () => {
        try {
          const next = await savePptx(doc)
          const result = await window.extension.ura.page.saveBytes<{
            ok: boolean
            conflict?: unknown
          }>(sessionId, next)
          if (!result?.ok) {
            status.textContent = result?.conflict ? '文件已被外部修改' : '保存失败'
            apply.disabled = false
            return
          }
          doc = await openPptx(next)
          slides = buildSlides(doc)
          close()
          select(current)
          window.extension.bus.send({ type: 'dirty', dirty: false })
        } catch (error) {
          status.textContent = `保存失败：${error instanceof Error ? error.message : String(error)}`
          apply.disabled = false
        }
      })()
    })
  }

  stage.addEventListener('dblclick', (event) => {
    const target = (event.target as Element | null)?.closest('[data-source-id]')
    const id = target?.getAttribute('data-source-id')
    if (!id) return
    const slide = doc.deck.slides[current] as unknown as {
      elements: TextElementLike[]
    }
    const element = findElement(slide?.elements ?? [], id)
    if (element?.text) {
      event.preventDefault()
      openTextEditor(element)
    }
  })

  prev.addEventListener('click', () => select(current - 1))
  next.addEventListener('click', () => select(current + 1))

  window.addEventListener('keydown', (event) => {
    if (event.target instanceof HTMLTextAreaElement) return
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
    if (opened.deck.slides.length === 0) throw new Error('演示文稿没有可读幻灯片')
    mount(data, sessionId, opened)
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
