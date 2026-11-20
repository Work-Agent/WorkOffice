// ─── PDF.js 加载 ─────────────────────────────────────────────────────
//
// 使用 legacy build（与 GenOffice 一致）：新版 V8 API 会让部分字体的
// 字形解析失败。`isEvalSupported:false` 关闭 PDF 内嵌 JS；
// `useSystemFonts:false` 保证同一份 PDF 在不同机器上渲染一致。

import type { PDFDocumentLoadingTask } from 'pdfjs-dist'
import { GlobalWorkerOptions, getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import { RENDER_POLICY } from '../shared/render-policy'
import { ASSET_BASE } from './view-config'

GlobalWorkerOptions.workerSrc = workerUrl

export interface LoadOptions {
  password?: string
  onPassword?: (updatePassword: (password: string) => void, wrong: boolean) => void
  onProgress?: (loaded: number, total: number) => void
}

export function loadPdf(bytes: Uint8Array, options: LoadOptions = {}): PDFDocumentLoadingTask {
  const task = getDocument({
    data: bytes,
    ...(options.password ? { password: options.password } : {}),
    cMapUrl: `${ASSET_BASE}cmaps/`,
    cMapPacked: true,
    standardFontDataUrl: `${ASSET_BASE}standard_fonts/`,
    wasmUrl: `${ASSET_BASE}wasm/`,
    iccUrl: `${ASSET_BASE}iccs/`,
    // PDF.js 6 不再提供 isEvalSupported（内嵌 JS 始终不执行）。
    useSystemFonts: RENDER_POLICY.useSystemFonts,
    enableXfa: RENDER_POLICY.enableXfa,
  })
  task.onPassword = (updatePassword: (password: string) => void, reason: number) => {
    // reason 1 = NEED_PASSWORD，2 = INCORRECT_PASSWORD（PasswordResponses）。
    options.onPassword?.(updatePassword, reason === 2)
  }
  if (options.onProgress) {
    task.onProgress = ({ loaded, total }: { loaded: number; total: number }) =>
      options.onProgress?.(loaded, total)
  }
  return task
}
