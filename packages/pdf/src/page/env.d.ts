/// <reference types="vite/client" />

import type { PdfBoot, PdfPageMeta } from '../shared/pdf-types'

/** 页面桥返回：宿主 preload 不抛错，统一用 CallResult。 */
export interface CallError {
  code: string
  message: string
}

export type CallResult<T = unknown> = { ok: true; data: T } | { ok: false; error: CallError }

export type DomainUpdate =
  | { type: 'snapshot'; domain: string; resource: string; data: unknown }
  | { type: 'patch'; domain: string; resource: string; data: Record<string, unknown> }
  | { type: 'event'; domain: string; resource: string; event: string; data: unknown }

export interface WorkStudioBridge {
  readonly workspaceId: string | null
  command<T = unknown>(
    domain: string,
    resource: string,
    action: string,
    input?: unknown,
  ): Promise<CallResult<T>>
  subscribe(domain: string, resource: string): Promise<unknown>
  unsubscribe(domain: string, resource: string): void
  onUpdate(listener: (update: DomainUpdate) => void): () => void
}

/** pdf.load 的返回：小文件带 bytes（或 data），大文件只带元数据。 */
export interface LoadResult {
  ref: string
  name: string
  size: number
  hash: string
  meta: PdfPageMeta
  inline: boolean
  bytes?: Uint8Array
  data?: string
}

/** pdf.chunk 的返回：bytes 或 base64 兜底。 */
export interface ChunkResult {
  ref: string
  start: number
  end: number
  size: number
  hash: string
  bytes?: Uint8Array
  data?: string
}

export interface SaveFileResult {
  cancelled?: boolean
  token?: string
  path?: string
}

declare global {
  interface Window {
    workStudio: WorkStudioBridge
    __pdfBoot?: PdfBoot
  }
}

export {}
