/// <reference types="vite/client" />

import type { PdfBoot, PdfPageMeta } from '../shared/pdf-types'

export interface PageSessionInfo {
  id: string
  ref: string
  version: string
  epoch: number
  codec: string | null
  surface: string
  owner?: string
  page: string
  data: PdfPageMeta
  degraded: boolean
  write: 'none' | 'bytes' | 'plan'
  size?: number
  tier?: 'normal' | 'lite' | 'readOnly' | 'refused'
}

export interface ExtensionBridge {
  ura: {
    page: {
      get<T = unknown>(id: string): Promise<T>
      bytes(id: string, window?: { start?: number; end?: number }): Promise<Uint8Array>
      close(id: string): Promise<unknown>
    }
    onChange(
      cb: (event: { ref: string; type: string; version?: string; at: number }) => void,
    ): () => void
  }
  bus: {
    on(cb: (message: { type: string; [key: string]: unknown }) => void): () => void
    send(message: { type: string; [key: string]: unknown }): void
    loaded(): void
  }
}

declare global {
  interface Window {
    extension: ExtensionBridge
    appBridge?: {
      available(): Promise<boolean>
      call<T = unknown>(path: string, input?: unknown): Promise<T>
      peek<T = unknown>(path: string): Promise<T | null>
    }
    __pdfBoot?: PdfBoot
  }
}

export {}
