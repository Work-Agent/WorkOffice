/// <reference types="vite/client" />

export interface PageSessionInfo {
  id: string
  ref: string
  version: string
  docType: string | null
  surface: string
  owner?: string
  page: string
  data: unknown
  degraded: boolean
}

export interface PageCommitResult {
  ok: boolean
  version?: string
  conflict?: {
    base: string
    current: string
    mine: string
    encoding?: 'text' | 'base64'
  }
}

export interface ExtensionBridge {
  invoke<T = unknown>(command: string, payload?: unknown): Promise<T>
  ura: {
    page: {
      get<T = unknown>(id: string): Promise<T>
      open<T = unknown>(ref: string): Promise<T>
      save<T = unknown>(id: string, content: string): Promise<T>
      saveBytes<T = unknown>(id: string, data: Uint8Array): Promise<T>
      reload<T = unknown>(id: string): Promise<T>
      close(id: string): Promise<unknown>
    }
    watch(ref?: string): Promise<{ ref: string } | null>
    onChange(cb: (event: { ref?: string; version?: string }) => void): () => void
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
    desktop?: Record<string, unknown>
    wordDesktop?: Record<string, unknown>
    slidesApi?: Record<string, unknown>
    desktopApi?: Record<string, unknown>
    projectApi?: Record<string, unknown>
  }
}

export {}
