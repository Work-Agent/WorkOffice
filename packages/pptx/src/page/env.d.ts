/// <reference types="vite/client" />

export interface PageSessionInfo {
  id: string
  ref: string
  version: string
  /** 命中的 codec id；null = 未经解码（raw）。 */
  codec: string | null
  surface: string
  owner?: string
  page: string
  data: unknown
  degraded: boolean
}

export interface PageCommitResult {
  ok: true
  version: string
}

export interface UraChangeEvent {
  ref: string
  type: 'content' | 'structure' | 'state'
  version?: string
  at: number
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
    onChange(cb: (event: UraChangeEvent) => void): () => void
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
