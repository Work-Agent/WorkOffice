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
  conflict?: unknown
}

export interface ExtensionBridge {
  invoke<T = unknown>(command: string, payload?: unknown): Promise<T>
  ura: {
    page: {
      get<T = unknown>(id: string): Promise<T>
      saveBytes<T = unknown>(id: string, data: Uint8Array): Promise<T>
      reload<T = unknown>(id: string): Promise<T>
    }
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
  }
}

export {}
