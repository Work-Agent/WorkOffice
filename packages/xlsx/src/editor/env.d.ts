/// <reference types="vite/client" />

declare module '*.md?raw' {
  const content: string
  export default content
}

import type { DesktopApi } from '../shared/desktop-api'
import type { ProjectApi } from '@genoffice/project-store'

interface WorkStudioBus {
  on(handler: (message: any) => void): () => void
  send(message: any): void
  loaded(): void
}

interface WorkStudioExtension {
  bus: WorkStudioBus
  ura: {
    page: {
      get<T = unknown>(sessionId?: string): Promise<T>
      saveBytes<T = unknown>(sessionId: string, bytes: Uint8Array): Promise<T>
    }
  }
}

interface WorkStudioAppBridge {
  available(): Promise<boolean>
  call<T = unknown>(path: string, input?: unknown): Promise<T>
  peek<T = unknown>(path: string): Promise<T | null>
}

declare global {
  interface Window {
    readonly desktopApi: DesktopApi
    readonly projectApi: ProjectApi
    readonly extension: WorkStudioExtension
    readonly appBridge: WorkStudioAppBridge
  }
}

export {}
