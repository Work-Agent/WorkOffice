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
      bytes(sessionId: string, window?: { start?: number; end?: number }): Promise<Uint8Array>
      commit<T = unknown>(sessionId: string, request: unknown): Promise<T>
      agentResult(sessionId: string, message: unknown): Promise<unknown>
      /// 只读工具通道（页面 invoke）。老宿主没有 → 页面回退到本地解析。
      invoke?<T = unknown>(sessionId: string, tool: string, input?: unknown): Promise<T>
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
