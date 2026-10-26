/// <reference types="vite/client" />
import type { SlidesApi } from '../shared/ipc'
import type { ProjectApi } from '@genoffice/project-store'

interface WorkStudioBus {
  on(handler: (message: any) => void): () => void
  send(message: any): void
  loaded(): void
}

interface WorkStudioExtension {
  invoke<T = unknown>(command: string, payload?: unknown): Promise<T>
  bus: WorkStudioBus
  ura: {
    page: {
      get<T = unknown>(sessionId?: string): Promise<T>
      saveBytes<T = unknown>(sessionId: string, bytes: Uint8Array): Promise<T>
    }
  }
}

interface WorkStudioDesktop {
  getPathForFile(file: File): string
  [key: string]: unknown
}

declare global {
  interface Window {
    slidesApi: SlidesApi
    projectApi: ProjectApi
    extension: WorkStudioExtension
    desktop: WorkStudioDesktop
  }
}

export {}
