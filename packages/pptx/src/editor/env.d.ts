/// <reference types="vite/client" />
import type { SlidesApi } from '../shared/ipc'
import type { ProjectApi } from '@genoffice/project-store'

interface WorkStudioDesktop {
  getPathForFile(file: File): string
  [key: string]: unknown
}

declare global {
  interface Window {
    slidesApi: SlidesApi
    projectApi: ProjectApi
    desktop: WorkStudioDesktop
  }
}

export {}
