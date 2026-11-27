/// <reference types="vite/client" />

declare module '*.md?raw' {
  const content: string
  export default content
}

import type { DesktopApi } from '../shared/desktop-api'
import type { ProjectApi } from '@genoffice/project-store'
import type { EditorBridge } from '@xlsx/page/bridge'
import type { SheetLike } from '@xlsx/shared/sheet-types'

/** 页面注入的启动数据（见 src/page/main.ts）。 */
export interface XlsxProBoot {
  sessionId: string
  data: {
    name: string
    ref: string
    version: string
    size?: number
    sheets: SheetLike[]
  }
  epoch: number
  bytes?: Uint8Array
}

/** App.tsx 暴露给页面代理的编辑器句柄（P3 由 agent 桥使用）。 */
export interface XlsxProEditorBridge {
  context(): unknown
  readCells(addresses: string[], sheetId?: string): unknown
  sheets(): ReadonlyArray<{ id: string; name: string }>
  focusSheet(sheetId: string, address?: string): void
  applyOps(ops: unknown[], dryRun: boolean): Promise<unknown>
  save(): Promise<{ ok: boolean; path?: string; error?: string }>
  docRev(): number
  structureEpoch(): number
  undo(): boolean
  undoDepth(): number
}

declare global {
  interface Window {
    readonly desktopApi: DesktopApi
    readonly projectApi: ProjectApi
    /** 编辑器唯一的 I/O 出口：取字节、保存、主题、脏态。 */
    __xlsxPlus?: EditorBridge
    __xlsxProBoot?: XlsxProBoot
    __xlsxPro?: any
  }
}

export {}
