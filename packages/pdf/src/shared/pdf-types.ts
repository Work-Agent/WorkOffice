// ─── PDF 插件共享类型 ────────────────────────────────────────────────

export interface PdfDoc {
  kind: 'pdf'
  name: string
  size: number
  version: string
  headerVersion?: string
  pageCount?: number
  encrypted?: boolean
  title?: string
  author?: string
  firstPageSize?: { width: number; height: number }
  metadataWarning?: string
}

/** 页面拿到的元数据：不携带字节，字节经 load / chunk 单独取。 */
export interface PdfPageMeta {
  kind: 'pdf'
  name: string
  ref: string
  size: number
  version: string
  pageCount?: number | null
  encrypted?: boolean | null
  title?: string | null
  author?: string | null
  headerVersion?: string
  metadataWarning?: string
}

export interface PageGeom {
  /** 显示宽度（CropBox，已乘 /UserUnit）。 */
  pw: number
  /** 显示高度（CropBox，已乘 /UserUnit）。 */
  ph: number
  /** 页面自身 /Rotate（视图旋转由页面层叠加）。 */
  rot: number
  /** CropBox 左下角（PDF user space）。 */
  x0?: number
  y0?: number
  /** PDF /UserUnit；布局坐标乘它，PDF API 坐标不乘。 */
  userUnit?: number
  mediaBox?: [number, number, number, number]
  cropBox?: [number, number, number, number]
}

/**
 * 页面启动数据：page/main.ts 注入，viewer 只读。
 *
 * 没有旧宿主的页面会话：元数据直接来自 pdf.load。
 */
export interface PdfBoot {
  ref: string
  data: PdfPageMeta
  bytes: Uint8Array
  version: string
  /** 资源打开时带的定位（pdf:page:12 / #p=12）。 */
  locator: string
  /** 重新取整份字节（另存副本 / 外部变更重载用）。 */
  refetch(): Promise<Uint8Array>
  /** 重新取元数据与字节。 */
  reload(): Promise<{ bytes: Uint8Array; data: PdfPageMeta; version: string }>
  /** 订阅同一 ref 的 changed 事件（只读：只提示，不产生写冲突）。 */
  watchExternal(handler: () => void): () => void
  /** 用页面当前字节写宿主 save token；返回 false 表示用户取消。 */
  saveCopy(bytes: Uint8Array): Promise<boolean>
}
