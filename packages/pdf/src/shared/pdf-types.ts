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

export interface PdfPageMeta {
  kind: 'pdf'
  name: string
  ref: string
  size: number
  version: string
  pageCount?: number | null
  encrypted?: boolean | null
  title?: string | null
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

export interface PdfBoot {
  sessionId: string
  data: PdfPageMeta
  bytes: Uint8Array
  version: string
  /** 宿主能力降级（只读提示用）。 */
  degraded?: boolean
  /** 重新按当前会话拉取整份字节（另存副本 / 外部变更重载用）。 */
  refetch(): Promise<Uint8Array>
  /** 重新 get + 分块拉取；返回新字节与新元数据。 */
  reload(): Promise<{ bytes: Uint8Array; data: PdfPageMeta; version: string }>
}
