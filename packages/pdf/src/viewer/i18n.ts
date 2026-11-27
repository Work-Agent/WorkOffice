// ─── 一期 zh / en 内联字典（不引 @genoffice/i18n） ──────────────────

export type Lang = 'zh' | 'en'

type Dict = Record<string, string>

const zh: Dict = {
  loading: '正在打开 PDF…',
  loadingProgress: '正在载入 {percent}%',
  errorTitle: 'PDF 打开失败',
  retry: '重试',
  passwordRequired: '此 PDF 受密码保护',
  passwordWrong: '密码错误，请重试',
  passwordPlaceholder: '请输入打开密码',
  passwordSubmit: '打开',
  passwordCancel: '取消',
  passwordCancelled: '已取消密码输入',
  passwordHint: '密码只在页面内存中使用，不会写入文件或日志。',
  readOnly: '只读预览',
  pageInput: '页码',
  pageOf: '{current} / {total}',
  previousPage: '上一页',
  nextPage: '下一页',
  zoomOut: '缩小',
  zoomIn: '放大',
  zoomLevel: '缩放比例',
  fitWidth: '适合宽度',
  fitPage: '适合页面',
  rotate: '旋转视图（仅查看）',
  thumbnails: '缩略图',
  outline: '大纲',
  search: '搜索',
  print: '打印',
  saveCopy: '另存副本',
  dismiss: '关闭提示',
  fileUpdated: '文件已在外部更新，重载后保留当前页码与缩放。',
  reload: '重载',
  reloading: '正在重载…',
  largeFileMetadata: '文件较大，已跳过标题/作者等元数据（不影响阅读）。',
  metadataUnavailable: '无法读取文件元数据，已交给 PDF.js 恢复。',
  noOutline: '此 PDF 没有书签大纲。',
  noResults: '没有匹配结果',
  searching: '正在索引… {percent}%',
  searchPlaceholder: '在当前 PDF 中搜索',
  matchCase: '区分大小写',
  wholeWord: '全词匹配',
  results: '{count} 个结果',
  nextMatch: '下一个',
  prevMatch: '上一个',
  pageRenderFailed: '第 {page} 页无法渲染',
  printProgress: '正在准备打印 {done}/{total}',
  printFailed: '打印失败：{message}',
  saveCopyLarge: '文件超过 64 MB，已禁用「另存副本」。',
  saveCopyFailed: '另存副本失败：{message}',
  noTextLayer: '此页没有文本层（可能是扫描件）。',
  emptyTitle: '没有可显示的 PDF',
  emptyHint: '请从 WorkStudio 资源面板打开一个 .pdf 文件。',
  openFailed: 'PDF 打开失败：{message}',
  unsupportedEncryption: '不支持的加密方式（如证书加密）。',
  sidebar: '侧栏',
}

const en: Dict = {
  loading: 'Opening PDF…',
  loadingProgress: 'Loading {percent}%',
  errorTitle: 'Could not open PDF',
  retry: 'Retry',
  passwordRequired: 'This PDF is password protected',
  passwordWrong: 'Incorrect password, try again',
  passwordPlaceholder: 'Enter password',
  passwordSubmit: 'Open',
  passwordCancel: 'Cancel',
  passwordCancelled: 'Password entry cancelled',
  passwordHint: 'The password stays in page memory; it is never written or logged.',
  readOnly: 'Read-only preview',
  pageInput: 'Page number',
  pageOf: '{current} / {total}',
  previousPage: 'Previous page',
  nextPage: 'Next page',
  zoomOut: 'Zoom out',
  zoomIn: 'Zoom in',
  zoomLevel: 'Zoom level',
  fitWidth: 'Fit width',
  fitPage: 'Fit page',
  rotate: 'Rotate view (view only)',
  thumbnails: 'Thumbnails',
  outline: 'Outline',
  search: 'Search',
  print: 'Print',
  saveCopy: 'Save a copy',
  dismiss: 'Dismiss',
  fileUpdated: 'The file changed on disk. Reloading keeps the current page and zoom.',
  reload: 'Reload',
  reloading: 'Reloading…',
  largeFileMetadata: 'Large file: title/author metadata was skipped (reading is unaffected).',
  metadataUnavailable: 'Metadata could not be read; PDF.js will recover the document.',
  noOutline: 'This PDF has no outline.',
  noResults: 'No matches',
  searching: 'Indexing… {percent}%',
  searchPlaceholder: 'Search this PDF',
  matchCase: 'Match case',
  wholeWord: 'Whole word',
  results: '{count} results',
  nextMatch: 'Next',
  prevMatch: 'Previous',
  pageRenderFailed: 'Page {page} could not be rendered',
  printProgress: 'Preparing print {done}/{total}',
  printFailed: 'Print failed: {message}',
  saveCopyLarge: 'Files larger than 64 MB cannot be saved as a copy here.',
  saveCopyFailed: 'Save a copy failed: {message}',
  noTextLayer: 'This page has no text layer (possibly a scan).',
  emptyTitle: 'No PDF to display',
  emptyHint: 'Open a .pdf file from the WorkStudio resource panel.',
  openFailed: 'Could not open PDF: {message}',
  unsupportedEncryption: 'Unsupported encryption (e.g. certificate-based).',
  sidebar: 'Sidebar',
}

export interface I18n {
  lang: Lang
  t(key: string, vars?: Record<string, string | number>): string
}

export function createI18n(): I18n {
  const lang: Lang =
    typeof navigator !== 'undefined' && !navigator.language.toLowerCase().startsWith('zh')
      ? 'en'
      : 'zh'
  const dict = lang === 'zh' ? zh : en
  return {
    lang,
    t(key, vars) {
      const template = dict[key] ?? zh[key] ?? key
      if (!vars) return template
      return template.replace(/\{(\w+)\}/g, (_match, name: string) =>
        vars[name] === undefined ? `{${name}}` : String(vars[name]),
      )
    },
  }
}
