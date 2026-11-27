/// <reference types="vite/client" />

/** 宿主注入的编辑器句柄（App.tsx 设置）。 */
declare global {
  interface Window {
    __xlsxProHost?: {
      getVersion?: () => string
      getEpoch?: () => number
      getReceipt?: () => unknown
      getBytes?: () => Uint8Array
    }
  }
}

export {}
