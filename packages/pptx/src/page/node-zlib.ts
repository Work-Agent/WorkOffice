// 只读渲染不会压缩/解压；pptx-engine 的 media-insert 导入该函数，
// 浏览器侧保留一个无操作替身，避免 node:zlib 进入包。
export function deflateSync(data: Uint8Array): Uint8Array {
  return data
}
