import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// PDF 查看器不依赖 GenOffice 源码：页面只用 pdfjs-dist + React。
// CMaps / standard_fonts / wasm 由 scripts/build.mjs 复制到
// src/page/pdfjs（dev）与 dist/page/pdfjs（build）。
export default defineConfig({
  root: 'src/page',
  base: './',
  plugins: [react()],
  esbuild: { jsx: 'automatic' },
  optimizeDeps: { esbuildOptions: { jsx: 'automatic' } },
  build: {
    outDir: '../../dist/page',
    emptyOutDir: true,
    target: 'chrome140',
    assetsInlineLimit: 0,
  },
})
