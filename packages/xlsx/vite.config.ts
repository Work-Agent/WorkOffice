import { defineConfig } from 'vite'

export default defineConfig({
  root: 'src/page',
  base: './',
  build: {
    outDir: '../../dist/page',
    emptyOutDir: true,
    target: 'chrome140',
    assetsInlineLimit: 0,
  },
})
