import { existsSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vite'
// @ts-ignore 解析器是 .mjs：构建脚本与 vite 共用同一份规则
import { viteAliases, viteGenoffice } from './scripts/genoffice-resolve.mjs'

const here = dirname(fileURLToPath(import.meta.url))
const pageDir = join(here, 'src/page')

if (!existsSync(join(pageDir, 'index.html'))) {
  throw new Error(`页面入口不存在：${join(pageDir, 'index.html')}`)
}

export default defineConfig({
  root: pageDir,
  base: './',
  plugins: [viteGenoffice()],
  resolve: {
    alias: [
      // 页面里没有 Node：gateway 的少数 node: 依赖走浏览器 shim（与历史 xlsx 实现同一套）。
      { find: /^node:crypto$/, replacement: join(here, 'src/editor/shims/node-crypto.ts') },
      { find: /^node:fs\/promises$/, replacement: join(here, 'src/editor/shims/node-fs-promises.ts') },
      { find: /^node:fs$/, replacement: join(here, 'src/editor/shims/node-fs.ts') },
      { find: /^node:path$/, replacement: join(here, 'src/editor/shims/node-path.ts') },
      ...viteAliases(),
    ],
  },
  esbuild: { jsx: 'automatic' },
  optimizeDeps: { esbuildOptions: { jsx: 'automatic' } },
  build: {
    outDir: resolve(here, 'dist/page'),
    emptyOutDir: true,
    target: 'chrome140',
    assetsInlineLimit: 0,
  },
})
