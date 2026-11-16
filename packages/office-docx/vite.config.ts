import { existsSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';
import react from '@vitejs/plugin-react';
import { defineConfig, type Plugin } from 'vite';

// 页面 v2（React + 本地渲染器 fork）。
// 与 packages/docx 的 vite 配置同构：@genoffice/* 走 vendor 源码；渲染器源码在
// `src/page-v2/genoffice/`（**复制**，不跨包引用），所以不需要额外 alias。

const vendor = resolve(__dirname, '../../vendor/genoffice');

/** 官方 GenOffice 包以源码形式在 monorepo 内互相引用；按包名映射到 vendor 源码。 */
function genofficeSource(): Plugin {
  return {
    name: 'genoffice-source',
    enforce: 'pre',
    resolveId(source) {
      if (!source.startsWith('@genoffice/')) return null;
      const rest = source.slice('@genoffice/'.length);
      const slash = rest.indexOf('/');
      const pkg = slash < 0 ? rest : rest.slice(0, slash);
      const sub = slash < 0 ? '' : rest.slice(slash + 1);
      const base = join(vendor, 'packages', pkg, 'src', sub);
      for (const candidate of [
        base,
        `${base}.ts`,
        `${base}.tsx`,
        `${base}.css`,
        join(base, 'index.ts'),
        join(base, 'index.tsx'),
      ]) {
        if (existsSync(candidate) && statSync(candidate).isFile()) return candidate;
      }
      return null;
    },
  };
}

export default defineConfig({
  root: 'src/page-v2',
  base: './',
  plugins: [genofficeSource(), react()],
  resolve: {
    alias: [
      {
        find: /^@genoffice\/([^/]+)\/(.*)$/,
        replacement: join(vendor, 'packages/$1/src/$2'),
      },
      {
        find: /^@genoffice\/([^/]+)$/,
        replacement: join(vendor, 'packages/$1/src/index.ts'),
      },
    ],
  },
  esbuild: { jsx: 'automatic' },
  optimizeDeps: { esbuildOptions: { jsx: 'automatic' } },
  build: {
    outDir: '../../dist/page-v2',
    emptyOutDir: true,
    target: 'chrome140',
    assetsInlineLimit: 0,
  },
});
