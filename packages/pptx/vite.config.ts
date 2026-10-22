import { existsSync, statSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { defineConfig, type Plugin } from 'vite'

const vendor = resolve(__dirname, '../../vendor/genoffice')

function genofficeSource(): Plugin {
  return {
    name: 'genoffice-source',
    enforce: 'pre',
    resolveId(source) {
      if (!source.startsWith('@genoffice/')) return null
      const rest = source.slice('@genoffice/'.length)
      const slash = rest.indexOf('/')
      const pkg = slash < 0 ? rest : rest.slice(0, slash)
      const sub = slash < 0 ? '' : rest.slice(slash + 1)
      const base = join(vendor, 'packages', pkg, 'src', sub)
      for (const candidate of [
        base,
        `${base}.ts`,
        `${base}.tsx`,
        `${base}.css`,
        join(base, 'index.ts'),
        join(base, 'index.tsx'),
      ]) {
        if (existsSync(candidate) && statSync(candidate).isFile()) return candidate
      }
      return null
    },
  }
}

export default defineConfig({
  root: 'src/page',
  base: './',
  plugins: [genofficeSource()],
  resolve: {
    alias: [
      {
        find: /^node:crypto$/,
        replacement: join(__dirname, 'src/page/node-crypto.ts'),
      },
      {
        find: /^node:zlib$/,
        replacement: join(__dirname, 'src/page/node-zlib.ts'),
      },
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
  build: {
    outDir: '../../dist/page',
    emptyOutDir: true,
    target: 'chrome140',
    assetsInlineLimit: 0,
  },
})
