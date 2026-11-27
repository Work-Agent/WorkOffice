#!/usr/bin/env node
// pptx 构建：esbuild 引擎 → 静态命令描述 → vite 页面。
//
// 产物：
//   dist/engine.js     definePlugin（宿主 exports['./workstudio']）
//   dist/seed.pptx     create 的空白模板（与 engine.js 同目录）
//   dist/page/**       ?ref= 页面
//   commands.json / api.json / api.d.ts   （gen-commands 生成，提交进仓库）

import { build } from 'esbuild'
import { copyFile, mkdir, rm } from 'node:fs/promises'
import { existsSync, readFileSync, statSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const vendor = resolve(root, '../../vendor/genoffice')
const argv = process.argv.slice(2)
const engineOnly = argv.includes('--engine-only')
const pageOnly = argv.includes('--page-only')

/** `?raw` 导入：把文本文件当字符串打进 bundle。 */
function rawPlugin() {
  return {
    name: 'raw-text',
    setup(build) {
      build.onResolve({ filter: /\?raw$/ }, (args) => ({
        path: resolve(dirname(args.importer), args.path.replace(/\?raw$/, '')),
        namespace: 'raw-text',
      }))
      build.onLoad({ filter: /.*/, namespace: 'raw-text' }, (args) => ({
        contents: readFileSync(args.path, 'utf8'),
        loader: 'text',
      }))
    },
  }
}

/** `@genoffice/*` 直接吃 vendor 源码（与 vite.config.ts 同一映射）。 */
function genofficeSource() {
  return {
    name: 'genoffice-source',
    setup(build) {
      build.onResolve({ filter: /^@genoffice\// }, (args) => {
        const rest = args.path.slice('@genoffice/'.length)
        const slash = rest.indexOf('/')
        const pkg = slash < 0 ? rest : rest.slice(0, slash)
        const sub = slash < 0 ? '' : rest.slice(slash + 1)
        const base = join(vendor, 'packages', pkg, 'src', sub)
        for (const candidate of [
          base,
          `${base}.ts`,
          `${base}.tsx`,
          join(base, 'index.ts'),
          join(base, 'index.tsx'),
        ]) {
          if (existsSync(candidate) && statSync(candidate).isFile()) return { path: candidate }
        }
        return null
      })
    },
  }
}

const NODE = {
  bundle: true,
  format: 'esm',
  platform: 'node',
  target: 'node22',
  sourcemap: false,
  logLevel: 'info',
  legalComments: 'none',
  banner: {
    js: "import { createRequire as __createRequire } from 'node:module'; const require = __createRequire(import.meta.url);",
  },
  plugins: [rawPlugin(), genofficeSource()],
}

async function buildEngine() {
  await mkdir(join(root, 'dist'), { recursive: true })
  await build({
    ...NODE,
    entryPoints: [join(root, 'src/engine/index.ts')],
    outfile: join(root, 'dist/engine.js'),
  })
  await copyFile(join(root, 'assets/seed.pptx'), join(root, 'dist/seed.pptx'))
  console.log('built dist/engine.js + seed.pptx')
  const generator = await import(pathToFileURL(join(root, 'scripts/gen-commands.mjs')).href)
  await generator.generate?.()
}

async function buildPage() {
  const vite = await import('vite')
  await vite.build({ configFile: join(root, 'vite.config.ts') })
  console.log('built dist/page')
}

if (pageOnly) {
  await buildPage()
} else if (engineOnly) {
  await buildEngine()
} else {
  await rm(join(root, 'dist'), { recursive: true, force: true })
  await buildEngine()
  await buildPage()
}
