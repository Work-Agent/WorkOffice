#!/usr/bin/env node
// docx 构建：esbuild 引擎 → workstudio-plugin 等价的 commands.json → vite 页面。
import { build } from 'esbuild'
import { copyFile, mkdir, rm } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const argv = process.argv.slice(2)
const engineOnly = argv.includes('--engine-only')
const pageOnly = argv.includes('--page-only')

const NODE = {
  bundle: true,
  format: 'esm',
  platform: 'node',
  target: 'node22',
  sourcemap: false,
  logLevel: 'info',
  legalComments: 'none',
}

async function buildEngine() {
  await mkdir(join(root, 'dist'), { recursive: true })
  await build({
    ...NODE,
    entryPoints: [join(root, 'src/engine/index.ts')],
    outfile: join(root, 'dist/engine.js'),
  })
  await copyFile(join(root, 'assets/seed.docx'), join(root, 'dist/seed.docx'))
  console.log('built dist/engine.js + seed.docx')
  const generator = await import(pathToFileURL(join(root, 'scripts/gen-commands.mjs')).href)
  void generator
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
