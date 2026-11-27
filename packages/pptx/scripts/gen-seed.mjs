#!/usr/bin/env node
// 用 pptx-engine 的 createBlankPptx() 生成 assets/seed.pptx。
// vendor 是 TS 源码；用 esbuild 临时打包后 import。

import { build } from 'esbuild'
import { mkdir, writeFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const blank = resolve(root, '../../vendor/genoffice/packages/pptx-engine/src/blank.ts')

const result = await build({
  entryPoints: [blank],
  absWorkingDir: root,
  bundle: true,
  platform: 'node',
  format: 'esm',
  target: 'node22',
  write: false,
  logLevel: 'silent',
})

const code = result.outputFiles?.[0]?.text
if (!code) throw new Error('seed bundle 为空')
const mod = await import(
  `data:text/javascript;base64,${Buffer.from(code).toString('base64')}`
)
const bytes = await mod.createBlankPptx()
await mkdir(join(root, 'assets'), { recursive: true })
await writeFile(join(root, 'assets/seed.pptx'), bytes)
console.log(`generated assets/seed.pptx (${bytes.byteLength} bytes)`)
