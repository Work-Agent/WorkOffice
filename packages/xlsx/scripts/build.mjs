/**
 * 构建：
 *   dist/engine.js           definePlugin（Node ESM，SDK 与网关打进 bundle）
 *   dist/page/**             vite 页面（?ref= 启动）
 *   dist/seed.xlsx           空白工作簿模板（新建用）
 *   commands.json / api.*    workstudio-plugin 从 definePlugin 抽取
 */

import { existsSync } from 'node:fs'
import { mkdir, copyFile, rm } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { build } from 'esbuild'
import { buildPlugin } from '@work-studio/sdk/build'
import { esbuildGenoffice } from './genoffice-resolve.mjs'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const argv = process.argv.slice(2)
const engineOnly = argv.includes('--engine-only')
const pageOnly = argv.includes('--page-only')

/** wasm 只在缺失时构建（首次 ~4 min，之后复用）；失败降级为整份字节模式。 */
async function ensureWasm() {
  if (argv.includes('--no-wasm')) return
  if (existsSync(join(root, 'dist/xlsx-engine.wasm'))) return
  const { spawnSync } = await import('node:child_process')
  const result = spawnSync(process.execPath, [join(root, 'scripts/build-wasm.mjs')], {
    stdio: 'inherit',
  })
  if (result.status !== 0)
    console.warn('wasm 未构建：页面将回退到「读整份字节」模式（build:wasm 重试）')
}

async function buildEngine() {
  await mkdir(join(root, 'dist'), { recursive: true })
  await ensureWasm()
  await build({
    entryPoints: [join(root, 'src/engine/index.ts')],
    outfile: join(root, 'dist/engine.js'),
    bundle: true,
    format: 'esm',
    platform: 'node',
    target: 'node22',
    external: ['node:*'],
    plugins: [esbuildGenoffice()],
    logLevel: 'info',
  })
  await copyFile(join(root, 'assets/seed.xlsx'), join(root, 'dist/seed.xlsx'))
}

async function buildPage() {
  const vite = await import('vite')
  await vite.build({ configFile: join(root, 'vite.config.ts') })
}

if (!pageOnly) await buildEngine()
if (!engineOnly) await buildPage()

if (!pageOnly) {
  const result = await buildPlugin(root)
  console.log(`commands: ${result.commands.length}`)
  for (const file of result.files) console.log(`wrote ${file.replace(root, '.')}`)
}

if (argv.includes('--clean')) await rm(join(root, 'dist/page'), { recursive: true, force: true })
