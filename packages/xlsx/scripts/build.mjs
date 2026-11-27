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

/**
 * wasm 引擎就位，见 scripts/ensure-wasm.mjs：
 *   默认 XLSX_ENGINE_WASM=auto —— 先拿 wasm.lock.json 固定的产物（缓存 / 下载），
 *   拿不到再用 Rust 源码编译（编译器版本见 wasm.lock.json 的 compile.rustc）。
 *   XLSX_ENGINE_WASM=off / --no-wasm 跳过；XLSX_REQUIRE_WASM=1 拿不到就失败。
 */
async function ensureWasm() {
  if (argv.includes('--no-wasm')) return
  const { spawnSync } = await import('node:child_process')
  const args = [join(root, 'scripts/ensure-wasm.mjs')]
  if (process.env['XLSX_REQUIRE_WASM'] === '1') args.push('--require')
  const result = spawnSync(process.execPath, args, { stdio: 'inherit' })
  if (result.status !== 0) throw new Error('wasm 引擎缺失（node scripts/ensure-wasm.mjs 重试）')
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
