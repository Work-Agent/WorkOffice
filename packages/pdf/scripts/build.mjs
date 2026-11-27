#!/usr/bin/env node
// PDF 插件构建：
//   1. esbuild src/engine/{index,text}.ts → dist/engine.js + dist/text.js
//      （code splitting；text.js 由 worker_threads 动态加载）
//   2. 复制 pdfjs 的 fake worker → dist/pdf.worker.mjs
//   3. 从 definePlugin 生成 commands.json / api.json / api.d.ts
//   4. esbuild 纯函数 → dist/shared/*.mjs（node --test 用）
//   5. vite build → dist/page/**
//   6. 复制 pdfjs cmaps / standard_fonts / wasm / iccs → dist/page/assets/pdfjs
//
// 参数：
//   --engine-only  只构建 engine + text + worker + 静态描述
//   --page-only    只构建页面（仍复制静态资源）
//   --assets-only  只复制静态资源到 src/page/assets（vite dev 用）
//   --shared-only  只构建 dist/shared（测试用）

import { execFileSync } from 'node:child_process'
import {
  cpSync,
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { build } from 'esbuild'

const pkgRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const repoRoot = resolve(pkgRoot, '../..')
const args = process.argv.slice(2)
const has = (flag) => args.includes(flag)
const engineOnly = has('--engine-only')
const pageOnly = has('--page-only')
const assetsOnly = has('--assets-only')
const sharedOnly = has('--shared-only')

function resolvePdfjsDist() {
  const candidates = [
    join(repoRoot, 'node_modules/pdfjs-dist'),
    join(pkgRoot, 'node_modules/pdfjs-dist'),
  ]
  for (const candidate of candidates) {
    if (existsSync(join(candidate, 'package.json'))) return candidate
  }
  throw new Error('找不到 pdfjs-dist：请先运行 pnpm install')
}

function copyDir(from, to) {
  if (!existsSync(from)) return
  mkdirSync(dirname(to), { recursive: true })
  cpSync(from, to, {
    recursive: true,
    // 运行时不需要 sourcemap，zip 体积也小一半。
    filter: (file) => !file.endsWith('.map'),
  })
}

function copyPdfjsAssets(targetPage) {
  const dist = resolvePdfjsDist()
  const out = join(targetPage, 'assets', 'pdfjs')
  rmSync(out, { recursive: true, force: true })
  for (const dir of ['cmaps', 'standard_fonts', 'wasm', 'iccs']) {
    copyDir(join(dist, dir), join(out, dir))
  }
  console.log(`copied pdfjs assets → ${out}`)
}

function copyPdfjsWorker() {
  const dist = resolvePdfjsDist()
  mkdirSync(join(pkgRoot, 'dist'), { recursive: true })
  cpSync(join(dist, 'legacy/build/pdf.worker.mjs'), join(pkgRoot, 'dist/pdf.worker.mjs'))
  console.log('copied pdf.worker.mjs → dist/pdf.worker.mjs')
}

async function buildEngine() {
  mkdirSync(join(pkgRoot, 'dist'), { recursive: true })
  await build({
    entryPoints: {
      engine: join(pkgRoot, 'src/engine/index.ts'),
      text: join(pkgRoot, 'src/engine/text.ts'),
    },
    outdir: join(pkgRoot, 'dist'),
    bundle: true,
    splitting: true,
    format: 'esm',
    platform: 'node',
    target: 'node22',
    entryNames: '[name]',
    chunkNames: 'chunks/[name]-[hash]',
    sourcemap: false,
    logLevel: 'info',
    legalComments: 'none',
  })
  copyPdfjsWorker()
  console.log('built dist/engine.js + dist/text.js')
  const generator = await import(
    `${pathToFileURL(join(pkgRoot, 'scripts/gen-commands.mjs')).href}?v=${Date.now()}`
  )
  await generator.generate?.()
}

/** 测试与诊断入口：dist/shared/** 不进安装包（pack.mjs 会跳过）。 */
async function buildShared() {
  const entries = {
    'pdf-meta': 'src/shared/pdf-meta.ts',
    chunk: 'src/shared/chunk.ts',
    print: 'src/shared/print.ts',
    'render-policy': 'src/shared/render-policy.ts',
    geometry: 'src/viewer/geometry.ts',
    search: 'src/viewer/search.ts',
  }
  const outdir = join(pkgRoot, 'dist/shared')
  mkdirSync(outdir, { recursive: true })
  await build({
    entryPoints: Object.values(entries).map((file) => join(pkgRoot, file)),
    outdir,
    bundle: true,
    format: 'esm',
    platform: 'neutral',
    target: 'es2022',
    logLevel: 'warning',
    entryNames: '[name]',
    outExtension: { '.js': '.mjs' },
  })
}

function buildPage() {
  execFileSync('pnpm', ['exec', 'vite', 'build'], {
    cwd: pkgRoot,
    stdio: 'inherit',
    shell: process.platform === 'win32',
  })
}

function writeBuildInfo() {
  const pkg = JSON.parse(readFileSync(join(pkgRoot, 'package.json'), 'utf8'))
  let upstreamCommit = 'unknown'
  try {
    upstreamCommit = execFileSync(
      'git',
      ['-C', join(repoRoot, 'vendor/genoffice'), 'rev-parse', 'HEAD'],
      { encoding: 'utf8' },
    ).trim()
  } catch {
    /* vendor 未初始化时保留 unknown */
  }
  writeFileSync(
    join(pkgRoot, 'dist/build.json'),
    JSON.stringify(
      {
        name: pkg.name,
        version: pkg.version,
        upstreamCommit,
        builtAt: new Date().toISOString(),
      },
      null,
      2,
    ) + '\n',
  )
}

function assertDist() {
  const required = [
    'dist/engine.js',
    'dist/text.js',
    'dist/pdf.worker.mjs',
    'dist/page/index.html',
    'dist/page/assets/pdfjs/cmaps',
    'dist/page/assets/pdfjs/standard_fonts',
    'dist/page/assets/pdfjs/wasm',
    'dist/page/assets/pdfjs/iccs',
    'commands.json',
  ]
  for (const file of required) {
    if (!existsSync(join(pkgRoot, file))) throw new Error(`构建产物缺失：${file}`)
  }
  // 旧 URA 注册面不应再出现。
  if (existsSync(join(pkgRoot, 'dist/ura.js'))) {
    throw new Error('产物不应该包含 dist/ura.js（PDF 已迁到 SDK 契约）')
  }
  if (existsSync(join(pkgRoot, 'dist/pdf.worker.min.mjs'))) {
    throw new Error('产物不应该包含 dist/pdf.worker.min.mjs')
  }
}

async function main() {
  if (assetsOnly) {
    copyPdfjsAssets(join(pkgRoot, 'src/page'))
    return
  }
  if (sharedOnly) {
    await buildShared()
    return
  }
  // 全量构建先清 dist：避免旧的 engine chunk / page 资产混进安装包。
  if (!engineOnly && !pageOnly) {
    rmSync(join(pkgRoot, 'dist'), { recursive: true, force: true })
  }
  if (!pageOnly) await buildEngine()
  if (engineOnly) return
  if (!pageOnly) await buildShared()
  buildPage()
  copyPdfjsAssets(join(pkgRoot, 'dist/page'))
  writeBuildInfo()
  assertDist()
  console.log('pdf build complete')
}

void main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error))
  process.exit(1)
})
