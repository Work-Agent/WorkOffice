#!/usr/bin/env node
// PDF 插件构建：
//   1. esbuild src/ura.ts → dist/ura.js（pdf-lib 进 bundle，pdfjs-dist 不进）
//   2. esbuild 纯函数 → dist/shared/*.mjs（node --test 用）
//   3. vite build → dist/page/**
//   4. 复制 pdfjs 的 cmaps / standard_fonts / wasm → dist/page/pdfjs/**
//
// 参数：
//   --ura-only     只构建注册面
//   --page-only    只构建页面（仍复制静态资源）
//   --assets-only  只复制静态资源到 src/page/pdfjs（vite dev 用）

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
import { fileURLToPath } from 'node:url'
import { build } from 'esbuild'

const pkgRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const repoRoot = resolve(pkgRoot, '../..')
const args = process.argv.slice(2)
const has = (flag) => args.includes(flag)
const uraOnly = has('--ura-only')
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
  const out = join(targetPage, 'pdfjs')
  rmSync(out, { recursive: true, force: true })
  for (const dir of ['cmaps', 'standard_fonts', 'wasm', 'iccs']) {
    copyDir(join(dist, dir), join(out, dir))
  }
  console.log(`copied pdfjs assets → ${out}`)
}

async function buildUra() {
  mkdirSync(join(pkgRoot, 'dist'), { recursive: true })
  await build({
    entryPoints: [join(pkgRoot, 'src/ura.ts')],
    outfile: join(pkgRoot, 'dist/ura.js'),
    bundle: true,
    format: 'esm',
    platform: 'node',
    target: 'node22',
    external: ['electron'],
    logLevel: 'info',
  })
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
    upstreamCommit = execFileSync('git', ['-C', join(repoRoot, 'vendor/genoffice'), 'rev-parse', 'HEAD'], {
      encoding: 'utf8',
    }).trim()
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
    'dist/ura.js',
    'dist/page/index.html',
    'dist/page/pdfjs/cmaps',
    'dist/page/pdfjs/standard_fonts',
    'dist/page/pdfjs/wasm',
    'dist/page/pdfjs/iccs',
  ]
  for (const file of required) {
    if (!existsSync(join(pkgRoot, file)))
      throw new Error(`构建产物缺失：${file}`)
  }
  // sidecar 零解析：产物里不应再有主进程 worker。
  if (existsSync(join(pkgRoot, 'dist/pdf.worker.mjs')))
    throw new Error('产物不应该包含 dist/pdf.worker.mjs（sidecar 不解析 PDF）')
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
  if (!pageOnly) await buildUra()
  if (uraOnly) return
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
