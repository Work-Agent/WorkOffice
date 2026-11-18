#!/usr/bin/env node
// 初始化官方 GenOffice 子模块，只取页面层需要的源码。
//
//   git submodule update --init --depth 1 --filter=blob:none
//   node scripts/setup-genoffice.mjs
//
// 默认只取 packages/*（引擎 + UI 基础包）；docs 页面层已 fork 到
// packages/docx/src/editor。唯一需要的 apps/* 是 xlsx 的 Rust 引擎源码
// （编成 wasm，不是原生 sidecar）。新增能力时再往 SPARSE_PATHS 里加。

import { execFileSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const vendor = join(root, 'vendor/genoffice')

const SPARSE_PATHS = [
  'packages/agent-core',
  'packages/ai-provider',
  'packages/docx-engine',
  'packages/electron-utils',
  'packages/font-metrics',
  'packages/i18n',
  'packages/pipelines',
  'packages/pptx-engine',
  'packages/pptx-ops',
  'packages/pptx-render',
  'packages/project-store',
  'packages/ui',
  'packages/xlsx-gateway',
  'packages/zip-gate',
  // wasm 引擎源码（xlsx-pro 编译 wasm32-wasip1 reactor，见
  // packages/xlsx-pro/scripts/build-wasm.mjs）。不做原生 sidecar，只要源码。
  'apps/sheets/native/xlsx-engine',
]

if (!existsSync(join(vendor, 'package.json'))) {
  execFileSync(
    'git',
    [
      'submodule',
      'update',
      '--init',
      '--depth',
      '1',
      '--filter=blob:none',
      'vendor/genoffice',
    ],
    { cwd: root, stdio: 'inherit' },
  )
}

// vendor 不是 git 工作区（例如整包复制进来的离线副本）时跳过稀疏设置，
// 只提示一句 —— 直接抛 execFileSync 的堆栈对使用者没有信息量。
if (!existsSync(join(vendor, '.git'))) {
  console.log(
    `GenOffice vendor 不是 git 工作区，跳过 sparse-checkout：${vendor}\n` +
      '  需要 wasm 引擎源码时请手动放入 vendor/genoffice/apps/sheets/native/xlsx-engine',
  )
  process.exit(0)
}

execFileSync('git', ['sparse-checkout', 'init', '--cone'], {
  cwd: vendor,
  stdio: 'inherit',
})
execFileSync('git', ['sparse-checkout', 'set', ...SPARSE_PATHS], {
  cwd: vendor,
  stdio: 'inherit',
})
console.log(`GenOffice vendor ready: ${SPARSE_PATHS.length} packages`)
