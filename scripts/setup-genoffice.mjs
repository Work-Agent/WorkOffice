#!/usr/bin/env node
// 初始化官方 GenOffice 子模块，只取页面层需要的源码。
//
//   git submodule update --init --depth 1 --filter=blob:none
//   node scripts/setup-genoffice.mjs
//
// 只需要 packages/*（引擎 + UI 基础包）；apps/* 不拉，docs 页面层已 fork
// 到 packages/docx/src/editor。新增页面能力时再往 SPARSE_PATHS 里加。

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

execFileSync('git', ['sparse-checkout', 'init', '--cone'], {
  cwd: vendor,
  stdio: 'inherit',
})
execFileSync('git', ['sparse-checkout', 'set', ...SPARSE_PATHS], {
  cwd: vendor,
  stdio: 'inherit',
})
console.log(`GenOffice vendor ready: ${SPARSE_PATHS.length} packages`)
