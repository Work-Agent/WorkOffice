#!/usr/bin/env node
// 校验构建产物是否符合 WorkStudio 新版 URA 注册契约。
//
//   node scripts/verify-contract.mjs
//
// 只检查注册面（codec / presenter / sidecar）与键；不打开真实文档。
// 需要先运行 `pnpm build:codec`。

import { access } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')

const PLUGINS = [
  {
    id: 'docx',
    mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    surface: 'docx-editor',
    extensions: ['docx', 'docm', 'dotx'],
  },
  {
    id: 'xlsx',
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    surface: 'xlsx-viewer',
    extensions: ['xlsx', 'xlsm', 'xltx'],
  },
  {
    id: 'pptx',
    mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    surface: 'pptx-viewer',
    extensions: ['pptx', 'pptm', 'potx'],
  },
]

let failed = 0

function check(condition, message) {
  if (condition) return
  failed++
  console.error(`  ✗ ${message}`)
}

for (const plugin of PLUGINS) {
  const entry = join(root, 'packages', plugin.id, 'dist', 'ura.js')
  try {
    await access(entry)
  } catch {
    failed++
    console.error(`✗ ${plugin.id}: 缺少 dist/ura.js，请先运行 pnpm build:codec`)
    continue
  }

  const seen = { codecs: [], presenters: [], sidecars: [] }
  globalThis.__workstudioSdk = {
    registerTransport() {},
    registerStore() {},
    registerCollection() {},
    registerViewAction() {},
    registerCodec(codec) {
      seen.codecs.push(codec)
    },
    registerPresenter(presenter) {
      seen.presenters.push(presenter)
    },
    registerSidecar(provider) {
      seen.sidecars.push(provider)
    },
  }

  try {
    await import(`${pathToFileURL(entry).href}?verify=${Date.now()}`)
  } catch (error) {
    failed++
    console.error(
      `✗ ${plugin.id}: 注册入口求值失败 · ${error instanceof Error ? error.message : String(error)}`,
    )
    continue
  }

  const failuresBefore = failed
  const codec = seen.codecs.find((item) => item.id === plugin.id)
  const presenter = seen.presenters.find(
    (item) => item.key?.codec === plugin.id,
  )
  const sidecar = seen.sidecars.find((item) => item.key?.codec === plugin.id)

  console.log(`\n${plugin.id}`)
  check(seen.codecs.length === 1, `codec 注册数量应为 1，实际 ${seen.codecs.length}`)
  check(Boolean(codec), `缺少 codec id=${plugin.id}`)
  if (codec) {
    check(
      Array.isArray(codec.keys) &&
        codec.keys.some(
          (key) => key.transport === '*' && key.mime === plugin.mime,
        ),
      `codec 缺少 keys: [{ transport: '*', mime: '${plugin.mime}' }]`,
    )
    check(
      plugin.extensions.every((ext) => codec.extensions?.includes(ext)),
      `codec.extensions 缺少 ${plugin.extensions.join(' / ')}`,
    )
    check(typeof codec.decode === 'function', 'codec.decode 缺失')
    check(typeof codec.locate === 'function', 'codec.locate 缺失')
    check(typeof codec.encode === 'function', 'codec.encode 缺失')
    check(codec.patchType === 'bytes', 'Office codec 的 patchType 应为 bytes')
  }

  check(Boolean(presenter), `缺少 presenter（codec=${plugin.id}）`)
  if (presenter) {
    check(presenter.type === 'page', 'presenter.type 应为 page')
    check(
      presenter.surface === plugin.surface,
      `presenter.surface 应为 ${plugin.surface}`,
    )
    check(typeof presenter.present === 'function', 'presenter.present 缺失')
  }

  check(Boolean(sidecar), `缺少 sidecar（codec=${plugin.id}）`)
  if (sidecar) {
    check(typeof sidecar.sidecar === 'function', 'sidecar.sidecar 缺失')
  }

  if (codec && presenter && sidecar && failed === failuresBefore) {
    console.log(`  ✓ codec / presenter / sidecar 契约完整`)
  }
}

if (failed > 0) {
  console.error(`\n契约校验失败：${failed} 项`)
  process.exit(1)
}

console.log('\n契约校验通过。')
