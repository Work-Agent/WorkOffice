// ─── 构建：ura / engine / page ─────────────────────────────────────────
//
// 宿主不编译 TS；每个执行面入口各自 bundle：
//   dist/ura.js
//   dist/engine/{tools,plan,exec,read,runtime}.mjs
//   dist/page/**
//   dist/seed.xlsx

import { build } from 'esbuild'
import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { existsSync, statSync } from 'node:fs'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const vendor = resolve(root, '../../vendor/genoffice')
const argv = process.argv.slice(2)
const engineOnly = argv.includes('--engine-only')
const pageOnly = argv.includes('--page-only')
// wasm 是引擎产物的一部分；源码不在（未 vendor）时自动 skip。
const noWasm = argv.includes('--no-wasm')

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

async function buildWasm() {
  if (noWasm) {
    console.log('skip wasm（--no-wasm）')
    return
  }
  const { spawnSync } = await import('node:child_process')
  const result = spawnSync(process.execPath, [join(root, 'scripts/build-wasm.mjs')], {
    stdio: 'inherit',
  })
  if (result.status !== 0) throw new Error(`wasm 构建失败（exit ${result.status}）`)
}

const NODE = {
  bundle: true,
  format: 'esm',
  platform: 'node',
  target: 'node22',
  sourcemap: false,
  logLevel: 'info',
  plugins: [genofficeSource()],
}

async function out(entry, file) {
  await mkdir(dirname(join(root, file)), { recursive: true })
  await build({ ...NODE, entryPoints: [join(root, entry)], outfile: join(root, file) })
  console.log(`built ${file}`)
}

async function buildEngine() {
  await mkdir(join(root, 'dist'), { recursive: true })
  await cp(join(root, 'assets/seed.xlsx'), join(root, 'dist/seed.xlsx'))
  await out('src/ura.ts', 'dist/ura.js')
  await out('src/engine/tools.ts', 'dist/engine/tools.mjs')
  await out('src/engine/plan.ts', 'dist/engine/plan.mjs')
  await out('src/engine/exec.ts', 'dist/engine/exec.mjs')
  await out('src/engine/read.ts', 'dist/engine/read.mjs')
  await out('src/engine/runtime.ts', 'dist/engine/runtime.mjs')
  // 供测试直接引用（不属于宿主清单的任何入口）
  await out('src/shared/chunks.ts', 'dist/shared/chunks.mjs')
  await out('src/page/session-bytes.ts', 'dist/shared/session-bytes.mjs')
  await out('src/shared/zip-raw.ts', 'dist/shared/zip-raw.mjs')
  // 诊断/测试用：wasm 桥单独出一个入口（不在宿主清单里）
  await out('src/engine/wasm.ts', 'dist/shared/wasm-bridge.mjs')
  await out('src/editor/workstudio-save.ts', 'dist/shared/page-save.mjs')
  await out('src/editor/save-plan-budget.ts', 'dist/shared/save-plan-budget.mjs')

  const manifest = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))
  await writeFile(
    join(root, 'dist/build.json'),
    `${JSON.stringify(
      {
        name: manifest.name,
        version: manifest.version,
        id: manifest.workstudio.id,
        tools: 'compact',
        builtAt: new Date().toISOString(),
      },
      null,
      2,
    )}\n`,
  )
  await buildWasm()
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
