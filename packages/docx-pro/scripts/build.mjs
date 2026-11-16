// ─── 构建：把 TS 打成宿主可直接求值的 ESM 产物 ─────────────────────────
//
// 宿主不编译 TS（docs/plugin-dev.md §11）：清单里的入口必须是 .js / .mjs。
// 因此每个执行面入口各自 bundle 成一个文件：
//
//   dist/ura.js                L0 注册面（主进程求值）
//   dist/engine/tools.mjs      L1 工具声明（主进程只声明，纯数据）
//   dist/engine/plan.mjs       L3 计划（执行点）
//   dist/engine/exec.mjs       L3 字节（执行点，唯一字节产生者）
//   dist/engine/read.mjs       L3 只读投影（执行点）
//   dist/engine/runtime.mjs    L3 探活（执行点）
//   dist/page/**               L4 页面（本包 src/editor：从 packages/docx fork 的副本）

import { build } from 'esbuild'
import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const argv = process.argv.slice(2)
const engineOnly = argv.includes('--engine-only')
const pageOnly = argv.includes('--page-only')
// 工具面两种形态：
//   默认        3 个工具（docx_guide / docx_read / docx_apply）
//   --legacy-tools  35 个逐工具声明（保留做对照，命令实现完全相同）
const legacyTools = argv.includes('--legacy-tools')
const toolsEntry = legacyTools ? 'src/engine/tools-legacy.ts' : 'src/engine/tools.ts'

const NODE = {
  bundle: true,
  format: 'esm',
  platform: 'node',
  target: 'node22',
  sourcemap: false,
  logLevel: 'info',
}

async function out(entry, file) {
  await mkdir(dirname(join(root, file)), { recursive: true })
  await build({ ...NODE, entryPoints: [join(root, entry)], outfile: join(root, file) })
  console.log(`built ${file}`)
}

async function buildEngine() {
  await mkdir(join(root, 'dist'), { recursive: true })
  await cp(join(root, 'assets/seed.docx'), join(root, 'dist/seed.docx'))
  await out('src/ura.ts', 'dist/ura.js')
  await out(toolsEntry, 'dist/engine/tools.mjs')
  await out('src/engine/plan.ts', 'dist/engine/plan.mjs')
  await out('src/engine/exec.ts', 'dist/engine/exec.mjs')
  await out('src/engine/read.ts', 'dist/engine/read.mjs')
  await out('src/engine/runtime.ts', 'dist/engine/runtime.mjs')

  const manifest = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))
  await writeFile(
    join(root, 'dist/build.json'),
    `${JSON.stringify(
      {
        name: manifest.name,
        version: manifest.version,
        id: manifest.workstudio.id,
        tools: legacyTools ? 'legacy' : 'compact',
        builtAt: new Date().toISOString(),
      },
      null,
      2,
    )}\n`,
  )
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
  // 页面入口的 html 由 vite 产出；styles 复用 docs 编辑器的样式，无需额外拷贝。
  for (const name of ['build.json']) {
    const from = join(root, 'dist', name)
    try {
      await cp(from, join(root, 'dist', name))
    } catch {
      /* 已存在 */
    }
  }
}
