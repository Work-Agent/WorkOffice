// ─── 构建：ura + engine（P0a） ───────────────────────────────────────
//
// 宿主不编译 TS；每个执行面入口各自 bundle：
//   dist/ura.js
//   dist/engine/{tools,plan,exec,read,runtime}.mjs
//   dist/seed.pptx

import { build } from 'esbuild'
import { execFileSync } from 'node:child_process'
import { mkdir, cp, readFile, rm, writeFile } from 'node:fs/promises'
import { existsSync, readFileSync, statSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const vendor = resolve(root, '../../vendor/genoffice')
const argv = process.argv.slice(2)
const engineOnly = argv.includes('--engine-only')
const pageOnly = argv.includes('--page-only')

function rawPlugin() {
  return {
    name: 'raw-text',
    setup(build) {
      build.onResolve({ filter: /\?raw$/ }, (args) => ({
        path: resolve(dirname(args.importer), args.path.replace(/\?raw$/, '')),
        namespace: 'raw-text',
      }))
      build.onLoad({ filter: /.*/, namespace: 'raw-text' }, (args) => ({
        contents: readFileSync(args.path, 'utf8'),
        loader: 'text',
      }))
    },
  }
}

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

const NODE = {
  bundle: true,
  format: 'esm',
  platform: 'node',
  target: 'node22',
  sourcemap: false,
  logLevel: 'info',
  banner: {
    js: "import { createRequire as __createRequire } from 'node:module'; const require = __createRequire(import.meta.url);",
  },
  plugins: [rawPlugin(), genofficeSource()],
}

async function out(entry, file) {
  await mkdir(dirname(join(root, file)), { recursive: true })
  await build({ ...NODE, entryPoints: [join(root, entry)], outfile: join(root, file) })
  console.log(`built ${file}`)
}

function vendorHash() {
  try {
    return execFileSync('git', ['-C', vendor, 'rev-parse', 'HEAD'], {
      encoding: 'utf8',
    }).trim()
  } catch {
    return 'unknown'
  }
}

async function buildEngine() {
  await mkdir(join(root, 'dist'), { recursive: true })
  await cp(join(root, 'assets/seed.pptx'), join(root, 'dist/seed.pptx'))
  await out('src/ura.ts', 'dist/ura.js')
  await out('src/engine/tools.ts', 'dist/engine/tools.mjs')
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
        tools: 'compact',
        vendor: vendorHash(),
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
  await rm(join(root, 'dist'), { recursive: true, force: true })
  await buildEngine()
} else {
  await rm(join(root, 'dist'), { recursive: true, force: true })
  await buildEngine()
  await buildPage()
}
