// ─── 测试/自检共用的引擎 harness ──────────────────────────────────────
//
// 引擎不依赖页面：这里用内存 fs + 假 scope 就能跑通全部只读命令与写路径，
// 只需要在 apply 时扮演一次"页面"（pending → runTxn → save → ack）。

import { existsSync, readFileSync, statSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { unzipSync, strFromU8 } from 'fflate'

const here = dirname(fileURLToPath(import.meta.url))
const vendor = resolve(here, '../../../vendor/genoffice')

export async function loadEngine() {
  const mod = await import(`../dist/engine.js?v=${Date.now()}`)
  const contribution = mod.workstudio ?? mod.default
  if (!contribution?.commands) throw new Error('engine 没有导出 workstudio')
  return contribution
}

export function createHarness(contribution, seed) {
  const files = new Map()
  const events = []
  const written = []
  let savePath = null

  const log = { info() {}, warn() {}, error() {} }
  const scope = {
    pluginId: 'pptx',
    dataDir: '',
    tempDir: '',
    settings: {
      get: (_key, fallback) => fallback,
      set() {},
      watch: () => () => {},
    },
    secrets: {
      availability: () => ({ available: false }),
      get: async () => null,
      set: async () => {},
    },
    credentials: {},
    events: {
      emit(name, data) {
        events.push({ name, data })
      },
    },
    files: {
      resolveSave: () => savePath ?? '',
      resolveInput: () => '',
    },
    log,
  }

  const ctx = {
    fs: {
      async read(ref) {
        const bytes = files.get(ref)
        if (!bytes) throw new Error('not_found: 文件不存在')
        return bytes
      },
      async write(ref, bytes) {
        files.set(ref, new Uint8Array(bytes))
        written.push(ref)
      },
    },
    access: { ls: async () => [], read: async () => [], grep: async () => [], find: async () => [] },
    settings: { get: () => ({}), watch: () => () => {} },
    credentials: {},
    changed() {},
    log: { info() {}, warn() {}, error() {} },
  }

  contribution.lifecycle?.activate?.(scope)

  return {
    files,
    events,
    written,
    setSavePath(path) {
      savePath = path
    },
    seed() {
      return new Uint8Array(seed)
    },
    async run(action, input) {
      const command = contribution.commands[`pptx.${action}`]
      if (!command) throw new Error(`没有命令 pptx.${action}`)
      return command.run(input, ctx)
    },
    /** 起一条 apply，返回 id / promise（页面由 nextRequest 取正文）。 */
    async startApply(input) {
      const from = events.length
      const promise = this.run('apply', input)
      promise.catch(() => undefined)
      const id = await waitForRequest(events, from)
      return { id, promise }
    },
    async pending(id) {
      return this.run('pending', { id })
    },
    async ack(id, ok, payload) {
      return this.run('ack', { id, ok, ...(payload ? { payload: JSON.stringify(payload) } : {}) })
    },
    stop() {
      contribution.lifecycle?.deactivate?.(scope)
    },
  }
}

export async function waitForRequest(events, from = 0, timeoutMs = 2000) {
  const started = Date.now()
  while (Date.now() - started < timeoutMs) {
    for (let i = from; i < events.length; i += 1) {
      const entry = events[i]
      if (entry?.name !== 'request') continue
      return String(entry.data?.id ?? '')
    }
    await new Promise((resolve) => setTimeout(resolve, 5))
  }
  throw new Error('没有收到页面请求事件')
}

// ─── 扮演页面 ───────────────────────────────────────────────────────

/** 测试里要用 pptx-engine/pptx-ops：node 解析不了 @genoffice/*，打成 bundle 再 import。 */
let opsBundle = null
async function vendorOps() {
  if (opsBundle) return opsBundle
  const { build } = await import('esbuild')
  const outfile = join(tmpdir(), 'pptx-pro-test-ops.mjs')
  await build({
    stdin: {
      contents: [
        "export { openPptx, savePptx } from '@genoffice/pptx-engine'",
        "export { runTxn } from '@genoffice/pptx-ops'",
      ].join('\n'),
      resolveDir: vendor,
      loader: 'ts',
    },
    bundle: true,
    format: 'esm',
    platform: 'node',
    target: 'node22',
    outfile,
    logLevel: 'error',
    plugins: [
      {
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
      },
      {
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
      },
    ],
  })
  opsBundle = await import(`file://${outfile.replace(/\\/g, '/')}?v=${Date.now()}`)
  return opsBundle
}

export async function applyOnPageBytes(bytes, commands) {
  const { openPptx, savePptx, runTxn } = await vendorOps()
  const opened = await openPptx(bytes)
  const result = runTxn(opened, { ops: commands })
  if (!result.applied)
    return { ok: false, error: result.failures?.[0]?.error ?? 'runTxn 失败' }
  return { ok: true, bytes: await savePptx(opened), applied: result.records?.length ?? commands.length }
}

export function slideXmlOf(bytes, index = 1) {
  const files = unzipSync(bytes)
  return strFromU8(files[`ppt/slides/slide${index}.xml`] ?? new Uint8Array())
}

export function slidePathsOf(bytes) {
  return Object.keys(unzipSync(bytes)).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
}
