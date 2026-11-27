/**
 * 端到端冒烟（不用打开 WorkStudio 窗口）：
 *
 *   构建产物 dist/engine.js + dist/xlsx-engine.wasm + dist/seed.xlsx
 *   → 建一个临时工作簿 → 流式打开（meta/range）→ 写值/公式 → 保存 → 重开校验
 *   → 租约（attach/report/detach）→ 清理
 *
 * 只用对外命令面（guide/read/apply/page），不 import 内部模块。
 */

import assert from 'node:assert/strict'
import { copyFile, mkdtemp, readFile, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const dist = join(root, 'dist')

const steps = []
function step(name, fn) {
  steps.push({ name, fn })
}

const { workstudio } = await import(pathToFileURL(join(dist, 'engine.js')).href)
const work = await mkdtemp(join(tmpdir(), 'xlsx-smoke-'))
workstudio.lifecycle.activate({
  pluginId: 'xlsx',
  tempDir: work,
  dataDir: work,
  files: { resolveSave: () => join(work, 'unused.xlsx') },
})

const path = join(work, 'smoke.xlsx')
const ref = pathToFileURL(path).href
const page = workstudio.commands['xlsx.page']
const read = workstudio.commands['xlsx.read']
const apply = workstudio.commands['xlsx.apply']
const guide = workstudio.commands['xlsx.guide']
const pageId = 'smoke-page'
const ctx = {
  fs: {
    read: async () => new Uint8Array(await readFile(path)),
    write: async (target, bytes) => {
      await (await import('node:fs/promises')).writeFile(new URL(target), bytes)
    },
  },
  changed: () => undefined,
}
const call = (op, input = {}) =>
  page.run({ ref, pageId, op, input }, ctx)

step('准备临时工作簿', async () => {
  await copyFile(join(dist, 'seed.xlsx'), path)
  assert.ok((await stat(path)).size > 0)
})

step('命令面（guide → 5 个命令）', async () => {
  const commands = Object.keys(workstudio.commands).sort()
  assert.deepEqual(commands, [
    'xlsx.apply',
    'xlsx.create',
    'xlsx.guide',
    'xlsx.page',
    'xlsx.read',
  ])
  const text = (await guide.run({ ref }, {})).text
  assert.match(text, /ops=56/)
})

step('流式打开：status → meta → range', async () => {
  const status = await call('status')
  assert.match(status.hash, /^[0-9a-f]{64}$/)
  const meta = await call('meta')
  assert.equal(meta.source, 'wasm', '需要 dist/xlsx-engine.wasm（build:wasm）')
  assert.ok(meta.sheets.length >= 1)
  const range = await call('range', {
    sheetId: meta.sheets[0].id,
    range: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 2 },
  })
  assert.equal(range.rows.length, 10)
})

step('写入：值 + 公式 + 格式', async () => {
  const result = await apply.run(
    {
      ref,
      commands: JSON.stringify([
        {
          op: 'set_range',
          sheet: 'Sheet1',
          start: 'A1',
          values: [
            ['项目', '金额'],
            ['差旅', 1200],
            ['住宿', 800],
            ['合计', '=SUM(B2:B3)'],
          ],
        },
        { op: 'format_range', sheet: 'Sheet1', range: 'A1:B1', format: { bold: true } },
      ]),
    },
    ctx,
  )
  assert.equal(result.applied, 2)
  assert.match(result.receipt.hash, /^[0-9a-f]{64}$/)
  assert.ok(!JSON.stringify(result.notices ?? []).includes('save-archive-failed'))
})

step('重开：读回落盘内容', async () => {
  const values = await read.run({ ref, what: 'range', sheet: 'Sheet1', range: 'A1:B4' }, {})
  assert.match(values.text, /项目/)
  assert.match(values.text, /1200/)
  assert.match(values.text, /SUM\(B2:B3\)/)
  const opened = await call('open')
  const meta = await call('meta')
  assert.equal(meta.hash, opened.hash)
})

step('租约：attach → report → detach', async () => {
  const attached = await call('attach', { pageId })
  assert.match(attached.token, /^[0-9a-f-]{36}$/)
  await call('report', { token: attached.token, docRev: 2, structureEpoch: 1, dirty: true })
  await call('detach', { token: attached.token })
})

let failed = 0
for (const { name, fn } of steps) {
  try {
    await fn()
    console.log(`✓ ${name}`)
  } catch (error) {
    failed += 1
    console.error(`✗ ${name}\n  ${error instanceof Error ? error.message : String(error)}`)
  }
}
await rm(work, { recursive: true, force: true })
console.log(failed === 0 ? '\nsmoke 通过' : `\n${failed} 步失败`)
process.exitCode = failed === 0 ? 0 : 1
