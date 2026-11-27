import assert from 'node:assert/strict'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { randomBytes } from 'node:crypto'
import { after, test } from 'node:test'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const { workstudio } = await import(pathToFileURL(join(root, 'dist/engine.js')).href)
const page = workstudio.commands['xlsx.page']
const seedRef = pathToFileURL(join(root, 'assets/seed.xlsx')).href

let tempDir = ''
// 引擎命令依赖插件作用域（wasm 会话落在 scope.tempDir）；测试按宿主顺序先激活。
tempDir = await mkdtemp(join(tmpdir(), 'xlsx-scope-'))
workstudio.lifecycle.activate({
  pluginId: 'xlsx',
  tempDir,
  dataDir: tempDir,
  files: { resolveSave: () => path.join(tempDir, 'unused.xlsx') },
})

const call = (ref, op, input = {}, extra = {}) =>
  page.run({ ref, pageId: 'page-A', op, input: JSON.stringify(input), ...extra }, {})

after(async () => {
  if (tempDir) await rm(tempDir, { recursive: true, force: true })
})

test('status / open：单文件身份 + 哈希', async () => {
  const status = await call(seedRef, 'status')
  assert.equal(status.ref, seedRef)
  assert.ok(status.size > 0)
  assert.match(status.hash, /^[0-9a-f]{64}$/)
  assert.equal(status.tier, 'normal')
  const opened = await call(seedRef, 'open')
  assert.equal(opened.hash, status.hash)
  assert.equal(opened.size, status.size)
})

test('chunk：按哈希绑定、分块可拼回原文件', async () => {
  tempDir = await mkdtemp(join(tmpdir(), 'xlsx-'))
  const bytes = randomBytes(2 * 1024 * 1024 + 137)
  const path = join(tempDir, 'big.xlsx')
  await writeFile(path, bytes)
  const ref = pathToFileURL(path).href

  const opened = await call(ref, 'open')
  assert.equal(opened.size, bytes.byteLength)

  const parts = []
  let offset = 0
  for (;;) {
    const slice = await call(ref, 'chunk', { hash: opened.hash, offset, length: 512 * 1024 })
    assert.equal(slice.bytes.byteLength, Math.min(512 * 1024, bytes.byteLength - offset))
    parts.push(Buffer.from(slice.bytes))
    offset = slice.next
    if (slice.eof) break
  }
  assert.equal(offset, bytes.byteLength)
  assert.deepEqual(Buffer.concat(parts), bytes)

  await assert.rejects(() => call(ref, 'chunk', { hash: 'f'.repeat(64), offset: 0 }), /stale|已变化/)
  await assert.rejects(() => call(ref, 'chunk', { hash: opened.hash, offset: bytes.byteLength + 1 }), /越界/)
})

test('meta / range：wasm 元数据 + 结构化范围', async () => {
  const meta = await call(seedRef, 'meta')
  assert.ok(meta.sheets.length >= 1)
  assert.equal(meta.source, 'wasm', 'dist/xlsx-engine.wasm 应在 PATH 中（先跑 build）')
  const sheetId = meta.sheets[0].id
  assert.ok(Number(meta.sheets[0].rowCount) > 0)
  assert.match(meta.hash, /^[0-9a-f]{64}$/)

  // 编辑器传的是 {startRow,endRow,startColumn,endColumn} 对象
  const range = await call(seedRef, 'range', {
    sheetId,
    range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 },
  })
  assert.ok(Array.isArray(range.cells))
  assert.equal(range.rows.length, 3)
  assert.ok(Number(range.indexedThroughRow) > 0)

  await assert.rejects(() => call(seedRef, 'range', { sheet: '不存在的表', range: 'A1:B2' }), /工作表不存在/)
  await assert.rejects(
    () => call(seedRef, 'range', { sheetId, range: 'A1:C3', baseHash: 'a'.repeat(64) }),
    /stale/,
  )
})

test('租约：令牌发放、抢占拒绝、报告与撤销', async () => {
  workstudio.lifecycle.deactivate()
  workstudio.lifecycle.activate({ tempDir, dataDir: tempDir, pluginId: 'xlsx' })

  const attached = await call(seedRef, 'attach', { pageId: 'page-A' })
  assert.match(attached.token, /^[0-9a-f-]{36}$/)
  assert.equal(attached.docRev, 0)
  assert.deepEqual(attached.pending, [])

  await assert.rejects(
    () => page.run({ ref: seedRef, pageId: 'page-B', op: 'attach', input: '{}' }, {}),
    /lease\.busy/,
  )

  const takeover = await page.run(
    { ref: seedRef, pageId: 'page-B', op: 'attach', input: JSON.stringify({ takeover: true }) },
    {},
  )
  assert.notEqual(takeover.token, attached.token)
  await assert.rejects(
    () => call(seedRef, 'report', { token: attached.token, docRev: 1 }),
    /denied|令牌/,
  )

  const reported = await call(seedRef, 'report', {
    token: takeover.token,
    docRev: 3,
    structureEpoch: 1,
    dirty: true,
  })
  assert.equal(reported.docRev, 3)
  assert.equal(reported.structureEpoch, 1)
  assert.equal(reported.dirty, true)

  await call(seedRef, 'detach', { token: takeover.token })
  await assert.rejects(
    () => call(seedRef, 'report', { token: takeover.token }),
    /lease\.missing|租约/,
  )
})

test('guide：56 条 op 与四组', async () => {
  const result = await workstudio.commands['xlsx.guide'].run({ ref: seedRef }, {})
  assert.match(result.text, /ops=56/)
  assert.match(result.text, /cells/)
  assert.match(result.text, /objects/)

  const group = await workstudio.commands['xlsx.guide'].run({ ref: seedRef, group: 'cells' }, {})
  assert.match(group.text, /set_cell/)
})

test('read：首行令牌 + 内容投影', async () => {
  const context = await workstudio.commands['xlsx.read'].run(
    { ref: seedRef, what: 'context' },
    {},
  )
  assert.match(context.text, /^hash=[0-9a-f]{64} size=\d+/)
  assert.equal(context.what, 'context')

  const range = await workstudio.commands['xlsx.read'].run(
    { ref: seedRef, what: 'range', sheet: 'Sheet1', range: 'A1:B2' },
    {},
  )
  assert.match(range.text, /Sheet1!A1:B2/)

  await assert.rejects(
    () => workstudio.commands['xlsx.read'].run({ ref: seedRef, what: 'nope' }, {}),
    /what 非法/,
  )
})

test('单文件守卫：目录 / 非 file:// / 错误扩展名', async () => {
  const adapter = workstudio.resources.xlsx
  assert.deepEqual(await adapter.ls('file:///w'), [])
  await assert.rejects(() => adapter.read('https://example.com/a.xlsx'), /file:|单个文件/)
  await assert.rejects(() => adapter.read(pathToFileURL(root).href), /单个文件|扩展名|不存在/)
  await assert.rejects(() => adapter.read(pathToFileURL(join(root, 'package.json')).href), /扩展名|单个文件|只支持/)

  const item = (await adapter.read(seedRef))[0]
  assert.equal(item.locator, 'xlsx:file')
  assert.match(item.hash, /^[0-9a-f]{64}$/)
})

test('save：令牌 + baseHash + 请求产字节', async () => {
  const fs = await import('node:fs/promises')
  const dir = await fs.mkdtemp(join(tmpdir(), 'xlsx-save-'))
  const path = join(dir, 'book.xlsx')
  await fs.copyFile(join(root, 'assets/seed.xlsx'), path)
  const ref = pathToFileURL(path).href
  const before = await fs.readFile(path)

  const writes = []
  const ctx = {
    fs: {
      read: async () => new Uint8Array(before),
      write: async (target, bytes) => {
        writes.push({ target, size: bytes.byteLength })
        await fs.writeFile(path, bytes)
      },
    },
    changed: () => undefined,
  }

  const attached = await call(ref, 'attach', { pageId: 'page-save' })
  const token = attached.token
  const opened = await call(ref, 'open')

  // 令牌错 → 拒
  // 页面真实路径传对象（Uint8Array 走结构化克隆），测试也一样
  await assert.rejects(
    () => page.run({ ref, pageId: 'page-save', op: 'save', token: 'bad', input: { bytes: new Uint8Array(before) } }, ctx),
    /denied|令牌/,
  )

  // baseHash 过期 → stale
  await assert.rejects(
    () => page.run({ ref, pageId: 'page-save', op: 'save', token, input: { bytes: new Uint8Array(before), baseHash: 'f'.repeat(64) } }, ctx),
    /stale|已变化/,
  )

  // 整份字节写回
  const receipt = await page.run(
    { ref, pageId: 'page-save', op: 'save', token, input: { bytes: new Uint8Array(before), baseHash: opened.hash } },
    ctx,
  )
  assert.equal(receipt.hash, opened.hash)
  assert.equal(receipt.size, before.byteLength)
  assert.equal(writes.length, 1)

  // 保存请求 → 引擎用 gateway 产字节（空改批 = 结构合法的原样重打包）
  const request = await page.run(
    { ref, pageId: 'page-save', op: 'save', token, input: { request: { edits: [], structuralOps: [] }, baseHash: opened.hash } },
    ctx,
  )
  assert.match(request.hash, /^[0-9a-f]{64}$/)
  assert.equal(writes.length, 2)
  const after = await fs.readFile(path)
  assert.ok(after.byteLength > 0)
  await assert.rejects(
    () => page.run({ ref, pageId: 'page-save', op: 'save', token, input: {} }, ctx),
    /bytes 或 request/,
  )
  await fs.rm(dir, { recursive: true, force: true })
})

test('save：分块提交与另存为', async () => {
  const fs = await import('node:fs/promises')
  const dir = await fs.mkdtemp(join(tmpdir(), 'xlsx-chunk-'))
  const path = join(dir, 'book.xlsx')
  await fs.copyFile(join(root, 'assets/seed.xlsx'), path)
  const ref = pathToFileURL(path).href
  const payload = new Uint8Array(await fs.readFile(path))

  const writes = []
  const ctx = {
    fs: {
      read: async () => payload,
      write: async (target, bytes) => {
        writes.push({ target, size: bytes.byteLength })
        await fs.writeFile(new URL(target), bytes)
      },
    },
    changed: () => undefined,
  }

  const attached = await call(ref, 'attach', { pageId: 'page-chunk' })
  const token = attached.token
  const opened = await call(ref, 'open')

  const begin = await page.run({ ref, pageId: 'page-chunk', op: 'saveBegin', token, input: { size: payload.byteLength } }, ctx)
  const half = Math.ceil(payload.byteLength / 2)
  await page.run({ ref, pageId: 'page-chunk', op: 'saveChunk', token, input: { transferId: begin.transferId, seq: 0, bytes: payload.subarray(0, half) } }, ctx)
  await page.run({ ref, pageId: 'page-chunk', op: 'saveChunk', token, input: { transferId: begin.transferId, seq: 1, bytes: payload.subarray(half) } }, ctx)
  const committed = await page.run({ ref, pageId: 'page-chunk', op: 'saveCommit', token, input: { transferId: begin.transferId, baseHash: opened.hash } }, ctx)
  assert.equal(committed.hash, opened.hash)
  assert.equal(writes.length, 1)

  // 另存为：save: token → 目标文件
  const target = join(dir, 'copy.xlsx')
  workstudio.lifecycle.deactivate()
  workstudio.lifecycle.activate({
    pluginId: 'xlsx',
    tempDir: dir,
    dataDir: dir,
    files: { resolveSave: (token) => {
      assert.match(token, /^save:/)
      return target
    } },
  })
  const attached2 = await call(ref, 'attach', { pageId: 'page-saveas' })
  const before = await call(ref, 'open')
  const savedAs = await page.run(
    { ref, pageId: 'page-saveas', op: 'saveAs', token: attached2.token, input: { bytes: payload, saveToken: 'save:test', baseHash: before.hash } },
    ctx,
  )
  assert.equal(savedAs.ref, pathToFileURL(target).href)
  assert.deepEqual(new Uint8Array(await fs.readFile(target)), payload)
  await fs.rm(dir, { recursive: true, force: true })
})

test('apply / create：无页面写路径', async () => {
  const fs = await import('node:fs/promises')
  const dir = await fs.mkdtemp(join(tmpdir(), 'xlsx-apply-'))
  const path = join(dir, 'book.xlsx')
  await fs.copyFile(join(root, 'assets/seed.xlsx'), path)
  const ref = pathToFileURL(path).href

  const applyCmd = workstudio.commands['xlsx.apply']
  const readCmd = workstudio.commands['xlsx.read']
  const createCmd = workstudio.commands['xlsx.create']

  let writes = 0
  const ctx = {
    fs: {
      read: async () => new Uint8Array(await fs.readFile(path)),
      write: async (target, bytes) => {
        writes += 1
        await fs.writeFile(new URL(target), bytes)
      },
    },
    changed: () => undefined,
  }

  // dryRun 不落盘
  const dry = await applyCmd.run(
    { ref, commands: JSON.stringify([{ op: 'set_cell', sheet: 'Sheet1', address: 'A1', value: 42 }]), dryRun: true },
    ctx,
  )
  assert.equal(dry.dryRun, true)
  assert.equal(dry.applied, 0)
  assert.equal(writes, 0)

  // 值 + 公式（公式经 recalc_cells 写入缓存值）
  const result = await applyCmd.run(
    {
      ref,
      commands: JSON.stringify([
        { op: 'set_cell', sheet: 'Sheet1', address: 'A1', value: 42 },
        { op: 'set_formula', sheet: 'Sheet1', address: 'B1', formula: '=A1*2' },
        { op: 'format_range', sheet: 'Sheet1', range: 'A1:B1', format: { bold: true } },
      ]),
    },
    ctx,
  )
  assert.equal(result.requested, 3)
  assert.equal(result.applied, 3)
  assert.match(result.receipt.hash, /^[0-9a-f]{64}$/)
  assert.equal(writes, 1)
  // wasm 可用时保存应走 save_archive（不回退 JSZip）
  assert.ok(!JSON.stringify(result.notices ?? []).includes('save-archive-failed'))

  const values = await readCmd.run({ ref, what: 'range', sheet: 'Sheet1', range: 'A1:B1' }, {})
  assert.match(values.text, /42/)
  assert.match(values.text, /A1\*2/)

  // 结构命令
  const structural = await applyCmd.run(
    { ref, commands: JSON.stringify([{ op: 'insert_rows', sheet: 'Sheet1', row: 1, count: 1 }]) },
    ctx,
  )
  assert.equal(structural.applied, 1)

  // objects 组必须打开页面
  await assert.rejects(
    () => applyCmd.run({ ref, commands: JSON.stringify([{ op: 'set_freeze', sheet: 'Sheet1', rows: 1, columns: 0 }]) }, ctx),
    /page-unavailable|需页面|打开/,
  )

  // 未知 op / 非法批
  await assert.rejects(
    () => applyCmd.run({ ref, commands: JSON.stringify([{ op: 'nope' }]) }, ctx),
    /未知 op/,
  )
  await assert.rejects(() => applyCmd.run({ ref, commands: JSON.stringify([]) }, ctx), /不能为空/)

  // create：新文件 + 两张表
  const created = join(dir, '预算.xlsx')
  const createdRef = pathToFileURL(created).href
  const receipt = await createCmd.run(
    {
      ref: createdRef,
      sheets: JSON.stringify([
        { name: '预算', rows: [['项目', '金额'], ['差旅', 1200], ['合计', '=SUM(B2:B3)']] },
        { name: '说明', rows: [['口径', '含税']] },
      ]),
    },
    ctx,
  )
  assert.equal(receipt.ref, createdRef)
  assert.match(receipt.hash, /^[0-9a-f]{64}$/)
  const context = await readCmd.run({ ref: createdRef, what: 'context' }, {})
  assert.match(context.text, /预算/)

  // 已存在 / 父目录缺失
  await assert.rejects(
    () => createCmd.run({ ref: createdRef, sheets: [] }, ctx),
    /已存在/,
  )
  await assert.rejects(
    () => createCmd.run({ ref: pathToFileURL(join(dir, 'missing', 'x.xlsx')).href }, ctx),
    /父目录不存在/,
  )

  await fs.rm(dir, { recursive: true, force: true })
})

test('宿主参数校验：页面与 agent 的载荷都能通过', async () => {
  const fs = await import('node:fs/promises')
  const declared = JSON.parse(await fs.readFile(join(root, 'commands.json'), 'utf8')).commands

  // 复刻宿主 validateCommandInput：只校验声明过的参数，类型不符即拒。
  const validate = (name, input) => {
    const command = declared.find((entry) => entry.name === name)
    assert.ok(command, `commands.json 缺少 ${name}`)
    for (const [key, spec] of Object.entries(command.args ?? {})) {
      const value = input[key]
      if (value === undefined) {
        assert.ok(!spec.required, `缺少参数：${key}`)
        continue
      }
      if ('enum' in spec) {
        assert.ok(typeof value === 'string' && spec.enum.includes(value), `参数非法：${key}`)
        continue
      }
      if (spec.type === 'locator') continue
      assert.equal(typeof value, spec.type, `参数类型非法：${key}`)
    }
  }

  // 页面：input 是对象（含 Uint8Array 的场景靠结构化克隆）
  validate('xlsx.page', { ref: seedRef, pageId: 'p', op: 'status', input: {} })
  validate('xlsx.page', { ref: seedRef, pageId: 'p', op: 'save', input: { bytes: new Uint8Array(4) } })

  // agent：commands 传数组（对象）或 JSON 字符串都要能过
  validate('xlsx.apply', { ref: seedRef, commands: [{ op: 'set_cell', sheet: 'Sheet1', address: 'A1', value: 1 }] })
  validate('xlsx.apply', { ref: seedRef, commands: '[]' })
  validate('xlsx.create', { ref: seedRef, sheets: [{ name: 'S', rows: [[1]] }] })

  // 回归锁：结构化载荷不得出现在声明里（否则宿主按 string 校验会拒）
  for (const [name, key] of [
    ['xlsx.page', 'input'],
    ['xlsx.apply', 'commands'],
    ['xlsx.create', 'sheets'],
    ['xlsx.create', 'rows'],
  ]) {
    const command = declared.find((entry) => entry.name === name)
    assert.ok(!(key in (command.args ?? {})), `${name} 不应声明 ${key}（宿主会按 string 校验）`)
  }
})
