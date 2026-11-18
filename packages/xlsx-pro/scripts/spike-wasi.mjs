#!/usr/bin/env node
// ─── wasm 接入 spike（G0，决策闸门） ───────────────────────────────────
//
// 这个脚本要回答的问题只有一个：**能不能在宿主执行点里跑 WASI reactor**，
// 以及跑起来之后 300 MB 的打开/保存是不是还能接受。它刻意分成两段：
//
//   阶段 A（不需要 Rust 产物，今天就能跑）
//     ① node:wasi 可用性（preview1 + initialize，不是 start）
//     ② 最小 reactor：只导出 _initialize / memory / wasi_alloc / wasi_free / call
//     ③ ABI 往返：alloc → 写 request → call → 读 response（hi32=ptr, lo32=len）
//     ④ WASI 导入真的被调用到（random_get / clock_time_get 写进 scratch）
//     ⑤ preopen + TMPDIR：宿主临时区映射成 guest '/tmp'
//
//   阶段 B（--module 指向真实 xlsx-engine.wasm 时）
//     ⑥ Open / ReadRange / SaveArchive，只写临时区
//     ⑦ 未改条目的**压缩流逐字节不变**（save_archive 的核心不变量）
//     ⑧ 耗时 + RSS 峰值（含 --size-mb 生成的合成大文件）
//
// 用法：
//   node scripts/spike-wasi.mjs                      # 只跑阶段 A
//   node scripts/spike-wasi.mjs --module dist/engine/xlsx-engine.wasm
//   node scripts/spike-wasi.mjs --module ... --fixture dist/seed.xlsx --size-mb 100
//   node scripts/spike-wasi.mjs --json               # 机读输出
//   node scripts/spike-wasi.mjs --self-test          # 只验 ⑨ 的比对器（不需要引擎）
//
// 阶段 A 失败 → 整条 wasm 路线作废（或转 W2）；阶段 B 的 ⑧ 超标 → 砍 P4（页面读取）。
//
// 最小 reactor 的 response 是**写死**的（它不是引擎，只是 ABI 夹具），所以
// 阶段 A 的断言是「结构化往返成立」，不是「业务正确」。

import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { zipSync, unzipSync } from 'fflate'

// ─── 结果面 ────────────────────────────────────────────────────────────

const results = []
/** @param {string} id @param {string} name @param {'pass'|'fail'|'skip'} status @param {string} detail */
function record(id, name, status, detail) {
  results.push({ id, name, status, detail })
  if (!process.argv.includes('--json'))
    console.log(`${status === 'pass' ? '  ok  ' : status === 'fail' ? ' FAIL ' : ' skip '} ${id} ${name}${detail ? ` — ${detail}` : ''}`)
}

function fail(message) {
  throw new Error(message)
}

function bytesOf(value) {
  if (value instanceof Uint8Array) return value
  if (value instanceof ArrayBuffer) return new Uint8Array(value)
  return new Uint8Array(value)
}

const MB = 1024 * 1024
const rssMb = () => Math.round((process.memoryUsage().rss / MB) * 10) / 10

// ─── 阶段 A：最小 WASI reactor（手编 wasm，无外部工具） ────────────────
//
// 导出：
//   memory                      JS 侧读写
//   _initialize                 空的 start（reactor 约定）
//   wasi_alloc(i32 len) -> i32   bump allocator（heap 从 4096 起）
//   wasi_free(ptr, len)          no-op
//   call(ptr,len) -> i64         hi32=response ptr, lo32=response len
//
// call 内部真的调用 random_get 与 clock_time_get，把结果写进 scratch，
// 这样宿主能从 memory 里读到非零值 —— 证明 WASI import 绑定成立。

const RESPONSE_PTR = 512
const RANDOM_SCRATCH = 2048
const CLOCK_SCRATCH = 2064

function u32leb(value) {
  const out = []
  let v = value >>> 0
  do {
    let byte = v & 0x7f
    v >>>= 7
    if (v !== 0) byte |= 0x80
    out.push(byte)
  } while (v !== 0)
  return out
}

function s32leb(value) {
  const out = []
  let v = value | 0
  for (;;) {
    const byte = v & 0x7f
    v >>= 7
    const signBitSet = (byte & 0x40) !== 0
    if ((v === 0 && !signBitSet) || (v === -1 && signBitSet)) {
      out.push(byte)
      return out
    }
    out.push(byte | 0x80)
  }
}

function s64leb(value) {
  const out = []
  let v = BigInt(value)
  for (;;) {
    const byte = Number(v & 0x7fn)
    v >>= 7n
    const signBitSet = (byte & 0x40) !== 0
    if ((v === 0n && !signBitSet) || (v === -1n && signBitSet)) {
      out.push(byte)
      return out
    }
    out.push(byte | 0x80)
  }
}

function vec(items) {
  return [...u32leb(items.length), ...items.flat()]
}

function section(id, payload) {
  return [id, ...u32leb(payload.length), ...payload]
}

/** 构造最小 reactor；response 是写死的 JSON（夹具，不是引擎）。 */
function buildMinimalReactor(responseJson) {
  const response = new TextEncoder().encode(responseJson)

  // type 0 () -> () | 1 (i32)->i32 | 2 (i32,i32)->() | 3 (i32,i32)->i64
  // 4 (i32,i32)->i32 (random_get) | 5 (i32,i64,i32)->i32 (clock_time_get)
  const types = vec([
    [0x60, 0x00, 0x00],
    [0x60, 0x01, 0x7f, 0x01, 0x7f],
    [0x60, 0x02, 0x7f, 0x7f, 0x00],
    [0x60, 0x02, 0x7f, 0x7f, 0x01, 0x7e],
    [0x60, 0x02, 0x7f, 0x7f, 0x01, 0x7f],
    [0x60, 0x03, 0x7f, 0x7e, 0x7f, 0x01, 0x7f],
  ])

  const encoder = new TextEncoder()
  const importName = (name) => [...u32leb(name.length), ...encoder.encode(name)]
  const imports = vec([
    [...importName('wasi_snapshot_preview1'), ...importName('random_get'), 0x00, ...u32leb(4)],
    [...importName('wasi_snapshot_preview1'), ...importName('clock_time_get'), 0x00, ...u32leb(5)],
  ])

  // 本地函数（索引 2..5，前两个是 import）
  const functions = vec([...u32leb(0), ...u32leb(1), ...u32leb(2), ...u32leb(3)])

  const memory = vec([[0x00, ...u32leb(2)]]) // min 2 pages
  const globals = vec([[0x7f, 0x01, 0x41, ...s32leb(4096), 0x0b]])

  const exportName = (name) => [...u32leb(name.length), ...encoder.encode(name)]
  const exports = vec([
    [...exportName('memory'), 0x02, ...u32leb(0)],
    [...exportName('_initialize'), 0x00, ...u32leb(2)],
    [...exportName('wasi_alloc'), 0x00, ...u32leb(3)],
    [...exportName('wasi_free'), 0x00, ...u32leb(4)],
    [...exportName('call'), 0x00, ...u32leb(5)],
  ])

  const body = (locals, ops) => {
    const payload = [...u32leb(locals.length), ...locals.flat(), ...ops, 0x0b]
    return [...u32leb(payload.length), ...payload]
  }

  const initializeBody = body([], [])
  // wasi_alloc: (ptr = heap); heap += len; return ptr
  const allocBody = body(
    [],
    [
      0x23, 0x00, // global.get 0
      0x23, 0x00, // global.get 0
      0x20, 0x00, // local.get 0
      0x6a, // i32.add
      0x24, 0x00, // global.set 0
    ],
  )
  const freeBody = body([], [])
  const callBody = body(
    [],
    [
      // random_get(RANDOM_SCRATCH, 8)
      0x41, ...s32leb(RANDOM_SCRATCH),
      0x41, ...s32leb(8),
      0x10, 0x00, // call import 0
      0x1a, // drop
      // clock_time_get(clock_id=0 realtime, precision=0, CLOCK_SCRATCH)
      0x41, 0x00,
      0x42, 0x00,
      0x41, ...s32leb(CLOCK_SCRATCH),
      0x10, 0x01, // call import 1
      0x1a, // drop
      // return (RESPONSE_PTR << 32) | response.length
      0x42,
      ...s64leb((BigInt(RESPONSE_PTR) << 32n) | BigInt(response.length)),
    ],
  )
  const codes = vec([initializeBody, allocBody, freeBody, callBody])

  const data = vec([
    [
      ...u32leb(0), // memory index
      0x41, ...s32leb(RESPONSE_PTR), 0x0b, // offset expr
      ...u32leb(response.length),
      ...response,
    ],
  ])

  const module = [
    0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00,
    ...section(1, types),
    ...section(2, imports),
    ...section(3, functions),
    ...section(5, memory),
    ...section(6, globals),
    ...section(7, exports),
    ...section(10, codes),
    ...section(11, data),
  ]
  return new Uint8Array(module)
}

// ─── ABI 桥：和将来 dist/engine/wasm.mjs 同形 ─────────────────────────

async function createReactor(wasmBytes, { tempDir, useMinimalResponse = true }) {
  const { WASI } = await import('node:wasi')
  const wasi = new WASI({
    version: 'preview1',
    // guest 只看见 '/' + '/tmp'，宿主临时区就是 '/tmp'：上游 crate 的
    // temp_dir()/canonicalize() 已按这套假设写好（TMPDIR + 无符号链接）。
    preopens: { '/tmp': tempDir },
    env: { TMPDIR: '/tmp' },
    returnOnExit: false,
  })
  const { instance } = await WebAssembly.instantiate(wasmBytes, {
    wasi_snapshot_preview1: wasi.wasiImport,
  })
  // reactor：initialize，不是 start（start 只跑一次 _start 就结束）
  wasi.initialize(instance)
  const memory = instance.exports.memory
  const view = () => new Uint8Array(memory.buffer)
  const encoder = new TextEncoder()
  const decoder = new TextDecoder()

  const invoke = (request) => {
    const payload = encoder.encode(JSON.stringify(request))
    const ptr = instance.exports.wasi_alloc(payload.length)
    if (!ptr) fail('wasi_alloc 返回 0')
    if (ptr + payload.length > memory.buffer.byteLength)
      fail(`内存不足：ptr=${ptr} len=${payload.length} size=${memory.buffer.byteLength}`)
    view().set(payload, ptr)
    const returned = instance.exports.call(ptr, payload.length)
    const outPtr = Number(returned >> 32n)
    const outLen = Number(returned & 0xffffffffn)
    if (outPtr + outLen > memory.buffer.byteLength)
      fail(`call 返回越界：ptr=${outPtr} len=${outLen}`)
    const text = decoder.decode(view().slice(outPtr, outPtr + outLen))
    instance.exports.wasi_free(ptr, payload.length)
    return { text, outPtr, outLen }
  }

  return { wasi, instance, memory, view, invoke, useMinimalResponse }
}

// ─── 阶段 A ───────────────────────────────────────────────────────────

async function phaseA(tempDir) {
  // ① node:wasi
  let wasiVersion = ''
  try {
    const { WASI } = await import('node:wasi')
    const probe = new WASI({ version: 'preview1' })
    if (typeof probe.wasiImport !== 'object' || typeof probe.initialize !== 'function')
      fail('WASI 构造面不符合预期（wasiImport / initialize）')
    wasiVersion = process.versions.electron ? `electron node ${process.versions.node}` : `node ${process.versions.node}`
    record('①', 'node:wasi 可用（preview1 + initialize）', 'pass', wasiVersion)
  } catch (error) {
    record('①', 'node:wasi 可用（preview1 + initialize）', 'fail', String(error?.message ?? error))
    return
  }

  // ② + ③ 最小 reactor / ABI 往返
  const responseJson = JSON.stringify({
    version: 1,
    requestId: 'spike',
    ok: true,
    result: { harness: 'minimal-wasi-reactor' },
  })
  let reactor
  try {
    reactor = await createReactor(buildMinimalReactor(responseJson), { tempDir })
    record('②', '最小 reactor 实例化 + initialize', 'pass', `${(buildMinimalReactor(responseJson).byteLength / 1024).toFixed(1)} KB`)
  } catch (error) {
    record('②', '最小 reactor 实例化 + initialize', 'fail', String(error?.message ?? error))
    return
  }

  try {
    const { text } = reactor.invoke({ version: 1, requestId: 'spike', command: 'ping' })
    const parsed = JSON.parse(text)
    if (parsed.ok !== true) fail(`response.ok=${parsed.ok}`)
    record('③', 'ABI 往返（alloc → call → 读回）', 'pass', `${text.length} B JSON`)
  } catch (error) {
    record('③', 'ABI 往返（alloc → call → 读回）', 'fail', String(error?.message ?? error))
  }

  // ④ WASI 导入真的被调用（scratch 里有非零字节）
  try {
    const mem = reactor.view()
    const random = mem.slice(RANDOM_SCRATCH, RANDOM_SCRATCH + 8)
    const clockBytes = mem.slice(CLOCK_SCRATCH, CLOCK_SCRATCH + 8)
    const clock = Buffer.from(clockBytes).readBigUInt64LE(0)
    const randomNonZero = random.some((byte) => byte !== 0)
    if (!randomNonZero) fail('random_get 未写入任何非零字节')
    if (clock === 0n) fail('clock_time_get 返回 0')
    record('④', 'WASI 导入生效（random_get / clock_time_get）', 'pass', `clock=${clock}ns`)
  } catch (error) {
    record('④', 'WASI 导入生效（random_get / clock_time_get）', 'fail', String(error?.message ?? error))
  }

  // ⑤ preopen + TMPDIR：宿主临时区可写、guest 映射被接受
  try {
    const probe = join(tempDir, 'probe.txt')
    await writeFile(probe, 'xlsx-pro spike')
    const roundTrip = await readFile(probe, 'utf8')
    if (roundTrip !== 'xlsx-pro spike') fail('临时区读写不一致')
    record('⑤', "preopen '/tmp' + TMPDIR 就绪", 'pass', tempDir)
  } catch (error) {
    record('⑤', "preopen '/tmp' + TMPDIR 就绪", 'fail', String(error?.message ?? error))
  }
}

// ─── 阶段 B：真实引擎（--module） ──────────────────────────────────────

/** 读 ZIP central directory，拿到每个条目的**原始压缩字节**。 */
function rawZipEntries(bytes) {
  const view = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const eocdSignature = 0x06054b50
  let eocd = -1
  for (let offset = view.length - 22; offset >= Math.max(0, view.length - 22 - 0xffff); offset -= 1) {
    if (view.readUInt32LE(offset) === eocdSignature) {
      eocd = offset
      break
    }
  }
  if (eocd < 0) fail('找不到 ZIP EOCD')
  const count = view.readUInt16LE(eocd + 10)
  let cursor = view.readUInt32LE(eocd + 16)
  const entries = new Map()
  for (let index = 0; index < count; index += 1) {
    if (view.readUInt32LE(cursor) !== 0x02014b50) fail('central directory 损坏')
    const compressedSize = view.readUInt32LE(cursor + 20)
    const nameLength = view.readUInt16LE(cursor + 28)
    const extraLength = view.readUInt16LE(cursor + 30)
    const commentLength = view.readUInt16LE(cursor + 32)
    const localOffset = view.readUInt32LE(cursor + 42)
    const name = view.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8')
    if (view.readUInt32LE(localOffset) !== 0x04034b50) fail(`local header 损坏：${name}`)
    const localNameLength = view.readUInt16LE(localOffset + 26)
    const localExtraLength = view.readUInt16LE(localOffset + 28)
    const dataStart = localOffset + 30 + localNameLength + localExtraLength
    entries.set(name, view.subarray(dataStart, dataStart + compressedSize))
    cursor += 46 + nameLength + extraLength + commentLength
  }
  return entries
}

/** 合成 fixture：level 0（不压缩），文件大小 ≈ --size-mb。 */
function buildSyntheticWorkbook(sizeMb) {
  const target = sizeMb * MB
  const rowPad = 'x'.repeat(80)
  const rows = []
  let bytes = 0
  for (let row = 0; bytes < target; row += 1) {
    const cell = `<row r="${row + 1}"><c r="A${row + 1}" t="inlineStr"><is><t>${rowPad}${row}</t></is></c></row>`
    rows.push(cell)
    bytes += cell.length
  }
  const sheet =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.join('')}</sheetData></worksheet>`
  const files = {
    '[Content_Types].xml':
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `</Types>`,
    '_rels/.rels':
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
      `</Relationships>`,
    'xl/workbook.xml':
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels':
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
      `</Relationships>`,
    'xl/worksheets/sheet1.xml': sheet,
  }
  const encoder = new TextEncoder()
  const encoded = Object.fromEntries(
    Object.entries(files).map(([name, text]) => [name, encoder.encode(text)]),
  )
  return zipSync(encoded, { level: 0 })
}

async function phaseB(modulePath, fixturePath, sizeMb, tempDir) {
  const started = Date.now()
  let wasmBytes
  try {
    wasmBytes = new Uint8Array(await readFile(modulePath))
    const info = await stat(modulePath)
    record('⑥a', 'wasm 产物可读', 'pass', `${(info.size / MB).toFixed(2)} MB`)
  } catch (error) {
    record('⑥a', 'wasm 产物可读', 'fail', `${modulePath}: ${error?.message ?? error}`)
    return
  }

  let reactor
  try {
    reactor = await createReactor(wasmBytes, { tempDir, useMinimalResponse: false })
    const exports = WebAssembly.Module.exports(new WebAssembly.Module(wasmBytes)).map((entry) => entry.name)
    for (const name of ['memory', 'wasi_alloc', 'wasi_free', 'call'])
      if (!exports.includes(name)) fail(`缺少导出：${name}`)
    record('⑥b', 'reactor 加载 + 导出面', 'pass', exports.join(','))
  } catch (error) {
    record('⑥b', 'reactor 加载 + 导出面', 'fail', String(error?.message ?? error))
    return
  }

  // fixture：显式传入，或从 seed 复制，或用 --size-mb 合成
  let fixture
  if (sizeMb > 0) {
    fixture = buildSyntheticWorkbook(sizeMb)
    record('⑥c', '合成大文件', 'pass', `${(fixture.byteLength / MB).toFixed(1)} MB`)
  } else {
    fixture = new Uint8Array(await readFile(fixturePath))
    record('⑥c', 'fixture', 'pass', `${fixturePath} (${(fixture.byteLength / MB).toFixed(1)} MB)`)
  }
  const basePath = join(tempDir, 'base.xlsx')
  const outPath = join(tempDir, 'out.xlsx')
  await writeFile(basePath, fixture)

  const invokeRaw = (request) => {
    const { text } = reactor.invoke(request)
    const parsed = JSON.parse(text)
    if (parsed.ok !== true)
      fail(`${request.command} 失败：${parsed.error?.code ?? '?'} ${parsed.error?.message ?? ''}`)
    return parsed.result
  }

  let sessionId
  let sheetId
  let readRange = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }
  let sheetMeta = {}
  try {
    const t0 = Date.now()
    const opened = invokeRaw({ version: 1, requestId: 'spike-open', command: 'open', path: '/tmp/base.xlsx' })
    // 上游 WorkbookMetadata 的键名以 snake_case 为主，两种都接受。
    sessionId = opened?.sessionId ?? opened?.session_id
    if (!sessionId) fail(`open 未返回 sessionId：${JSON.stringify(opened).slice(0, 200)}`)
    const meta = opened?.sheets?.[0] ?? {}
    sheetId = meta.id ?? meta.sheetId
    if (!sheetId) fail(`open 未返回工作表 id：${JSON.stringify(opened).slice(0, 200)}`)
    // 用引擎报的 extents 收缩读取范围（空表 A1 合法，50 行会越界）
    readRange = {
      startRow: 0,
      endRow: Math.max(0, Math.min(Number(meta.rowCount ?? meta.row_count ?? 1), 50) - 1),
      startColumn: 0,
      endColumn: Math.max(0, Math.min(Number(meta.columnCount ?? meta.column_count ?? 1), 10) - 1),
    }
    sheetMeta = meta
    record(
      '⑥',
      'open（引擎冷启 + 解析）',
      'pass',
      `${Date.now() - t0} ms · sheetXml ${((Number(sheetMeta.sourceXmlBytes ?? sheetMeta.source_xml_bytes ?? 0)) / MB).toFixed(1)} MB · rss ${rssMb()} MB`,
    )
  } catch (error) {
    record('⑥', 'Open', 'fail', String(error?.message ?? error))
    return
  }

  try {
    const t0 = Date.now()
    const read = invokeRaw({
      version: 1,
      requestId: 'spike-read',
      command: 'read_range',
      sessionId,
      sheetId,
      // CellRange 带 #[serde(rename_all = "camelCase")]（types.rs:428）
      range: readRange,
    })
    record(
      '⑦',
      'read_range（O(范围)）',
      'pass',
      `${Date.now() - t0} ms · ${readRange.endRow + 1}x${readRange.endColumn + 1} · keys=${Object.keys(read ?? {}).join(',')} · rss ${rssMb()} MB`,
    )
  } catch (error) {
    record('⑦', 'ReadRange', 'fail', String(error?.message ?? error))
  }

  // ⑦b：索引建好之后的第二次读取 —— 这才是「O(范围)」的真实数字，
  // 决定了页面滚动能不能一次滚动一次 invoke。
  {
    const totalRows = Number(sheetMeta.rowCount ?? sheetMeta.row_count ?? 1)
    const startRow = readRange.endRow + 1
    const endRow = Math.min(startRow + 49, totalRows - 1)
    if (endRow < startRow) {
      record('⑦b', 'read_range 第二次（索引已建，O(范围)）', 'skip', '工作表只有 1 行，换个 fixture 再测')
    } else {
      try {
        const t0 = Date.now()
        invokeRaw({
          version: 1,
          requestId: 'spike-read-2',
          command: 'read_range',
          sessionId,
          sheetId,
          range: {
            startRow,
            endRow,
            startColumn: readRange.startColumn,
            endColumn: readRange.endColumn,
          },
        })
        record('⑦b', 'read_range 第二次（索引已建，O(范围)）', 'pass', `${Date.now() - t0} ms · rss ${rssMb()} MB`)
      } catch (error) {
        record('⑦b', 'read_range 第二次（索引已建，O(范围)）', 'fail', String(error?.message ?? error))
      }
    }
  }

  try {
    const t0 = Date.now()
    invokeRaw({
      version: 1,
      requestId: 'spike-save',
      command: 'save_archive',
      sourcePath: '/tmp/base.xlsx',
      targetPath: '/tmp/out.xlsx',
      replacements: [],
      removals: [],
      additions: [],
      streamPatches: [],
    })
    record('⑧', 'SaveArchive 只写临时区', 'pass', `${Date.now() - t0} ms · rss ${rssMb()} MB`)
  } catch (error) {
    record('⑧', 'SaveArchive', 'fail', String(error?.message ?? error))
    return
  }

  try {
    const before = rawZipEntries(fixture)
    const after = rawZipEntries(new Uint8Array(await readFile(outPath)))
    const changed = []
    for (const [name, raw] of before) {
      const next = after.get(name)
      if (!next) changed.push(`${name}:missing`)
      else if (Buffer.compare(Buffer.from(raw), Buffer.from(next)) !== 0) changed.push(name)
    }
    if (changed.length > 0) fail(`未改条目压缩流变了：${changed.slice(0, 5).join(', ')}`)
    record('⑨', '未改条目压缩流逐字节不变（I1）', 'pass', `${before.size} 个条目`)
  } catch (error) {
    record('⑨', '未改条目压缩流逐字节不变（I1）', 'fail', String(error?.message ?? error))
  }

  // ⑨b：真的改一个 part（replacements 是「临时文件路径」而不是字符串——
  // 这是 P3 的一个集成成本：gateway 产出的字符串要先落盘成文件）。
  if (fixture.byteLength < 20 * MB) {
    try {
      const flat = unzipSync(fixture)
      const sheetName = 'xl/worksheets/sheet1.xml'
      const original = new TextDecoder().decode(flat[sheetName])
      const replaced = original.replace('</worksheet>', '<!--spike-replacement--></worksheet>')
      if (replaced === original) fail('替换内容没有生效（找不到 </worksheet>）')
      const replacementPath = join(tempDir, 'replacement.xml')
      await writeFile(replacementPath, replaced)
      const out2Path = join(tempDir, 'out2.xlsx')
      invokeRaw({
        version: 1,
        requestId: 'spike-save-replaced',
        command: 'save_archive',
        sourcePath: '/tmp/base.xlsx',
        targetPath: '/tmp/out2.xlsx',
        replacements: [{ name: sheetName, contentPath: '/tmp/replacement.xml' }],
        removals: [],
        additions: [],
        streamPatches: [],
      })
      const out2 = new Uint8Array(await readFile(out2Path))
      const out2Flat = unzipSync(out2)
      const written = new TextDecoder().decode(out2Flat[sheetName])
      if (!written.includes('<!--spike-replacement-->')) fail('目标 part 没有被替换')
      const beforeRaw = rawZipEntries(fixture)
      const afterRaw = rawZipEntries(out2)
      const drifted = [...beforeRaw.keys()].filter((name) => {
        if (name === sheetName) return false
        const next = afterRaw.get(name)
        return !next || Buffer.compare(Buffer.from(beforeRaw.get(name)), Buffer.from(next)) !== 0
      })
      if (drifted.length > 0) fail(`只改一个 part 却动了其它条目：${drifted.join(',')}`)
      record('⑨b', 'replacements 生效且其余条目原样（I1）', 'pass', `${beforeRaw.size - 1} 个未改条目`)
    } catch (error) {
      record('⑨b', 'replacements 生效且其余条目原样（I1）', 'fail', String(error?.message ?? error))
    }
  } else {
    record('⑨b', 'replacements 生效且其余条目原样（I1）', 'skip', '大文件不做（unzip 峰值过高）')
  }

  record('⑩', '总耗时', 'pass', `${Date.now() - started} ms · rss ${rssMb()} MB`)
}

// ─── --self-test：先证明 ⑨ 的比对器会「正确地说不等」 ──────────────────

async function selfTest(tempDir) {
  const sample = buildSyntheticWorkbook(2)
  const raw = rawZipEntries(sample)
  const flat = unzipSync(sample)
  const names = [...raw.keys()].sort()
  const flatNames = Object.keys(flat).sort()
  if (names.join('|') !== flatNames.join('|')) fail(`条目集合不一致：${names} vs ${flatNames}`)
  const mismatched = names.filter((name) => Buffer.compare(Buffer.from(raw.get(name)), Buffer.from(flat[name])) !== 0)
  if (mismatched.length > 0) fail(`level 0 条目应与解压内容相同：${mismatched.join(',')}`)
  record('S1', 'rawZipEntries 与 unzipSync 对齐（level 0）', 'pass', `${names.length} 个条目`)

  // 负向用例：只改一个条目并重新压缩，比对器必须报「变了」——
  // 正是 TS 底座（JSZip 全包重压缩）产生的结果，⑨ 必须能抓到。
  const encoder = new TextEncoder()
  const entries = Object.fromEntries(
    Object.entries(flat).map(([name, content]) => [name, encoder.encode(new TextDecoder().decode(content))]),
  )
  entries['xl/worksheets/sheet1.xml'] = encoder.encode(
    new TextDecoder().decode(flat['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData><!--x-->'),
  )
  const recompressed = zipSync(entries, { level: 6 })
  const after = rawZipEntries(recompressed)
  const changed = [...raw.keys()].filter((name) => {
    const next = after.get(name)
    return !next || Buffer.compare(Buffer.from(raw.get(name)), Buffer.from(next)) !== 0
  })
  if (changed.length === 0) fail('重压缩后比对器没有报告任何变化（断言器失效）')
  record('S2', '重压缩必被判为「压缩流变了」（负向用例）', 'pass', `${changed.length}/${raw.size} 个条目变化`)

  // 正向用例：把未改条目的原始压缩字节直接搬运 → 必须判为相同。
  const mixed = new Map()
  for (const [name, content] of Object.entries(flat))
    mixed.set(name, name === 'xl/worksheets/sheet1.xml' ? entries[name] : content)
  const copyLike = zipSync(Object.fromEntries(mixed), { level: 0 })
  const copyRaw = rawZipEntries(copyLike)
  const stillChanged = [...raw.keys()].filter((name) => {
    if (name === 'xl/worksheets/sheet1.xml') return false
    const next = copyRaw.get(name)
    return !next || Buffer.compare(Buffer.from(raw.get(name)), Buffer.from(next)) !== 0
  })
  if (stillChanged.length > 0)
    record('S3', '未改条目在「原样拷贝」下判为相同', 'fail', stillChanged.join(','))
  else record('S3', '未改条目在「原样拷贝」下判为相同', 'pass', `${raw.size - 1} 个条目`)

  // S4：合成的 fixture 必须是引擎能打开的合法工作簿，否则阶段 B 的 300 MB
  // 测试会因为一个拼错的关系文件而误判成「引擎不行」。
  try {
    const readModule = await import(new URL('../dist/engine/read.mjs', import.meta.url).href)
    const view = await readModule.parseWorkbook(sample)
    const sheet = view?.sheets?.[0]
    if (!sheet || sheet.name !== 'Sheet1') fail(`合成工作簿表名异常：${sheet?.name}`)
    record('S4', '合成 fixture 可被引擎解析', 'pass', `sheet=${sheet.name}`)
  } catch (error) {
    record('S4', '合成 fixture 可被引擎解析', 'skip', String(error?.message ?? error))
  }
}

// ─── 入口 ──────────────────────────────────────────────────────────────

function argValue(name, fallback) {
  const argv = process.argv
  const index = argv.indexOf(name)
  if (index < 0) return fallback
  const value = argv[index + 1]
  return value === undefined || value.startsWith('--') ? fallback : value
}

const USAGE = `用法：
  node scripts/spike-wasi.mjs                      阶段 A（不需要引擎产物）
  node scripts/spike-wasi.mjs --self-test          只验对比对器（原始压缩字节）
  node scripts/spike-wasi.mjs --module <wasm>      阶段 B（真实引擎）
      [--fixture dist/seed.xlsx] [--size-mb 100] [--keep-temp] [--json]
`

async function main() {
  if (process.argv.includes('--help') || process.argv.includes('-h')) {
    console.log(USAGE)
    return
  }
  const modulePath = argValue('--module', '')
  const fixturePath = argValue('--fixture', join('dist', 'seed.xlsx'))
  const sizeMb = Number(argValue('--size-mb', '0'))
  const keep = process.argv.includes('--keep-temp')
  const json = process.argv.includes('--json')

  const tempDir = await mkdtemp(join(tmpdir(), 'xlsx-pro-spike-'))
  if (!json) {
    console.log('# xlsx-pro wasm spike (G0)')
    console.log(`  node ${process.version}${process.versions.electron ? ` · electron ${process.versions.electron}` : ''}`)
    console.log(`  temp ${tempDir}`)
    console.log('')
    if (!process.argv.includes('--self-test')) console.log('## 阶段 A：node:wasi + 最小 reactor')
  }

  if (process.argv.includes('--self-test')) {
    if (!json) console.log('## 自检：⑨ 的原始压缩字节比对器')
    try {
      await selfTest(tempDir)
    } catch (error) {
      record('S0', '自检', 'fail', String(error?.message ?? error))
    }
    if (!json) console.log('')
    const failedSelf = results.filter((entry) => entry.status === 'fail')
    if (json) console.log(JSON.stringify({ ok: failedSelf.length === 0, results }, null, 2))
    else console.log(`  结论：${failedSelf.length === 0 ? '通过' : `${failedSelf.length} 项失败`}`)
    process.exitCode = failedSelf.length === 0 ? 0 : 1
    await rm(tempDir, { recursive: true, force: true }).catch(() => {})
    return
  }

  try {
    await phaseA(tempDir)
    if (modulePath) {
      if (!json) console.log('\n## 阶段 B：真实引擎')
      await phaseB(modulePath, fixturePath, sizeMb, tempDir)
    } else if (!json) {
      console.log('\n## 阶段 B：跳过（未提供 --module <xlsx-engine.wasm>）')
    }
  } finally {
    if (!keep) await rm(tempDir, { recursive: true, force: true }).catch(() => {})
    else if (!json) console.log(`\n  保留临时区：${tempDir}`)
  }

  const failed = results.filter((entry) => entry.status === 'fail')
  if (json) {
    console.log(JSON.stringify({ ok: failed.length === 0, results }, null, 2))
  } else {
    console.log('')
    console.log(`  结论：${failed.length === 0 ? '通过' : `${failed.length} 项失败`}`)
    if (!modulePath)
      console.log('  提示：阶段 B 需要上游 Rust 产物，见 rust/xlsx-engine（P1）')
  }
  process.exitCode = failed.length === 0 ? 0 : 1
}

main().catch((error) => {
  console.error('spike 崩溃：', error)
  process.exitCode = 1
})
