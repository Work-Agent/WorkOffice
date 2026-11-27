#!/usr/bin/env node
// 装载前自检：清单 / 静态命令 / 引擎命令 / 无页面写拒绝 / 一次 load → save。
import { readFile, stat, unlink, writeFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { buildDocx } from '../tests/fixtures.mjs'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const problems = []
const check = (ok, message) => { if (!ok) problems.push(message) }

const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))
const ws = pkg.workstudio ?? {}
check(ws.id === 'docx-pro', `workstudio.id 必须是 docx-pro（当前 ${ws.id}）`)
check(!ws.apiVersion, '不应再声明 apiVersion')
check(pkg.exports?.['./workstudio'] === './dist/engine.js', "exports['./workstudio'] 必须是 ./dist/engine.js")
for (const file of [ws.pages?.content, 'dist/engine.js', 'commands.json']) {
  check(Boolean(file) && Boolean(await stat(join(root, file)).catch(() => null)), `缺少文件：${file}`)
}
for (const dir of ws.skills ?? []) {
  check(Boolean(await stat(join(root, dir, 'SKILL.md')).catch(() => null)), `技能缺少 SKILL.md：${dir}`)
}

const mod = await import(pathToFileURL(join(root, 'dist/engine.js')).href)
const contribution = mod.workstudio ?? mod.default
const commands = contribution?.commands ?? {}
const staticCommands = JSON.parse(await readFile(join(root, 'commands.json'), 'utf8')).commands
check(Object.keys(commands).length === staticCommands.length, 'commands.json 与动态命令数量不一致')
for (const entry of staticCommands) check(Boolean(commands[entry.name]), `静态命令没有实现：${entry.name}`)
for (const name of Object.keys(commands)) check(name.startsWith('docx-pro.'), `命令名必须以 docx-pro. 开头：${name}`)
for (const name of ['docx-pro.guide', 'docx-pro.read', 'docx-pro.apply']) {
  check(commands[name]?.expose?.agent === true, `${name} 应对助手可见`)
}
for (const name of ['docx-pro.saveAs', 'docx-pro.status']) {
  check(!commands[name], `不应存在 ${name}`)
}

const file = join(root, '.verify.docx')
const ref = pathToFileURL(file).href
await writeFile(file, buildDocx())
const scope = { pluginId: 'docx-pro', events: { emit: () => {} }, log: console }
await contribution.lifecycle.activate(scope)
const ctx = {
  fs: {
    read: async (r) => new Uint8Array(await readFile(fileURLToPath(r))),
    write: async (r, bytes) => writeFile(fileURLToPath(r), bytes),
  },
  changed: () => {},
  log: console,
}
try {
  const guide = await commands['docx-pro.guide'].run({ ref }, ctx)
  check(typeof guide === 'string' && guide.includes('blocks'), 'guide 没有返回命令目录')
  const read = await commands['docx-pro.read'].run({ ref, what: 'context' }, ctx)
  check(typeof read === 'string' && read.includes('blocks='), 'read 没有返回块投影')
  const loaded = await commands['docx-pro.load'].run({ ref }, ctx)
  check(loaded?.bytes instanceof Uint8Array, 'load 没有返回 bytes')
  check(typeof loaded?.hash === 'string' && loaded.hash.length > 0, 'load 没有返回 hash')
  const createdFile = join(root, '.verify-created.docx')
  const createdRef = pathToFileURL(createdFile).href
  const created = await commands['docx-pro.create'].run({ ref: createdRef }, ctx)
  check(created?.created === true, 'create 没有返回 created')
  check(Boolean(await stat(createdFile).catch(() => null)), 'create 没有落盘')
  let existsRejected = false
  try {
    await commands['docx-pro.create'].run({ ref: createdRef }, ctx)
  } catch (error) {
    existsRejected = /resource\.exists/.test(String(error?.message ?? error))
  }
  check(existsRejected, 'create 对已存在目标应拒绝')
  await unlink(createdFile).catch(() => {})
  let rejected = false
  try {
    await commands['docx-pro.apply'].run({ ref, commands: '[]' }, ctx)
  } catch (error) {
    rejected = /command\.page-unavailable/.test(String(error?.message ?? error))
  }
  check(rejected, '无页面 apply 应返回 command.page-unavailable')
  await commands['docx-pro.attach'].run({ ref, pageId: 'verify' }, ctx)
  const saved = await commands['docx-pro.save'].run(
    { ref, pageId: 'verify', bytes: loaded.bytes, baseHash: loaded.hash, type: 'page' },
    ctx,
  )
  check(saved?.hash === loaded.hash, 'save 后 hash 变化')
} finally {
  await unlink(file).catch(() => {})
  await contribution.lifecycle.deactivate(scope)
}

if (problems.length > 0) {
  for (const problem of problems) console.error(`error ${problem}`)
  console.error(`\nverify failed: ${problems.length} 个问题`)
  process.exit(1)
}
console.log(`verify ok · ${pkg.name}@${pkg.version} · ${Object.keys(commands).length} commands`)
