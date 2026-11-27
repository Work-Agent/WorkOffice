#!/usr/bin/env node
// 从 dist/engine.js 的 definePlugin 贡献生成宿主静态描述。
// 不依赖 SDK CLI，产物形状与 workstudio-plugin 一致。
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const mod = await import(pathToFileURL(join(root, 'dist/engine.js')).href)
const contribution = mod.workstudio ?? mod.default
if (!contribution?.commands) throw new Error('engine 没有导出 workstudio.commands')
const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))
const id = pkg.workstudio.id
const commands = Object.entries(contribution.commands).map(([name, command]) => ({
  name,
  summary: command.summary,
  args: command.args ?? {},
  exitCodes: { ok: 0, failed: 1, invalid: 2, notFound: 3 },
}))
for (const command of commands) {
  if (!command.name.startsWith(`${id}.`)) throw new Error(`命令名必须以 ${id}. 开头：${command.name}`)
}
await writeFile(join(root, 'commands.json'), `${JSON.stringify({ commands }, null, 2)}\n`)
await writeFile(
  join(root, 'api.json'),
  `${JSON.stringify({ commands: commands.map(({ name, args }) => ({ name, args })) }, null, 2)}\n`,
)
const lines = ['export interface PluginApi {', '  commands: {']
for (const command of commands) lines.push(`    ${JSON.stringify(command.name)}: { input: Record<string, unknown>; output: unknown };`)
lines.push('  };', '}', '')
await writeFile(join(root, 'api.d.ts'), lines.join('\n'))
console.log(`generated commands.json (${commands.length} commands)`)
