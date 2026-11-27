#!/usr/bin/env node
// 从 dist/engine.js 的 definePlugin 贡献生成宿主静态描述。
//
// 宿主装载插件时读 commands.json（扫描期），并要求它与动态 commands 双向对齐；
// api.json / api.d.ts 供工具链使用。

import { readFile, writeFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')

export async function generate() {
  const mod = await import(
    `${pathToFileURL(join(root, 'dist/engine.js')).href}?v=${Date.now()}`
  )
  const contribution = mod.workstudio ?? mod.default
  if (!contribution?.commands) throw new Error('engine 没有导出 workstudio.commands')
  const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))
  const id = pkg.workstudio.id

  const commands = Object.entries(contribution.commands).map(([name, command]) => ({
    name,
    summary: String(command.summary ?? ''),
    args: command.args ?? {},
  }))
  for (const command of commands) {
    if (!command.name.startsWith(`${id}.`)) {
      throw new Error(`命令名必须以 ${id}. 开头：${command.name}`)
    }
  }

  await writeFile(
    join(root, 'commands.json'),
    `${JSON.stringify(
      {
        commands: commands.map((command) => ({
          name: command.name,
          summary: command.summary,
          args: command.args,
        })),
      },
      null,
      2,
    )}\n`,
  )
  await writeFile(
    join(root, 'api.json'),
    `${JSON.stringify(
      {
        id,
        resources: Object.keys(contribution.resources ?? {}),
        events: Boolean(contribution.events),
        commands: commands.map(({ name, args }) => ({ name, args })),
      },
      null,
      2,
    )}\n`,
  )
  const lines = ['export interface PluginApi {', '  commands: {']
  for (const command of commands) {
    lines.push(
      `    ${JSON.stringify(command.name)}: { input: Record<string, unknown>; output: unknown };`,
    )
  }
  lines.push('  };', '}', '')
  await writeFile(join(root, 'api.d.ts'), lines.join('\n'))
  console.log(`generated commands.json (${commands.length} commands)`)
  return commands
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
  await generate()
}
