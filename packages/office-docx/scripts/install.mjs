#!/usr/bin/env node
// 把构建结果安装到 WorkStudio 的插件目录（只这一个插件）。
//
//   node scripts/install.mjs
//     → <userData>/plugins/installed/office-docx/
//
//   node scripts/install.mjs --target <dir>
//     → 指定目录（本地开发 / 随应用构建）
//
// 只复制 package.json 与 dist/：源码、tests、node_modules 不进入宿主。
// 也可以完全不复制——WorkStudio「设置 → 插件 → 安装来源」填本包目录即可原地加载。

import { cpSync, existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

function defaultPluginsRoot() {
  if (process.platform === 'darwin')
    return join(homedir(), 'Library', 'Application Support', 'WorkStudio', 'plugins');
  if (process.platform === 'win32')
    return join(process.env['APPDATA'] ?? join(homedir(), 'AppData', 'Roaming'), 'WorkStudio', 'plugins');
  return join(process.env['XDG_CONFIG_HOME'] ?? join(homedir(), '.config'), 'WorkStudio', 'plugins');
}

const args = process.argv.slice(2);
const targetArg = args.includes('--target') ? args[args.indexOf('--target') + 1] : undefined;
const target = targetArg
  ? resolve(targetArg.replace(/^~(?=$|\/|\\)/, homedir()))
  : join(defaultPluginsRoot(), 'installed');

const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
const ws = pkg.workstudio ?? {};
if (!pkg.version) throw new Error('package.json 缺少 version');
if (ws.apiVersion !== 2) throw new Error('workstudio.apiVersion 必须为 2');
if (ws.id !== 'office-docx') throw new Error('workstudio.id 必须为 office-docx');
if (!existsSync(join(root, 'dist'))) {
  console.error('缺少 dist/：先运行 node scripts/build.mjs');
  process.exit(1);
}
for (const rel of [
  pkg.exports?.['./ura'],
  ws.page,
  ...Object.values(ws.engine ?? {}),
]) {
  if (!rel || !existsSync(join(root, rel))) throw new Error(`产物缺失：${String(rel)}`);
}

const pluginTarget = join(target, ws.id);
rmSync(pluginTarget, { recursive: true, force: true });
mkdirSync(pluginTarget, { recursive: true });
cpSync(join(root, 'package.json'), join(pluginTarget, 'package.json'));
cpSync(join(root, 'dist'), join(pluginTarget, 'dist'), { recursive: true });
if (existsSync(join(root, 'skills')))
  cpSync(join(root, 'skills'), join(pluginTarget, 'skills'), { recursive: true });
console.log(`installed ${ws.id}@${pkg.version} → ${pluginTarget}`);
console.log('重启 WorkStudio（或插件列表「重新加载」）后生效。');
