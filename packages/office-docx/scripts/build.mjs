// ─── 构建：把 TS 打成宿主可直接求值的 ESM 产物 ─────────────────────────
//
// 宿主不编译 TS（docs/plugin-dev.md §11）：清单里的入口必须是 .js / .mjs。
// 因此这里把每个执行面入口各自 bundle 成一个文件：
//
//   dist/ura.js                L0 注册面（主进程求值）
//   dist/engine/tools.mjs      L1 工具声明（主进程只声明）
//   dist/engine/plan.mjs       L3 计划（执行点）
//   dist/engine/exec.mjs       L3 字节（执行点，唯一字节产生者）
//   dist/engine/read.mjs       L3 只读投影（执行点）
//   dist/engine/runtime.mjs    L3 探活（执行点）
//   dist/page/assets/index.js  L4 页面
//   dist/page/{index.html,styles.css}
//   dist/lib/{shared,zip}.mjs  单测用（不安装进宿主）
//
// 页面与执行点的共享模块来自同一份源码 → I17「两侧引擎同源」在构建期成立。

import { build } from 'esbuild';
import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const libOnly = process.argv.includes('--lib-only');

const NODE = {
  bundle: true,
  format: 'esm',
  platform: 'node',
  target: 'node22',
  sourcemap: false,
  logLevel: 'info',
};

const BROWSER = {
  bundle: true,
  format: 'esm',
  platform: 'browser',
  target: 'chrome120',
  sourcemap: false,
  logLevel: 'info',
};

async function out(entry, file, options) {
  await mkdir(dirname(join(root, file)), { recursive: true });
  await build({ ...options, entryPoints: [join(root, entry)], outfile: join(root, file) });
  console.log(`built ${file}`);
}

async function buildLib() {
  await out('src/shared/index.ts', 'dist/lib/shared.mjs', NODE);
  await out('src/engine/zip.ts', 'dist/lib/zip.mjs', NODE);
}

async function buildPlugin() {
  await buildLib();
  await out('src/ura.ts', 'dist/ura.js', NODE);
  await out('src/main/tools.ts', 'dist/engine/tools.mjs', NODE);
  await out('src/engine/plan.ts', 'dist/engine/plan.mjs', NODE);
  await out('src/engine/exec.ts', 'dist/engine/exec.mjs', NODE);
  await out('src/engine/read.ts', 'dist/engine/read.mjs', NODE);
  await out('src/engine/runtime.ts', 'dist/engine/runtime.mjs', NODE);
  await out('src/page/main.ts', 'dist/page/assets/index.js', BROWSER);

  const pageDir = join(root, 'dist/page');
  await mkdir(join(pageDir, 'assets'), { recursive: true });
  await cp(join(root, 'src/page/index.html'), join(pageDir, 'index.html'));
  await cp(join(root, 'src/page/styles.css'), join(pageDir, 'styles.css'));

  const manifest = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'));
  await writeFile(
    join(root, 'dist/build.json'),
    `${JSON.stringify(
      { name: manifest.name, version: manifest.version, id: manifest.workstudio.id, builtAt: new Date().toISOString() },
      null,
      2,
    )}\n`,
  );
  console.log('build complete');
}

if (libOnly) {
  await buildLib();
} else {
  // 只清理本脚本自己的产物：`dist/page-v2` 归 vite（`npm run build` 的第二步），
  // 在这里整目录 rm 会把渲染器版页面一起删掉，导致 `verify` 的入口检查失败。
  for (const entry of ['ura.js', 'engine', 'page', 'lib', 'build.json'])
    await rm(join(root, 'dist', entry), { recursive: true, force: true });
  await buildPlugin();
}
