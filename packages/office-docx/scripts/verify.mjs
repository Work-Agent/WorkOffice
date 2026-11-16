#!/usr/bin/env node
// 装载前自检：按 WorkStudio 的安装闸门（docs/plugin-dev.md §2 / §4.8）检查本包。
//
//   node scripts/verify.mjs
//
// 宿主不编译 TS，也不读源码：这里断言 dist 产物与清单一致，避免出现
// 「装得上但执行点 import 失败」这类运行期才发现的问题。

import { existsSync, readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const failures = [];
const checks = [];
const pending = [];

function check(name, fn) {
  pending.push([name, fn]);
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
const ws = pkg.workstudio ?? {};
const engine = ws.engine ?? {};

check('清单：id / version / apiVersion', () => {
  assert(ws.id === 'office-docx', `workstudio.id=${String(ws.id)}`);
  assert(typeof pkg.version === 'string' && pkg.version, '缺少 version');
  assert(ws.apiVersion === 2, 'apiVersion 必须为 2');
});

check('清单：声明位组合约束', () => {
  assert(ws.write === 'plan', "write 必须为 'plan'");
  assert(Boolean(engine.exec), "write: 'plan' 必须声明 engine.exec");
  assert(Boolean(engine.plan), 'engine.plan / engine.exec 必须同时出现');
  assert(ws.content === 'pulled' || ws.content === undefined, "content 必须为 'pulled'");
  assert(Boolean(ws.page), "content: 'pulled' 必须有 page");
  assert(ws.model === 'none' || ws.model === undefined, "model 必须为 'none'");
  assert(Boolean(engine.tools), '缺少 engine.tools');
});

check('入口文件存在（相对插件根）', () => {
  for (const rel of [
    pkg.exports?.['./ura'],
    ws.page,
    engine.tools,
    engine.plan,
    engine.exec,
    engine.read,
    engine.runtime,
  ]) {
    assert(typeof rel === 'string' && rel, `清单缺少入口：${String(rel)}`);
    assert(existsSync(join(root, rel)), `文件不存在：${rel}`);
  }
});

check('页面产物引用本地资源', () => {
  const page = join(root, ws.page);
  const pageDir = dirname(page);
  const html = readFileSync(page, 'utf8');
  // 两种页面产物都合法：旧块列表页（esbuild 固定名）与渲染器版页（vite 带 hash）。
  // 统一做法：把 html 里所有相对引用抽出来，逐个断言文件存在。
  const refs = [...html.matchAll(/(?:src|href)="(\.\/[^"]+)"/g)].map((match) => match[1]);
  assert(refs.length > 0, 'index.html 没有任何本地资源引用');
  for (const ref of refs) {
    if (ref.includes('://')) continue;
    assert(
      existsSync(join(pageDir, ref.replace(/^\.\//, ''))),
      `页面引用的资源不存在：${ref}`,
    );
  }
  const js = refs.some((ref) => /^\.\/assets\/index[^/]*\.js$/.test(ref));
  assert(js, `index.html 未引用 ./assets/index*.js（实际：${refs.join(', ')}）`);
});

check('技能声明（dir 内必须有 SKILL.md）', () => {
  for (const skill of ws.skills ?? []) {
    assert(skill.name && skill.dir, '技能声明缺少 name/dir');
    assert(existsSync(join(root, skill.dir, 'SKILL.md')), `${skill.dir}/SKILL.md 不存在`);
    const text = readFileSync(join(root, skill.dir, 'SKILL.md'), 'utf8');
    assert(text.split('\n').length <= 80, 'SKILL.md 过长（设计上限 ≤60 行，含空行容差 80）');
  }
});

check('工具声明：形状 / 归属 / 令牌', async () => {
  const module = await import(pathToFileURL(join(root, engine.tools)).href);
  const tools = module.tools;
  assert(Array.isArray(tools) && tools.length === 2, '必须声明 docx_guide / docx_apply');
  const names = tools.map((tool) => tool.name);
  assert(names.includes('docx_guide') && names.includes('docx_apply'), `工具名不符：${names.join(', ')}`);
  for (const tool of tools) {
    assert(/^[a-z][a-z0-9._-]*$/i.test(tool.name), `工具名非法：${tool.name}`);
    assert(tool.label && tool.description, `${tool.name} 缺少 label/description`);
    assert(Array.isArray(tool.kinds) && tool.kinds.length > 0, `${tool.name} 缺少 kinds`);
    assert(tool.parameters && typeof tool.parameters === 'object', `${tool.name} 缺少 parameters`);
  }
  const apply = tools.find((tool) => tool.name === 'docx_apply');
  assert(apply.readOnly === false, 'docx_apply 必须是写工具');
  assert(apply.requiresBaseVersion === true, 'docx_apply 必须声明 requiresBaseVersion');
  assert(apply.tokens === 'docRev+epoch', 'docx_apply 的页面内令牌必须是 docRev+epoch');
  assert(apply.maxItems === 200, 'docx_apply 批上限必须是 200');
  assert(apply.plan === 'docx.apply', 'docx_apply 的 plan 名必须是 docx.apply');
  const guide = tools.find((tool) => tool.name === 'docx_guide');
  assert(guide.readOnly === true, 'docx_guide 必须是只读工具');
});

check('执行点模块导出齐全', async () => {
  const plan = await import(pathToFileURL(join(root, engine.plan)).href);
  assert(typeof plan.plan === 'function', 'engine.plan 缺少 plan()');
  const exec = await import(pathToFileURL(join(root, engine.exec)).href);
  assert(typeof exec.exec === 'function', 'engine.exec 缺少 exec()');
  assert(typeof exec.commitModel === 'function', 'engine.exec 缺少 commitModel()');
  assert(typeof exec.abortModel === 'function', 'engine.exec 缺少 abortModel()');
  const read = await import(pathToFileURL(join(root, engine.read)).href);
  assert(typeof read.project === 'function', 'engine.read 缺少 project()');
  assert(typeof read.query === 'function', 'engine.read 缺少 query()');
  const runtime = await import(pathToFileURL(join(root, engine.runtime)).href);
  assert(typeof runtime.runtime === 'function', 'engine.runtime 缺少 runtime()');
  const report = await runtime.runtime();
  assert(report.write === 'plan', 'RuntimeReport.write 必须为 plan');
  assert(report.features.tokens.baselineSync === 'manifest', 'baselineSync 必须为 manifest');
  assert(report.features.tokens.inPage === 'docRev+epoch', 'inPage 必须为 docRev+epoch');
  assert(report.features.undo === 'page', 'undo 必须为 page');
  assert(report.features.isolation === 'batch', 'isolation 必须为 batch');
  assert(report.features.requiresBaseVersion === true, 'requiresBaseVersion 必须为 true');
});

check('URA 入口：同步登记 + 编码通道关闭', async () => {
  const registered = {};
  globalThis.__workstudioSdk = {
    registerCodec: (value) => (registered.codec = value),
    registerSidecar: (value) => (registered.sidecar = value),
    registerStore: (value) => (registered.store = value),
    registerPresenter: (value) => (registered.presenter = value),
  };
  await import(`${pathToFileURL(join(root, pkg.exports['./ura'])).href}?v=${Date.now()}`);
  assert(registered.codec, '未注册 codec');
  assert(registered.sidecar, '未注册 sidecar');
  assert(registered.store, '未注册 store');
  assert(registered.presenter, '未注册 presenter');
  assert(registered.codec.patchType === 'bytes', 'patchType 必须为 bytes');
  let threw = false;
  try {
    registered.codec.encode({}, {});
  } catch {
    threw = true;
  }
  assert(threw, 'codec.encode 必须拒绝（字节只由 exec 产生）');
  let encodeError = '';
  try {
    registered.codec.encode({}, {});
  } catch (error) {
    encodeError = String(error && error.message);
  }
  assert(encodeError.includes('command.forbidden'), 'encode 必须抛 command.forbidden');
  assert(registered.presenter.type === 'page', 'presenter 必须是 page 出口');
  assert(Array.isArray(registered.presenter.requires), 'presenter.requires 缺失');
  delete globalThis.__workstudioSdk;
});

for (const [name, fn] of pending) {
  try {
    const result = await fn();
    if (result === false) throw new Error('返回 false');
    checks.push(name);
    console.log(`ok  ${name}`);
  } catch (error) {
    failures.push(`${name}：${error instanceof Error ? error.message : String(error)}`);
  }
}
if (failures.length > 0) {
  for (const failure of failures) console.error(`fail ${failure}`);
  process.exit(1);
}
console.log(`\n${checks.length} 项自检通过（${pkg.name}@${pkg.version}）。`);
