import assert from 'node:assert/strict';
import { existsSync, mkdirSync, rmSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import test, { after } from 'node:test';

// 页面层的纯函数没有运行时环境可跑，但「快速样式库」的判定是纯数据逻辑，
// 值得钉住：它决定 Ribbon 样式库里能看见哪些样式。用 esbuild 现场打包后导入。

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const outDir = join(root, '.tmp-test');
const outFile = join(outDir, 'style-gallery.mjs');

async function loadGallery() {
  mkdirSync(outDir, { recursive: true });
  const { build } = await import('esbuild');
  await build({
    entryPoints: [join(root, 'src/page-v2/genoffice/editor/style-gallery.ts')],
    outfile: outFile,
    bundle: true,
    format: 'esm',
    platform: 'node',
    target: 'node22',
    logLevel: 'silent',
  });
  const module = await import(`${pathToFileURL(outFile).href}?t=${Date.now()}`);
  return module;
}

/** 构造一份 StyleMap（只填判定用得到的字段）。 */
function styleMap(styles) {
  const map = new Map();
  for (const style of styles) {
    map.set(style.styleId, { type: 'paragraph', ...style });
  }
  return map;
}

test('样式库：只有默认样式带 qFormat 时，回退列出全部可见样式（生成器文档）', async () => {
  const { quickStyleEntries } = await loadGallery();
  const styles = styleMap([
    { styleId: 'Normal', name: 'Normal', isDefault: true, qFormat: true },
    { styleId: 'Title', name: 'Title', qFormat: false },
    { styleId: 'Heading1', name: 'heading 1', qFormat: false },
    { styleId: 'Heading2', name: 'heading 2', qFormat: false },
  ]);
  const ids = quickStyleEntries(styles, new Set()).map((info) => info.styleId);
  assert.deepEqual(ids.sort(), ['Heading1', 'Heading2', 'Normal', 'Title']);
});

test('样式库：一个 qFormat 都没有时同样回退（转换器产物）', async () => {
  const { quickStyleEntries } = await loadGallery();
  const styles = styleMap([
    { styleId: 'Normal', name: 'Normal', isDefault: true, qFormat: false },
    { styleId: 'Title', name: 'Title', qFormat: false },
    { styleId: 'Heading1', name: 'heading 1', qFormat: false },
  ]);
  const ids = quickStyleEntries(styles, new Set()).map((info) => info.styleId);
  assert.deepEqual(ids.sort(), ['Heading1', 'Normal', 'Title']);
});

test('样式库：正常 Word 文档只列 qFormat 样式（不回退、不混入正文样式变体）', async () => {
  const { quickStyleEntries } = await loadGallery();
  const styles = styleMap([
    { styleId: 'Normal', name: 'Normal', isDefault: true, qFormat: true },
    { styleId: 'Title', name: 'Title', qFormat: true, uiPriority: 2 },
    { styleId: 'Heading1', name: 'heading 1', qFormat: true, uiPriority: 9 },
    { styleId: 'BodyText', name: 'Body Text', qFormat: false, uiPriority: 60 },
  ]);
  const ids = quickStyleEntries(styles, new Set()).map((info) => info.styleId);
  assert.deepEqual(ids.sort(), ['Heading1', 'Normal', 'Title']);
});

test('样式库：半隐藏样式只在文档用过后出现（Heading 3+ 的 Word 行为）', async () => {
  const { quickStyleEntries } = await loadGallery();
  const styles = styleMap([
    { styleId: 'Normal', name: 'Normal', isDefault: true, qFormat: true },
    { styleId: 'Title', name: 'Title', qFormat: true },
    { styleId: 'Heading3', name: 'heading 3', qFormat: true, semiHidden: true, unhideWhenUsed: true },
  ]);
  const unused = quickStyleEntries(styles, new Set()).map((info) => info.styleId);
  assert.equal(unused.includes('Heading3'), false);
  const used = quickStyleEntries(styles, new Set(['Heading3'])).map((info) => info.styleId);
  assert.equal(used.includes('Heading3'), true);
});

after(() => {
  if (existsSync(outDir)) rmSync(outDir, { recursive: true, force: true });
});
