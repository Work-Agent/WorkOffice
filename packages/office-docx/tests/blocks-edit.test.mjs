import assert from 'node:assert/strict';
import test from 'node:test';
import {
  applyCommands,
  baselineSlices,
  mergeCommands,
  paragraphTextOf,
  parseDocxBytes,
  pasteCommands,
  planFromModel,
  rangeCommands,
} from '../dist/lib/shared.mjs';
import { buildDocx } from './fixtures.mjs';

function baseline() {
  const parsed = parseDocxBytes(buildDocx());
  return { parsed, slices: baselineSlices(parsed) };
}

function model() {
  return baseline().parsed.blocks.map((block) => ({ ...block }));
}

function texts(blocks) {
  return blocks.filter((block) => !block.hidden).map((block) => block.text);
}

// 夹具（fixtures.mjs）的可见块：0 标题段 / 1 正文段 / 2 表格 / 3 空段。

test('合并：段首退格并入上一段，光标停在接缝处', () => {
  const blocks = model();
  const batch = mergeCommands(blocks, 1, 'back');
  assert.equal(batch.kind, 'apply');
  assert.deepEqual(batch.focus, { index: 0, offset: 4 });
  const applied = applyCommands(blocks, batch.commands);
  assert.equal(applied.model.length, blocks.length - 1);
  assert.equal(applied.model[0].text, '年度报告本年度公司营收 增长 20%');
});

test('合并：段尾 Delete 并入下一段', () => {
  const blocks = model();
  const batch = mergeCommands(blocks, 1, 'forward');
  // 下一块是表格 → 拒绝，提示交给页面。
  assert.equal(batch.kind, 'refuse');
  assert.match(batch.reason, /表格/);
});

test('合并：首段退格 / 末段 Delete 无事可做', () => {
  const blocks = model();
  assert.equal(mergeCommands(blocks, 0, 'back').kind, 'skip');
  assert.equal(mergeCommands(blocks, 3, 'forward').kind, 'skip');
});

test('跨段选区删除：首段留头部 + 末段留尾部，中间段整段消失', () => {
  const blocks = model();
  const batch = rangeCommands(
    blocks,
    { from: 0, fromOffset: 2, to: 1, toOffset: 3 },
    '',
  );
  assert.equal(batch.kind, 'apply');
  assert.deepEqual(batch.focus, { index: 0, offset: 2 });
  const applied = applyCommands(blocks, batch.commands);
  assert.equal(applied.model.length, blocks.length - 1);
  assert.equal(applied.model[0].text, '年度公司营收 增长 20%');
});

test('跨段选区被输入覆盖：新文本落在接缝处', () => {
  const blocks = model();
  const batch = rangeCommands(blocks, { from: 0, fromOffset: 2, to: 1, toOffset: 3 }, '【新】');
  const applied = applyCommands(blocks, batch.commands);
  assert.equal(applied.model[0].text, '年度【新】公司营收 增长 20%');
});

test('跨段选区含表格 → 拒绝（不静默删表格内容）', () => {
  const blocks = model();
  const batch = rangeCommands(blocks, { from: 1, fromOffset: 0, to: 3, toOffset: 0 }, '');
  assert.equal(batch.kind, 'refuse');
});

test('多行粘贴：每行一段，首行接光标前、末行接光标后，顺序正确', () => {
  const blocks = model();
  const body = blocks[1].text;
  const batch = pasteCommands(blocks, 1, body, 3, ['新一', '新二', '新三']);
  assert.equal(batch.kind, 'apply');
  assert.deepEqual(batch.tail, {
    anchor: 1,
    count: 3,
    offset: '新三公司营收 增长 20%'.length,
  });
  const applied = applyCommands(blocks, batch.commands);
  assert.deepEqual(texts(applied.model).slice(0, 4), [
    '年度报告',
    '本年度新一',
    '新二',
    '新三公司营收 增长 20%',
  ]);
  // 新段沿用锚点段的大纲级别（这里都是正文 → 0）。
  assert.equal(applied.model[2].heading, 0);
});

test('多行粘贴进表格块 → 拒绝；空行仍是段落', () => {
  const blocks = model();
  assert.equal(pasteCommands(blocks, 2, '', 0, ['a', 'b']).kind, 'refuse');
  const batch = pasteCommands(blocks, 1, '', 0, ['a', '', 'b']);
  const applied = applyCommands(blocks, batch.commands);
  assert.deepEqual(texts(applied.model).slice(1, 4), ['a', '', 'b']);
});

test('跨块编辑产出的仍是普通 SavePlan：未改块 original，块数随插入增长', () => {
  const { parsed, slices } = baseline();
  const blocks = parsed.blocks.map((block) => ({ ...block }));
  const body = blocks[1].text;
  const batch = pasteCommands(blocks, 1, body, 3, ['A', 'B']);
  const applied = applyCommands(blocks, batch.commands);
  const plan = planFromModel(applied.model, slices, applied.applied);
  assert.equal(plan.kind, 'save');
  // 两行 → 锚点块改写 + 1 个新块（尾部 sectPr 仍在）。
  assert.equal(plan.saveBlocks.length, parsed.blocks.length + 1);
  assert.equal(plan.saveBlocks[0].kind, 'original');
  assert.equal(plan.saveBlocks[1].kind, 'xml');
  assert.equal(paragraphTextOf(plan.saveBlocks[1].xml), '本年度A');
});
