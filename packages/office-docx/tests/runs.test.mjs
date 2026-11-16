import assert from 'node:assert/strict';
import test from 'node:test';
import {
  CommandRejected,
  applyCommands,
  applyRunStyle,
  isWellFormedElement,
  baselineSlices,
  paragraphTextOf,
  parseDocxBytes,
  patchRunProps,
  planFromModel,
  runSpans,
  replaceRunText,
  runStyleAt,
  splitRunAt,
} from '../dist/lib/shared.mjs';
import { buildDocx } from './fixtures.mjs';

// 夹具第 1 块是两个 run：
//   run1 = <w:rPr><w:i/></w:rPr><w:t xml:space="preserve">本年度公司营收 </w:t>   → 字符 0..8
//   run2 = <w:t>增长 20%</w:t>                                              → 字符 8..14
const SECOND_RUN = '<w:r><w:t>增长 20%</w:t></w:r>';

function baseline() {
  const parsed = parseDocxBytes(buildDocx());
  const slices = baselineSlices(parsed);
  const blocks = parsed.blocks.map((block) => ({ ...block }));
  return { parsed, slices, blocks, body: blocks[1].xml, text: blocks[1].text };
}

function model() {
  return baseline().blocks;
}

function runElements(xml) {
  return xml.match(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g) ?? [];
}

function hasItalic(run) {
  return /<w:i\s*\/>/.test(run);
}

function hasBold(run) {
  return /<w:b(?:\s+w:val="(?:1|true)")?\s*\/>/.test(run);
}

test('runSpans：字符区间与 paragraphTextOf 同口径（含 run 间空格）', () => {
  const { body, text } = baseline();
  const spans = runSpans(body);
  assert.equal(text.length, 14);
  assert.deepEqual(
    spans.map((span) => [span.charStart, span.charEnd]),
    [
      [0, 8],
      [8, 14],
    ],
  );
  // 区间首尾与元素字符串区间对得上
  assert.equal(body.slice(spans[0].start, spans[0].end).startsWith('<w:r'), true);
  assert.equal(body.slice(spans[1].start, spans[1].end), SECOND_RUN);
});

test('区间格式：只给区间内的 run 加 rPr，其余 run 逐字节不变', () => {
  const { body } = baseline();
  const styled = applyRunStyle(body, 0, 3, { bold: true });
  assert.equal(paragraphTextOf(styled), paragraphTextOf(body));
  const runs = runElements(styled);
  assert.equal(runs.length, 3, '端点落在 run 内部时必须切开');
  assert.equal(hasBold(runs[0]), true);
  assert.equal(hasBold(runs[1]), false);
  assert.equal(hasBold(runs[2]), false);
  // 未覆盖的 run 原样保留（含未切开的第二个 run）
  assert.equal(styled.includes(SECOND_RUN), true);
});

test('区间格式：两端都落在 run 中间 → 三段，只有中段带格式', () => {
  const { body, text } = baseline();
  const styled = applyRunStyle(body, 8, 10, { bold: true });
  assert.equal(paragraphTextOf(styled), text);
  const runs = runElements(styled);
  assert.equal(runs.length, 3);
  assert.deepEqual(
    runs.map(hasBold),
    [false, true, false],
  );
  // 切开的右半保留原 rPr 语义（run2 原本没有 rPr）
  assert.equal(runs[2].includes('<w:rPr>'), false);
  assert.equal(runs[1].includes('<w:rPr>'), true);
});

test('区间格式：覆盖整段 ≡ patchRunProps（旧行为回归）', () => {
  const { body, text } = baseline();
  const ranged = applyRunStyle(body, 0, text.length, { bold: true, size: 14 });
  assert.equal(ranged, patchRunProps(body, { bold: true, size: 14 }));
});

test('splitRunAt：边界在 run 边界上 / 越界时原样返回', () => {
  const { body, text } = baseline();
  assert.equal(splitRunAt(body, 0), body);
  assert.equal(splitRunAt(body, text.length), body);
  assert.equal(splitRunAt(body, 8), body, '段 8 正好是 run1/run2 的分界');
  assert.notEqual(splitRunAt(body, 3), body);
});

test('runStyleAt：按字符偏移取所在 run 的字符属性（工具栏 toggle 基准）', () => {
  const { body } = baseline();
  assert.equal(runStyleAt(body, 0).italic, true);
  assert.deepEqual(runStyleAt(body, 10), {});
  // 空段落没有 run
  const { parsed } = baseline();
  const empty = parsed.blocks[3].xml;
  assert.deepEqual(runSpans(empty), []);
  assert.deepEqual(runStyleAt(empty, 0), {});
});

test('ops：setTextStyle 带 start/end 只改区间，非法区间被拒', () => {
  const blocks = model();
  const applied = applyCommands(blocks, [
    { op: 'setTextStyle', index: 1, start: 0, end: 3, bold: true },
  ]);
  const runs = runElements(applied.model[1].xml);
  assert.deepEqual(
    runs.map(hasBold),
    [true, false, false],
  );

  const invalid = [
    { op: 'setTextStyle', index: 1, start: 3, end: 3, bold: true },
    { op: 'setTextStyle', index: 1, start: 4, end: 2, bold: true },
    { op: 'setTextStyle', index: 1, start: -1, end: 2, bold: true },
    { op: 'setTextStyle', index: 1, start: 0, end: 99, bold: true },
    { op: 'setTextStyle', index: 1, start: 0, bold: true },
  ];
  for (const command of invalid) {
    assert.throws(
      () => applyCommands(blocks, [command]),
      (error) => error instanceof CommandRejected && error.code === 'command.invalid',
      JSON.stringify(command),
    );
  }
});

test('ops：区间格式仍是普通 SavePlan——其他块保持 original', () => {
  const { parsed, slices, blocks } = baseline();
  const applied = applyCommands(blocks, [
    { op: 'setTextStyle', index: 1, start: 8, end: 10, bold: true },
  ]);
  const plan = planFromModel(applied.model, slices, applied.applied);
  assert.equal(plan.kind, 'save');
  assert.equal(plan.changedCount, 1);
  assert.equal(plan.saveBlocks[1].kind, 'xml');
  assert.equal(plan.saveBlocks[0].kind, 'original');
  assert.equal(plan.saveBlocks.length, parsed.blocks.length);
});

test('区间改字：插入沿用光标前 run 的 rPr（区间格式才不会被打字冲掉）', () => {
  const { body, text } = baseline();
  const tail = replaceRunText(body, text.length, text.length, '。');
  assert.equal(paragraphTextOf(tail), `${text}。`);
  // 末尾插入：前一个 run 是 run2（无 rPr），插入的 run 也不该有 rPr
  const appended = runElements(tail)[runElements(tail).length - 1];
  assert.equal(appended.includes('<w:rPr>'), false);
  assert.equal(tail.includes(SECOND_RUN), true);

  const head = replaceRunText(body, 0, 0, '【序】');
  assert.equal(paragraphTextOf(head), `【序】${text}`);
  // 段首插入：没有前一个 run → 用后一个 run（run1，带 <w:i/>）的格式
  assert.equal(hasItalic(runElements(head)[0]), true);
});

test('区间改字：段中间插入会切开 run，且不动两侧文本', () => {
  const { body, text } = baseline();
  const edited = replaceRunText(body, 10, 10, 'X');
  assert.equal(paragraphTextOf(edited), `${text.slice(0, 10)}X${text.slice(10)}`);
  const runs = runElements(edited);
  assert.equal(runs.length, 4, 'run1 + run2 左半 + 新 run + run2 右半');
  assert.equal(runs[2].includes('X'), true);
  assert.equal(runs[2].includes('<w:rPr>'), false, '沿用 run2 的无格式');
});

test('区间改字：跨 run 删除只删区间内的字符，两侧 rPr 保留', () => {
  const { body, text } = baseline();
  const edited = replaceRunText(body, 2, 9, '');
  assert.equal(text, '本年度公司营收 增长 20%');
  assert.equal(paragraphTextOf(edited), '本年长 20%');
  const runs = runElements(edited);
  assert.equal(runs.length, 2);
  assert.equal(hasItalic(runs[0]), true, '左侧仍是原 <w:i/> run');
  assert.equal(hasItalic(runs[1]), false);
});

test('区间格式 + 后续打字：半段加粗能存活', () => {
  const { body, text } = baseline();
  const bolded = applyRunStyle(body, 0, 3, { bold: true });
  const edited = replaceRunText(bolded, 3, 3, '！');
  assert.equal(paragraphTextOf(edited), `本年度！${text.slice(3)}`);
  const runs = runElements(edited);
  assert.equal(hasBold(runs[0]), true, '加粗的前半段仍在');
  // 插入的字符落在加粗区间内 → 也跟着加粗（Word 行为）
  assert.equal(runs.filter(hasBold).length, 2);
  assert.equal(hasBold(runs[runs.length - 1]), false, '区间外的尾巴保持原样');
});

test('ops：setText 带 start/end 只改区间，非法区间被拒', () => {
  const blocks = model();
  const applied = applyCommands(blocks, [
    { op: 'setText', index: 1, start: 8, end: 14, text: '收尾' },
  ]);
  assert.equal(applied.model[1].text, '本年度公司营收 收尾');
  // 文本变了但左侧 run 的 <w:i/> 还在
  assert.equal(hasItalic(runElements(applied.model[1].xml)[0]), true);

  const invalid = [
    { op: 'setText', index: 1, start: 5, end: 2, text: 'x' },
    { op: 'setText', index: 1, start: -1, end: 2, text: 'x' },
    { op: 'setText', index: 1, start: 0, end: 99, text: 'x' },
    { op: 'setText', index: 1, start: 0, text: 'x' },
  ];
  for (const command of invalid) {
    assert.throws(
      () => applyCommands(blocks, [command]),
      (error) => error instanceof CommandRejected && error.code === 'command.invalid',
      JSON.stringify(command),
    );
  }
});

test('区间改字：空段落（<w:p/>）走整段重写，产出良构元素', () => {
  const { parsed } = baseline();
  const empty = parsed.blocks[3].xml;
  assert.equal(empty, '<w:p/>');
  const typed = replaceRunText(empty, 0, 0, '开头');
  assert.equal(paragraphTextOf(typed), '开头');
  assert.equal(isWellFormedElement(typed), true);
  assert.equal(typed.includes('<w:p/>'), false, '自闭合开标签必须被展开');
  // 二段合成（先打字再追加）也不坏
  const again = replaceRunText(typed, 2, 2, '。');
  assert.equal(paragraphTextOf(again), '开头。');
});
