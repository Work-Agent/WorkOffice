import assert from 'node:assert/strict';
import test from 'node:test';
import {
  CommandRejected,
  applyCommands,
  baselineSlices,
  parseDocxBytes,
  paragraphHeadingOf,
  paragraphTextOf,
  planFromModel,
  firstRunStyleOf,
} from '../dist/lib/shared.mjs';
import { buildDocx } from './fixtures.mjs';

function baseline() {
  const parsed = parseDocxBytes(buildDocx());
  return { parsed, slices: baselineSlices(parsed) };
}

test('setText：只改目标块，保留段属性与首个 run 的字符属性', () => {
  const { parsed, slices } = baseline();
  const model = parsed.blocks.map((block) => ({ ...block }));
  const result = applyCommands(model, [{ op: 'setText', index: 1, text: '新的正文' }]);
  const changed = result.model[1].xml;
  assert.equal(paragraphTextOf(changed), '新的正文');
  assert.match(changed, /<w:jc w:val="left"\/>/);
  assert.match(changed, /<w:i\/>/);
  const plan = planFromModel(result.model, slices, result.applied);
  assert.equal(plan.kind, 'save');
  assert.equal(plan.changedCount, 1);
  assert.equal(plan.applied, 1);
  assert.equal(plan.saveBlocks[0].kind, 'original');
  assert.equal(plan.saveBlocks[1].kind, 'xml');
  assert.equal(plan.saveBlocks[4].kind, 'original');
});

test('未改动模型 → 计划全部 original（逆计划语义）', () => {
  const { parsed, slices } = baseline();
  const plan = planFromModel(parsed.blocks, slices, 0);
  assert.equal(plan.changedCount, 0);
  assert.equal(plan.deletedCount, 0);
  assert.ok(plan.saveBlocks.every((block) => block.kind === 'original'));
});

test('结构命令：插入（负坐标）/ 删除 / 移动', () => {
  const { parsed } = baseline();
  const model = parsed.blocks.map((block) => ({ ...block }));
  const inserted = applyCommands(model, [
    { op: 'insertParagraph', after: 1, text: '新增', heading: 2 },
  ]);
  assert.equal(inserted.model.length, parsed.blocks.length + 1);
  const fresh = inserted.model[2];
  assert.ok(fresh.docxIndex < 0, '新块使用负坐标占位');
  assert.equal(fresh.heading, 2);

  const deleted = applyCommands(inserted.model, [{ op: 'deleteBlock', index: 0 }]);
  assert.equal(deleted.model.length, inserted.model.length - 1);

  const moved = applyCommands(inserted.model, [{ op: 'moveBlock', index: -1, to: 0 }]);
  assert.equal(moved.model[1].text, '新增');
  assert.equal(moved.model[0].text, '年度报告');
});

test('格式命令：标题 / 粗斜体 / 字号 / 对齐', () => {
  const { parsed } = baseline();
  const model = parsed.blocks.map((block) => ({ ...block }));
  const result = applyCommands(model, [
    { op: 'setHeadingLevel', index: 1, level: 2 },
    { op: 'setTextStyle', index: 1, bold: true, italic: false, size: 14, align: 'center' },
  ]);
  assert.equal(paragraphHeadingOf(result.model[1].xml), 2);
  const style = firstRunStyleOf(result.model[1].xml);
  assert.equal(style.bold, true);
  assert.equal(style.italic, false);
  assert.equal(style.size, 14);
  assert.match(result.model[1].xml, /<w:jc w:val="center"\/>/);
  // 两段都应有 rPr 补丁（原文 2 个 run）
  assert.equal((result.model[1].xml.match(/<w:sz w:val="28"\/>/g) ?? []).length, 2);
});

test('findReplace 是选择器型命令：命中块重写，未命中即拒绝', () => {
  const { parsed } = baseline();
  const model = parsed.blocks.map((block) => ({ ...block }));
  const result = applyCommands(model, [{ op: 'findReplace', find: '20%', replace: '30%' }]);
  assert.equal(result.model[1].text, '本年度公司营收 增长 30%');
  assert.throws(
    () => applyCommands(model, [{ op: 'findReplace', find: '不存在', replace: 'x' }]),
    (error) => error instanceof CommandRejected && error.code === 'command.not-found',
  );
});

test('批级原子：一条非法即整批抛错，并带命令下标', () => {
  const { parsed } = baseline();
  const model = parsed.blocks.map((block) => ({ ...block }));
  assert.throws(
    () =>
      applyCommands(model, [
        { op: 'setText', index: 1, text: 'ok' },
        { op: 'setText', index: 999, text: 'bad' },
      ]),
    (error) =>
      error instanceof CommandRejected &&
      error.index === 1 &&
      error.code === 'block.not-found',
  );
  assert.throws(
    () => applyCommands(model, [{ op: 'deleteBlock', index: 4 }]),
    (error) => error.code === 'block.hidden',
  );
  assert.throws(
    () => applyCommands(model, [{ op: 'setText', index: 2, text: '表格' }]),
    (error) => error.code === 'block.not-paragraph',
  );
  assert.throws(
    () => applyCommands(model, [{ op: 'nope' }]),
    (error) => error.code === 'command.invalid',
  );
});

test('自闭合段落 <w:p/> 改写后仍是良构元素', () => {
  const { parsed } = baseline();
  const model = parsed.blocks.map((block) => ({ ...block }));
  const result = applyCommands(model, [{ op: 'setText', index: 3, text: '填充空段' }]);
  assert.equal(paragraphTextOf(result.model[3].xml), '填充空段');
  assert.match(result.model[3].xml, /^<w:p><w:r>/);
  assert.match(result.model[3].xml, /<\/w:r><\/w:p>$/);
  const styled = applyCommands(model, [
    { op: 'setTextStyle', index: 3, bold: true },
    { op: 'setHeadingLevel', index: 3, level: 1 },
  ]);
  assert.match(styled.model[3].xml, /^<w:p><w:pPr>/);
  assert.match(styled.model[3].xml, /<\/w:p>$/);
});

test('尾部 sectPr 丢失时计划自动补回（防分节信息丢失）', () => {
  const { parsed, slices } = baseline();
  const model = parsed.blocks.filter((block) => block.kind !== 'sectPr');
  const plan = planFromModel(model, slices, 0);
  const last = plan.saveBlocks[plan.saveBlocks.length - 1];
  assert.equal(last.kind, 'original');
  assert.equal(last.docxIndex, 4);
  assert.equal(plan.trailingSectPrChange, true);
});
