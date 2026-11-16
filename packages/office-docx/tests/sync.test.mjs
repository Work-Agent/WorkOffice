import assert from 'node:assert/strict';
import test from 'node:test';
import {
  applyCommands,
  applyManifest,
  baselineSlices,
  buildManifest,
  parseDocumentXml,
  parseDocxBytes,
  planFromModel,
  spliceBody,
} from '../dist/lib/shared.mjs';
import { buildDocx } from './fixtures.mjs';

test('计划 → 新 document.xml → manifest → 页面 rebase 后可以继续建计划', () => {
  const parsed = parseDocxBytes(buildDocx());
  const slices = baselineSlices(parsed);
  const model = parsed.blocks.map((block) => ({ ...block }));
  const result = applyCommands(model, [
    { op: 'setText', index: 1, text: '改过的正文' },
    { op: 'insertParagraph', after: 0, text: '插在标题后' },
  ]);
  const plan = planFromModel(result.model, slices, result.applied);

  const nextXml = spliceBody(parsed, plan);
  const nextParsed = parseDocumentXml(nextXml);
  assert.equal(nextParsed.blocks.length, parsed.blocks.length + 1);
  assert.equal(nextParsed.blocks[1].text, '插在标题后');
  assert.equal(nextParsed.blocks[2].text, '改过的正文');
  assert.equal(nextParsed.blocks[nextParsed.blocks.length - 1].kind, 'sectPr');

  const manifest = buildManifest(parsed, plan);
  assert.equal(manifest.blocks.length, nextParsed.blocks.length);
  assert.ok(manifest.blocks.some((block) => typeof block.xml === 'string'));

  const applied = applyManifest(parsed.blocks, manifest);
  assert.equal(applied.mismatched, false);
  assert.equal(applied.model.length, nextParsed.blocks.length);
  applied.model.forEach((block, index) => {
    assert.equal(block.docxIndex, index);
    assert.equal(block.text, nextParsed.blocks[index].text);
  });
  // 新块拿到了真实下标（页面可以继续用 docxIndex 作为坐标）。
  assert.equal(applied.model[1].docxIndex, 1);

  // 页面路径：manifest 应用在「产出计划的模型」上，id 保留（rebase 不破坏会话身份）。
  const pageApplied = applyManifest(result.model, manifest);
  assert.equal(pageApplied.mismatched, false);
  assert.equal(pageApplied.model[1].id, result.model[1].id);
  assert.equal(pageApplied.model[1].docxIndex, 1);

  // rebase 后的模型对新基线生成「全 original」计划（不重复写）。
  const newBaseline = {
    hashes: new Map(manifest.blocks.map((block) => [block.docxIndex, block.hash])),
    total: manifest.blocks.length,
    trailingSectPrIndex: manifest.blocks.length - 1,
    trailingSectPrHash: manifest.blocks[manifest.blocks.length - 1].hash,
  };
  const again = planFromModel(applied.model, newBaseline, 0);
  assert.equal(again.changedCount, 0);
  assert.equal(again.deletedCount, 0);
});

test('manifest 对不上时不猜：返回 mismatched', () => {
  const parsed = parseDocxBytes(buildDocx());
  const applied = applyManifest(parsed.blocks, {
    blocks: [
      { docxIndex: 0, hash: 'deadbeefdeadbeefdeadbeefdeadbeef' },
      { docxIndex: 1, hash: 'feedfacefeedfacefeedfacefeedface' },
    ],
  });
  assert.equal(applied.mismatched, true);
});

test('删除块后 manifest 不再包含它（deletedCount 反映基线丢失）', () => {
  const parsed = parseDocxBytes(buildDocx());
  const slices = baselineSlices(parsed);
  const model = parsed.blocks.map((block) => ({ ...block }));
  const result = applyCommands(model, [{ op: 'deleteBlock', index: 2 }]);
  const plan = planFromModel(result.model, slices, result.applied);
  assert.equal(plan.deletedCount, 1);
  const manifest = buildManifest(parsed, plan);
  assert.equal(manifest.blocks.length, parsed.blocks.length - 1);
  const applied = applyManifest(parsed.blocks, manifest);
  assert.equal(applied.mismatched, false);
  assert.equal(applied.model.length, parsed.blocks.length - 1);
});
