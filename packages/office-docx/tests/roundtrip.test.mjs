import assert from 'node:assert/strict';
import test from 'node:test';
import {
  applyCommands,
  applyManifest,
  baselineSlices,
  parseDocxBytes,
  planFromModel,
} from '../dist/lib/shared.mjs';
import { exec } from '../dist/engine/exec.mjs';
import { buildDocx } from './fixtures.mjs';

// 端到端的「页面两次提交」：验证 I15 的核心风险——
// 第一次提交重排 docxIndex 后，第二次必须用 rebase 后的坐标，且未改块字节不变。

const ref = 'file:///w/report.docx';

function ctx(bytes) {
  return {
    readBytes: async () => ({ bytes, version: 'v1' }),
    tempDir: '',
    importEpoch: 1,
    budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 60_000 },
  };
}

function baselineOf(manifest) {
  const hashes = new Map(manifest.blocks.map((block) => [block.docxIndex, block.hash]));
  const sectPr = (manifest.hidden ?? []).filter((item) => item.kind === 'sectPr').pop();
  return {
    hashes,
    total: manifest.blocks.length,
    trailingSectPrIndex: sectPr?.docxIndex ?? null,
    trailingSectPrHash: sectPr ? (hashes.get(sectPr.docxIndex) ?? null) : null,
  };
}

test('两次连续提交：rebase 后的坐标写对块，未改块切片不变', async () => {
  let bytes = buildDocx();
  let parsed = parseDocxBytes(bytes);
  let baseline = baselineSlices(parsed);
  let model = parsed.blocks.map((block) => ({ ...block }));

  // 第一次：开头插入 + 改第二段
  const first = applyCommands(model, [
    { op: 'insertParagraph', after: 0, text: '插在标题后' },
    { op: 'setText', index: 3, text: '第一次改的正文' },
  ]);
  const plan1 = planFromModel(first.model, baseline, first.applied);
  const out1 = await exec({ ref, baseVersion: 'v1', plan: plan1, bytes }, ctx(bytes));
  bytes = out1.bytes;

  // 页面应用 manifest（rebase）
  const applied1 = applyManifest(first.model, out1.result.sync);
  assert.equal(applied1.mismatched, false);
  model = applied1.model;
  baseline = baselineOf(out1.result.sync);

  // 第二次：用 rebase 后的 docxIndex 改「原第 1 段」
  const target = model.find((block) => block.text === '本年度公司营收 增长 20%');
  assert.ok(target, 'rebase 后仍能按文本找到目标块');
  assert.ok(target.docxIndex >= 0, '新坐标必须是真实下标');
  const second = applyCommands(model, [
    { op: 'setText', index: target.docxIndex, text: '第二次改的正文' },
  ]);
  const plan2 = planFromModel(second.model, baseline, second.applied);
  // 第二次计划里只有目标块是 xml，其余（含第一次插入的块）都指回基线切片。
  const changed = plan2.saveBlocks.filter((block) => block.kind === 'xml');
  assert.equal(changed.length, 1);
  assert.equal(changed[0].docxIndex, target.docxIndex);

  const out2 = await exec({ ref, baseVersion: 'v1', plan: plan2, bytes }, ctx(bytes));
  const parsed2 = parseDocxBytes(out2.bytes);
  const texts = parsed2.blocks.map((block) => block.text);
  assert.ok(texts.includes('插在标题后'), '第一次插入的段落仍在');
  assert.ok(texts.includes('第一次改的正文'), '第一次改的段落仍在');
  assert.ok(texts.includes('第二次改的正文'), '第二次改的是正确的块');
  assert.ok(!texts.includes('本年度公司营收 增长 20%'), '原文本已被替换');
  assert.equal(texts.filter((text) => text === '第二次改的正文').length, 1);

  // 未改块的切片逐字节相同（标题 / 表格 / 尾部 sectPr）。
  const before = parseDocxBytes(buildDocx()).blocks;
  assert.equal(parsed2.blocks[0].xml, before[0].xml);
  assert.equal(
    parsed2.blocks.find((block) => block.kind === 'tbl').xml,
    before.find((block) => block.kind === 'tbl').xml,
  );
  assert.equal(
    parsed2.blocks[parsed2.blocks.length - 1].xml,
    before[before.length - 1].xml,
  );
});

test('空批与只读 check：不产字节、version 不变', async () => {
  const bytes = buildDocx();
  const parsed = parseDocxBytes(bytes);
  const plan = planFromModel(parsed.blocks, baselineSlices(parsed), 0);
  const outcome = await exec({ ref, baseVersion: 'v1', plan, bytes }, ctx(bytes));
  // 全 original 的计划仍会产字节（内容等价），但没有人会提交它：页面只在 dirty 时 flush。
  assert.ok(outcome.bytes instanceof Uint8Array);
  assert.equal(outcome.result.applied, 0);

  const checked = await exec(
    {
      ref,
      baseVersion: 'v1',
      plan: {
        kind: 'check',
        changedCount: 0,
        deletedCount: 0,
        totalOriginal: parsed.blocks.length,
        trailingSectPrChange: false,
        applied: 0,
      },
      bytes,
    },
    ctx(bytes),
  );
  assert.equal(checked.bytes, undefined);
  assert.equal(checked.result.applied, 0);
});
