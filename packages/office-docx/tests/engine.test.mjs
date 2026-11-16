import assert from 'node:assert/strict';
import test from 'node:test';
import { parseDocxBytes, parseDocumentXml } from '../dist/lib/shared.mjs';
import { plan } from '../dist/engine/plan.mjs';
import { exec } from '../dist/engine/exec.mjs';
import { project, query } from '../dist/engine/read.mjs';
import { runtime } from '../dist/engine/runtime.mjs';
import { buildDocx } from './fixtures.mjs';

const bytes = buildDocx();

function ctx(source = bytes) {
  return {
    readBytes: async () => ({ bytes: source, version: 'v1' }),
    tempDir: '',
    importEpoch: 1,
    budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 },
  };
}

const ref = 'file:///w/%E6%8A%A5%E5%91%8A.docx';

test('engine.plan：guide 三级发现（回给模型的 summary 是完整目录文本）', async () => {
  const top = await plan({ ref, tool: 'docx.guide', params: {} }, ctx());
  assert.match(top.summary, /# docx 命令目录/);
  assert.match(top.summary, /setText/);
  assert.ok(Array.isArray(top.plan.groups));

  const group = await plan({ ref, tool: 'docx.guide', params: { group: 'text' } }, ctx());
  assert.match(group.summary, /### setText/);

  const single = await plan({ ref, tool: 'docx.guide', params: { op: 'setText' } }, ctx());
  assert.match(single.summary, /example:/);
  assert.equal(single.plan.op.op, 'setText');
});

test('engine.plan：命令批 → SavePlan；非法命令 → rejected 计划', async () => {
  const outcome = await plan(
    { ref, tool: 'docx.apply', params: { commands: [{ op: 'setText', index: 1, text: '新正文' }] } },
    ctx(),
  );
  assert.equal(outcome.plan.kind, 'save');
  assert.equal(outcome.plan.changedCount, 1);
  assert.match(outcome.summary, /plan: 1 条命令/);

  const rejected = await plan(
    { ref, tool: 'docx.apply', params: { commands: [{ op: 'deleteBlock', index: 99 }] } },
    ctx(),
  );
  assert.equal(rejected.plan.kind, 'reject');
  assert.equal(rejected.plan.rejected[0].code, 'block.not-found');
  assert.equal(rejected.plan.rejected[0].index, 0);

  const mixed = await plan(
    {
      ref,
      tool: 'docx.apply',
      params: { commands: [{ op: 'check' }, { op: 'setText', index: 1, text: 'x' }] },
    },
    ctx(),
  );
  assert.equal(mixed.plan.kind, 'reject');
  assert.match(mixed.plan.rejected[0].message, /check 必须单独成批/);

  const toolDryRun = await plan(
    { ref, tool: 'docx.apply', params: { dryRun: true, commands: [{ op: 'setText', index: 1, text: 'x' }] } },
    ctx(),
  );
  assert.equal(toolDryRun.plan.dryRun, true);

  const tooMany = await plan(
    { ref, tool: 'docx.apply', params: { commands: Array.from({ length: 201 }, () => ({ op: 'check' })) } },
    ctx(),
  );
  assert.equal(tooMany.plan.rejected[0].code, 'command.too-large');
});

test('engine.exec：唯一字节产生者，产出 manifest 与 audit', async () => {
  const planned = await plan(
    { ref, tool: 'docx.apply', params: { commands: [{ op: 'setText', index: 1, text: '新正文' }] } },
    ctx(),
  );
  const outcome = await exec({ ref, baseVersion: 'v1', plan: planned.plan, bytes }, ctx());
  assert.ok(outcome.bytes instanceof Uint8Array);
  const parsed = parseDocxBytes(outcome.bytes);
  assert.equal(parsed.blocks[1].text, '新正文');
  assert.equal(outcome.result.applied, 1);
  assert.equal(outcome.result.sync.kind, 'manifest');
  assert.equal(outcome.result.sync.blocks.length, parsed.blocks.length);
  assert.match(outcome.result.audit.resultDocHash, /^[0-9a-f]{64}$/);
});

test('engine.exec：dryRun 不产字节；check 只回 audit', async () => {
  const planned = await plan(
    { ref, tool: 'docx.apply', params: { commands: [{ op: 'setText', index: 1, text: 'x' }] } },
    ctx(),
  );
  const dry = await exec({ ref, baseVersion: 'v1', plan: planned.plan, bytes, dryRun: true }, ctx());
  assert.equal(dry.bytes, undefined);
  assert.equal(dry.result.applied, 1);

  const checked = await plan({ ref, tool: 'docx.apply', params: { commands: [{ op: 'check' }] } }, ctx());
  const checkOutcome = await exec({ ref, baseVersion: 'v1', plan: checked.plan, bytes }, ctx());
  assert.equal(checkOutcome.bytes, undefined);
  assert.equal(checkOutcome.result.audit.blocks, 5);
});

test('engine.exec：undo / 非法计划 / 拒绝计划都不产字节', async () => {
  await assert.rejects(
    () => exec({ ref, baseVersion: 'v1', plan: null, bytes, command: 'undo' }, ctx()),
    (error) => error.code === 'command.unsupported',
  );
  await assert.rejects(
    () =>
      exec(
        { ref, baseVersion: 'v1', plan: { kind: 'save', saveBlocks: [{ kind: 'xml', xml: '<w:p>' }] }, bytes },
        ctx(),
      ),
    (error) => error.code === 'command.invalid',
  );
  const rejected = await exec(
    {
      ref,
      baseVersion: 'v1',
      plan: {
        kind: 'reject',
        changedCount: 0,
        deletedCount: 0,
        totalOriginal: 5,
        trailingSectPrChange: false,
        rejected: [{ index: 0, code: 'block.not-found', message: 'no' }],
      },
      bytes,
    },
    ctx(),
  );
  assert.equal(rejected.bytes, undefined);
  assert.equal(rejected.result.rejected.length, 1);
});

test('engine.read：project 打出与命令一致的索引；query 过滤表格', async () => {
  const text = await project({ ref, tool: 'docx.project' }, ctx());
  assert.match(text, /\[0\] h1 年度报告/);
  assert.match(text, /\[1\] p 本年度公司营收/);
  assert.match(text, /\[2\] table 1x2/);
  assert.match(text, /# 用 docx_apply/);

  const result = await query({ ref, query: { where: { type: 'tbl' } } }, ctx());
  assert.equal(result.total, 1);
  assert.equal(result.rows[0].index, 2);
});

test('engine.runtime：能力报告声明 manifest sync / page undo / batch 原子', async () => {
  const report = await runtime();
  assert.equal(report.write, 'plan');
  assert.equal(report.execution, 'isolated');
  assert.equal(report.features.isolation, 'batch');
  assert.equal(report.features.undo, 'page');
  assert.equal(report.features.tokens.baselineSync, 'manifest');
  assert.equal(report.features.tokens.inPage, 'docRev+epoch');
  assert.equal(report.probe.ok, true);
  assert.ok(report.limits.maxFileBytes > 0);
});

test('页面路径与执行点路径产同一计划（契约 §14.1 验收 2）', async () => {
  const parsed = parseDocumentXml(
    (await import('./fixtures.mjs')).DOCUMENT_XML,
  );
  const { applyCommands, baselineSlices, planFromModel } = await import('../dist/lib/shared.mjs');
  const commands = [
    { op: 'setText', index: 1, text: '新正文' },
    { op: 'setHeadingLevel', index: 3, level: 2 },
  ];
  const pageResult = applyCommands(parsed.blocks, commands);
  const pagePlan = planFromModel(pageResult.model, baselineSlices(parsed), pageResult.applied);

  const enginePlan = await plan({ ref, tool: 'docx.apply', params: { commands } }, ctx());
  assert.deepEqual(enginePlan.plan, pagePlan);
});
