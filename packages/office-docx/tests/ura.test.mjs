import assert from 'node:assert/strict';
import test from 'node:test';
import { buildDocx } from './fixtures.mjs';

const registered = {};
globalThis.__workstudioSdk = {
  registerCodec: (value) => (registered.codec = value),
  registerSidecar: (value) => (registered.sidecar = value),
  registerStore: (value) => (registered.store = value),
  registerPresenter: (value) => (registered.presenter = value),
};

await import(`${new URL('../dist/ura.js', import.meta.url).href}?t=${Date.now()}`);

const MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const ref = 'file:///w/%E6%8A%A5%E5%91%8A%E6%96%87%E6%A1%A3.docx';
const bytes = buildDocx();
const decoded = registered.codec.decode({ bytes, version: 'abcdef0123456789', mime: MIME }, ref);

test('codec 注册与解码：块投影 / 大纲 / 分级', () => {
  assert.equal(registered.codec.id, 'office-docx');
  assert.deepEqual(registered.codec.extensions, ['docx', 'docm', 'dotx']);
  assert.equal(decoded.doc.name, '报告文档.docx');
  assert.equal(decoded.doc.blockCount, 4);
  assert.equal(decoded.doc.tableCount, 1);
  assert.equal(decoded.doc.tier, 'normal');
  assert.match(decoded.anchorMap, /年度报告/);
  assert.deepEqual(decoded.outline, [{ level: 1, title: '年度报告' }]);
  // 整份字节绝不进结构化文档。
  assert.equal('bytes' in decoded.doc, false);
});

test('codec.encode 拒绝整份写（字节只由 exec 产生）', () => {
  assert.throws(() => registered.codec.encode(decoded, { type: 'replace-all' }), /command\.forbidden/);
});

test('sidecar 打印 docxIndex 与命令目标一致', () => {
  const sidecar = registered.sidecar.sidecar(decoded, { bytes, version: 'abcdef0123456789', mime: MIME }, ref);
  assert.equal(sidecar.type, 'text');
  assert.match(sidecar.content, /^# 报告文档\.docx blocks=4/);
  assert.match(sidecar.content, /\[0\] h1 年度报告/);
  assert.match(sidecar.content, /\[2\] table 1x2/);
  assert.match(sidecar.content, /\[n\] 即命令的 index/);
});

test('store 查询结构化块', () => {
  const result = registered.store.query(decoded, { where: { type: 'tbl' } });
  assert.equal(result.total, 1);
  const selected = registered.store.query(decoded, { select: ['index', 'text'], limit: 2 });
  assert.deepEqual(Object.keys(selected.rows[0]).sort(), ['index', 'text']);
});

test('presenter：只放元数据，不放整份内容', () => {
  const presentation = registered.presenter.present(decoded, { bytes, version: 'abcdef0123456789', mime: MIME }, ref);
  assert.equal(presentation.type, 'page');
  assert.equal(presentation.surface, 'office-docx');
  assert.equal(presentation.page, '');
  assert.equal(presentation.data.kind, 'docx');
  assert.equal(presentation.data.blocks, 4);
  assert.equal('bytes' in presentation.data, false);
  assert.equal('text' in presentation.data, false);
  assert.ok(presentation.data.size > 0);
  assert.deepEqual(registered.presenter.requires, ['write']);
});
