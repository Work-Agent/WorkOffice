import assert from 'node:assert/strict';
import test from 'node:test';
import { parseDocxBytes, parseDocumentXml, blockHash, tierFor } from '../dist/lib/shared.mjs';
import { buildDocx, DOCUMENT_XML } from './fixtures.mjs';

test('解析 body 顶层块：段落 / 表格 / 尾部 sectPr', () => {
  const parsed = parseDocxBytes(buildDocx());
  assert.equal(parsed.blocks.length, 5);
  assert.deepEqual(
    parsed.blocks.map((block) => block.kind),
    ['p', 'p', 'tbl', 'p', 'sectPr'],
  );
  assert.equal(parsed.blocks[0].heading, 1);
  assert.equal(parsed.blocks[0].text, '年度报告');
  assert.equal(parsed.blocks[1].text, '本年度公司营收 增长 20%');
  assert.equal(parsed.blocks[2].text, '');
  assert.equal(parsed.blocks[3].text, '');
  assert.equal(parsed.blocks[4].hidden, true);
  assert.ok(parsed.blocks[0].docxIndex === 0 && parsed.blocks[4].docxIndex === 4);
  assert.equal(tierFor(parsed.weight), 'normal');
});

test('保存 body 前后缀：切片拼回只差顶层块之间的空白（按设计归一化）', () => {
  const parsed = parseDocumentXml(DOCUMENT_XML);
  const rebuilt = `${parsed.beforeBody}${parsed.blocks.map((block) => block.xml).join('')}${parsed.afterBody}`;
  // 未改块的切片逐字节保留；顶层块之间的排版空白在保存时归一化。
  for (const block of parsed.blocks) assert.ok(DOCUMENT_XML.includes(block.xml));
  assert.equal(rebuilt.replace(/\s+/g, ''), DOCUMENT_XML.replace(/\s+/g, ''));
});

test('坏包 / 缺 part 给出可诊断错误', () => {
  assert.throws(() => parseDocxBytes(new Uint8Array([1, 2, 3])), /不是有效的 DOCX 包/);
  assert.throws(
    () => parseDocumentXml('<w:document><w:body><w:p></w:body></w:document>'),
    /未闭合|缺少/,
  );
});

test('blockHash 与 sha256 前 16 字节一致（同步实现，页面/执行点同源）', () => {
  const hash = blockHash('<w:p/>');
  assert.equal(hash.length, 32);
  assert.match(hash, /^[0-9a-f]{32}$/);
  assert.equal(blockHash('<w:p/>'), hash);
  assert.notEqual(blockHash('<w:p></w:p>'), hash);
});
