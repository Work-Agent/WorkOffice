import assert from 'node:assert/strict';
import test from 'node:test';
import { strFromU8, strToU8, unzipSync } from 'fflate';
import { crc32, rewriteZipPart, zipEntryNames, zipEntryRecords } from '../dist/lib/zip.mjs';
import { DOCUMENT_XML, buildDocx } from './fixtures.mjs';

test('crc32 标准向量', () => {
  assert.equal(crc32(strToU8('hello')), 0x3610a686);
  assert.equal(crc32(new Uint8Array(0)), 0);
});

test('只重写 document.xml；其余条目压缩流逐字节不变', () => {
  const original = buildDocx();
  const before = zipEntryRecords(original);
  const nextXml = DOCUMENT_XML.replace('增长 20%', '增长 42%');
  const rewritten = rewriteZipPart(original, 'word/document.xml', strToU8(nextXml));
  const after = zipEntryRecords(rewritten);

  assert.deepEqual(zipEntryNames(rewritten).sort(), zipEntryNames(original).sort());
  for (const name of zipEntryNames(original)) {
    if (name === 'word/document.xml') continue;
    assert.deepEqual(
      Array.from(after.get(name)),
      Array.from(before.get(name)),
      `${name} 的原始 local record 必须逐字节不变`,
    );
  }
  assert.notDeepEqual(
    Array.from(after.get('word/document.xml')),
    Array.from(before.get('word/document.xml')),
  );

  const unpacked = unzipSync(rewritten);
  assert.equal(strFromU8(unpacked['word/document.xml']), nextXml);
  assert.equal(
    strFromU8(unpacked['word/styles.xml']),
    strFromU8(unzipSync(original)['word/styles.xml']),
  );
});

test('缺少目标 part / 非法 zip 时报错（不产出半成品）', () => {
  assert.throws(
    () => rewriteZipPart(buildDocx(), 'word/missing.xml', strToU8('<x/>')),
    /缺少 word\/missing.xml/,
  );
  assert.throws(() => rewriteZipPart(new Uint8Array([0, 1, 2, 3]), 'word/document.xml', strToU8('<x/>')), /ZIP/);
});
