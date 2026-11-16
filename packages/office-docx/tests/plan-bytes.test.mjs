import assert from 'node:assert/strict';
import test from 'node:test';
import {
  describeForeignParts,
  describePart,
  parseDocxBytes,
  planFromBytes,
  planFromModel,
  baselineSlices,
} from '../dist/lib/shared.mjs';
import { exec } from '../dist/engine/exec.mjs';
import { buildDocx, DOCUMENT_XML } from './fixtures.mjs';

// 页面保存通道（P2）的计划生成：把「引擎产出的新字节」还原成块级 SavePlan。
// 夹具的 body 顶层块：0 标题段 / 1 正文段（两个 run）/ 2 表格 / 3 空段 / 4 sectPr。

const BASELINE = buildDocx();

function produced(documentXml) {
  return buildDocx({ documentXml });
}

/** 改一个段落的文字（模拟引擎重新序列化这一块）。 */
function edited() {
  return produced(
    DOCUMENT_XML.replace('本年度公司营收 ', '本年度公司营收（改） '),
  );
}

/** 删掉空段落。 */
function deleted() {
  return produced(DOCUMENT_XML.replace('<w:p/>', ''));
}

/** 在表格后插入一个新段落。 */
function inserted() {
  return produced(
    DOCUMENT_XML.replace(
      '</w:tbl>',
      '</w:tbl>\n<w:p><w:r><w:t>新增段</w:t></w:r></w:p>',
    ),
  );
}

function texts(bytes) {
  return parseDocxBytes(bytes).blocks.map((block) => block.text);
}

test('planFromBytes：字节未变 → 全部 original（字节保真的零改动情形）', () => {
  const built = planFromBytes({ baseline: BASELINE, produced: BASELINE });
  assert.equal(built.ok, true);
  assert.equal(built.plan.kind, 'save');
  assert.equal(built.plan.changedCount, 0);
  assert.equal(built.plan.deletedCount, 0);
  assert.equal(built.stats.kept, parseDocxBytes(BASELINE).blocks.length);
  assert.ok(built.plan.saveBlocks.every((block) => block.kind === 'original'));
});

test('planFromBytes：改了一段 → 只有那一段是 xml，其余保持 original', () => {
  const built = planFromBytes({ baseline: BASELINE, produced: edited() });
  assert.equal(built.ok, true);
  assert.equal(built.plan.changedCount, 1);
  assert.equal(built.plan.deletedCount, 0);
  const xml = built.plan.saveBlocks.filter((block) => block.kind === 'xml');
  assert.equal(xml.length, 1);
  assert.match(xml[0].xml, /本年度公司营收（改）/);
  assert.equal(built.plan.saveBlocks[0].kind, 'original');
  assert.equal(built.plan.saveBlocks[2].kind, 'original');
});

test('planFromBytes：删了一段 → deletedCount 计数，后续块仍按顺序命中', () => {
  const built = planFromBytes({ baseline: BASELINE, produced: deleted() });
  assert.equal(built.ok, true);
  assert.equal(built.plan.changedCount, 0);
  assert.equal(built.plan.deletedCount, 1);
  assert.equal(built.plan.saveBlocks.length, parseDocxBytes(BASELINE).blocks.length - 1);
});

test('planFromBytes：插了一段 → 只有新段是 xml', () => {
  const built = planFromBytes({ baseline: BASELINE, produced: inserted() });
  assert.equal(built.ok, true);
  assert.equal(built.plan.changedCount, 1);
  assert.equal(built.plan.deletedCount, 0);
  const xml = built.plan.saveBlocks.filter((block) => block.kind === 'xml');
  assert.match(xml[0].xml, /新增段/);
});

test('planFromBytes：删一段 + 改一段 混合（顺序匹配不能错位）', () => {
  const mixed = produced(
    DOCUMENT_XML.replace('<w:p/>', '').replace('本年度公司营收 ', '改后的正文 '),
  );
  const built = planFromBytes({ baseline: BASELINE, produced: mixed });
  assert.equal(built.ok, true);
  assert.equal(built.plan.deletedCount, 1);
  assert.equal(built.plan.changedCount, 1);
  const xml = built.plan.saveBlocks.filter((block) => block.kind === 'xml');
  assert.equal(xml.length, 1);
  assert.match(xml[0].xml, /改后的正文/);
});

test('planFromBytes：body 外层变化（分节 / 命名空间）→ 明确拒绝', () => {
  const changedWrapper = produced(
    DOCUMENT_XML.replace('<w:document ', '<w:document w:foo="1" '),
  );
  const built = planFromBytes({ baseline: BASELINE, produced: changedWrapper });
  assert.equal(built.ok, false);
  assert.match(built.error, /正文外层结构/);
});

test('planFromBytes：非法字节 → 明确拒绝（不产生半成品计划）', () => {
  const built = planFromBytes({ baseline: BASELINE, produced: new Uint8Array([1, 2, 3]) });
  assert.equal(built.ok, false);
  assert.match(built.error, /序列化产物不可解析（3B:010203）/);
  // 不是 docx 的基线也不能静默通过
  const bad = planFromBytes({ baseline: new Uint8Array([9]), produced: BASELINE });
  assert.equal(bad.ok, false);
});

test('端到端：planFromBytes 的计划交给 exec，产出与引擎字节等价（未改块逐字节不变）', async () => {
  const p = edited();
  const built = planFromBytes({ baseline: BASELINE, produced: p });
  assert.equal(built.ok, true);

  const ref = 'file:///w/%E6%8A%A5%E5%91%8A.docx';
  const outcome = await exec(
    { ref, baseVersion: 'v1', plan: built.plan, bytes: BASELINE },
    {
      readBytes: async () => ({ bytes: BASELINE, version: 'v1' }),
      tempDir: '',
      importEpoch: 1,
      budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 },
    },
  );
  assert.ok(outcome.bytes instanceof Uint8Array);
  const written = parseDocxBytes(outcome.bytes);
  // 块序 / 文本与引擎产出完全一致
  assert.deepEqual(
    written.blocks.map((block) => block.text),
    texts(p),
  );
  // 未改块逐字节不变：与基线切片相同
  const baseBlocks = parseDocxBytes(BASELINE).blocks;
  assert.equal(written.blocks[0].xml, baseBlocks[0].xml);
  assert.equal(written.blocks[2].xml, baseBlocks[2].xml);
  // manifest 正常给出
  assert.equal(outcome.result.sync.kind, 'manifest');
  assert.equal(outcome.result.sync.blocks.length, written.blocks.length);
});

test('端到端：planFromBytes 与 planFromModel 在「只用 setText 改一段」时产同一形状', () => {
  const parsed = parseDocxBytes(BASELINE);
  const model = parsed.blocks.map((block) => ({ ...block }));
  // 用我们的命令层改同一段文字，得到参考计划
  const reference = planFromModel(
    model.map((block, index) =>
      index === 1
        ? { ...block, xml: block.xml.replace('本年度公司营收 ', '本年度公司营收（改） ') }
        : block,
    ),
    baselineSlices(parsed),
    1,
  );
  const built = planFromBytes({ baseline: BASELINE, produced: edited() });
  assert.equal(built.ok, true);
  assert.deepEqual(
    built.plan.saveBlocks.map((block) => block.kind),
    reference.saveBlocks.map((block) => block.kind),
  );
  assert.equal(built.plan.changedCount, reference.changedCount);
});

test('P5 整包通道：新增图片 + rels + contentTypes → 计划带 parts，exec 写入', async () => {
  const { strToU8 } = await import('fflate');
  const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);
  const p = buildDocx({
    documentXml: DOCUMENT_XML,
    extra: {
      'word/media/image1.png': png,
      'word/_rels/document.xml.rels': strToU8(
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
      ),
    },
  });
  const built = planFromBytes({ baseline: BASELINE, produced: p });
  assert.equal(built.ok, true, built.ok ? '' : built.error);
  const paths = built.plan.parts.map((part) => part.path).sort();
  assert.deepEqual(paths, ['word/_rels/document.xml.rels', 'word/media/image1.png']);
  assert.equal(built.stats.parts, 2);
  // 二进制走 base64，XML/rels 走 xml
  assert.equal('base64' in built.plan.parts.find((part) => part.path.endsWith('.png')), true);

  const outcome = await exec(
    { ref: 'file:///w/a.docx', baseVersion: 'v1', plan: built.plan, bytes: BASELINE },
    {
      readBytes: async () => ({ bytes: BASELINE, version: 'v1' }),
      tempDir: '',
      importEpoch: 1,
      budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 },
    },
  );
  // 图片字节原样落盘
  const { unzipSync } = await import('fflate');
  const archive = unzipSync(outcome.bytes);
  assert.deepEqual([...archive['word/media/image1.png']], [...png]);
  assert.match(new TextDecoder().decode(archive['word/_rels/document.xml.rels']), /rId7/);
  // 未涉及的条目逐字节不变
  const before = unzipSync(BASELINE);
  assert.deepEqual([...archive['word/styles.xml']], [...before['word/styles.xml']]);
  assert.equal(outcome.result.audit.changedParts, 2);
});

test('P5 整包通道：样式表替换 → parts 带 xml；无关条目保持原字节', async () => {
  const { strToU8 } = await import('fflate');
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="New"/></w:styles>
`;
  const baseline = buildDocx({ extra: { 'word/media/keep.txt': strToU8('same') } });
  const p = buildDocx({ stylesXml, extra: { 'word/media/keep.txt': strToU8('same') } });
  const built = planFromBytes({ baseline, produced: p });
  assert.equal(built.ok, true, built.ok ? '' : built.error);
  assert.deepEqual(built.plan.parts.map((part) => part.path), ['word/styles.xml']);
  assert.match(built.plan.parts[0].xml, /styleId="New"/);

  const outcome = await exec(
    { ref: 'file:///w/a.docx', baseVersion: 'v1', plan: built.plan, bytes: baseline },
    {
      readBytes: async () => ({ bytes: baseline, version: 'v1' }),
      tempDir: '',
      importEpoch: 1,
      budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 },
    },
  );
  const { unzipSync } = await import('fflate');
  const archive = unzipSync(outcome.bytes);
  assert.match(new TextDecoder().decode(archive['word/styles.xml']), /styleId="New"/);
  assert.deepEqual([...archive['word/media/keep.txt']], [...strToU8('same')]);
});

test('能力闸门：docProps 元数据变化放行（引擎每次保存都会写修改时间）', async () => {
  const { strToU8 } = await import('fflate');
  const baseline = buildDocx({ extra: { 'docProps/core.xml': strToU8('<core v="1"/>') } });
  const produced = buildDocx({
    documentXml: DOCUMENT_XML.replace('本年度公司营收 ', '本年度公司营收（改） '),
    extra: { 'docProps/core.xml': strToU8('<core v="2"/>') },
  });
  const built = planFromBytes({ baseline, produced });
  assert.equal(built.ok, true);
  assert.equal(built.plan.changedCount, 1);
});

// ─── P4-1 布局组：页面设置 / 分节落在尾部 w:sectPr（仍在 document.xml 内） ───

function withSectPr(replacement) {
  return produced(
    DOCUMENT_XML.replace(
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440"/></w:sectPr>',
      replacement,
    ),
  );
}

test('P4-1 布局：改纸张尺寸（尾部 sectPr）→ 计划带出 xml 块，exec 落盘生效', async () => {
  const p = withSectPr('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440"/></w:sectPr>');
  const built = planFromBytes({ baseline: BASELINE, produced: p });
  assert.equal(built.ok, true, built.ok ? '' : built.error);
  // 只有尾部 sectPr 变了
  assert.equal(built.plan.changedCount, 1);
  const last = built.plan.saveBlocks[built.plan.saveBlocks.length - 1];
  assert.equal(last.kind, 'xml');
  assert.match(last.xml, /w:w="12240"/);

  const outcome = await exec(
    { ref: 'file:///w/a.docx', baseVersion: 'v1', plan: built.plan, bytes: BASELINE },
    {
      readBytes: async () => ({ bytes: BASELINE, version: 'v1' }),
      tempDir: '',
      importEpoch: 1,
      budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 },
    },
  );
  const written = parseDocxBytes(outcome.bytes);
  const sectPr = written.blocks[written.blocks.length - 1];
  assert.equal(sectPr.kind, 'sectPr');
  assert.match(sectPr.xml, /w:w="12240"/);
  // 其余块逐字节不变
  const baseBlocks = parseDocxBytes(BASELINE).blocks;
  assert.deepEqual(
    written.blocks.slice(0, -1).map((block) => block.xml),
    baseBlocks.slice(0, -1).map((block) => block.xml),
  );
});

test('P4-1 布局：加分节符（sectPr 加 type/列）也不影响其它块', async () => {
  const p = withSectPr(
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440"/><w:type w:val="continuous"/><w:cols w:num="2"/></w:sectPr>',
  );
  const built = planFromBytes({ baseline: BASELINE, produced: p });
  assert.equal(built.ok, true, built.ok ? '' : built.error);
  assert.equal(built.plan.changedCount, 1);
  assert.match(built.plan.saveBlocks.at(-1).xml, /w:cols w:num="2"/);
});

test('P5 整包通道：页眉改动 → parts 带 xml（不再是"请在 Word 里做"）', async () => {
  const { strToU8 } = await import('fflate');
  const baseline = buildDocx({ extra: { 'word/header1.xml': strToU8('<w:hdr/>') } });
  const p = buildDocx({
    documentXml: DOCUMENT_XML,
    extra: { 'word/header1.xml': strToU8('<w:hdr><w:p><w:r><w:t>1</w:t></w:r></w:p></w:hdr>') },
  });
  const built = planFromBytes({ baseline, produced: p });
  assert.equal(built.ok, true, built.ok ? '' : built.error);
  assert.deepEqual(built.plan.parts.map((part) => part.path), ['word/header1.xml']);
  assert.match(built.plan.parts[0].xml, />1</);
});

test('部件 → 功能名（超限 / 诊断提示用）', () => {
  assert.equal(describePart('word/header1.xml'), '页眉');
  assert.equal(describePart('word/footer2.xml'), '页脚');
  assert.equal(describePart('word/media/image9.png'), '图片 / 媒体');
  assert.equal(describePart('word/numbering.xml'), '编号 / 列表定义');
  assert.equal(describePart('word/comments.xml'), '批注');
  assert.equal(describePart('unknown/thing.bin'), 'unknown/thing.bin');
  assert.equal(
    describeForeignParts(['word/header1.xml', 'word/media/a.png']),
    '页眉（word/header1.xml）、图片 / 媒体（word/media/a.png）',
  );
});

test('P4 安全：产出的块顺序里 sectPr 不在末尾 → 计划被 exec 拒绝（不产坏文件）', async () => {
  // 反例：有人把段落排在 sectPr 之后（非法 document.xml 结构）
  const bad = produced(
    DOCUMENT_XML.replace(
      '</w:sectPr>',
      '</w:sectPr>\n<w:p><w:r><w:t>越界段</w:t></w:r></w:p>',
    ),
  );
  const built = planFromBytes({ baseline: BASELINE, produced: bad });
  if (!built.ok) {
    assert.match(built.error, /正文外层|无法解析/);
    return;
  }
  // 计划能生成时，exec 必须自己挡住（不得写出 sectPr 之后的块）
  await assert.rejects(
    exec(
      { ref: 'file:///w/a.docx', baseVersion: 'v1', plan: built.plan, bytes: BASELINE },
      {
        readBytes: async () => ({ bytes: BASELINE, version: 'v1' }),
        tempDir: '',
        importEpoch: 1,
        budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 },
      },
    ),
    (error) => {
      assert.match(String(error.code ?? ''), /command\./);
      return true;
    },
  );
});

test('P4 表格组：表格加一行 → 只有表格块是 xml，其它块逐字节不变', async () => {
  const p = produced(
    DOCUMENT_XML.replace(
      '</w:tbl>',
      '<w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    ),
  );
  const built = planFromBytes({ baseline: BASELINE, produced: p });
  assert.equal(built.ok, true, built.ok ? '' : built.error);
  assert.equal(built.plan.changedCount, 1);
  const xmlBlock = built.plan.saveBlocks.find((block) => block.kind === 'xml');
  assert.match(xmlBlock.xml, /<w:tbl>/);
  assert.match(xmlBlock.xml, />C</);

  const outcome = await exec(
    { ref: 'file:///w/a.docx', baseVersion: 'v1', plan: built.plan, bytes: BASELINE },
    {
      readBytes: async () => ({ bytes: BASELINE, version: 'v1' }),
      tempDir: '',
      importEpoch: 1,
      budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 },
    },
  );
  const written = parseDocxBytes(outcome.bytes);
  assert.deepEqual(
    written.blocks.map((block) => block.text),
    texts(p),
  );
  const baseBlocks = parseDocxBytes(BASELINE).blocks;
  assert.equal(written.blocks[0].xml, baseBlocks[0].xml);
  assert.equal(written.blocks[3].xml, baseBlocks[3].xml);
});

// ─── 回读校验：磁盘上的到底是不是这次提交的内容 ─────────────────────

test('sameBlockSlices：块切片一致 → true（重新打包 / 块间空白差异也算一致）', async () => {
  const { sameBlockSlices } = await import('../dist/lib/shared.mjs');
  const { strToU8 } = await import('fflate');
  assert.equal(sameBlockSlices(BASELINE, BASELINE), true);
  // 重新打包（不同压缩、多一个无关部件）：块切片没变 → 视为「磁盘上就是这次的内容」
  const repacked = buildDocx({ extra: { 'word/media/x.bin': strToU8('x') } });
  assert.equal(sameBlockSlices(BASELINE, repacked), true);
  // 块间空白差异（exec 归一化导致）同样算一致
  const spaced = produced(DOCUMENT_XML.replace('<w:p/>', '\n\n<w:p/>'));
  assert.equal(sameBlockSlices(BASELINE, spaced), true);
});

test('sameBlockSlices：改了一段 / 删了一段 / 非法字节 → false', async () => {
  const { sameBlockSlices } = await import('../dist/lib/shared.mjs');
  assert.equal(
    sameBlockSlices(BASELINE, produced(DOCUMENT_XML.replace('本年度公司营收 ', '改 '))),
    false,
  );
  assert.equal(sameBlockSlices(BASELINE, deleted()), false);
  assert.equal(sameBlockSlices(BASELINE, new Uint8Array([1, 2, 3])), false);
});

// ─── P5：删除 part / 计划校验 ─────────────────────────────────────────

function ctxOf(bytes) {
  return {
    readBytes: async () => ({ bytes, version: 'v1' }),
    tempDir: '',
    importEpoch: 1,
    budget: { maxBytes: 64 * 1024 * 1024, deadlineMs: 30_000 },
  };
}

const PLAN_REF = 'file:///w/a.docx';

test('P5 整包通道：删除 part（图片被移除）→ 条目消失，其余不变', async () => {
  const { strToU8, unzipSync } = await import('fflate');
  const baseline = buildDocx({ extra: { 'word/media/gone.png': strToU8('BYE') } });
  const produced = buildDocx();
  const built = planFromBytes({ baseline, produced });
  assert.equal(built.ok, true, built.ok ? '' : built.error);
  assert.deepEqual(built.plan.parts, [{ path: 'word/media/gone.png', delete: true }]);

  const outcome = await exec(
    { ref: PLAN_REF, baseVersion: 'v1', plan: built.plan, bytes: baseline },
    ctxOf(baseline),
  );
  const archive = unzipSync(outcome.bytes);
  assert.equal(archive['word/media/gone.png'], undefined);
  // 正文块逐字节不变（document.xml 整体的块间空白是既有归一化，见 README §5）
  assert.deepEqual(
    parseDocxBytes(outcome.bytes).blocks.map((block) => block.xml),
    parseDocxBytes(baseline).blocks.map((block) => block.xml),
  );
});

test('P5 校验：parts 的非法路径 / 重复 / 缺内容 / 数量超限 → command.invalid', async () => {
  const basePlan = {
    kind: 'save',
    saveBlocks: parseDocxBytes(BASELINE).blocks.map((block) => ({
      kind: 'original',
      docxIndex: block.docxIndex,
    })),
    changedCount: 0,
    deletedCount: 0,
    totalOriginal: parseDocxBytes(BASELINE).blocks.length,
    trailingSectPrChange: false,
  };
  const cases = [
    [{ parts: [{ path: 'word/document.xml', xml: '<w:p/>' }] }, /正文必须走 saveBlocks/],
    [{ parts: [{ path: '../evil.xml', xml: '<w:p/>' }] }, /path 非法/],
    [{ parts: [{ path: '/abs.xml', xml: '<w:p/>' }] }, /path 非法/],
    [{ parts: [{ path: 'word/a.xml', xml: '<a/>' }, { path: 'word/a.xml', xml: '<b/>' }] }, /重复的 path/],
    [{ parts: [{ path: 'word/a.xml' }] }, /xml \/ base64 \/ delete 之一/],
    [{ parts: [{ path: 'word/a.xml', xml: '<a/>', base64: 'AA==' }] }, /xml \/ base64 \/ delete 之一/],
    [
      { parts: Array.from({ length: 25 }, (_, i) => ({ path: `word/p${i}.xml`, xml: '<a/>' })) },
      /parts 超过上限/,
    ],
  ];
  for (const [extra, pattern] of cases) {
    await assert.rejects(
      exec({ ref: PLAN_REF, baseVersion: 'v1', plan: { ...basePlan, ...extra }, bytes: BASELINE }, ctxOf(BASELINE)),
      (error) => {
        assert.equal(error.code, 'command.invalid', String(error.message));
        assert.match(error.message, pattern);
        return true;
      },
    );
  }
});

test('P5：合法的 parts 计划不改变正文块（document.xml 由块序列决定）', async () => {
  const plan = {
    kind: 'save',
    saveBlocks: parseDocxBytes(BASELINE).blocks.map((block) => ({
      kind: 'original',
      docxIndex: block.docxIndex,
    })),
    changedCount: 0,
    deletedCount: 0,
    totalOriginal: parseDocxBytes(BASELINE).blocks.length,
    trailingSectPrChange: false,
    parts: [{ path: 'docProps/core.xml', xml: '<cp:coreProperties xmlns:cp="x"/>' }],
  };
  const outcome = await exec({ ref: PLAN_REF, baseVersion: 'v1', plan, bytes: BASELINE }, ctxOf(BASELINE));
  const { unzipSync } = await import('fflate');
  const archive = unzipSync(outcome.bytes);
  assert.match(new TextDecoder().decode(archive['docProps/core.xml']), /coreProperties/);
  assert.deepEqual(
    parseDocxBytes(outcome.bytes).blocks.map((block) => block.xml),
    parseDocxBytes(BASELINE).blocks.map((block) => block.xml),
  );
  assert.equal(outcome.result.sync.kind, 'manifest');
});

test('P5：只改页眉（正文块完全没变）也要产生计划，不能被当成空保存', async () => {
  const { strToU8 } = await import('fflate');
  const baseline = buildDocx({ extra: { 'word/header1.xml': strToU8('<w:hdr/>') } });
  const p = buildDocx({
    documentXml: DOCUMENT_XML,
    extra: { 'word/header1.xml': strToU8('<w:hdr><w:p><w:r><w:t>页眉</w:t></w:r></w:p></w:hdr>') },
  });
  const built = planFromBytes({ baseline, produced: p });
  assert.equal(built.ok, true, built.ok ? '' : built.error);
  // 正文块 0 改动，但有 1 个 part —— 页面据此必须继续提交
  assert.equal(built.plan.changedCount, 0);
  assert.equal(built.plan.deletedCount, 0);
  assert.equal(built.plan.parts.length, 1);
  assert.equal(built.stats.parts, 1);
});
