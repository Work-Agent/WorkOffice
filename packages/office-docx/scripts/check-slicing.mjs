#!/usr/bin/env node
// ─── 切分一致性校验（P0） ─────────────────────────────────────────────
//
// 迁移到 docs 渲染器后，页面的 `docxIndex` 来自 `@genoffice/docx-engine` 的
// `parsed.blocks`，而本包的 `exec` 用 `shared/document.ts` 的 `parseDocxBytes`。
// 两者只要对「body 顶层怎么切」不一致，`{kind:'original', docxIndex}` 就会回写
// 错误的基线切片（静默写坏文件）。这个脚本把差异量化出来。
//
//   node scripts/check-slicing.mjs              # 只跑内置合成夹具
//   node scripts/check-slicing.mjs a.docx b.docx  # 追加真实文档
//
// 期望：
//   - 不含 `w:sdt` / `w:altChunk` 的文档 → 必须逐块对齐（exit 1 表示失败）
//   - 含上述结构的文档 → 记录为「已知分歧」，迁移期由只读门槛兜底（P3 解决）
//
// vendor/genoffice 未初始化时只打印提示，不算失败。

import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const repo = resolve(root, '../..');
const vendor = join(repo, 'vendor/genoffice');
const engineEntry = join(vendor, 'packages/docx-engine/src/index.ts');

if (!existsSync(engineEntry)) {
  console.log('skip: vendor/genoffice 未初始化（先跑仓库根的 scripts/setup-genoffice.mjs）');
  process.exit(0);
}

// ── 1. 生成探针入口（esbuild 打包后由 node 运行） ─────────────────────
const work = join(root, '.tmp');
mkdirSync(work, { recursive: true });
const entry = join(work, 'slicing-entry.ts');
writeFileSync(
  entry,
  `
import { readFileSync } from 'node:fs';
import { strToU8, zipSync } from 'fflate';
import { parseDocx } from '@genoffice/docx-engine';
import { blockHash, parseDocxBytes } from ${JSON.stringify(join(root, 'src/shared/index.ts'))};

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function docx(bodyXml: string, extra: Record<string, Uint8Array> = {}) {
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' + W + '" xmlns:r="' + R + '"><w:body>' +
    bodyXml +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' +
    '</w:body></w:document>';
  const entries = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
    ),
    'word/document.xml': strToU8(documentXml),
    'word/_rels/document.xml.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk.htm"/>' +
      '</Relationships>',
    ),
    'word/afchunk.htm': strToU8('<html><body><p>altChunk</p></body></html>'),
    ...extra,
  };
  return zipSync(entries, { level: 6 });
}

const PARA = (text: string) => '<w:p><w:r><w:t>' + text + '</w:t></w:r></w:p>';

const FIXTURES: Array<{ name: string; expectAlign: boolean; bytes: Uint8Array }> = [
  { name: 'plain', expectAlign: true, bytes: docx(PARA('第一段') + PARA('第二段')) },
  {
    name: 'runs-and-table',
    expectAlign: true,
    bytes: docx(
      '<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>斜体</w:t></w:r><w:r><w:t>正文</w:t></w:r></w:p>' +
        '<w:tbl><w:tblPr/><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
        '<w:p/>',
    ),
  },
  {
    name: 'sdt-block',
    expectAlign: false,
    bytes: docx(
      PARA('前') +
        '<w:sdt><w:sdtPr><w:alias w:val="控件"/></w:sdtPr><w:sdtContent>' +
        PARA('控件内') +
        '</w:sdtContent></w:sdt>' +
        PARA('后'),
    ),
  },
  {
    name: 'sdt-multi-block',
    expectAlign: false,
    bytes: docx(
      PARA('前') +
        '<w:sdt><w:sdtPr><w:alias w:val=多段控件/></w:sdtPr><w:sdtContent>' +
        PARA('控件段一') +
        PARA('控件段二') +
        '</w:sdtContent></w:sdt>' +
        PARA('后'),
    ),
  },
  {
    name: 'sdt-table-and-para',
    expectAlign: false,
    bytes: docx(
      '<w:sdt><w:sdtPr><w:tag w:val=表/></w:sdtPr><w:sdtContent>' +
        '<w:tbl><w:tblPr/><w:tr><w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
        PARA('表后段') +
        '</w:sdtContent></w:sdt>',
    ),
  },
  {
    name: 'nested-sdt',
    expectAlign: false,
    bytes: docx(
      '<w:sdt><w:sdtPr><w:tag w:val="外"/></w:sdtPr><w:sdtContent>' +
        '<w:sdt><w:sdtPr><w:tag w:val="内"/></w:sdtPr><w:sdtContent>' +
        PARA('嵌套') +
        '</w:sdtContent></w:sdt>' +
        '</w:sdtContent></w:sdt>' +
        PARA('尾'),
    ),
  },
  {
    name: 'altchunk',
    expectAlign: false,
    bytes: docx(PARA('前') + '<w:altChunk r:id="rId9"/>' + PARA('后')),
  },
];

export interface SliceReport {
  name: string;
  expectAlign: boolean;
  ours: number;
  engine: number;
  aligned: boolean;
  firstMismatch: string | null;
}

function summarize(name: string, expectAlign: boolean, ours: any[], engine: any[]): SliceReport {
  let firstMismatch: string | null = null;
  for (let i = 0; i < Math.min(ours.length, engine.length); i += 1) {
    const a = ours[i];
    const b = engine[i];
    const sameIndex = a.docxIndex === b.docxIndex;
    const sameHash = blockHash(a.xml) === blockHash(b.originalXml ?? '');
    if (!sameIndex || !sameHash) {
      firstMismatch =
        '#' + i + ' ours(' + a.docxIndex + ',' + blockHash(a.xml) + ') engine(' +
        String(b.docxIndex) + ',' + blockHash(b.originalXml ?? '') + ')';
      break;
    }
  }
  if (!firstMismatch && ours.length !== engine.length) {
    firstMismatch = 'count ' + ours.length + ' vs ' + engine.length;
  }
  return {
    name,
    expectAlign,
    ours: ours.length,
    engine: engine.length,
    aligned: firstMismatch === null,
    firstMismatch,
  };
}

export async function probeFile(name: string, expectAlign: boolean, bytes: Uint8Array): Promise<SliceReport> {
  const ours = parseDocxBytes(bytes).blocks;
  let engineBlocks: any[] = [];
  let engineError: string | null = null;
  try {
    const parsed = await parseDocx(bytes);
    engineBlocks = parsed.blocks;
  } catch (error) {
    engineError = error instanceof Error ? error.message : String(error);
  }
  if (engineError) {
    return {
      name,
      expectAlign,
      ours: ours.length,
      engine: -1,
      aligned: false,
      firstMismatch: 'engine parse error: ' + engineError,
    };
  }
  return summarize(name, expectAlign, ours, engineBlocks);
}

export const fixtures = FIXTURES;
export function readDocx(path: string): Uint8Array {
  return new Uint8Array(readFileSync(path));
}
`,
  { encoding: 'utf8' },
);

// ── 2. esbuild 打包（@genoffice/* 映射到 vendor 源码） ────────────────
const { build } = await import('esbuild');
const genoffice = {
  name: 'genoffice-source',
  setup(builder) {
    builder.onResolve({ filter: /^@genoffice\// }, (args) => {
      const rest = args.path.slice('@genoffice/'.length);
      const slash = rest.indexOf('/');
      const pkg = slash < 0 ? rest : rest.slice(0, slash);
      const sub = slash < 0 ? '' : rest.slice(slash + 1);
      const base = join(vendor, 'packages', pkg, 'src', sub);
      for (const candidate of [
        `${base}.ts`,
        `${base}.tsx`,
        join(base, 'index.ts'),
        join(base, 'index.tsx'),
      ]) {
        if (existsSync(candidate)) return { path: candidate };
      }
      return null;
    });
  },
};

const outfile = join(work, 'slicing-entry.mjs');
await build({
  entryPoints: [entry],
  outfile,
  bundle: true,
  format: 'esm',
  platform: 'node',
  target: 'node22',
  plugins: [genoffice],
  logLevel: 'silent',
  banner: { js: "import { createRequire as __cr } from 'node:module'; const require = __cr(import.meta.url);" },
});

const probe = await import(pathToFileURL(outfile).href);

// ── 3. 跑夹具 + 命令行文档 ───────────────────────────────────────────
const reports = [];
for (const fixture of probe.fixtures) {
  reports.push(await probe.probeFile(fixture.name, fixture.expectAlign, fixture.bytes));
}
for (const path of process.argv.slice(2)) {
  const name = path.split(/[\\/]/).pop() ?? path;
  const bytes = probe.readDocx(resolve(process.cwd(), path));
  const text = Buffer.from(bytes).toString('latin1');
  const sensitive = /<w:sdt[ >]/.test(text) || /<w:altChunk[ >]/.test(text);
  reports.push(await probe.probeFile(name, !sensitive, bytes));
}

const pad = (text, width) => String(text).padEnd(width);
console.log(
  pad('夹具 / 文档', 26) + pad('ours', 6) + pad('engine', 8) + pad('对齐', 6) + '首个不一致',
);
console.log('-'.repeat(96));
let failed = 0;
let known = 0;
for (const report of reports) {
  const verdict = report.aligned ? (report.expectAlign ? '✓' : '✓(超预期)') : report.expectAlign ? '✗' : '已知';
  if (!report.aligned && report.expectAlign) failed += 1;
  if (!report.aligned && !report.expectAlign) known += 1;
  console.log(
    pad(report.name, 26) +
      pad(report.ours, 6) +
      pad(report.engine, 8) +
      pad(verdict, 6) +
      (report.firstMismatch ?? '—'),
  );
}
console.log('-'.repeat(96));
console.log(
  `${reports.length} 个样本：${reports.length - failed - known} 对齐，${known} 已知分歧（sdt/altChunk，P3 解决），${failed} 失败`,
);

rmSync(work, { recursive: true, force: true });
process.exit(failed > 0 ? 1 : 0);
