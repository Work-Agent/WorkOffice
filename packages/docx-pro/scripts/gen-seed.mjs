#!/usr/bin/env node
// 重新生成 assets/seed.docx（新建文档的空白模板）。
//
// 模板来自 vendor/genoffice 的 buildBlankDocx：A4、简体中文默认字体、
// Heading1-6 / ListParagraph / bullet+ordered 编号定义。平时不需要跑；
// 只有升级 vendor 或想换纸张 / 字体时才重新生成。

import { build } from 'esbuild'
import { mkdir, rm, writeFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const temp = join(root, '.tmp-seed.mjs')

try {
  await build({
    stdin: {
      contents:
        "export { buildBlankDocx } from '../../../vendor/genoffice/packages/docx-engine/src/blank'\n",
      resolveDir: join(root, 'scripts'),
      sourcefile: 'seed-entry.ts',
      loader: 'ts',
    },
    bundle: true,
    platform: 'node',
    format: 'esm',
    outfile: temp,
    logLevel: 'silent',
  })

  const { buildBlankDocx } = await import(pathToFileURL(temp).href)
  const bytes = await buildBlankDocx({
    eastAsiaFont: '宋体',
    paperSize: 'A4',
  })
  await mkdir(join(root, 'assets'), { recursive: true })
  await writeFile(join(root, 'assets/seed.docx'), bytes)
  console.log(`wrote assets/seed.docx (${bytes.length} bytes)`)
} finally {
  await rm(temp, { force: true })
}
