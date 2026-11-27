#!/usr/bin/env node
// 生成 docx-pro ↔ GenOffice 的逐文件处置清单：
//
//   node scripts/gen-genoffice-inventory.mjs
//
// 输出：
//   docs/docx-pro-genoffice-inventory.csv   逐文件（可直接给脚本/表格用）
//   docs/docx-pro-genoffice-inventory.md    按处置分组的清单
//
// 判定口径见 docs/docx-pro-genoffice-audit.md §1：
//   direct-*        直接 git：submodule 包 / 数据 / 资产，不维护本地副本
//   upstream-ready  文件 + 相对 import 闭包全一致；可整流取，但先别混用路径
//   fork-*          必须保留本地副本并维护改动
//   delete-*        不取/删除
//   upstream-*      上游有、本地没有；未来取 = 一次 fork 变更

import { existsSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const FORK = resolve(root, 'packages/docx-pro/src/editor')
const UP = resolve(root, 'vendor/genoffice/apps/docs/src/renderer')
const OUT_CSV = resolve(root, 'docs/docx-pro-genoffice-inventory.csv')
const OUT_MD = resolve(root, 'docs/docx-pro-genoffice-inventory.md')

if (!existsSync(FORK) || !existsSync(UP)) {
  console.error('缺少 fork 或 vendor/genoffice/apps/docs；先跑 pnpm setup。')
  process.exit(1)
}

const walk = (dir, base = dir, out = []) => {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name)
    if (statSync(p).isDirectory()) walk(p, base, out)
    else out.push(relative(base, p).replaceAll('\\', '/'))
  }
  return out
}
const read = (p) => readFileSync(p)
const text = (p) => readFileSync(p, 'utf8').replaceAll('\r\n', '\n')
const isText = (p) => !read(p).includes(0)
const isFile = (p) => {
  try {
    return statSync(p).isFile()
  } catch {
    return false
  }
}
const isBinaryAsset = (f) => /\.(png|jpe?g|gif|ttf|woff2?)$/.test(f)
const lineSet = (p) => new Set(text(p).split('\n').map((l) => l.trim()).filter(Boolean))

const forkFiles = walk(FORK)
const upFiles = new Set(walk(UP))
const same = new Set()
for (const f of forkFiles) {
  if (!upFiles.has(f)) continue
  const a = join(FORK, f)
  const b = join(UP, f)
  if (read(a).equals(read(b)) || (isText(a) && text(a) === text(b))) same.add(f)
}

// 相对 import 闭包：返回是否一致 + 第一个阻断点
const resolveSpec = (from, spec) => {
  if (!spec.startsWith('.')) return null
  const base = resolve(dirname(join(FORK, from)), spec)
  for (const c of [base, `${base}.ts`, `${base}.tsx`, join(base, 'index.ts'), join(base, 'index.tsx')]) {
    if (isFile(c)) return relative(FORK, c).replaceAll('\\', '/')
  }
  return null
}
const depsOf = (f) => {
  const out = []
  for (const m of text(join(FORK, f)).matchAll(/(?:from\s+|import\s*\(\s*)['"]([^'"]+)['"]/g)) {
    const dep = resolveSpec(f, m[1])
    if (dep) out.push(dep)
  }
  return out
}
const memo = new Map()
const closureInfo = (f, stack = new Set()) => {
  if (memo.has(f)) return memo.get(f)
  if (stack.has(f)) return { safe: true, blockedBy: null }
  stack.add(f)
  let result = { safe: true, blockedBy: null }
  if (!same.has(f)) result = { safe: false, blockedBy: f }
  else {
    for (const dep of depsOf(f)) {
      if (!upFiles.has(dep)) {
        result = { safe: false, blockedBy: `${dep} (only-local)` }
        break
      }
      const r = closureInfo(dep, stack)
      if (!r.safe) {
        result = { safe: false, blockedBy: r.blockedBy ?? dep }
        break
      }
    }
  }
  stack.delete(f)
  memo.set(f, result)
  return result
}

// —— 固定规则表（优先于通用规则）——
const DEAD_AI = new Set([
  'ai/chat-retention.ts',
  'ai/continuation.ts',
  'ai/docs-skill.ts',
  'ai/edit-queue.ts',
  'ai/EditQueueCard.tsx',
  'ai/files-skill.ts',
  'ai/skills-skill.ts',
  'ai/transport.ts',
])
const DELETE_UNUSED = new Set([
  'assets/file-document.png',
  'assets/file-excel.png',
  'assets/file-general.png',
  'assets/file-image.png',
  'assets/file-pdf.png',
  'assets/file-ppt.png',
  'assets/file-video.png',
  'assets/file-voice.png',
  'assets/file-word.png',
])
const FORK_THIN = new Set(['i18n/locale.tsx', 'i18n/strings.ts'])
const TAKE_DATA = (f) =>
  /^i18n\/(app|ribbon)\//.test(f) ||
  /^i18n\/strings-(app|editor|ribbon|table|zotero)\.ts$/.test(f)

function classifyFork(f) {
  if (DEAD_AI.has(f)) return ['delete-ai', 'GenOffice AI 聊天/队列/传输遗留；入口不可达']
  if (f === 'i18n/strings-ai.ts' || f.startsWith('i18n/ai/'))
    return ['delete-ai', 'AI 文案；全仓库已无 t.ai* 调用']
  if (DELETE_UNUSED.has(f)) return ['delete-unused', '本地文件类型图标；src 内无引用']
  if (f === 'ai/mind-map.ts')
    return ['delete-ai', 'GenOffice agent 的 create_mindmap 工具；文档编辑不需要']
  if (FORK_THIN.has(f)) return ['fork-thin', 'i18n 组合层 + 8 键 overlay；保留几十行本地文件']
  if (TAKE_DATA(f)) return ['take-data', '上游语言包是超集；删本地副本（ribbon 需 8 键 overlay）']
  if (f === 'fonts/fonts.css' || f === 'fonts/README.md')
    return ['fork-changed', '字体注册/CSS 与本地 line-metrics、分页版本绑定']
  if (f.startsWith('fonts/') && same.has(f))
    return ['take-asset', '与上游逐字节一致；可从 vendor 取']
  if (f.startsWith('fonts/'))
    return ['fork-changed', '本地重切的 CJK/KR 子集；与本地排版度量绑定']
  if (f.startsWith('assets/') && same.has(f))
    return ['take-asset', '与上游逐字节一致；可从 vendor 取']
  if (f === 'editor/block-id.ts') return ['fork-local', 'WorkStudio 新增的持久 blockId；上游没有']
  if (!upFiles.has(f)) return ['fork-local', '本地新增']

  if (f === 'ai/tools.ts')
    return [
      'fork-changed',
      '保留文档动作分发（insert_content/replace_blocks/...）；删除 web_search/generate_image/generate_mindmap/create_document 等 agent-only 分支',
    ]
  // 与上游有差异
  if (!same.has(f)) {
    if (isBinaryAsset(f)) return ['fork-changed', '本地重切的二进制子集；与本地排版绑定']
    const local = lineSet(join(FORK, f))
    const up = lineSet(join(UP, f))
    const localOnly = [...local].filter((l) => !up.has(l)).length
    const upOnly = [...up].filter((l) => !local.has(l)).length
    if (localOnly === 0) {
      return [
        'fork-stale',
        `本地 ${local.size} 行 ⊂ 上游 ${up.size} 行（本地独有 0 行）：版本落后，重 fork 时可直接换上游`,
      ]
    }
    return ['fork-changed', `本地独有 ${localOnly} 行 / 上游独有 ${upOnly} 行：宿主适配 / AI 剥离 / 本地功能`]
  }

  const c = closureInfo(f)
  if (c.safe) return ['upstream-ready', '文件 + 相对 import 闭包与上游全一致']
  return ['fork-blocked', `文件相同，但依赖 ${c.blockedBy} 已改；混用会把上游兄弟模块带进来`]
}

// —— 已直接来自 submodule 的包（文档能力）——
const PACKAGES = [
  ['@genoffice/docx-engine', 'DOCX 解析 / 写回 / zip / math / metafile', 'doc-engine'],
  ['@genoffice/ui', 'Dropdown / ColorPicker / ShapePreview / 字体候选 / ImageViewer（非 AI 基础件）', 'doc-ui'],
  ['@genoffice/i18n', 'createI18n / htmlLang / htmlDir / ribbonCommon*', 'doc-data'],
  ['@genoffice/zip-gate', 'ZIP 安全解包', 'doc-engine'],
  ['@genoffice/font-metrics', 'FaceVerticalMetrics', 'doc-data'],
  ['@genoffice/electron-utils', 'safe-external-url / headless-export', 'doc-host'],
  ['@genoffice/project-store', 'ProjectApi 类型', 'doc-host'],
  ['@genoffice/pptx-render', 'preset-geometry（跨格式形状几何）', 'doc-editor'],
  ['@genoffice/agent-core/redact-*', 'redact-core / redact-range（文档脱敏；不是 agent 循环）', 'doc-privacy'],
  ['@genoffice/agent-core (loop/skills/transport)', 'GenOffice agent 循环、skills、IPC transport；不取', 'ai-agent'],
  ['@genoffice/ai-provider (runtime)', '多 provider LLM 客户端 / AI_PROVIDERS；不取', 'ai-agent'],
]
const PACKAGE_AI = new Set([
  '@genoffice/agent-core (loop/skills/transport)',
  '@genoffice/ai-provider (runtime)',
])

// —— 上游有、fork 没有 ——
function classifyUpstream(f) {
  if (f === 'ai/AiPanel.tsx' || f === 'ai/MindMapCard.tsx' || f === 'ai/revision-author.ts')
    return ['upstream-ai', 'GenOffice 自有 AI UI；永不取']
  if (f === 'components/AiAskPopover.tsx') return ['upstream-ai', 'GenOffice 自有 AI 入口；永不取']
  if (f.startsWith('i18n/ai/')) return ['upstream-ai', 'GenOffice AI 文案；永不取']
  if (/^i18n\/(app|ribbon)\/(tr|hu)\.ts$/.test(f))
    return ['upstream-data', '新语言包；随 i18n 一起取']
  if (f.startsWith('fonts/')) return ['upstream-asset', '新字体（度量别名生成）；随排版同步取']
  if (f === 'index.html') return ['upstream-skip', '上游 Electron 页面入口；WorkStudio 用 page/index.html']
  return ['upstream-sync', '上游新功能/重构；取 = 一次 fork 变更，按文件移植']
}

function scopeOf(f, category, side) {
  if (category === 'delete-ai' || category === 'upstream-ai') return 'ai-agent'
  if (side === 'upstream') {
    if (category === 'upstream-data' || category === 'upstream-asset') return 'doc-data'
    if (category === 'upstream-skip') return 'build'
    return 'doc-upstream'
  }
  if (f.startsWith('i18n/') || f.startsWith('fonts/') || f.startsWith('assets/')) return 'doc-data'
  if (f === 'ai/redact.ts' || f === 'ai/redact-view.ts') return 'doc-privacy'
  if (f.startsWith('ai/')) return 'doc-command'
  if (
    f.startsWith('components/') ||
    f === 'App.tsx' ||
    f === 'main.tsx' ||
    f === 'styles.css' ||
    f === 'ui-theme.ts'
  )
    return 'doc-ui'
  if (f.startsWith('editor/')) return 'doc-editor'
  if (f.startsWith('zotero/')) return 'doc-integration'
  return 'doc-core'
}
const rows = []
for (const [path, reason, scope] of PACKAGES)
  rows.push({
    side: 'package',
    path,
    category: PACKAGE_AI.has(path) ? 'delete-ai' : 'direct-package',
    scope,
    reason,
  })
for (const f of forkFiles) {
  const [category, reason] = classifyFork(f)
  rows.push({ side: 'fork', path: f, category, scope: scopeOf(f, category, 'fork'), reason })
}
for (const f of [...upFiles].filter((f) => !forkFiles.includes(f))) {
  const [category, reason] = classifyUpstream(f)
  rows.push({ side: 'upstream', path: f, category, scope: scopeOf(f, category, 'upstream'), reason })
}

// —— CSV（带引号转义；Excel/脚本都能读）——
const esc = (s) => `"${String(s).replaceAll('"', '""')}"`
writeFileSync(
  OUT_CSV,
  [
    'scope,side,path,category,reason',
    ...rows.map((r) => [r.scope, r.side, r.path, r.category, r.reason].map(esc).join(',')),
  ].join('\n') + '\n',
)

// —— Markdown ——
const ORDER = [
  'direct-package',
  'take-data',
  'take-asset',
  'upstream-ready',
  'fork-stale',
  'fork-blocked',
  'fork-changed',
  'fork-thin',
  'fork-local',
  'delete-ai',
  'delete-unused',
  'upstream-data',
  'upstream-asset',
  'upstream-sync',
  'upstream-ai',
  'upstream-skip',
]
const TITLES = {
  'direct-package': '直接 git：已在 submodule 里（零本地副本）',
  'take-data': '直接 git：删本地副本，改为上游 import（数据）',
  'take-asset': '直接 git：删本地副本，改为上游资产',
  'upstream-ready': '可整流取：文件 + 依赖闭包全一致（建议随重 fork 再做）',
  'fork-stale': '必须 fork：文件是上游子集（版本落后；重 fork 时可换上游）',
  'fork-blocked': '必须 fork：文件相同但依赖已改',
  'fork-changed': '必须 fork：本地有独立改动',
  'fork-thin': '必须 fork（轻）：保留几十行本地组合/overlay',
  'fork-local': '必须 fork：本地新增',
  'delete-ai': '删除：GenOffice AI 残留',
  'delete-unused': '删除：无引用资产',
  'upstream-data': '上游有：新语言包（随 i18n 取）',
  'upstream-asset': '上游有：新字体（随排版同步取）',
  'upstream-sync': '上游有：新功能/重构（取 = fork 变更）',
  'upstream-ai': '上游有：AI 部分（永不取）',
  'upstream-skip': '上游有：不需要',
}
const DESC = {
  'direct-package': '不落本地副本，直接从 `vendor/genoffice` 打包',
  'take-data': '删除本地副本，改为从上游 import；ribbon 补 8 键 overlay',
  'take-asset': '资源从 vendor 取；不再复制',
  'upstream-ready': '可整体取上游；但先别混用路径',
  'fork-stale': '本地只是上游旧快照；现在仍留本地，重 fork 时替换',
  'fork-blocked': '文件一样，依赖闭包已改；保留本地',
  'fork-changed': '本地 patch；必须保留并随上游同步',
  'fork-thin': '保留很薄的本地层（组合 + overlay）',
  'fork-local': '上游没有，必须本地维护',
  'delete-ai': '入口不可达或已无调用；删',
  'delete-unused': '无引用；删',
  'upstream-data': '随对应数据一起取',
  'upstream-asset': '随排版同步取',
  'upstream-sync': '取它就要改 fork',
  'upstream-ai': '永不取',
  'upstream-skip': '不需要',
}
const byCat = new Map()
for (const r of rows) {
  if (!byCat.has(r.category)) byCat.set(r.category, [])
  byCat.get(r.category).push(r)
}
const count = (c) => (byCat.get(c) ?? []).length
const lines = []
lines.push('# docx-pro ↔ GenOffice 逐文件处置清单')
lines.push('')
lines.push('> 由 `node scripts/gen-genoffice-inventory.mjs` 生成，请勿手改。')
lines.push('> 判定口径与原因见 [`docx-pro-genoffice-audit.md`](./docx-pro-genoffice-audit.md)。')
lines.push('')
lines.push('## 怎么用')
lines.push('')
lines.push('- **想少维护代码**：只看 `direct-package` / `take-*`，这些不需要 fork。')
lines.push('- **只看 Office DOC**：筛选 CSV 的 `scope` 列，排除 `ai-agent`。')
lines.push('- **准备重 fork**：`fork-stale` 是“本地只是旧快照”的文件，可优先用上游替换；`fork-changed` 才是需要重放的 patch。')
lines.push('- **只想删**：`delete-ai` + `delete-unused`，不改行为。')
lines.push('- **想加新功能**：`upstream-sync` 是上游已有、本地没有的清单；每一项都要按文件移植（= fork 变更）。')
lines.push('')
lines.push('## 按范围（scope）')
lines.push('')
lines.push('| scope | 数量 | 含义 |')
lines.push('| --- | ---: | --- |')
const SCOPE_DESC = {
  'doc-engine': 'DOCX 格式引擎（直接 submodule）',
  'doc-editor': 'ProseMirror 编辑器 / 分页 / 渲染',
  'doc-ui': 'Ribbon、对话框、导航、App 壳',
  'doc-command': '文档命令实现（`ai/ops`、`ai/tools` 等，名字带 ai 但不是 agent）',
  'doc-privacy': '脱敏 / withheld spans（服务任何模型，含我们自己的）',
  'doc-data': 'i18n / 字体 / 图标',
  'doc-host': '保存、打印、导出、项目存储',
  'doc-integration': 'Zotero 等外部集成',
  'doc-upstream': '上游有、本地没有的文档功能',
  'build': '构建入口 / 不需要的页面壳',
  'ai-agent': 'GenOffice AI 智能体（不要）',
  'doc-core': '其他本地适配 / 共享逻辑',
}
const scopeCount = new Map()
for (const r of rows) scopeCount.set(r.scope, (scopeCount.get(r.scope) ?? 0) + 1)
for (const [s, n] of [...scopeCount].sort((a, b) => b[1] - a[1]))
  lines.push(`| \`${s}\` | ${n} | ${SCOPE_DESC[s] ?? ''} |`)
lines.push('')
lines.push('## 汇总')
lines.push('')
lines.push('| 处置 | 数量 | 含义 |')
lines.push('| --- | ---: | --- |')
for (const c of ORDER) lines.push(`| \`${c}\` | ${count(c)} | ${DESC[c]} |`)
lines.push('')
lines.push('逐文件 CSV：[`docx-pro-genoffice-inventory.csv`](./docx-pro-genoffice-inventory.csv)')
lines.push('')
for (const c of ORDER) {
  const list = (byCat.get(c) ?? []).sort((a, b) => a.path.localeCompare(b.path))
  lines.push(`## ${TITLES[c]}（${list.length}）`)
  lines.push('')
  if (!list.length) {
    lines.push('(none)')
    lines.push('')
    continue
  }
  lines.push('| 文件 | scope | 原因 |')
  lines.push('| --- | --- | --- |')
  for (const r of list) lines.push(`| \`${r.path}\` | \`${r.scope}\` | ${r.reason} |`)
  lines.push('')
}
writeFileSync(OUT_MD, lines.join('\n'))

console.log(`wrote ${relative(root, OUT_CSV)} (${rows.length} rows)`)
console.log(`wrote ${relative(root, OUT_MD)}`)
for (const c of ORDER) console.log(`  ${c.padEnd(16)} ${count(c)}`)
