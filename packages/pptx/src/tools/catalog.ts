// ─── 统一命令目录（pptx_guide / pptx_apply 的唯一来源） ───────────────
//
// 上游 6 组 75 条注册 op；目录只暴露 aiCallable 的 62 条，再补本插件的
// `editChart`。目录里没有的 op 就是不存在。

import {
  OP_DOCS,
  OP_GROUPS,
  OP_GUIDES,
  OP_SUMMARIES,
  OP_TITLES,
  type OpDoc,
} from './op-catalog.generated'
import { HIDDEN_OPS, PAGE_ONLY_OPS } from '../shared/ops'
import { runtimeReport } from '../shared/runtime'

export const GROUPS = [...OP_GROUPS, 'chart'] as const
export type Group = (typeof GROUPS)[number]

export interface CustomOpDoc extends OpDoc {
  group: 'chart'
}

const EDIT_CHART: CustomOpDoc = {
  sig: '{kind?,categories?,series?,title?,legendPos?,dataLabels?,gridlines?,catAxisTitle?,valAxisTitle?,gapWidthPct?,switchRowCol?,pointColors?} — target:{slide,el}',
  group: 'chart',
  body: [
    '编辑已有图表。高层字段会映射到上游 `setChart` 的 `patch`。',
    '至少给一个字段；只改提供的字段。',
    '无页面可执行。',
  ].join('\n'),
  examples: [
    JSON.stringify(
      {
        op: 'editChart',
        target: { slide: 's_1', el: 'e_ab12' },
        title: '季度收入',
        categories: ['Q1', 'Q2', 'Q3'],
        series: [{ name: '收入', values: [120, 150, 180] }],
      },
      null,
      2,
    ),
  ],
}

const CHART_GUIDE = `# Chart ops

> Chart editing

### editChart
\`${EDIT_CHART.sig}\`

${EDIT_CHART.body}
`

export const CUSTOM_DOCS: Record<string, CustomOpDoc> = { editChart: EDIT_CHART }

export function isAiCallable(name: string): boolean {
  if (name in CUSTOM_DOCS) return true
  const doc = OP_DOCS[name]
  return Boolean(doc && doc.aiCallable !== false && !doc.pending)
}

export function docOf(name: string): OpDoc | undefined {
  return CUSTOM_DOCS[name] ?? (isAiCallable(name) ? OP_DOCS[name] : undefined)
}

export function groupOf(name: string): Group | undefined {
  if (name in CUSTOM_DOCS) return 'chart'
  const doc = OP_DOCS[name]
  if (!doc || !isAiCallable(name)) return undefined
  return doc.group as Group
}

export function isPageOnly(name: string): boolean {
  return PAGE_ONLY_OPS.has(name) || HIDDEN_OPS.has(name)
}

export function isKnownOp(name: string): boolean {
  return name in CUSTOM_DOCS || name in OP_DOCS
}

export function opsInGroup(group: Group): string[] {
  if (group === 'chart') return Object.keys(CUSTOM_DOCS)
  return Object.entries(OP_DOCS)
    .filter(([name, doc]) => doc.group === group && isAiCallable(name))
    .map(([name]) => name)
}

export function allOps(): string[] {
  return [...Object.keys(OP_DOCS).filter(isAiCallable), ...Object.keys(CUSTOM_DOCS)]
}

export function groupGuide(group: Group): string {
  if (group === 'chart') return CHART_GUIDE
  return OP_GUIDES[group] ?? `# ${group}\n\n(no guide)`
}

export function groupSummary(group: Group): string {
  if (group === 'chart') return '图表编辑（editChart）'
  return OP_SUMMARIES[group] ?? group
}

export function groupTitle(group: Group): string {
  if (group === 'chart') return 'Chart ops'
  return OP_TITLES[group] ?? group
}

export function counts(): { total: number; byGroup: Record<string, number>; pageOnly: number } {
  const byGroup: Record<string, number> = {}
  for (const group of GROUPS) byGroup[group] = opsInGroup(group).length
  const names = allOps()
  return {
    total: names.length,
    byGroup,
    pageOnly: names.filter(isPageOnly).length,
  }
}

export function guideText(input: {
  group?: string
  op?: string
  version?: string
  epoch?: number
}): string {
  const report = runtimeReport()
  const runtime = [
    `runtime: write=${report.write} model=${report.model} exec=${report.exec} render=${report.render}`,
    `undo=${report.features.undo} inPage=${report.features.tokens?.inPage} sync=${report.features.tokens?.baselineSync}`,
    `limits: file=${report.limits.maxFileBytes} commit=${report.limits.maxCommitBytes} cache=${report.limits.cacheModels}/${report.limits.cacheBytes}`,
  ].join('\n')

  if (input.op) {
    const doc = docOf(input.op)
    if (!doc)
      return [
        `未知 op：${input.op}`,
        `可用组：${GROUPS.join(' / ')}`,
        '目录里没有的 op 就是不存在；用 group 查组内 schema。',
      ].join('\n')
    return [
      `op: ${input.op}`,
      `group: ${doc.group}`,
      `signature: ${doc.sig}`,
      '',
      doc.body,
      ...(doc.examples.length
        ? ['', 'example:', ...doc.examples.map((example) => `\`\`\`json\n${example}\n\`\`\``)]
        : []),
      '',
      runtime,
      ...(input.version ? [`version: ${input.version}`] : []),
      ...(input.epoch !== undefined ? [`epoch: ${input.epoch}`] : []),
    ].join('\n')
  }

  if (input.group) {
    if (!GROUPS.includes(input.group as Group))
      return `未知组：${input.group}（可用：${GROUPS.join(' / ')}）`
    const group = input.group as Group
    const ops = opsInGroup(group)
    return [
      groupGuide(group),
      '',
      `schema: ${ops.map((name) => `${name} ${docOf(name)?.sig ?? ''}`.trim()).join(' · ')}`,
      `ops=${ops.length} pageOnly=${ops.filter(isPageOnly).length}`,
      '',
      runtime,
    ].join('\n')
  }

  const c = counts()
  return [
    'pptx 命令目录（先 guide，再 apply；不要凭记忆写 op）',
    ...GROUPS.map(
      (group) =>
        `- ${group}: ${c.byGroup[group] ?? 0} 条 — ${groupSummary(group)}`,
    ),
    `ops=${c.total} pageOnly=${c.pageOnly}`,
    '无页面可执行：text / element（除图片填充）/ table / slide（除剪贴板与生成页）/ deck / chart.editChart / insert 的大多数。',
    '需页面：addPicture / replacePicture / setImageFill / addMedia / addNarration / clearNarration / addModel3d / setModel3dView / pasteElements / pasteSlide / insertSlidePptx / setAnimations。',
    '',
    runtime,
    ...(input.version ? [`version: ${input.version}`] : []),
    ...(input.epoch !== undefined ? [`epoch: ${input.epoch}`] : []),
  ].join('\n')
}

export { OP_DOCS }
