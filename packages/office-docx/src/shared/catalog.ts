// ─── 命令目录（catalog 即可用面） ──────────────────────────────────────
//
// 未实现的 op 不出现在这里 → 分阶段交付不破契约（契约 §9.4）。目录同时是
// `docx_guide` 的单一来源：分组 / schema / example 都从这里生成，三级发现
// 只切片、不手写（避免与命令实现分叉）。

export interface OpSpec {
  op: string;
  summary: string;
  schema: Record<string, unknown>;
  example: Record<string, unknown>;
  /** 坐标层级：L2 = 位置/选区型（`command.stale-index` 后必须重读）。 */
  level: 'L0' | 'L1' | 'L2';
}

export interface GroupSpec {
  name: string;
  summary: string;
  ops: OpSpec[];
}

const INDEX = {
  type: 'number',
  description: '目标块的 docxIndex（sidecar 打印的索引；read 之后再取）',
};

const HEADING = {
  type: 'number',
  description: '0 = 正文，1..6 = Heading N',
};

export const GROUPS: GroupSpec[] = [
  {
    name: 'blocks',
    summary: '段落结构：插入 / 删除 / 移动',
    ops: [
      {
        op: 'insertParagraph',
        level: 'L2',
        summary: '在指定块之后插入段落；after 省略或 null = 文档开头',
        schema: {
          type: 'object',
          properties: { after: INDEX, text: { type: 'string' }, heading: HEADING },
          required: [],
        },
        example: { op: 'insertParagraph', after: 2, text: '新增一段', heading: 0 },
      },
      {
        op: 'deleteBlock',
        level: 'L2',
        summary: '删除一个顶层块（段落或表格）；尾部 sectPr 不可删',
        schema: { type: 'object', properties: { index: INDEX }, required: ['index'] },
        example: { op: 'deleteBlock', index: 3 },
      },
      {
        op: 'moveBlock',
        level: 'L2',
        summary: '把块移动到 to 所指块之后（to=null 移到开头）',
        schema: {
          type: 'object',
          properties: { index: INDEX, to: { ...INDEX, description: '目标块之后；null = 开头' } },
          required: ['index'],
        },
        example: { op: 'moveBlock', index: 5, to: 1 },
      },
    ],
  },
  {
    name: 'text',
    summary: '文本内容：整段替换 / 查找替换',
    ops: [
      {
        op: 'setText',
        level: 'L2',
        summary: '替换段落文本；带 start/end 时只替换段内区间（不动区间外 run）',
        schema: {
          type: 'object',
          properties: {
            index: INDEX,
            text: { type: 'string' },
            start: {
              type: 'number',
              description: '段内起始字符偏移（与 end 成对；缺省 = 整段重写）',
            },
            end: { type: 'number', description: '段内结束字符偏移（不含）' },
          },
          required: ['index', 'text'],
        },
        example: { op: 'setText', index: 1, text: '新的正文内容' },
      },
      {
        op: 'findReplace',
        level: 'L0',
        summary: '全文查找替换（选择器型；在已编辑文档上重新解析）',
        schema: {
          type: 'object',
          properties: {
            find: { type: 'string' },
            replace: { type: 'string' },
            all: { type: 'boolean', description: '缺省 true；false = 只替换第一处' },
          },
          required: ['find', 'replace'],
        },
        example: { op: 'findReplace', find: '2024', replace: '2025' },
      },
    ],
  },
  {
    name: 'format',
    summary: '段落与字符格式',
    ops: [
      {
        op: 'setHeadingLevel',
        level: 'L2',
        summary: '设置标题级别（0 = 正文）；依赖文档已有 Heading N 样式',
        schema: {
          type: 'object',
          properties: { index: INDEX, level: HEADING },
          required: ['index', 'level'],
        },
        example: { op: 'setHeadingLevel', index: 0, level: 1 },
      },
      {
        op: 'setTextStyle',
        level: 'L2',
        summary: '字符 / 段落格式：粗体、斜体、下划线、字号、颜色、对齐；带 start/end 时只改段内区间',
        schema: {
          type: 'object',
          properties: {
            index: INDEX,
            start: {
              type: 'number',
              description: '段内起始字符偏移（与 end 成对；缺省 = 整段）',
            },
            end: { type: 'number', description: '段内结束字符偏移（不含）' },
            bold: { type: 'boolean' },
            italic: { type: 'boolean' },
            underline: { type: 'boolean' },
            size: { type: 'number', description: '字号（磅）' },
            color: { type: 'string', description: 'RRGGBB' },
            align: { type: 'string', enum: ['left', 'center', 'right', 'both'] },
          },
          required: ['index'],
        },
        example: { op: 'setTextStyle', index: 1, bold: true, size: 12 },
      },
    ],
  },
  {
    name: 'check',
    summary: '只读诊断：块数 / 标题 / 空段 / 体积分级',
    ops: [
      {
        op: 'check',
        level: 'L0',
        summary: '统计与自检（不落盘，回执带 audit）',
        schema: { type: 'object', properties: {}, required: [] },
        example: { op: 'check' },
      },
    ],
  },
];

export const OP_INDEX = new Map<string, { group: GroupSpec; spec: OpSpec }>();
for (const group of GROUPS)
  for (const spec of group.ops) OP_INDEX.set(spec.op, { group, spec });

export const OP_NAMES = [...OP_INDEX.keys()];

export interface GuideInput {
  version?: string;
  group?: string;
  op?: string;
  epoch?: number;
  runtime?: unknown;
}

export interface GuideResult {
  /** 回给模型的可读文本（宿主的工具摘要只用字符串，见 host/plugins/tools.ts）。 */
  summary: string;
  /** 结构化版本（进调用详情 / 页面 `invoke` 的结果）。 */
  data: unknown;
}

/** 三级发现：无参 → 分组；group → 该组 schema；op → 单条 schema + example。 */
export function buildGuide(input: GuideInput): GuideResult {
  const header = [
    `# docx 命令目录`,
    input.version ? `version=${short(input.version)}` : '',
    input.epoch !== undefined ? `epoch=${input.epoch}` : '',
    `isolation=batch tokens=docRev+epoch sync=manifest`,
  ]
    .filter(Boolean)
    .join(' ');

  if (input.op) {
    const found = OP_INDEX.get(input.op);
    if (!found)
      return {
        summary: `${header}\n未知 op：${input.op}（可用：${OP_NAMES.join(', ')}）`,
        data: { version: input.version, error: 'op.not-found', ops: OP_NAMES },
      };
    const text = [
      header,
      `## ${found.spec.op} · ${found.group.name}`,
      found.spec.summary,
      `level=${found.spec.level}`,
      'schema:',
      JSON.stringify(found.spec.schema, null, 2),
      'example:',
      JSON.stringify(found.spec.example, null, 2),
    ].join('\n');
    return {
      summary: text,
      data: { version: input.version, op: found.spec, group: found.group.name },
    };
  }

  if (input.group) {
    const group = GROUPS.find((item) => item.name === input.group);
    if (!group)
      return {
        summary: `${header}\n未知分组：${input.group}（可用：${GROUPS.map((item) => item.name).join(', ')}）`,
        data: { version: input.version, error: 'group.not-found' },
      };
    const text = [
      header,
      `## ${group.name} · ${group.summary}`,
      ...group.ops.map((spec) =>
        [ `### ${spec.op} (${spec.level})`, spec.summary, JSON.stringify(spec.schema) ].join('\n'),
      ),
    ].join('\n');
    return { summary: text, data: { version: input.version, group } };
  }

  const text = [
    header,
    ...GROUPS.map(
      (group) =>
        `## ${group.name} · ${group.summary}\n${group.ops.map((spec) => `- ${spec.op} (${spec.level}) ${spec.summary}`).join('\n')}`,
    ),
    '',
    '下一步：guide({ group }) 取 schema；guide({ op }) 取 schema + example。写前先 dryRun。',
  ].join('\n');
  return {
    summary: text,
    data: {
      version: input.version,
      isolation: 'batch',
      tokens: { commit: 'version', inPage: 'docRev+epoch', baselineSync: 'manifest' },
      groups: GROUPS.map((group) => ({
        name: group.name,
        summary: group.summary,
        ops: group.ops.map((spec) => ({ op: spec.op, level: spec.level, summary: spec.summary })),
      })),
    },
  };
}

function short(version: string): string {
  return version.length > 12 ? `${version.slice(0, 12)}…` : version;
}
