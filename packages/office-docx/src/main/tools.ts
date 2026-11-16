// ─── engine.tools：主进程只声明（不导出执行函数） ──────────────────────
//
// 执行体由宿主生成（契约 §7.2）：只读 → 执行点 plan；写 → 有页面转发页面
// 代理，无页面走权威层锁内 plan → exec → write。

export const tools = [
  {
    name: 'docx_guide',
    label: 'Word 命令目录',
    description:
      '取 .docx 的命令词汇：无参给分组与 op 名；group 给该组 schema；op 给单条 schema + example。写命令前必须先调它，不要凭记忆写 op。',
    kinds: ['docx', 'docm', 'dotx'],
    readOnly: true,
    parameters: {
      type: 'object',
      properties: {
        ref: { type: 'string', description: '资源 id，如 file:///w/报告.docx' },
        group: { type: 'string', description: '分组名：blocks / text / format / check' },
        op: { type: 'string', description: '单个 op 名' },
      },
      required: ['ref'],
    },
    plan: 'docx.guide',
  },
  {
    name: 'docx_apply',
    label: 'Word 命令批',
    description:
      '原子执行一批命令（批上限 200）。一条被拒则整批不落盘，回执给出 index 与原因；改那一条后整批重发。index 用 docx_guide / read 得到的块索引。禁止用 write/edit 改 .docx。',
    kinds: ['docx', 'docm', 'dotx'],
    readOnly: false,
    requiresBaseVersion: true,
    tokens: 'docRev+epoch',
    maxItems: 200,
    parameters: {
      type: 'object',
      properties: {
        ref: { type: 'string' },
        commands: {
          type: 'array',
          description: '命令批；op 见 docx_guide',
          items: { type: 'object' },
        },
        version: {
          type: 'string',
          description: '无页面写必填；有页面时由页面会话持有，可省略',
        },
        epoch: {
          type: 'number',
          description: 'agent 读到的基线代数；含位置型命令时需要',
        },
        dryRun: { type: 'boolean', description: '只校验，不落盘' },
      },
      required: ['ref', 'commands'],
    },
    plan: 'docx.apply',
  },
];

export default tools;
