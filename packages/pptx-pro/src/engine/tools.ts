// ─── engine.tools：主进程只声明（不导出执行函数） ──────────────────────
//
// 5 个工具：目录 + 读 + 检查 + 新建 + 批量写。
// 目录里的 63 条命令（上游 62 条 AI 可调用 + editChart）按需用 pptx_guide 取。

import type { ToolDecl } from '../tools/types'

const KINDS = ['pptx', 'pptm', 'potx']
const REF = { type: 'string', description: '资源 id，如 file:///w/2026年度汇报.pptx' }

const guide: ToolDecl = {
  name: 'pptx_guide',
  label: 'PPT 命令目录',
  kinds: KINDS,
  readOnly: true,
  description:
    '取 .pptx 的命令词汇。无参：分组与数量；group：该组全部 schema；op：单条 schema 与示例。' +
    'text / element / table / slide / deck / chart 多数无页面可执行；图片/媒体/剪贴板/3D/旁白需页面。' +
    '写命令前先调它，不要凭记忆写 op。',
  tokens: 'docRev',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      group: { type: 'string', enum: ['text', 'element', 'insert', 'table', 'slide', 'deck', 'chart'] },
      op: { type: 'string', description: '单个 op 名' },
    },
    required: ['ref'],
  },
}

const read: ToolDecl = {
  name: 'pptx_read',
  label: 'PPT 读取',
  kinds: KINDS,
  readOnly: true,
  description:
    '读演示文稿：概览与页列表（context / outline / slides）、单页元素（slide）、备注（notes）、' +
    '批注（comments）、图表（charts）、表格（tables）、主题（theme）、结构诊断（audit）、' +
    '全文查找（find）。只读，不产字节。target.slide 用 s_* 或 0-based 索引，target.el 用 e_*。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      what: {
        type: 'string',
        enum: ['context', 'outline', 'slides', 'slide', 'notes', 'comments', 'charts', 'tables', 'theme', 'audit', 'find'],
        description: '要读什么（默认 context）',
      },
      slide: { type: 'integer', description: 'slide：0-based 页索引' },
      query: { type: 'string', description: 'find：查询文本' },
      limit: { type: 'integer', description: 'find：上限' },
    },
    required: ['ref'],
  },
}

const create: ToolDecl = {
  name: 'pptx_create',
  label: 'PPT 新建',
  kinds: KINDS,
  readOnly: false,
  requiresBaseVersion: false,
  description:
    '从零新建 .pptx。ref 是尚不存在的目标资源 id。两条路径：\n' +
    '① 设计稿（推荐）：pages 是页面数组。每页写「版式 + 内容」即可 —— ' +
    '{layout, type?, title, kicker?, subtitle?, items?:[{title?,value?,note?,bullet?}], body?, footnote?, image?}，' +
    '坐标由引擎按版式算（共 16 种版式：cover_* / three_column_cards / kpi_cards_row / hero_big_number / ' +
    'two_column_comparison / two_by_two_grid / timeline_horizontal / left_text_right_image / title_and_bullets / closing_*）；' +
    '不写 layout 会按内容自动挑。需要精确控制时改用 {background?, elements:[…]}（1280×720 px，text/shape/image 三种元素）。' +
    'pptx_check({styles:true}) 列出风格预设，pptx_create({style:"consulting-navy"}) 一行换皮。\n' +
    '② 文字稿（快但素）：slides 是 [{title,body?,bullets?,notes?,background?}]，或 Markdown content。\n' +
    '空内容 = 空白演示文稿。目标已存在时拒绝；修改已有演示文稿请用 pptx_apply。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      pages: {
        type: 'array',
        description:
          '设计稿页面数组。每页写「版式 + 内容」（layout + title/items/body…），坐标由引擎算；' +
          '也可以给 {background?,elements:[…]} 直接摆 px 元素',
        items: {
          type: 'object',
          properties: {
            layout: {
              type: 'string',
              description:
                '版式名。cover_typography_hero / cover_dark_minimal / cover_split_color / cover_magazine / ' +
                'cover_split_image / cover_full_image_overlay / three_column_cards / kpi_cards_row / ' +
                'hero_big_number / two_column_comparison / two_by_two_grid / timeline_horizontal / ' +
                'left_text_right_image / full_image_text_overlay / title_and_bullets / closing_thank_you / closing_cta；' +
                '不写则按内容自动挑',
            },
            type: { type: 'string', enum: ['cover', 'content', 'data', 'closing'] },
            title: { type: 'string' },
            kicker: { type: 'string', description: '眉标：章节号 / 分类 / 日期，小字在标题上方' },
            subtitle: { type: 'string' },
            items: {
              type: 'array',
              description: '卡片 / 数字 / 节点的内容；数量决定栏数（3 项 → 三栏）',
              items: {
                type: 'object',
                properties: {
                  title: { type: 'string', description: '卡片标题' },
                  value: { type: 'string', description: '大数字 / 关键值，如 +31%' },
                  note: { type: 'string', description: '说明；可多行（\n）' },
                  bullet: { type: 'string' },
                },
              },
            },
            body: {
              oneOf: [{ type: 'string' }, { type: 'array', items: { type: 'string' } }],
              description: '项目符号正文；每项一条',
            },
            footnote: { type: 'string', description: '来源 / 备注，放页面底部' },
            image: {
              type: 'string',
              description: '图片：http(s) / file:// / 相对 deck 目录的路径 / data: URL',
            },
            imageCaption: { type: 'string' },
            background: { type: 'string', description: '#RRGGBB 或 #RRGGBBAA（不给按版式定）' },
            style: { type: 'string', description: '本页换一套预设（只影响这一页）' },
            elements: {
              type: 'array',
              description: '低层：px 元素数组（给了就不再用 layout），先画的在底层',
              items: { type: 'object' },
            },
          },
        },
      },
      style: {
        oneOf: [{ type: 'string' }, { type: 'object' }],
        description:
          '风格：预设 id（consulting-navy / tech-dark / editorial-warm / minimal-mono / fresh-teal / bold-red）、' +
          '中文名，或 {preset, colors:{accent,bg,card,…}}；不传默认 consulting-navy。' +
          'pptx_check({styles:true}) 可列出全部',
      },
      slides: {
        type: 'array',
        description: '文字稿页规格列表（只有标题 + 正文两段文字）',
        items: {
          type: 'object',
          properties: {
            title: { type: 'string' },
            body: { oneOf: [{ type: 'string' }, { type: 'array', items: { type: 'string' } }] },
            bullets: { type: 'array', items: { type: 'string' } },
            notes: { type: 'string' },
            background: { type: 'string', description: '#RRGGBB' },
          },
        },
      },
      content: { type: 'string', description: '轻量 Markdown' },
      theme: { type: 'object', description: 'applyTheme 的 {name,colors,majorFont?,minorFont?}' },
      size: { oneOf: [{ type: 'string', enum: ['16:9', '4:3'] }, { type: 'object' }] },
    },
    required: ['ref'],
  },
}

const check: ToolDecl = {
  name: 'pptx_check',
  label: 'PPT 检查',
  kinds: KINDS,
  readOnly: true,
  description:
    '写之前/写之后查一遍，不产字节：\n' +
    '① outline：核心 hook、每页 type+layout、占位符、版式重复（design.md 的硬规则）；\n' +
    '② pages：逐页解析页面 spec，并和 outline 条目、style.md 调色板对账；\n' +
    '③ 不给 outline/pages（或 audit:true）：对现有文件跑几何审计（越界/文字溢出/重叠/图片变形），每条附可直接套用的 setTransform；\n' +
    '④ skeleton:true：抽出这张 deck 的可复用版面骨架（背景/标题框/品牌图槽/装饰），拄进 style.md 下次复用。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      outline: {
        oneOf: [{ type: 'object' }, { type: 'string' }],
        description: 'outline 对象（或它的 JSON 文本）：{core_hook,pages:[{title,type,layout,brief,image_queries}]}',
      },
      pages: {
        oneOf: [{ type: 'array' }, { type: 'object' }],
        description: '页面 spec 数组（或 {pages:[…]}）；给了就和 outline/style 对账',
      },
      style: {
        oneOf: [{ type: 'string' }, { type: 'object' }],
        description:
          '预设 id / {preset}/{colors}，或 style.md 的 Markdown 文本（解析 #RRGGBB 当调色板）。' +
          '高层页面按预设上色；自定义 style.md 只做调色板检查',
      },
      styles: { type: 'boolean', description: '列出可用风格预设（按内容挑一个）' },
      audit: { type: 'boolean', description: '对 ref 跑几何审计（默认在没给 outline/pages 时自动开）' },
      skeleton: { type: 'boolean', description: '打印可复用版面骨架' },
      slide: { type: 'integer', description: '只审某一页（0-based）' },
    },
    required: ['ref'],
  },
}

const apply: ToolDecl = {
  name: 'pptx_apply',
  label: 'PPT 命令批',
  kinds: KINDS,
  readOnly: false,
  requiresBaseVersion: true,
  maxItems: 200,
  tokens: 'docRev',
  description:
    '按顺序执行一批命令（上限 200，op 见 pptx_guide）。' +
    '多数 text / element / insert / table / slide / deck / chart 命令不需要页面；' +
    '图片/媒体/剪贴板/3D/旁白需要页面，否则整批回 command.page-unavailable。' +
    '默认 atomic：一条被拒整批回滚；isolation:"per_op" 可逐条保留。' +
    '被拒时按回执的 index + code 修那一条后重发。禁止用 write/edit 改 .pptx。',
  parameters: {
    type: 'object',
    properties: {
      ref: REF,
      commands: {
        type: 'array',
        description: '命令批：{ op, target?, ...字段 }；op 与字段用 pptx_guide 取',
        items: { type: 'object' },
      },
      version: {
        type: 'string',
        description: '无页面写必填；有页面时由页面会话持有，可省略',
      },
      docRev: {
        type: 'string',
        description: 'agent 读到的页面模型代数；页面存在时用于 stale-index 判定',
      },
      epoch: {
        type: 'number',
        description: '宿主透传的基线代数；pptx 不消费（保留兼容）',
      },
      isolation: {
        type: 'string',
        enum: ['atomic', 'per_op'],
        description: '默认 atomic；per_op 逐条保留成功项',
      },
      dryRun: { type: 'boolean', description: '只校验，不落盘' },
    },
    required: ['ref', 'commands'],
  },
}

export const tools: ToolDecl[] = [guide, read, check, create, apply]

export default tools
