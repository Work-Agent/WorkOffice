// ─── 风格预设：内容 → 风格，不依赖模板 ────────────────────────────────
//
// 没有模板不等于没有风格。这里给的是"设计系统"而不是版式文件：
// 配色槽位 + 字体对 + 字号阶梯 + 圆角/装饰规则 + 适合什么内容。
//
// 模型按内容挑一个预设（`pptx_check({styles:true})` 能列出全部），
// 版式函数（shared/layouts.ts）再按这套 token 把内容摆成页面。
// 想换皮只换预设名，版式和内容不动。

export interface StylePalette {
  /** 页面主背景（封面/结尾） */
  bg: string
  /** 内容页背景 */
  bgContent: string
  /** 主文字色（在 bg 上） */
  ink: string
  /** 主文字色（在 bgContent 上） */
  inkOnLight: string
  /** 次级文字（注释、来源） */
  muted: string
  /** 次级文字（浅底上） */
  mutedOnLight: string
  /** 强调色：全篇只用一个 */
  accent: string
  /** 强调色的浅底（卡片/数字底色） */
  accentSoft: string
  /** 卡片底 */
  card: string
  /** 卡片描边 / 分隔线 */
  border: string
  /** 深底上的文字（强调块上） */
  onAccent: string
}

export interface StyleScale {
  /** 封面主标题 */
  cover: number
  /** 页标题 */
  h1: number
  /** 卡片主数字 */
  number: number
  /** 副标题 / 卡片标题 */
  h2: number
  /** 正文 */
  body: number
  /** 注释 / 来源 */
  caption: number
}

export interface StylePreset {
  id: string
  label: string
  /** 一句话说明适合什么内容——这就是"按内容选风格"的依据 */
  when: string
  palette: StylePalette
  font: { title: string; body: string }
  scale: StyleScale
  /** 卡片圆角（px）；0 = 直角 */
  radius: number
  /** 装饰风格：line = 细线，block = 色块，none = 不用装饰 */
  decor: 'line' | 'block' | 'none'
  /** 这套风格最常用的版式（给模型的建议） */
  layouts: string[]
}

export const STYLE_PRESETS: readonly StylePreset[] = [
  {
    id: 'consulting-navy',
    label: '咨询深蓝',
    when: '汇报 / 复盘 / 商业分析：数据多、结论重、需要可信感',
    palette: {
      bg: '#0E1A2B',
      bgContent: '#FFFFFF',
      ink: '#FFFFFF',
      inkOnLight: '#0E1A2B',
      muted: '#9FB0C6',
      mutedOnLight: '#5A6473',
      accent: '#FFB000',
      accentSoft: '#FFF4DE',
      card: '#F4F6F8',
      border: '#D8DEE6',
      onAccent: '#0E1A2B',
    },
    font: { title: 'Helvetica Neue', body: 'PingFang SC' },
    scale: { cover: 64, h1: 36, number: 44, h2: 20, body: 18, caption: 12 },
    radius: 12,
    decor: 'line',
    layouts: ['cover_typography_hero', 'three_column_cards', 'kpi_cards_row', 'two_by_two_grid'],
  },
  {
    id: 'tech-dark',
    label: '科技暗色',
    when: '产品发布 / 技术方案 / 路演：要科技感、暗场投影',
    palette: {
      bg: '#0B0F14',
      bgContent: '#11161D',
      ink: '#F2F5F8',
      inkOnLight: '#F2F5F8',
      muted: '#8494A7',
      mutedOnLight: '#8494A7',
      accent: '#3DDC97',
      accentSoft: '#12241D',
      card: '#151B23',
      border: '#25303C',
      onAccent: '#0B0F14',
    },
    font: { title: 'Inter', body: 'PingFang SC' },
    scale: { cover: 68, h1: 36, number: 56, h2: 20, body: 17, caption: 12 },
    radius: 16,
    decor: 'block',
    layouts: ['cover_dark_minimal', 'hero_big_number', 'kpi_cards_row', 'timeline_horizontal'],
  },
  {
    id: 'editorial-warm',
    label: '杂志暖橙',
    when: '品牌 / 市场 / 内容策划：要质感和叙事，图多',
    palette: {
      bg: '#FBF7F0',
      bgContent: '#FBF7F0',
      ink: '#2B2118',
      inkOnLight: '#2B2118',
      muted: '#8A7B6B',
      mutedOnLight: '#8A7B6B',
      accent: '#E4572E',
      accentSoft: '#FBE6DE',
      card: '#FFFFFF',
      border: '#E7DCCB',
      onAccent: '#FFFFFF',
    },
    font: { title: 'Georgia', body: 'PingFang SC' },
    scale: { cover: 62, h1: 34, number: 48, h2: 19, body: 17, caption: 12 },
    radius: 4,
    decor: 'line',
    layouts: ['cover_magazine', 'cover_split_image', 'left_text_right_image', 'two_column_comparison'],
  },
  {
    id: 'minimal-mono',
    label: '极简黑白',
    when: '学术 / 内部评审 / 制度说明：信息优先，不要装饰',
    palette: {
      bg: '#FFFFFF',
      bgContent: '#FFFFFF',
      ink: '#111111',
      inkOnLight: '#111111',
      muted: '#7A7A7A',
      mutedOnLight: '#7A7A7A',
      accent: '#111111',
      accentSoft: '#F0F0F0',
      card: '#FAFAFA',
      border: '#DCDCDC',
      onAccent: '#FFFFFF',
    },
    font: { title: 'Helvetica Neue', body: 'PingFang SC' },
    scale: { cover: 58, h1: 32, number: 44, h2: 18, body: 16, caption: 11 },
    radius: 0,
    decor: 'none',
    layouts: ['cover_typography_hero', 'two_column_comparison', 'timeline_horizontal', 'hero_big_number'],
  },
  {
    id: 'fresh-teal',
    label: '清新青绿',
    when: '培训 / 产品介绍 / 团队分享：轻松、要亲和力',
    palette: {
      bg: '#0F3D3E',
      bgContent: '#FFFFFF',
      ink: '#FFFFFF',
      inkOnLight: '#12333A',
      muted: '#A8C7C8',
      mutedOnLight: '#5E7B80',
      accent: '#2EC4B6',
      accentSoft: '#E3F8F6',
      card: '#F2FAF9',
      border: '#CFE7E4',
      onAccent: '#0F3D3E',
    },
    font: { title: 'PingFang SC', body: 'PingFang SC' },
    scale: { cover: 60, h1: 34, number: 46, h2: 20, body: 18, caption: 12 },
    radius: 20,
    decor: 'block',
    layouts: ['cover_split_color', 'three_column_cards', 'hero_big_number', 'closing_cta'],
  },
  {
    id: 'bold-red',
    label: '冲击红',
    when: '发布会 / 招商 / 竞选：一句话主张，冲击力优先',
    palette: {
      bg: '#1A0F12',
      bgContent: '#FFFFFF',
      ink: '#FFFFFF',
      inkOnLight: '#1A0F12',
      muted: '#C9AEB3',
      mutedOnLight: '#6B5A5E',
      accent: '#E01A3C',
      accentSoft: '#FCE4E8',
      card: '#F7F2F3',
      border: '#E6D6D9',
      onAccent: '#FFFFFF',
    },
    font: { title: 'Helvetica Neue', body: 'PingFang SC' },
    scale: { cover: 76, h1: 40, number: 60, h2: 22, body: 18, caption: 12 },
    radius: 8,
    decor: 'block',
    layouts: ['cover_dark_minimal', 'hero_big_number', 'kpi_cards_row', 'closing_cta'],
  },
]

const BY_ID = new Map(STYLE_PRESETS.map((preset) => [preset.id, preset]))

/** 默认预设：商务汇报，最稳。 */
export const DEFAULT_PRESET = 'consulting-navy'

/** 预设名（大小写/空格宽容）→ 预设；未知返回 null。 */
export function findPreset(name: unknown): StylePreset | null {
  if (typeof name !== 'string') return null
  const key = name.trim().toLowerCase().replace(/\s+/g, '-')
  const direct = BY_ID.get(key)
  if (direct) return direct
  for (const preset of STYLE_PRESETS) {
    if (preset.label === name || preset.id.replace(/-/g, '') === key.replace(/-/g, '')) return preset
  }
  return null
}

/** 预设里所有会被画上去的颜色（off-palette 检查用）。 */
export function presetColors(preset: StylePreset): string[] {
  const p = preset.palette
  return [...new Set([p.bg, p.bgContent, p.ink, p.inkOnLight, p.muted, p.mutedOnLight, p.accent, p.accentSoft, p.card, p.border, p.onAccent])]
}

/** 给模型看的预设清单（一次列完，别让模型猜）。 */
export function presetsGuide(): string {
  const lines = [
    '风格预设（按内容选一个，`pptx_create({style})` 传 id；不传默认 ' + DEFAULT_PRESET + '）：',
  ]
  for (const preset of STYLE_PRESETS) {
    const p = preset.palette
    lines.push(
      `- ${preset.id}（${preset.label}）：${preset.when}\n` +
        `    背景 ${p.bg} / 内容页 ${p.bgContent} / 强调 ${p.accent} / 卡片 ${p.card} · ` +
        `标题 ${preset.scale.cover}/${preset.scale.h1}pt · 常用版式 ${preset.layouts.join(' ')}`,
    )
  }
  lines.push('要自定义就自己写 style.md（颜色用 #RRGGBB），再传 style 文本。')
  return lines.join('\n')
}
