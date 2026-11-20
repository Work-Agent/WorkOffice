---
name: pptx
description: 编辑与新建工作区里的 .pptx（PowerPoint 演示文稿）。五个工具：pptx_guide 查命令、pptx_read 读、pptx_check 检查、pptx_create 新建、pptx_apply 提交命令批。禁止用 write/edit 改 .pptx。
---

## 五个工具

| 工具 | 用途 |
| --- | --- |
| `pptx_guide(<ref>, { group?, op? })` | 命令目录：七组 op 名 → 组内 schema → 单条 schema。**改已有 deck 之前先调它** |
| `pptx_read(<ref>, { what })` | 读：`context` / `outline` / `slides` / `slide` / `notes` / `comments` / `charts` / `tables` / `theme` / `audit` / `find` |
| `pptx_check(<ref>, { styles? \| outline? \| pages? \| audit? \| skeleton? })` | 检查，不产字节：风格清单 / outline 校验 / pages 对账 / 几何审计 / 版面骨架 |
| `pptx_create(<ref>, { pages, style? \| slides \| content })` | **新建**：`ref` 必须尚不存在。`pages` = 设计稿（推荐） |
| `pptx_apply(<ref>, { commands, version?, docRev?, dryRun?, isolation? })` | 按顺序执行一批命令（≤200）。一条被拒 = 整批回滚 |

宿主另有 `copy(from, to)`：字节级复制资源，**沿用模板时先用它**再改文案。

## 新建：没有模板也能做，关键是"内容 → 风格 → 版式"

三种入口，按需求挑：

1. **工作区里有设计好的模板** → `copy` 一份再改（`findReplace` 换词 + 逐页 `setText`）。模板的母版/版式/图片/配色全保留。
2. **没有模板、要能交付** → **设计稿**（默认）：先按内容选风格预设，再逐页写"版式 + 内容"。**坐标由引擎算**，你只写内容和版式名。
3. **只要文字骨架** → `slides` / Markdown `content`。白底两段文字，只适合草稿。

**不要因为"没有模板"就退回文字稿。** 风格来自预设或 style.md，版式来自版式库，两者都不需要模板文件。

### 第 1 步：按内容选风格（一次，1 分钟）

`pptx_check({styles:true})` 列出全部预设。选择依据是**内容形态**，不是个人审美：

| 内容 | 预设 | 理由 |
| --- | --- | --- |
| 汇报 / 复盘 / 商业分析 | `consulting-navy` | 深蓝 + 单一琥珀强调色，数据可信感强 |
| 产品发布 / 技术方案 / 路演 | `tech-dark` | 暗场投影，亮青强调 |
| 品牌 / 市场 / 内容策划 | `editorial-warm` | 暖白纸感 + 衬线标题，图多不脏 |
| 学术 / 内部评审 / 制度说明 | `minimal-mono` | 无装饰，信息密度优先 |
| 培训 / 产品介绍 / 团队分享 | `fresh-teal` | 青绿 + 大圆角，亲和 |
| 发布会 / 招商 / 竞选 | `bold-red` | 主张优先，超大标题 |

要自己的配色就传 `{ preset: "consulting-navy", colors: { accent: "#E4572E", bg: "#101820" } }`，
或干脆用低层 `elements` 自己写颜色（见文末）。

### 第 2 步：写 outline（一次）

```json
{
  "core_hook": "毛利跌了 40%，货量却在涨",
  "pages": [
    { "title": "The 40% squeeze", "type": "cover", "layout": "cover_typography_hero",
      "brief": "主标题 + 时间段副标题 + 汇报人", "image_queries": [] },
    { "title": "Where the margin went", "type": "content", "layout": "three_column_cards",
      "brief": "三张卡：燃油 +31%、工资 +18%、即期运价 -22%，每张一条来源",
      "image_queries": ["container port cranes at dawn"] }
  ]
}
```

`pptx_check({outline})` 会在这些情况直接 **error**：`core_hook` 缺失、`type`/`layout` 不在版式库、
brief 出现 `XX%`/`TBD`/`lorem`、**连续两页同版式**。先改到 0 error 再写页面。

### 第 3 步：逐页写"版式 + 内容"（一页一个对象）

```json
{
  "layout": "three_column_cards",
  "type": "content",
  "title": "Where the margin went",
  "kicker": "成本结构",
  "items": [
    { "value": "+31%", "title": "燃油", "note": "2024 vs 2022" },
    { "value": "+18%", "title": "人力", "note": "2024 vs 2022" },
    { "value": "-22%", "title": "即期运价", "note": "2024 vs 2022" }
  ],
  "footnote": "来源：公司年报 2024、Drewry WCI"
}
```

引擎按版式算出所有坐标：卡片等宽、内边距、数字与说明的层级、强调线、来源行位置。

- `items` 的数量决定栏数（2 / 3 / 4 项都支持）；`items[].value` 是**大数字**，`items[].title` 是卡片标题，`items[].note` 是说明（可多行）
- 文字类页面用 `body`（字符串数组，每项一条项目符号）
- 不要为了"信息全"把 8 条挤进一页：一页一个结论，多的拆页
- 需要图就传 `image`（见下节），版式会用它

### 第 4 步：先 check 再 create

```text
pptx_check({ pages, outline, style: "consulting-navy" })   # 逐页对账，0 error 再继续
pptx_create({ ref, pages, style: "consulting-navy" })
pptx_check({ audit: true })                                # 几何审计
```

`pptx_check({pages, outline, style})` 检查的是**展开后真正会落盘的那一页**：echo 的 `title/type/layout` 与 outline 不符 → error；页面文本缺 outline 标题、outline 计划了图但页面没图 → warning；颜色不在预设调色板 → warning。

审计有 finding 就 `pptx_apply` 把 `→` 后面的 op 原样提交。**最多两轮修**，只修客观缺陷。

`pptx_check` 在有页面时读的是**已保存的基线**：页面里有未保存改动就先保存，或提交后再查。

## 版式库（17 种，`layout` 只能取这些）

| 版式 | 用在什么内容上 | 必需字段 |
| --- | --- | --- |
| `cover_typography_hero` | 纯文字封面，靠字号和留白立住 | `title`，建议 `subtitle`/`kicker` |
| `cover_dark_minimal` | 暗色极简封面，一个强调块 | `title` |
| `cover_split_color` | 左文右色块，色块里放关键数字 | `title`，`items[0].value` 可选 |
| `cover_magazine` | 杂志式大标题 + 细线 | `title` |
| `cover_split_image` | 左文右图（需要 `image`） | `title` + `image` |
| `cover_full_image_overlay` | 全屏图 + 暗蒙版（需要 `image`） | `title` + `image` |
| `three_column_cards` | 三个并列点 / 三条成本 / 三种方案 | `items` 2–4 项 |
| `kpi_cards_row` | 一排 KPI + 一句结论 | `items` 2–4 项，`body` 可选 |
| `hero_big_number` | 一个关键数字 + 支撑论据 | `items[0].value`，`body` |
| `two_column_comparison` | 对比（之前/之后、我方/对手） | `items` 2 项 |
| `two_by_two_grid` | 四个并列要点 | `items` 4 项 |
| `timeline_horizontal` | 3–5 步流程 / 时间线 | `items` 3–5 项 |
| `left_text_right_image` | 图文并排（需要 `image`） | `body` + `image` |
| `full_image_text_overlay` | 整图 + 底部文字带（需要 `image`） | `title` + `image` |
| `title_and_bullets` | 纯项目符号页 | `body` |
| `closing_thank_you` | 居中致谢 | `title`，建议 `subtitle` |
| `closing_cta` | 结尾行动号召 + 按钮 | `title`，`items[0].value` 是按钮文字 |

**不写 `layout`** 时会按内容自动挑（首页封面 / 末页结尾 / 3 项→三栏 / 2 项→对比 / 1 个 value→大数字 / 图+正文→图文 / 其余→项目符号）。自动挑能跑，但**显式写版式更稳**，也才好让 outline 校验版式变化。

## 风格 = 色 + 字 + 尺度 + 装饰规则

预设已经写好 6 套 token。想在预设上改一两个色：

```json
{ "preset": "editorial-warm", "colors": { "accent": "#C2410C", "bg": "#FBF7F0" } }
```

要完全自定义（自定义配色/字号），两条路：

1. 传 `style` 为 style.md 文本（只做调色板检查），版式仍按默认预设上色；
2. 或用**低层 `elements`** 自己写颜色与坐标（见文末），此时 `layout` 不生效。

## 图片

`image` 支持四种来源，按优先级用：

1. **本地文件**：`"image": "./assets/hero.png"`（相对 deck 所在目录）或 `"file:///D:/w/assets/hero.png"` 或绝对路径
2. **网络图片**：`"image": "https://…/hero.jpg"`（≤20 MB）
3. **内联**：`"image": "data:image/png;base64,…"`（小图标可以，大图别用）
4. 没有真图就**不要放图**：版式会自动退回不带图的变体（`cover_split_image` → `cover_split_color`），不会画灰块

图片按 **cover 居中裁剪**，永不拉伸；盒子比例越接近原图，裁掉的越少。裁掉超过一半会回一条 warning。

## 改已有 deck（七组命令，用 guide 取 schema）

| 组 | 无页面 | 例子 |
| --- | --- | --- |
| `text` | ✅ | `setText` `setFont` `setParagraphFormat` `insertEquation` |
| `element` | ✅ | `setTransform` `setFill` `setStroke` `groupElements` `alignElements` `setEffects` |
| `insert` | 部分 | `addElement` `addTable` `addChart` `addSmartArt` `addConnector`；图片/媒体需页面 |
| `table` | ✅ | `setTableCell` `tableMerge` `tableStructure` `setTableRowHeight` `setTableStyle` |
| `slide` | ✅ | `addBlankSlide` `deleteSlide` `moveSlide` `setBackground` `setNotes` `addAnimation` |
| `deck` | ✅ | `setSlideSize` `findReplace` `applyTheme` `addSection` `setSections` |
| `chart` | ✅ | `editChart`（高层字段，内部映射 setChart） |

**需页面**：`addPicture` / `replacePicture` / `setImageFill` / `addMedia` /
`addNarration` / `clearNarration` / `addModel3d` / `setModel3dView` /
`pasteElements` / `pasteSlide` / `insertSlidePptx` / `setAnimations`。

三步循环：`pptx_read({what:"context"})` 取 `s_*` → `pptx_guide({group|op})` 取 schema →
`pptx_apply({commands})`，改完 `pptx_check({audit:true})`。

## 复用上一次的风格

`pptx_check({ skeleton: true })` 打印这张 deck 的**版面骨架**（背景、标题框盒位与字号、
品牌图槽、装饰形状）。和预设一起留在工作区，下次生成时按同样的盒位写页面，
标题和 logo 就不会逐页漂移。

## 低层 `elements`（精确控制 / 预设外的设计）

画布 **1280×720 px** 整数，原点左上，`x/y/w/h` 必填，**数组顺序 = z-order**。
页数 ≤60，每页 ≤48 元素、≤8 张图。三种元素：

- `text {type,x,y,w,h,paragraphs,valign?}`
- `shape {type,shape,x,y,w,h,fill?,stroke?:{color,widthPt},paragraphs?,valign?,flipH?,flipV?}`
  —— `shape` ∈ rect roundRect ellipse triangle rightArrow leftArrow upArrow downArrow chevron diamond parallelogram trapezoid hexagon pentagon pie donut star5 heart cloud line lineArrow
- `image {type,x,y,w,h,url}`

段落 `{runs:[…], align?, lineSpacingPct?, spaceBeforePt?, spaceAfterPt?, bullet?}`；
run `{text, sizePt?, bold?, italic?, color?, font?}`。

文本框**零内边距、按真实字体度量自动长高**（不会缩）；越界元素被 clamp；
背景色写成整页矩形即可，引擎会提升成原生背景。全字段见同目录 `spec.md`
（设计流程全文见 `design.md`）。

## 设计底线（这些是缺陷，不是风格）

- 每页都是"色块 + 粗副标题 + 描述"的列表；封面只有标题 + 副标题
- 卡片左边一条细竖条 / 顶上一条色条；装饰角块；每页都挂一条顶栏
- 彩虹卡片（每项一个颜色）；强调色超过两个
- 整页内容挤在上半页，下半页空着；正文默认 12pt 铺满
- 用灰块假装照片；用无关的通用素材图凑数
- 一页里两段话说同一件事（副标题复述图表标题）
- 同一层级用不同字号

## 铁律

- **禁止**对 .pptx 用 `write` / `edit` —— 会写成假文件
- 一批 ≤200 条；一条被拒整批回滚，按回执 `index` + `code` 修那一条后**整批重发**
- `target.slide` 用 `s_*` 或 0-based 索引；`target.el` 用 `e_*`。能用 id 就别用数字
- 结构命令（增删页、插入/删除行）会移动后面的索引；跨批修改请重新 `read`
- **插入 slide 后不要在同一批里编辑新页**：整批先校验，新页还不存在。拆两批
- `command.stale` = 文件被外部改过 → 重新 `read`，不要重放旧批次
- `command.stale-index` = 页面模型已变（docRev 不符）→ 重新 `read`
- `command.page-unavailable` = 让用户打开页面，**不要**回退别的写路径
- 只有回执里有 `applied` / `version` 才算真的改了；**要内容用 `pptx_read`**
- 不要自己拼"撤销"；撤销是页面的模型历史 / 引擎快照
- 坐标是 EMU（1 in = 914400，1 pt = 12700，1 px@96dpi = 9525）；`pages` 用 px，op 用 EMU
- 默认 16:9 页 = 12192000 × 6858000 EMU
- 表格是整块：`setTableCell` 按 `row/col`；图表用 `editChart`，不要用低层 `setChart`
- **目录里没有的 op 都不存在**；不要臆造批注/脚注/页眉页脚这类 op
