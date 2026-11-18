---
name: pptx
description: 编辑工作区里的 .pptx（PowerPoint 演示文稿）。四个工具：pptx_guide 查命令、pptx_read 读演示文稿、pptx_create 新建、pptx_apply 提交命令批。禁止用 write/edit 改 .pptx。
---

## 四个工具

| 工具 | 用途 |
| --- | --- |
| `pptx_guide(<ref>, { group?, op? })` | 命令目录：七组 op 名 → 组内 schema → 单条 schema。**写之前先调它** |
| `pptx_read(<ref>, { what })` | 读：`context` / `outline` / `slides` / `slide` / `notes` / `comments` / `charts` / `tables` / `theme` / `audit` / `find` |
| `pptx_create(<ref>, { slides \| content })` | **新建**：`ref` 是尚不存在的目标路径；空内容 = 空白演示文稿 |
| `pptx_apply(<ref>, { commands, version?, docRev?, dryRun?, isolation? })` | 按顺序执行一批命令（≤200）。一条被拒 = 整批回滚 |

## 新建文件

```text
pptx_create("file:///w/2026年度汇报.pptx", {
  slides: [
    { title: "年度汇报", bullets: ["收入增长", "成本下降"], notes: "讲稿" },
    { title: "第二页", body: "正文" }
  ]
})
```

- `ref` 必须指向工作区里**不存在**的文件
- 新建不需要先 `pptx_guide`；直接调 `pptx_create`
- 也支持轻量 Markdown `content`：`# 页标题`、`---` 分页、`-` 项目符号、`notes:` 备注
- 不支持图片 / 表格 / 复杂版式；需要精确结构时用 `slides`
- 新建已经落盘，**不要**再用 write 或重复创建

## 七组命令（用 guide 取 schema）

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
无页面时整批回 `command.page-unavailable`。

## 三步循环

1. `pptx_read(<ref>, { what: "context" })` → 页列表 `[0] s_1 "年度汇报"`；**`s_*` 就是 slide id**
2. `pptx_guide(<ref>)` → 分组与 op 名；`{ group }` / `{ op }` 取 schema
3. `pptx_apply(<ref>, { commands: [{ op: "setText", target: { slide: "s_1", el: "e_12" }, paragraphs: [...] }] })`
   （不确定时先 `dryRun: true`）

## 铁律

- **禁止**对 .pptx 用 `write` / `edit` —— 会写成假文件
- 一批 ≤200 条；一条被拒整批回滚，按回执里的 `index` + `code` 修那一条后**整批重发**
- `target.slide` 用 `s_*` 或 0-based 索引；`target.el` 用 `e_*`。**能用 s_*/e_* 就不要用数字**
- 结构命令（增删页、插入/删除行）会移动后面的索引；跨批修改请重新 `read`
- **插入 slide 后不要在同一批里编辑新页**：`runTxn` 先整体校验，新页还不存在。拆成两批
- `command.stale` = 文件被外部改过 → 重新 `read`，不要重放旧批次
- `command.stale-index` = 页面模型已变（docRev 不符）→ 重新 `read`
- `command.page-unavailable` = 页面不可用；让用户打开页面，**不要**回退别的写路径
- 只有回执里有 `applied` / `version` 才算真的改了；`applied` 只告诉数量，**要内容用 `pptx_read`**
- 不要自己拼“撤销”；撤销是页面的模型历史 / 引擎快照

## pptx 的坑

- 坐标是 EMU：1 in = 914400，1 pt = 12700，1 px(96dpi) = 9525
- 默认 16:9 页 = 12192000 × 6858000 EMU
- 文本用 `paragraphs[].runs[].text`；`setText` 会替换该元素的 run 结构
- 表格是整块：`setTableCell` 按 `row/col`；`tableStructure` 增删行列
- 图表用 `editChart`，不要用低层 `setChart`
- 动画：`addAnimation` 用元素 id；`removeAnimation` / `reorderAnimation` 用时间线 `seq`
- 组合内的子元素要带 `group` id（用 `pptx_read slide` 查）
- 主题色 / 字体在 `pptx_read what:"theme"`；`applyTheme` 会改整个 deck
- **目录里没有的 op 都不存在**；不要臆造批注/脚注/页眉页脚这类 op
