---
name: docx
description: 编辑工作区里的 .docx（Word 文档）。四个工具：docx_guide 查命令、docx_read 读文档、docx_create 新建、docx_apply 提交命令批。禁止用 write/edit 改 .docx。
---

## 四个工具

| 工具 | 用途 |
| --- | --- |
| `docx_guide(<ref>, { group?, op? })` | 命令目录：三组 op 名 → 组内 schema → 单条 schema。**写之前先调它** |
| `docx_read(<ref>, { what })` | 读：`context`（块列表）/`blocks`（范围全文）/`comments`/`notes`/`revisions`/`styles`/`audit` |
| `docx_create(<ref>, { content })` | **新建**：`ref` 是尚不存在的目标路径，`content` 是 Markdown 全文（空 = 空白文档） |
| `docx_apply(<ref>, { commands, version?, epoch?, dryRun? })` | 按顺序执行一批命令（≤200）。一条被拒 = 整批回滚 |

## 新建文件

```text
docx_create("file:///w/2026年度工作报告.docx", { content: "# 标题\n\n正文……" })
```

- `ref` 必须指向工作区里**不存在**的文件；文件名由 ref 决定
- 新建**不需要**先调 `docx_guide`（它需要已有文件）；直接调 `docx_create`
- `content` 是 Markdown：标题、段落、列表、表格、粗体/斜体/删除线、代码块、引用
- 目标已存在 → 拒绝，改用 `docx_apply`
- 图片和公式暂不支持；链接按纯文本写入
- 新建已经落盘，**不要**再用 write 或重复创建

## 三组命令（用 guide 取 schema）

| 组 | 数量 | 需要页面 | 例子 |
| --- | --- | --- | --- |
| `blocks` | 8 | 否 | `setText` `insertParagraph` `setHeadingLevel` `setTextStyle` `deleteBlock` `moveBlock` `findReplace` |
| `editor` | 23 | **是** | `setFont` `setParagraphFormat` `setMatchedFont` `deleteBlocks` `moveBlocks` `insertToc` `insertField` 表格系列 `applyStyle` |
| `actions` | 10 | **是** | `insert_content` `replace_blocks` `replace_selection` `insert_text_box` `insert_picture` `insert_image` `insert_chart` `edit_chart` `accept_changes` `reject_changes` |

页面没打开时只能发 `blocks` 组；混进别的组 → `command.page-unavailable`
（不要改用 write，要么让用户打开文档，要么只用 blocks 能表达的改动）。

## 三步循环

1. `docx_read(<ref>, { what: "context" })` → 块列表 `[0] h1 年度报告`；**`[n]` 就是 index**
2. `docx_guide(<ref>)` → 分组与 op 名；`{ group }` / `{ op }` 取 schema
3. `docx_apply(<ref>, { commands: [{ op: "setText", index: 3, text: "新标题" }] })`
   （不确定时先 `dryRun: true`）

## 铁律

- **禁止**对 .docx 用 `write` / `edit` —— 会写成假文件
- 一批 ≤200 条；一条被拒整批回滚，按回执里的 `index` + `code` 修那一条后**整批重发**
- `index` 用刚读到的块索引。结构命令（`insertParagraph`/`deleteBlock`/`moveBlock`）会移动后面的块，
  所以同一批里后面命令的 index 要按执行后的位置算；不确定就拆成两批、中间重读
- `command.stale` = 文件被外部改过 → 重新 `read`，不要重放旧批次
- `command.stale-index` = 块坐标失效（页面 epoch 变了）→ 重新 `read`
- `command.page-unavailable` = 页面不可用 → 让用户打开文档，**不要**回退别的写路径
- 只有回执里有 `applied` / `version` 才算真的改了；`applied` 只告诉数量，**要内容用 `docx_read`**
- 不要自己拼“撤销”：Ctrl+Z 是页面模型历史；无页面路径没有 undo

## 块命令速查

```text
insertParagraph { after, text, heading }   // after=null → 文档开头
setText         { index, text }            // 保留段落属性；改文本会合并该段 run
setHeadingLevel { index, level }           // 0=正文，1..6=Heading N
setTextStyle    { index, bold?, italic?, underline?, size?, color?, align? }
deleteBlock     { index }
moveBlock       { index, to }              // to=null → 移到开头
findReplace     { find, replace }
```

整段/整篇重写用 `actions` 的 `replace_blocks` / `insert_content`（受限 HTML，格式保留得更好）；
批量格式用 `editor` 的 `setFont` / `setParagraphFormat`（一批一个事务、一次撤销）。

## 坑

- 尾部 `sectPr`（分节 / 页边距 / 页眉页脚）不可见、不可删；计划会自动保序
- 表格是整块：`blocks` 组只能删/移；单元格编辑用 `editor` 的表格 op（需页面）
- 标题效果依赖文档已有 Heading 样式，没有定义时 Word 按正文显示
- 含域代码的段落避免 `setText`（会丢域），只改格式
- `docx_read` 的 `context` 在无页面时没有「当前选区」，按块索引定位
- **目录里没有的 op 都不存在**。批注 / 脚注 / 页眉页脚 / 页面设置 / 样式定义 / 水印暂未接线——不要臆造这些 op
- 新建独立文件用 `docx_create`；`write_document`（在已打开文档里流式长写）仍未接线，但不需要它来新建
