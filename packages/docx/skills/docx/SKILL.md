---
name: docx
description: 编辑工作区里已打开的 .docx（Word 文档）。四个命令：docx.guide 查命令、docx.read 读文档、docx.create 创建空白文件、docx.apply 提交命令批。正文写入必须由用户先打开文档页面；禁止用 write/edit 改 .docx。
---

# Word 文档

## 四个命令

| 命令 | 用途 |
| --- | --- |
| `docx.guide` | 命令目录：三组 op 名 → 组内 schema → 单条 schema。**写之前先调它** |
| `docx.read` | 读：`context` / `blocks` / `comments` / `notes` / `revisions` / `styles` / `audit` |
| `docx.create` | 只创建空白 seed 文件；不写正文；目标必须不存在 |
| `docx.apply` | 在已打开的文档页执行命令批；`commands` 必须是 **JSON 字符串** |

通过 `command` 调用：

```json
{
  "domain": "docx",
  "action": "apply",
  "input": {
    "ref": "file:///w/报告.docx",
    "commands": "[{\"op\":\"setText\",\"index\":3,\"text\":\"新标题\"}]",
    "baseHash": "…",
    "docRev": 12
  }
}
```

## 新建文档流程

```text
docx.create({ ref: "file:///w/新文档.docx" })
  → 创建空白 .docx
  → 用 open 工具打开这个 ref（或让用户打开）
  → docx.apply 写入正文
  → 页面 save
```

`create` 只创建空白模板，不写正文；正文仍然必须页面打开后 `apply`。

## 写入前提

- **对应文档页面必须已经打开**。页面没打开时 `apply` 会失败；不要尝试后台改文件，也不要用 `write` / `edit`。
- 新文档先 `docx.create`，再打开页面，再 `apply`。
- 每次写前先 `docx.read` 取 `hash` / `docRev` 和块索引；`apply` 带上它们。
- `command.stale`：文件已被外部修改 → 重新读，不要重放旧批次。
- `command.stale-index`：页面 `docRev` 已变 → 重新读。
- `command.page-unavailable`：让用户先打开文档；不要换通道。

## 三组命令

| 组 | 数量 | 说明 |
| --- | --- | --- |
| `blocks` | 8 | 段落结构 / 文本 / 格式 |
| `editor` | 23 | 字体、段落、域、表格等编辑器 op |
| `actions` | 10 | HTML 片段、图片、图表、修订 |

所有组都要求页面打开。用 `docx.guide({ group })` / `({ op })` 取 schema。

## 块命令速查

```text
insertParagraph { after, text, heading }   // after=null → 文档开头
setText         { index, text }
setHeadingLevel { index, level }           // 0=正文，1..6=Heading N
setTextStyle    { index, bold?, italic?, underline?, size?, color?, align? }
deleteBlock     { index }
moveBlock       { index, to }
findReplace     { find, replace }
```

## 铁律

- 一批 ≤200 条；一条被拒 = 整批回滚，按回执里的 `index` + `code` 修那一条后整批重发。
- 结构命令会移动后续块，同一批里后面命令的 `index` 要按执行后的位置算；不确定就拆批重读。
- 表格是整块：`blocks` 组只能删/移；单元格编辑用 `editor` 的表格 op。
- 标题效果依赖文档已有 Heading 样式。
- 含域代码的段落避免 `setText`（会丢域），只改格式。
- **另存为 / 文件列表 / 标签页不属于本插件**；新建只提供空白 seed，正文仍要页面写。
- 目录里没有的 op 都不存在；不要臆造批注、脚注、页眉页脚、水印等未接线能力。
