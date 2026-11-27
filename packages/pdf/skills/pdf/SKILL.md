---
name: pdf
description: 读取工作区里的 PDF（只读）。read 取元数据，grep 查页文本，pdf.read 读大纲 / 单页文本 / 全文查找，open 用查看器定位到页。禁止用 write/edit 改 .pdf。
---

# PDF 文档（只读）

PDF 插件只读：没有编辑命令，也不要尝试 `write` / `edit` 改 `.pdf`。
查看器提供渲染、翻页、缩放、搜索、大纲、缩略图、链接、打印和另存副本。

## 三个读取入口

| 入口 | 用途 |
| --- | --- |
| `read <ref>` | 元数据：name / size / pageCount / title / author / hash |
| `grep <ref> <query>` | 页文本检索，返回 `pdf:page:N` 与片段 |
| `command pdf/read {ref, what, page?, query?, limit?}` | `context` / `outline` / `pages` / `text` / `find` |

`pdf.read` 的 `what`：

| what | 参数 | 返回 |
| --- | --- | --- |
| `context` | — | 元数据 + hash + size |
| `outline` | — | 书签树（标题 → 页码） |
| `pages` | `limit?` | 页尺寸 / 旋转 |
| `text` | `page`（1-based） | 该页文本 |
| `find` | `query` | 逐页命中与片段 |

```json
{ "domain": "pdf", "action": "read", "input": { "ref": "file:///w/报告.pdf", "what": "text", "page": 3 } }
```

## 打开与定位

- `open(ref)` 打开查看器；
- `open(ref, "pdf:page:12")` 打开并跳到第 12 页；
- 资源面板点击 PDF 也走同一入口。

## 边界

- 只支持 `file://` 的单个 `.pdf`，上限 128 MB；
- 加密 PDF：查看器可输入密码打开；agent 文本读取会回 `command.page-unavailable`；
- 扫描件没有文本层：`grep` 为空，`pdf.read context` 仍可用；
- 大文件元数据可能缺页数 / 标题：以查看器加载后的页数为准。
