---
name: office-docx
description: 编辑工作区里的 .docx（Word 文档）：段落文本与格式、标题、插入/删除/移动块、查找替换；命令批提交，字节保真。禁止用 write/edit 改 .docx。页面打开时命令进入页面会话。
---

## 三步循环

1. `read(<ref>)` → 块投影（`[0] h1 年度报告`），**`[n]` 就是命令的 index**
2. `docx_guide(<ref>)` → 分组与 op 名；`{ group }` / `{ op }` 取 schema。**不要凭记忆写 op**
3. `docx_apply(<ref>, { version, commands: [...] })`（先 `dryRun: true` 走一遍）

## 铁律

- **禁止**对 .docx 用 `write` —— 会写成假文件
- 一次一批（≤200 条）；批内一条被拒 = 整批不落盘：只改那一条，整批重发
- `index` 用 read/guide 得到的 `docxIndex`；改完重新 `read`，不要复用旧坐标
- **无页面**写必须带 `version`；有页面时由页面会话持有，工具里的 version 不参与 CAS
- `command.stale` = 文件被外部改过 → 重新 `read`，不要重放旧批次
- `command.stale-index` = 页面坐标已失效（epoch 变了）→ 重新 `read`
- `command.page-unavailable` = 页面会话不可用（关闭 / 崩溃 / 未就绪）→ 让用户重开页面，**不要**回退到别的写路径
- 只有 `docx_apply` 返回了 `applied` / `version` 才算真的改了
- 不要自己拼“撤销”：页面 Ctrl+Z 走模型历史，无页面路径没有 undo
- guide 的 `runtime` 里为 `false` 的能力不要承诺（如 `recalc:false`）

## 命令速查（细节走 guide）

```text
insertParagraph { after, text, heading }   // after=null → 文档开头
setText        { index, text }            // 保留段属性与首个 run 的字符属性
setHeadingLevel{ index, level }           // 0=正文，1..6=Heading N
setTextStyle   { index, bold?, italic?, underline?, size?, color?, align? }
deleteBlock    { index }
moveBlock      { index, to }              // to=null → 移到开头
findReplace    { find, replace, all? }    // 选择器型，可原地重试
check          { }                        // 只读诊断，回执带 audit
```

## docx 的坑

- 尾部 `sectPr`（分节 / 页边距 / 页眉页脚）不可见、不可删；它由计划自动保序
- 表格是整块：只能删除 / 移动；单元格编辑尚未开放（不要臆造 `setCellText`）
- 标题效果依赖文档已有 Heading 样式；没有定义时 Word 按正文显示，不要承诺“已套用企业模板”
- 页眉页脚 / 页码 / 目录字段不在块模型里，任何 op 都不会修改它们
- 含域代码的段落避免 `setText`（会丢失域），只改格式
- 图片 / 批注 / 修订尚未实现；回执里出现 `rejected` 时按 `code` 处理，不要换别的通道绕

## 新建

本版本不支持通过工具新建 `.docx`（写权威要求基线版本）。让用户先在工作区创建空文档，再 `read` 后编辑。
