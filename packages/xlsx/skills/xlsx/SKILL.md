---
name: xlsx
description: 读写工作区里的单个 .xlsx（Excel 工作簿）。四个命令：xlsx.guide 查命令目录、xlsx.read 读、xlsx.create 新建、xlsx.apply 提交命令批。禁止用 write/edit 改 .xlsx。
---

## 范围

- 只处理**一个 `.xlsx` 文件**（`.xlsx` / `.xlsm` / `.xltx`）；
- 不支持目录、跨文件合并、批量转换；`ref` 必须是工作区里的 `file://` 地址。

## 四个命令

| 命令 | 用途 |
| --- | --- |
| `xlsx.guide({ ref })` | 命令目录：四组 op 名与数量；`group` / `op` 取 schema 与示例 |
| `xlsx.read({ ref, what })` | `context` / `sheets` / `range` / `formulas` / `find` / `audit` |
| `xlsx.create({ ref, sheets })` | **新建**：ref 必须不存在；`[{ name, rows }]`，也支持 `rows` / `csv` |
| `xlsx.apply({ ref, commands })` | 按顺序执行一批命令（≤200）；一条被拒 = 整批不动 |

**读结果首行是令牌**：`hash=… size=…`（有页面时还有 `docRev` / `structureEpoch`）。
写批里含 A1 坐标 / 结构类命令要 `structureEpoch`，内容型命令（fill/copy/sort/find）要 `docRev`：
**从最近一次 `read` 的首行取**（有页面时才有这两个值），写成功后的回执里会带新的令牌。

## 三步循环

1. `xlsx.read({ what: "context" })` → 令牌 + 工作表名（**用名称，不要用 id**）
2. `xlsx.guide()` → 分组与 op 名；`{ group }` / `{ op }` 取 schema
3. `xlsx.apply({ ref, commands: [{ op: "set_range", sheet: "Sheet1", start: "A1", values: [[...]] }] })`
   - 不确定时先 `dryRun: true`；
   - 改完结构后再 `read` 一次，别拿旧坐标继续写。

## 四组命令

| 组 | 数量 | 无页面可执行 | 例子 |
| --- | --- | --- | --- |
| `cells` | 12 | 是（`convert_to_values` 除外） | `set_cell` `set_range` `set_formula` `format_range` `find_replace` `fill_range` `copy_range` `sort_range` |
| `structure` | 10 | 是 | `insert_rows` `delete_rows` `insert_cols` `delete_cols` `merge_cells` `set_row_height` `set_col_width` |
| `sheets` | 6 | 是 | `add_sheet` `rename_sheet` `delete_sheet` `duplicate_sheet` `move_sheet` `set_sheet_hidden` |
| `objects` | 28 | **否（需页面）** | 图表、形状、图片、表格、透视表、迷你图、条件格式、数据验证、批注、筛选、冻结、保护、页面设置、定义名称 |

页面没打开时混进 `objects` → 整批回 `command.page-unavailable`。不要换 `write` 绕，
要么让用户打开文件，要么只用前 3 组能表达的改动。

## 新建文件

```text
xlsx.create("file:///w/2026预算.xlsx", {
  sheets: [
    { name: "预算", rows: [["项目", "金额"], ["差旅", 1200], ["合计", "=SUM(B2:B3)"]] },
    { name: "说明", rows: [["口径", "含税"]] }
  ]
})
```

- 字符串以 `=` 开头写成公式；空内容 = 空白工作簿；
- 目标已存在 → 拒绝，改用 `xlsx.apply`；
- 父目录必须已存在（本插件不做目录操作）。
