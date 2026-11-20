---
name: xlsx
description: 编辑工作区里的 .xlsx（Excel 工作簿）。四个工具：xlsx_guide 查命令、xlsx_read 读工作簿、xlsx_create 新建、xlsx_apply 提交命令批。禁止用 write/edit 改 .xlsx。
---

## 四个工具

| 工具 | 用途 |
| --- | --- |
| `xlsx_guide(<ref>, { group?, op? })` | 命令目录：四组 op 名 → 组内 schema → 单条 schema。**写之前先调它** |
| `xlsx_read(<ref>, { what })` | 读：`context` / `sheets` / `range` / `formulas` / `find` / `audit` |
| `xlsx_create(<ref>, { sheets })` | **新建**：`ref` 必须不存在；`sheets: [{ name, rows }]`，也支持 `rows` / `csv` |
| `xlsx_apply(<ref>, { commands, version?, structureEpoch?, docRev?, dryRun? })` | 按顺序执行一批命令（≤200）。一条被拒 = 整批回滚 |

**读结果首行是令牌**：无页面只有 `version=…`；有页面还有 `epoch=… docRev=… structureEpoch=…`。
写批里含 A1 坐标 / 结构类命令要 `structureEpoch`，内容型命令（fill/copy/sort/find）要 `docRev`；
有页面时宿主会自动带上，你显式传的话以你的为准。

## 新建文件

```text
xlsx_create("file:///w/2026预算.xlsx", {
  sheets: [
    { name: "预算", rows: [["项目", "金额"], ["差旅", 1200], ["合计", "=SUM(B2:B3)"]] },
    { name: "说明", rows: [["口径", "含税"]] }
  ]
})
```

- `ref` 必须指向工作区里**不存在**的文件
- 字符串以 `=` 开头会写成公式
- 空内容 = 空白工作簿
- 目标已存在 → 拒绝，改用 `xlsx_apply`

## 四组命令（用 guide 取 schema）

| 组 | 数量 | 需要页面 | 例子 |
| --- | --- | --- | --- |
| `cells` | 12 | 否（`convert_to_values` 除外） | `set_range` `set_cell` `set_formula` `format_range` `find_replace` `fill_range` `copy_range` `move_range` `sort_range` |
| `structure` | 10 | 否 | `insert_rows` `delete_rows` `insert_cols` `delete_cols` `merge_cells` `set_row_height` `set_col_width` `set_rows_hidden` |
| `sheets` | 6 | 否 | `add_sheet` `rename_sheet` `delete_sheet` `duplicate_sheet` `move_sheet` `set_sheet_hidden` |
| `objects` | 28 | **是** | 图表、形状、图片、表格、透视表、迷你图、条件格式、数据验证、批注、筛选、冻结、保护、页面设置、定义名称 |

页面没打开时只能发前 3 组；混进 `objects` → `command.page-unavailable`
（不要改用 write，要么让用户打开文件，要么只用前 3 组能表达的改动）。

## 三步循环

1. `xlsx_read(<ref>, { what: "context" })` → 令牌 + 工作表名与范围；**工作表用名称，不要用会话 id**
2. `xlsx_guide(<ref>)` → 分组与 op 名；`{ group }` / `{ op }` 取 schema
3. `xlsx_apply(<ref>, { commands: [{ op: "set_range", sheet: "Sheet1", start: "A1", values: [[...]] }] })`
   （不确定时先 `dryRun: true`）

读不到内容（只有令牌行、或 `notices=` 说页面未响应）就先停下确认文件是否开着；
不要靠换 `what` 反复重试。

## 沿用现有格式（“按当前格式补充”类需求）

`xlsx_read` **不返回**样式；`query` 也只有值。要沿用上一条记录的格式：

1. 先 `xlsx_read({ what: "context" })` 拿到令牌与 `merges=`、行范围
2. `copy_range` 把上一条记录的**整块**（含标题行、内容合并区、填充格）复制到新位置：
   `{ op: "copy_range", sheet: "Sheet1", source: "A25:C32", target: "A33" }`
   它会一起复制：值/公式、单元格样式（字体/填充/对齐/换行/数字格式）、
   完全落在源区域内的合并区、以及源行的行高。
3. 再用 `set_range` 只覆盖值（写值不动样式）；需要时补 `format_range`。

源区域要覆盖整块（比如 `A25:C32` 而不是 `A25:A26`），否则合并区/填充格不会被带过去；
边框与主题颜色不做推算，需要时用 `format_range` 显式补。
不要为了看格式去读 PDF / 截图 / 反复 `audit`。

## 常用写法

```text
set_range      { sheet, start, values }       // values 是矩形二维数组；null 清空
set_cell       { sheet, address, value }
set_formula    { sheet, address, formula }    // 必须带 =
format_range   { sheet, range, format }       // bold / fillColor / numberFormat / horizontalAlign ...
find_replace   { sheet, range, find, replace }
insert_rows    { sheet, row, count }          // row 是 1-based，插在它之前
delete_rows    { sheet, row, count }
rename_sheet   { sheet, name }
add_sheet      { name }
```

## 铁律

- **禁止**对 .xlsx 用 `write` / `edit` —— 会写成假文件
- 一批 ≤200 条；一条被拒整批回滚，按回执里的 `index` + `code` 修那一条后**整批重发**
- 工作表用**名称**（`sheet`），不要用 `sheet-1` 这类会话 id
- 结构命令（插行/删行/插列/删列）会移动后面的单元格；同一批里后面的命令按执行后的坐标算，
  不确定就拆成两批、中间重读
- `command.stale` = 文件被外部改过 → 重新 `read`，不要重放旧批次
- `command.stale-index` = 页面 baseline / revision 变了 → 重新 `read`
- `command.page-unavailable` = 页面不可用 → 让用户打开文件，**不要**回退别的写路径
- 只有回执里有 `applied` / `version` 才算真的改了；`applied` 只告诉数量，**要内容用 `xlsx_read`**
- 无页面路径**不重算公式**：公式文本会写入，但缓存值可能保持旧值；需要重算请打开页面
- 不要自己拼“撤销”：Ctrl+Z 是页面模型历史；无页面路径没有 undo

## 坑

- `xlsx_read` 读的是落盘内容；页面有未保存修改时，以页面上下文为准
- 大范围操作会展开成很多单元格；超过 2,000 个写入会被拒绝，拆成多次
- `format_range` 无页面只支持常用样式（粗体、颜色、字号、数字格式、对齐、边框）；
  复杂主题色/渐变填充请打开页面
- 图表、形状、图片、表格、透视表、条件格式、数据验证、批注、筛选、冻结等 `objects` 组
  需要页面；目录里没有的 op 都不存在
- 新建独立文件用 `xlsx_create`；不要用 `write` 生成 xlsx
