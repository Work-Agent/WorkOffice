# xlsx（P3：流式打开 + 保存 + agent 写）

单个 `.xlsx` 工作簿的 WorkStudio 新框架插件（`@work-studio/sdk`）。
P3 做到：**契约、单文件守卫、编辑器、wasm 流式打开、保存 / 另存 / 分块保存、自动保存、
崩溃恢复副本、agent 读与写（`apply` / `create`）**。只剩清理与发布（P4）。

## 现在有什么

| 层 | 内容 |
| --- | --- |
| 清单 | `claims: xlsx/xlsm/xltx`（priority 100）、`pages.content`、`skills: ["skills/xlsx"]` |
| resources.xlsx | `ls/find/grep` 恒空；`read` 返回单条 `{locator:'xlsx:file', …}` |
| `xlsx.guide` | 56 条 op 目录（四组：cells 12 / structure 10 / sheets 6 / objects 28） |
| `xlsx.read` | `context / sheets / range / formulas / find / audit`，只读磁盘 |
| `xlsx.apply` | 一批命令（≤200）：有页面进 Univer 事务（op 级令牌），无页面网关展开后写盘；`objects` 需页面 |
| `xlsx.create` | 单文件新建：`sheets/rows/csv`，目标必须不存在、父目录已存在 |
| `xlsx.page` | 页面通道：只读 op + `save*` / `saveAs` + `attach` 令牌 + `pending/ack` + `draft*` + `report/detach` |
| 页面 | `?ref=` 启动 → `@sheets/main`（GenOffice Sheets 渲染层）+ `__xlsxPlus` 桥；主题跟随 `shell/appearance`；15 s 租约心跳 |
| 打开 | wasm `open` 元数据 + `read_range`：只读可视区，不读整份字节；wasm 缺失时回退整份字节（≤64 MB） |
| 保存 | 页面交保存请求或整份字节 → 引擎先 `recalc_cells` 再 `planCellEditsToXlsx` 产字节 → `ctx.fs.write`（锁 + `baseHash`）；>8 MB 走 `saveBegin/saveChunk/saveCommit` |
| 另存 | `shell/native.saveFile` 签 `save:` token → 引擎写新文件，通道随之指向新 ref |
| 自动保存 | `getAutoSaveDefault → { on: true }`，编辑器自带的空闲保存直接落到 `save` op |
| agent 写 | `apply`：有租约 → 事件 + `pending/ack`，页面执行事务并保存；无租约 → 引擎展开写盘 |
| 崩溃恢复 | 编辑器写恢复副本到插件数据目录；下次打开页面显示「恢复 / 丢弃」条幅 |
| 依赖 | GenOffice 走 `vendor/genoffice` submodule + 构建期物化（`build/sheets/src`）；`src/editor` 只有 125 个改过的文件 |

宿主零改动：`ctx.origin`、脏态、语言通道都不依赖，页面身份用 `attach` 发放的租约令牌区分。

## 构建 / 测试

```bash
pnpm --filter @workstudio/plugin-xlsx build
pnpm --filter @workstudio/plugin-xlsx typecheck
pnpm --filter @workstudio/plugin-xlsx test     # 需要先 build
pnpm --filter @workstudio/plugin-xlsx verify
pnpm --filter @workstudio/plugin-xlsx build:wasm   # 首次 4 min；缺失时自动触发
```

产物：

```text
dist/engine.js                 definePlugin（Node ESM，SDK 与网关已 bundle）
dist/xlsx-engine.wasm          wasm 引擎（open / read_range / recalc_cells）
dist/xlsx-engine.json          产物自证（体积 / sha256 / 导出面）
dist/page/**                   ?ref= 页面
dist/seed.xlsx                 空白工作簿模板
commands.json / api.json / api.d.ts
```

## 安装到本机 WorkStudio

```bash
node scripts/install.mjs --only xlsx
```

或在 WorkStudio 设置 → 插件添加本地目录。重启后打开工作区里的 `.xlsx`。

## 下一步

| 阶段 | 内容 |
| --- | --- |
| P1 | fork 收敛：`src/editor` 只放改过的页面文件，其余回落 vendor；接回 Univer 编辑器 |
| P2 | 写通路：`page {op:'save'/'saveAs'/'saveBegin'…}`、自动保存、恢复副本 |
| P3 | agent 面：`apply` / `create` + 页面租约执行 + 令牌分级 |
| P4 | 注册表发布 |

方案与取舍见仓库 `docs/xlsx-migration.md`。
