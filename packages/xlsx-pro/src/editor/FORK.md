# 页面 fork 溯源（I17）

`src/editor/**` 是 GenOffice Sheets 渲染层的**原样 fork**。主线 I17 要求页面 bundle 与引擎
bundle 由同一份引擎依赖构建 —— 所以这里把来源钉死，`scripts/verify.mjs` 会校验它。

```yaml
# 机器可读：verify.mjs 解析这三行
vendorRepo: vendor/genoffice
vendorCommit: f6c2358c
sourcePath: apps/sheets/src/renderer
wasmEnginePath: apps/sheets/native/xlsx-engine
```

| 项 | 值 |
| --- | --- |
| fork 来源 | `vendor/genoffice` @ `f6c2358c`（upstream `main`） |
| 页面层路径 | `apps/sheets/src/renderer`（经 `packages/xlsx` fork） |
| 引擎（wasm）路径 | `apps/sheets/native/xlsx-engine`（同一 commit，编成 `wasm32-wasip1` reactor） |
| fork 日期 | 2026-11-17 |

## 升级 vendor 时必须同步做

1. 改上面的 `vendorCommit`；
2. 重新跑 `pnpm setup`（稀疏路径含 `apps/sheets/native/xlsx-engine`）；
3. `pnpm --filter @workspace/plugin-xlsx-pro build`（会重新编 wasm）；
4. 全插件 `typecheck + build + test`。上游历史上的一次真实破坏面：
   - `packages/i18n` 新增语言（`tr` / `hu`）→ fork 的三张 `strings-*.ts` 与 `locale.tsx` 要补；
   - `xlsx-gateway` 新增 `node:path` 的 `basename` / `extname` → 页面 shim 要补；
   - `PivotLayoutErrorCode` 新增 3 个码 → `workbook-ops.ts` 的文案映射要补。

## 已知的、故意保留的差异

| 差异 | 说明 |
| --- | --- |
| `tr` / `hu` 未翻译 | 直接映射到 `en` 字典；宿主 `main.tsx` 固定 `lang = 'zh'`，这两个语言不可达 |
| 3 个 pivot 错误码共用文案 | `baseFieldMissing` / `baseFieldNotAxis` → `appPivotFieldNotHeader`；`baseItemMissing` → `appPivotPageItemMissing`，避免为 3 条错误在 22 种语言里各加一行 |
| AI 面板 / AI Ribbon | 已移除（插件内不含 agent-core / provider） |
| WorkStudio 宿主适配 | 全部集中在 `workstudio-host.ts` 与 `src/page/*`，不在 `src/editor` 里加宿主逻辑 |
