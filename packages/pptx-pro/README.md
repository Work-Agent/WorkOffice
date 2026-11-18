# pptx-pro（PowerPoint · `.pptx`）

`pptx-pro` 是 `packages/pptx` 完整 Slides 编辑器的**计划型写版本**：

- **有页面**：完整 GenOffice Slides renderer，保存时提交 `TxnRequest`；
- **无页面**：`pptx_apply` 走执行点 `openPptx → runTxn → savePptxToFile`；
- **常驻模型**：`OpenedPptx` 按 `(ref, version)` 缓存，两阶段提交
  `artifact + modelToken` / `commitModel` / `abortModel`；
- **工具面**：4 个工具（`pptx_guide / pptx_read / pptx_create / pptx_apply`）
  + 7 个命令组（上游 62 条 AI 可调用 op + `editChart`）。

> 当前状态：**P0b（headless 引擎 + pulled 页面 + commit(TxnRequest)）** 已落地。
> 保存只写变更的 slide/part，保存后 bake 不重开；manifest 增量同步与 200 MB 流式仍在后续阶段。

## 当前能力

- `pptx_guide`：三级命令目录
- `pptx_read`：context / slide / notes / comments / charts / tables / theme / audit / find
- `pptx_apply`：无页面可执行 text / element / table / slide / deck / chart；
  图片 / 媒体 / 剪贴板 / 3D / 旁白 / `setAnimations` 回 `command.page-unavailable`
- `pptx_create`：结构化 slides + 轻量 Markdown；seed + setup/content 两阶段事务
- `exec`：`savePptxToFile` → `artifact` + `modelToken`
- `commitModel` / `abortModel`：两阶段模型提交；`abort` 恢复快照与 undo 栈
- 页面：pulled bytes 打开、`pageSave`（changed slide/part）、agent 与人类共用模型/undo、页面快照撤销

## 限制

- P0 上限 64 MB（宿主 `ENGINE_BUDGET` + `artifact` 整份读入）；
- artifact 目前不是流式落盘；200 MB 需要宿主改造；
- 页面提交后仍是整包重新拉取（manifest 增量同步尚未实现）；
- sidecar live、页面只读渲染侧属于后续阶段。

## 构建 / 测试

```bash
pnpm --filter @workstudio/plugin-pptx-pro build
pnpm --filter @workstudio/plugin-pptx-pro typecheck
pnpm --filter @workstudio/plugin-pptx-pro test
pnpm --filter @workstudio/plugin-pptx-pro verify
```

产物：

```text
dist/
├─ ura.js
├─ seed.pptx
└─ engine/{tools,plan,exec,read,runtime}.mjs
```
