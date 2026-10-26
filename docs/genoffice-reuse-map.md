# GenOffice 组件复用分析

> 目标：**尽可能原样复用 GenOffice 的 renderer / components / engines，只保留 WorkStudio 的壳层差异，并移除 AI。**
>
> 结论先行：不应该继续手写一个“精简版 Office”。应把 GenOffice 对应 App 的 renderer fork 进来，用 host adapter 替换 Electron IPC，再删掉 AI。这样 Ribbon、对话框、快捷键、状态栏、保存逻辑才能和 GenOffice 保持一致。

## 0. 结论

| 格式 | 现状 | 建议 |
| --- | --- | --- |
| DOCX | 已经 fork GenOffice Docs 编辑器 | **保持现状**，只做上游同步与 WorkStudio 集成 |
| XLSX | 刚接入 Univer，只有极简外壳 | **fork GenOffice Sheets renderer**，保留 `ExcelShell` / Ribbon / 对话框 / 保存链路，Univer 继续作为网格引擎 |
| PPTX | 目前是自绘 SVG + 模态文本编辑 | **fork GenOffice Slides renderer**，保留 Konva / SlideCanvas / TextEditOverlay / Ribbon / 幻灯片操作 |

当前 XLSX「上方工具栏没有了」不是样式问题，而是因为没有搬 GenOffice Sheets 的 `ExcelShell` Ribbon；`UniverSheetsCorePreset({ toolbar: false })` 本来就只保留名称框和公式栏。

---

## 1. 总体分层

### 1.1 GenOffice 原始结构

```
apps/{docs,sheets,slides}/src/renderer
├─ App.tsx                   应用状态、命令、快捷键、生命周期
├─ Ribbon / ExcelShell       顶部功能区、格式面板、状态栏
├─ components/               对话框、右键菜单、导航、打印等
├─ editor / SlideCanvas      ProseMirror / Univer / Konva 编辑面
├─ actions / journal         编辑操作、撤销、保存队列
├─ ai/                       AI 面板、工具、技能（要删除）
└─ window.desktopApi
   window.slidesApi          渲染层 → 主进程 IPC

apps/{docs,sheets,slides}/src/main
├─ 文件打开 / 保存 / 另存为
├─ print / PDF / CSV / 导出
├─ sidecar / 原生引擎 / 字体
├─ 窗口 / 标签 / 最近文件
└─ AI IPC（要删除）

packages/*
├─ docx-engine / pptx-engine / pptx-ops / pptx-render
├─ xlsx-gateway（OOXML 写回、结构、格式）
├─ ui（Ribbon 基础组件、颜色、下拉、查找、形状库）
├─ i18n / zip-gate / font-metrics
└─ agent-core / ai-provider / ai-search（AI，删除）
```

### 1.2 WorkStudio 插件目标结构

```
packages/{xlsx,pptx}/src/editor       ← 原样 fork GenOffice renderer
├─ App / Ribbon / components
├─ 编辑器画布
├─ actions（保留）
└─ ai/（删除）

packages/{xlsx,pptx}/src/page         ← WorkStudio 宿主层
├─ page session 启动
├─ host adapter（模拟 window.desktopApi / slidesApi）
├─ 主题 / 脏态 / 保存
└─ saveBytes / 字节写回

WorkStudio URA
├─ codec / presenter / sidecar
├─ page session
└─ 文件字节
```

关键原则：

> **渲染层尽量不改；把差异集中在 host adapter。**
>
> 不把 GenOffice 的 main/preload 搬进来，但保留它的 renderer、组件、快捷键、Ribbon、对话框和引擎调用语义。

---

## 2. 包级复用矩阵

| 包 | 作用 | 是否 Electron-free | 复用方式 | WorkStudio 用法 |
| --- | --- | --- | --- | --- |
| `@genoffice/docx-engine` | DOCX 解析 / OOXML / patch | 是 | **A 直接复用** | 已在 DOCX fork 中使用 |
| `@genoffice/pptx-engine` | PPTX 解析 / 元素树 / patch | 是 | **A 直接复用** | PPTX 页面/主进程引擎 |
| `@genoffice/pptx-ops` | PPTX 操作层、事务、journal | 是 | **A 直接复用** | PPTX 编辑动作 |
| `@genoffice/pptx-render` | RenderTree / 文本布局 / Konva 适配 | 是 | **A 直接复用** | PPTX 画布 |
| `@genoffice/xlsx-gateway` | XLSX OOXML patch / 结构 / 格式 / 公式 | 是（仅少数 Node API） | **A/B 复用** | 在页面内加浏览器 shim，或走宿主 RPC |
| `@genoffice/zip-gate` | ZIP 安全解压 | 是 | **A 直接复用** | 引擎依赖 |
| `@genoffice/i18n` | 语言 / 方向 / 快捷键格式化 | 是 | **A 直接复用** | 与 WorkStudio 语言同步 |
| `@genoffice/ui` | Ribbon 基础件、颜色、下拉、查找、形状库 | 是（但依赖 AI 包） | **B 复用** | 抽 `ui-lite`，只保留非 AI 组件 |
| `@genoffice/font-metrics` | 系统字体度量 | 主进程 | **B/C** | DOCX/PPTX 现有字体逻辑按需 |
| `@genoffice/electron-utils` | 主进程辅助 | 部分 | **B 复用** | 只取 headless/print 纯函数，不取 Electron 部分 |
| `@genoffice/project-store` | 项目 / 历史存储 | 是 | **C 不优先** | WorkStudio 自己管会话与文件 |
| `@genoffice/file-parse` | 附件解析 | 主进程 | **C 可选** | 仅附件/AI 需要，暂不需要 |
| `@genoffice/html2docx` | HTML → DOCX | 是 | **C 可选** | 后续导入能力 |
| `@genoffice/odf2ooxml` | ODF → DOCX/PPTX | 是 | **C 可选** | 后续导入能力 |
| `@genoffice/pdf2docx` | PDF → DOCX | 是 | **C 可选** | 后续转换能力 |
| `@genoffice/pipelines` | AI 生成流水线 | 是 | **D 删除** | AI |
| `@genoffice/cli` | 命令行 | 是 | **D 不复用** | 不在页面内 |
| `@genoffice/agent-core` | Agent 循环 | 是 | **D 删除** | AI |
| `@genoffice/ai-provider` | LLM 客户端 | 是 | **D 删除** | AI |
| `@genoffice/ai-search` | 网络搜索 | 主进程 | **D 删除** | AI |

> `@genoffice/ui` 的 `index.ts` 同时导出 AI 组件和通用组件，且依赖 `agent-core` / `ai-provider`。复用前应做 `ui-lite`（或扩展 subpath exports），否则 AI 依赖会重新进入 bundle。

---

## 3. App 级组件地图

### 3.1 DOCX

当前已经是 GenOffice Docs fork，不需要重新移植。

| 组件 | 数量 | 复用 |
| --- | --- | --- |
| `editor/` | 90 | 保留 |
| `components/` | 55 | 保留 |
| `i18n/` | 71 | 保留（按语言裁剪） |
| `ai/` | 32 | 已删除 |
| 字体/资源 | 35 | 保留 |
| `App.tsx` / `Ribbon` | — | 保留 |

DOCX 的主要工作不是“重写编辑器”，而是：

- 同步 GenOffice 上游改动；
- 保持 `window.desktop` 适配层；
- 把打开/保存/另存为/打印/导出/最近文件接到 WorkStudio；
- 保持 AI 入口删除。

### 3.2 XLSX

GenOffice Sheets renderer 共 **339** 个文件，其中 `ai/` **33** 个。核心分组：

| 分组 | 代表文件 | 复用方式 | 说明 |
| --- | --- | --- | --- |
| 应用壳 | `App.tsx`、`ExcelShell.tsx` | **B 复制 + adapter** | Ribbon、状态栏、布局、快捷键 |
| Univer | `create-univer.ts`、`univer-sync.ts`、`univer-state.ts` | **A/B** | 我们已有 `create-univer.ts`，应换成 GenOffice 版 |
| Ribbon | `ribbon-actions.ts`、`excel-format-shortcuts.ts`、`ExcelShell.tsx` | **A/B** | 恢复顶部工具栏/功能区 |
| 对话框 | `FormatCellsDialog`、`FindReplacePanel`、`GoToDialog`、`PasteSpecialDialog`、`PrintDialog` 等 | **A 复制** | 依赖 UI/i18n，少量 host |
| 编辑/历史 | `edit-journal.ts`、`op-executor.ts`、`plan-operations.ts`、`undo-carry.ts` | **A/B** | 保存与撤销核心 |
| 保存 | `save-actions.ts`、`save-edits-staging.ts`、`save-scheduler.ts` | **B 适配** | 改为 `page.saveBytes` |
| 结构/格式 | `format-cells.ts`、`sheet-tab-actions.ts`、`outline-actions.ts`、`pivot-*` | **A/B** | 依赖 `xlsx-gateway` |
| 查找/筛选/排序 | `find-replace-session.ts`、`large-filter.ts`、`filter-*` | **A/B** | Univer presets |
| 打印/导出 | `print-*`、`csv-export.ts`、`pdf-*` | **B 适配** | 页面打印或交给壳层 |
| 右键菜单 | `cell-context-menu.ts`、`sheet-context-menu.ts`、`SheetTabContextMenu.tsx` | **A/B** | 恢复 GenOffice 菜单 |
| AI | `ai/` 33 个 | **D 删除** | AI 面板、工具、技能 |

**XLSX 的主进程依赖**：

- `apps/sheets/src/main/sheets-main.ts`
- `xlsx-sidecar-client.ts`
- `save-edits-transfer.ts`
- `print-*` / `pdf-*` / `csv-*`
- `xlsx-gateway` + 原生 `native/xlsx-engine`

目标不是把 `sheets-main` 搬进 WorkStudio 主进程，而是：

1. 页面内直接使用 `xlsx-gateway` 的纯 TS 写回（浏览器 shim）；
2. 大文件/公式流式能力后续再接 WorkStudio 主进程或 sidecar；
3. 保存仍走 `ura.page.saveBytes`。

当前我们只接了 `UniverSheetsCorePreset({ toolbar:false })`，所以只有名称框、公式栏、网格；**GenOffice 的 Ribbon / 对话框 / 状态栏没有搬**。这就是工具栏消失的原因。

### 3.3 PPTX

GenOffice Slides renderer 共 **263** 个文件，其中 `ai/` **20** 个。核心分组：

| 分组 | 代表文件 | 复用方式 | 说明 |
| --- | --- | --- | --- |
| 应用壳 | `App.tsx`、`components/Ribbon.tsx` | **B 复制 + adapter** | 幻灯片工作区、Ribbon |
| 画布 | `SlideCanvas.tsx`、`NodeBody.tsx`、`konva-adapter.ts` | **A/B** | Konva 渲染、命中选择 |
| 文本编辑 | `TextEditOverlay.tsx` | **A/B** | 原位 contentEditable，正是当前缺的 |
| 格式 | `style-actions.ts`、`FormatPane.tsx`、`RibbonHomeTab.tsx` | **A/B** | 字体、填充、段落 |
| 操作 | `slide-actions.ts`、`arrange-actions.ts`、`clipboard-actions.ts`、`insert-actions.ts` | **A/B** | 增删页、排列、复制粘贴 |
| 表格/图表/媒体 | `table-*`、`ChartBody.tsx`、`insert-*` | **A/B** | 依赖 `pptx-engine` / `pptx-ops` |
| 放映 | `show-actions.ts`、`SlideShowView.tsx`、`PresenterView.tsx` | **B/C** | 需要 WorkStudio 窗口能力 |
| AI | `ai/` 20 个 | **D 删除** | AI 面板、工具 |
| 主进程 | `slides-main.ts`、`presenter-show.ts`、`font-*`、`media-*` | **B/C** | 大部分操作应在页面内用 `pptx-ops` 完成 |

PPTX 当前页面的 `SVG + 模态文本编辑` 是自研精简实现，和 GenOffice Slides 的 `Konva + TextEditOverlay` 不是同一套。继续在 SVG 上补交互会越走越远；应直接移植 GenOffice Slides 的 renderer 组件。

---

## 4. Host API 分类

GenOffice renderer 通过 `window.desktopApi`（Docs/Sheets）和 `window.slidesApi`（Slides）访问宿主。它约有：

| 类别 | 例子 | WorkStudio 处理 |
| --- | --- | --- |
| 主题 / 语言 | `getTheme`、`onThemeChanged`、`getLanguage` | 映射到 `app:ui.get` / `bus` |
| 文件生命周期 | `selectWorkbook`、`openPptx`、`closeWorkbook` | 页面会话 + 资源面板 |
| 保存 | `saveWorkbookEdits`、`writeRecovery` | `ura.page.saveBytes` |
| 工作簿数据 | `readWorkbookRange`、`readWorkbookFormulas`、`recalcWorkbook` | 页面内 `xlsx-gateway` / 预载快照 |
| 幻灯片编辑 | `slidesApi.editText`、`addSlide`、`batchEditTransform` | 页面内 `pptx-engine` / `pptx-ops`，或 Host RPC |
| 打印/导出 | `printWorkbook`、`exportPdf`、`exportCsv` | 页面打印 / 壳层保存对话框 |
| 字体/媒体 | `fontCatalog`、`readLocalImage`、`privateFontFaces` | 页面 API + 内置资源 |
| 系统窗口 | `listDocsTabs`、`setShowFullScreen`、`presenter-show` | WorkStudio 视图/窗口能力 |
| MCP / CLI / 恢复 | `onMcpCommand`、`consumeHeadlessExport`、`onRecoveryPrompt` | 暂不接，返回 no-op |
| AI | `aiStream`、`aiChat`、`webSearch`、`pickAttachments` | **删除** |

### 4.1 推荐做法：Host Adapter Shim

不要修改 renderer 里成百上千个 `window.desktopApi.xxx` 调用点。页面启动时安装一个兼容对象：

```ts
window.desktopApi = createGenOfficeDesktopAdapter({
  pageSession,
  engine,
  theme,
  saveBytes,
})

window.slidesApi = createGenOfficeSlidesAdapter({
  pageSession,
  deckEngine,
  theme,
})
```

这样：

- renderer 文件可以**原样复制**；
- 差异集中在 adapter；
- AI 方法直接不实现 / 抛出明确错误；
- 文件、主题、保存、打印等行为由 WorkStudio 决定。

### 4.2 两种引擎放置方案

| 方案 | 做法 | 优点 | 缺点 |
| --- | --- | --- | --- |
| A. 页面内引擎 | 把 `xlsx-gateway` / `pptx-engine` / `pptx-ops` 打进页面 | 不扩展 WorkStudio 协议；保存走 `saveBytes` | 页面包大；Node API 要 shim |
| B. 宿主 RPC | WorkStudio 增加插件 RPC 通道，主进程执行引擎 | 页面轻；可复用原生 sidecar | 要改 WorkStudio 框架协议 |

建议：

- **XLSX**：页面内 `xlsx-gateway` 写回；大文件/公式流式后续再考虑宿主 RPC。
- **PPTX**：页面内 `pptx-engine` / `pptx-ops`；媒体/字体/放映等原生能力走 WorkStudio 命令。
- **DOCX**：已经是页面内 ProseMirror。

---

## 5. 复用等级

| 等级 | 含义 | 处理 |
| --- | --- | --- |
| **A 直接复用** | 无 Electron / 无宿主 API | 原样复制，改 import 路径 |
| **B 适配复用** | renderer 可复制，但调用 `window.desktopApi` / `slidesApi` | 复制 + host adapter |
| **C 参考实现** | 依赖主进程、原生 sidecar、窗口系统 | 参考其行为，在页面/WorkStudio 重新接 |
| **D 删除** | AI、CLI、Electron main/preload、Home | 不进入插件 |

规则：

1. **渲染层不重写**：Ribbon、对话框、快捷键、右键菜单优先原样复制。
2. **差异进 adapter**：不要在每个组件里改 `window.desktopApi`。
3. **AI 全部剥离**：`ai/` 目录、AI 面板挂载、AI 设置、AI IPC、`agent-core` / `ai-provider` / `ai-search`。
4. **文件字节只走页面会话**：渲染层不直接碰路径。
5. **上游可同步**：保留 GenOffice 文件结构，便于后续 diff/merge。

---

## 6. 推荐目录与构建

```
packages/xlsx/
├─ src/editor/                  ← fork apps/sheets/src/renderer
│  ├─ App.tsx / ExcelShell.tsx
│  ├─ components/ / dialogs/
│  ├─ univer-* / edit-journal / op-executor
│  ├─ i18n/（按语言裁剪）
│  └─ ai/（删除）
├─ src/page/                    ← WorkStudio 宿主
│  ├─ main.ts
│  ├─ host-adapter.ts
│  ├─ workbook.ts
│  └─ save.ts
└─ vite.config.ts

packages/pptx/
├─ src/editor/                  ← fork apps/slides/src/renderer
│  ├─ App.tsx / SlideCanvas / NodeBody
│  ├─ TextEditOverlay / Ribbon / FormatPane
│  ├─ slide-actions / arrange-actions / clipboard-actions
│  ├─ i18n/
│  └─ ai/（删除）
├─ src/page/                    ← WorkStudio 宿主
│  ├─ main.ts
│  ├─ host-adapter.ts
│  └─ save.ts
└─ vite.config.ts
```

构建配置：

- vendor sparse checkout 增加：`packages/xlsx-gateway`、`packages/pptx-ops`（`pptx-engine`/`pptx-render` 已有）。
- Vite alias 映射 `@genoffice/*` 到 `vendor/genoffice/packages/*`。
- Node shim：`node:crypto`、`node:fs`、`node:path`、`Buffer`（仅页面内引擎需要）。
- AI 依赖检查：构建时禁止 `@genoffice/agent-core` / `ai-provider` / `ai-search` 进入页面 bundle。

---

## 7. 实施路线

### Phase 0：复用准备

- 增加 `@genoffice/xlsx-gateway`、`@genoffice/pptx-ops` 到 vendor。
- 建立 `ui-lite`（或 subpath imports），移除 AI 依赖。
- 建立 `GenOfficeHostAdapter` 契约和 no-op 实现。
- 建立 AI 剥离脚本/检查。

### Phase 1：XLSX 完整 renderer

1. 复制 `apps/sheets/src/renderer` → `packages/xlsx/src/editor`。
2. 删除 `ai/`，移除 `App.tsx` 的 AI 面板、AI 设置、AI hooks。
3. 复制 `ExcelShell` / Ribbon / 对话框 / 右键菜单 / 状态栏。
4. 用 `host-adapter` 实现：
   - `getTheme` / `getLanguage` / `getDocumentTheme`；
   - `readWorkbookRange` / `readWorkbookFormulas`（页面内 gateway）；
   - `saveWorkbookEdits` → `saveBytes`；
   - 打印/导出 → WorkStudio 命令或页面打印。
5. 保存接入 `xlsx-gateway`：cell / format / structural / layout。
6. 保留现有 Univer 作为网格引擎。

验收：GenOffice Sheets 的 Ribbon、格式对话框、右键菜单、查找、工作表操作、保存全部可用；AI 入口不存在。

### Phase 2：PPTX 完整 renderer

1. 复制 `apps/slides/src/renderer` → `packages/pptx/src/editor`。
2. 删除 `ai/` 与 AI 面板挂载。
3. 用 `host-adapter` 实现 `slidesApi`：
   - `editText` / `editTransform` / `addSlide` / `deleteSlide` / `arrange` / `table` → 页面内 `pptx-engine` / `pptx-ops`；
   - 字体/媒体 → 页面资源或 WorkStudio 命令；
   - 打印/放映 → 后续接 WorkStudio。
4. 保存 → `savePptx` / `saveBytes`。
5. 保留 Konva / TextEditOverlay / Ribbon / FormatPane。

验收：选择、拖动、缩放、旋转、原位文本编辑、幻灯片操作、格式、撤销、保存与 GenOffice Slides 一致；AI 入口不存在。

### Phase 3：DOCX 同步与收尾

- 同步 GenOffice Docs 上游；
- 保持 AI 剥离；
- 补齐打开/保存/打印/导出系统命令；
- 快捷键说明书与页面命令对齐。

---

## 8. 风险与决策

| 风险 | 影响 | 对策 |
| --- | --- | --- |
| `xlsx-gateway` 原生 sidecar | 大文件、公式流式、完整保存 | 先用纯 TS 路径；后续加 WorkStudio RPC/sidecar |
| `slidesApi` 主进程操作 | PPTX 编辑依赖主进程引擎 | 在页面内跑 `pptx-engine/pptx-ops`，或扩展 WorkStudio RPC |
| `@genoffice/ui` 依赖 AI 包 | AI 代码回流 bundle | 建 `ui-lite` 或 subpath exports |
| renderer 包体大 | 首屏慢 | 懒加载、worker、locale 裁剪；与 GenOffice 同量级 |
| 上游同步 | fork 漂移 | 保留目录结构，adapter 集中在 `src/page` |
| 原生窗口/放映/打印 | 页面无法完全实现 | 由 WorkStudio 视图/命令承接，或标记暂不支持 |
| 文件正确性 | 结构化保存缺失 | 在写回接入前拦截；不做“改了但保存丢失”的操作 |

---

## 9. 一句话判断

**能复制的就复制，能原样跑的就原样跑；只在 host adapter 里做 WorkStudio 适配，AI 全部删除。**

当前 XLSX 的 Univer 页面是“引擎接入”，不是“GenOffice Sheets 复用完成”。下一步应停止手写精简工具栏，直接移植 `apps/sheets/src/renderer` 的 `ExcelShell` / Ribbon / 对话框 / 保存链路。
