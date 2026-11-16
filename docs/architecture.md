# GenOffice 架构与 WorkStudio Office 插件

> 这份文档只回答三个问题：
>
> 1. GenOffice 的源代码为什么这样分层？
> 2. 这套分层如何把“Office 功能”拆成可复用的部件？
> 3. 我们怎样把这些部件改造成 WorkStudio 插件，同时不重写 Office？

---

## 0. 一页结论

GenOffice 不是一个“Office 页面”，而是一套 **格式引擎 + 操作层 + 渲染树 + 交互外壳** 的组合。

```
文件生命周期
    │
    ▼
包 / OOXML / 格式解析          ← codec / engine
    │
    ▼
文档模型（可编辑语义）          ← document / slide / workbook model
    │
    ▼
操作与事务（可撤销）            ← ops / journal / gateway
    │
    ▼
渲染树 / 画布 / 文本布局        ← render tree / Konva / ProseMirror / Univer
    │
    ▼
Ribbon / 对话框 / 快捷键 / 状态栏 ← application shell
    │
    ▼
宿主：文件、窗口、打印、导出、AI  ← Electron main / preload
```

WorkStudio 插件只接管**最上层和最下层**：

- 最上层：用 WorkStudio 的标签、主题、资源面板、保存语义替代 GenOffice 的应用壳；
- 最下层：用 `ura.page.saveBytes` 替代 GenOffice 的路径式文件写入；
- 中间四层尽量原样复用 GenOffice，尤其是引擎、操作层、渲染树和交互外壳。

因此我们做的不是“重新实现一个 Office”，而是：

> **把 GenOffice 的 Office 解构结果，重新组装到 WorkStudio 的文档生命周期里。**

---

## 1. GenOffice 源码结构

GenOffice 是 Electron monorepo，核心目录可以按职责直接映射到 Office 功能：

```
vendor/genoffice
├─ apps/
│  ├─ docs/     Word 应用：main + renderer + shared
│  ├─ sheets/   Excel 应用：main + renderer + shared
│  └─ slides/   PowerPoint 应用：main + renderer + shared
│
├─ packages/
│  ├─ docx-engine/       DOCX 解析、模型、patch
│  ├─ pptx-engine/       PPTX 解析、模型、patch
│  ├─ pptx-ops/          PPTX 操作注册表、事务、撤销记录
│  ├─ pptx-render/       PPTX → RenderTree → Konva
│  ├─ xlsx-gateway/      XLSX OOXML 读写、结构、格式、公式
│  ├─ ui/                Ribbon 基础件、颜色、下拉、形状库、查找
│  ├─ i18n/              语言、方向、快捷键格式化
│  ├─ zip-gate/          安全 ZIP 解包
│  ├─ electron-utils/    主进程辅助（打印、headless、字体）
│  └─ agent-core / ai-provider / ai-search / pipelines   AI
│
└─ native/               原生引擎（如大文件 XLSX）
```

### 1.1 一个应用内部的固定分层

每个 `apps/*` 都遵循同一结构：

```
apps/<app>/src/
├─ main/        Electron 主进程：文件、窗口、打印、导出、sidecar、AI IPC
├─ renderer/    React 页面：App、Ribbon、画布、对话框、快捷键
└─ shared/      主进程 ↔ renderer 的 IPC 类型与纯函数
```

这不是“目录习惯”，而是 Office 功能被切开后的边界：

| 层 | 放在哪里 | 负责什么 | 不负责什么 |
| --- | --- | --- | --- |
| 文件生命周期 | `main` | 打开、保存、另存为、最近文件、窗口 | 文档语义 |
| 包 / 格式 | `engine` / `gateway` | ZIP、OOXML、关系、内容类型 | 用户交互 |
| 文档模型 | `engine` | 段落、幻灯片、单元格、样式、公式 | 像素坐标、按钮 |
| 操作 / 事务 | `ops` / `gateway` | 修改模型、校验、撤销、journal | 菜单布局 |
| 渲染树 | `renderer` + `render` | 把模型变成可绘制的节点 | 文件写入 |
| 交互外壳 | `renderer` | 选择、拖动、双击、Ribbon、对话框 | OOXML 细节 |
| 宿主能力 | `main` / `preload` | 打印、PDF、字体、系统窗口、AI | 文档结构 |

### 1.2 三种应用的共同运行链路

```
字节
  │  main 读取文件 / sidecar / 原生引擎
  ▼
格式解析
  │  docx-engine / pptx-engine / xlsx-gateway
  ▼
文档模型
  │  段落 / 幻灯片元素 / 工作表单元格
  ▼
渲染树
  │  ProseMirror / Konva + pptx-render / Univer
  ▼
用户操作
  │  选择、输入、拖动、Ribbon、对话框
  ▼
操作 / 事务
  │  pptx-ops / xlsx-gateway / 编辑器 mutation
  ▼
重新渲染 + 写回字节
```

这条链路里，**只有“渲染树 → 用户操作 → 操作层”需要大量 UI 代码**；格式和模型可以独立测试，不需要 Electron。

---

## 2. Office 功能的解构

把 Office 功能按“用户看到的能力”拆开，会得到下面七层。GenOffice 的目录结构基本就是这七层的物化。

### 2.1 文件身份与生命周期

**能力**：打开、保存、另存为、最近文件、窗口标题、关闭前保存、外部变更。

**GenOffice**：`apps/*/src/main`、preload、`window.desktopApi` / `window.slidesApi`。

**本质**：文档在 Office 里首先是“一个路径 + 一个进程会话”。

**WorkStudio**：文档身份是 URA 资源，不是路径；页面只拿到 `sessionId` 和字节。文件生命周期归壳层。

### 2.2 包与格式

**能力**：ZIP、`[Content_Types].xml`、关系、主题、字体、媒体、工作表、幻灯片、段落。

**GenOffice**：`docx-engine`、`pptx-engine`、`xlsx-gateway`、`zip-gate`。

**本质**：OOXML 不是一个对象，而是一组互相引用的 part；编辑器只是这些 part 的语义视图。

**WorkStudio**：插件 codec 决定 MIME/codec；页面内部仍用 GenOffice 引擎读写 part。

### 2.3 文档模型

**能力**：段落/run、幻灯片元素树、单元格/样式/公式、节、批注、页眉页脚。

**GenOffice**：
- DOCX：`docx-engine` 的文档模型 + ProseMirror 文档；
- PPTX：`pptx-engine` 的 `Slide` / `SlideElement`；
- XLSX：`xlsx-gateway` 的 workbook/sheet/cell 模型 + Univer 工作簿。

**本质**：模型是保存的唯一来源；渲染和交互都只是模型的投影。

### 2.4 操作、事务与历史

**能力**：插入、删除、改文本、改格式、移动、组合、合并、撤销、重做。

**GenOffice**：
- PPTX：`pptx-ops` 的 `runTxn`、op registry、`OpRecord`；
- XLSX：`xlsx-gateway` 的 plan/patch + Univer mutation / edit journal；
- DOCX：ProseMirror transaction + 自定义命令。

**本质**：Office 的“编辑”不是直接改 DOM，而是“提交一个可验证、可逆的操作”。

这是 GenOffice 最值得复用的层，因为它同时解决：
- 操作合法性；
- 引导式错误；
- 事务原子性；
- 撤销记录；
- 保存时的字节写回。

### 2.5 渲染树

**能力**：文本换行、字体度量、形状、图片、表格、图表、幻灯片缩略图。

**GenOffice**：
- DOCX：ProseMirror + 分页/页眉页脚；
- PPTX：`pptx-render` 的 `RenderSlide` / `RenderNode` + Konva；
- XLSX：Univer 的 canvas 网格。

**本质**：把模型变成“屏幕坐标系里的节点”，但节点不反向拥有文档。

### 2.6 交互外壳

**能力**：Ribbon、右键菜单、对话框、状态栏、快捷键、选择模型、拖拽、原位编辑。

**GenOffice**：`apps/*/src/renderer` 的 `App.tsx`、`ExcelShell`、`Ribbon`、`components`、`*-actions.ts`。

**本质**：用户感知到的“Office”主要在这一层。它依赖引擎，但不属于引擎。

### 2.7 AI

**能力**：聊天、工具调用、生图、搜索、文档生成。

**GenOffice**：`apps/*/ai`、`agent-core`、`ai-provider`、`ai-search`、`pipelines`。

**本质**：AI 是宿主能力，不是 Office 格式能力。

**WorkStudio**：全部移除。AI 入口不进页面；页面不依赖 agent/provider 包。

---

## 3. 三种格式的具体映射

### 3.1 DOCX：ProseMirror 文档

```
DOCX bytes
  → docx-engine：ZIP / OOXML / 样式 / 编号 / 关系
  → ProseMirror 文档
  → 分页渲染 + 页眉页脚
  → 编辑 transaction
  → 重新序列化 DOCX bytes
```

关键复用：

| 部件 | 来源 | 我们怎么用 |
| --- | --- | --- |
| 文档解析 / 写回 | `@genoffice/docx-engine` | 原样复用 |
| 编辑器 | GenOffice Docs `apps/docs/src/renderer` | fork 到 `packages/docx/src/editor`；`docx-pro` 再复制一份到自己的 `src/editor` |
| 页面宿主 | GenOffice main/preload | 替换为 `packages/docx/src/page` |
| AI | `apps/docs/src/ai` | 页面入口移除；源码尚未完全清理 |

DOCX 的页面入口仍然保留“字节进、字节出”的模型：

```
ura.page.get → bytes → 编辑器
编辑器提交 → DOCX bytes → ura.page.saveBytes
```

### 3.2 XLSX：Univer 网格 + xlsx-gateway

```
XLSX bytes
  → xlsx-gateway：workbook / worksheet / styles / sharedStrings
  → WorkbookFile + 页面数据
  → Univer preset-sheets-core
  → 单元格编辑 / 选择 / 公式 / 结构 / 格式
  → WorkbookSaveRequest
  → planCellEditsToXlsx → assembleWithJsZip → saveBytes
```

关键复用：

| 部件 | 来源 | 我们怎么用 |
| --- | --- | --- |
| 网格 / 选择 / 公式栏 | Univer + GenOffice Sheets renderer | fork `apps/sheets/src/renderer` |
| OOXML 写回 | `@genoffice/xlsx-gateway` | 页面内直接调用纯 TS 路径 |
| Ribbon / 对话框 | GenOffice `ExcelShell` / `components` | 原样保留，删除 AI 入口 |
| 主进程 sidecar | GenOffice `apps/sheets/src/main` | 不搬；保存走页面内 `saveBytes` |

XLSX 保存必须使用**完整参数**：

```ts
const plan = await planCellEditsToXlsx(
  await createBufferEntrySource(sourceBuffer),
  edits, structuralOps, chartEdits, sheetPlan,
  filterStates, hyperlinkEdits, cfStates, dvStates,
  sheetProtections, definedNamesState, visualAdditions,
  pageSetupStates, noteStates, tableAdditions, pivotAdditions,
  pivotCacheRefreshPaths, pivotRefreshUpdates, visualEdits,
  sparklineAdditions, formulaValues, themeState,
  workbookProtectionState, protectedRangeStates,
  bulkConstantFills, tabColorStates, tableEdits,
  cellImageAdditions,
)
const bytes = await assembleWithJsZip(sourceBuffer, plan)
```

只调用 14 参数的旧包装会静默丢掉视觉对象、表格、透视、主题、保护、单元格图片等写回字段。

### 3.3 PPTX：pptx-engine + pptx-ops + pptx-render + Konva

```
PPTX bytes
  → pptx-engine：presentation / slides / layouts / masters / rels
  → Slide + SlideElement
  → pptx-render：RenderSlide / RenderNode
  → Konva + TextEditOverlay（原位 contentEditable）
  → pptx-ops：runTxn 修改模型
  → 重建 RenderSlide
  → savePptx → saveBytes
```

关键复用：

| 部件 | 来源 | 我们怎么用 |
| --- | --- | --- |
| 解析 / 写回 | `@genoffice/pptx-engine` | 原样复用 |
| 操作 / 事务 | `@genoffice/pptx-ops` | 原样复用 |
| 渲染树 | `@genoffice/pptx-render` | 原样复用 |
| 画布 / 文本 / 交互外壳 | GenOffice Slides `apps/slides/src/renderer` | fork 到 `packages/pptx/src/editor` |
| 主进程 | GenOffice `apps/slides/src/main` | 不搬；页面内实现 host adapter |

PPTX 的编辑原则：

- 所有模型修改都走 `runTxn`；
- 一次用户手势只产生一个撤销单元；
- 保存时 `savePptx` 生成新字节，再 `saveBytes`；
- 保存后重新 `openPptx(bytes)`，刷新元素 id 和字节锚点。

---

## 4. WorkStudio 的接入契约

Office 插件不改变 WorkStudio 的插件模型。它只要求三层接缝。

### 4.1 codec / presenter / sidecar

每个格式注册三个键：

```ts
registerCodec({
  id: 'pptx',
  keys: [{ transport: '*', mime: MIME.pptx }],
  extensions: ['pptx', 'pptm', 'potx'],
  patchType: 'bytes',
  decode,
  locate,
  encode,
})

registerPresenter({
  key: { transport: '*', mime: MIME.pptx, codec: 'pptx' },
  type: 'page',
  surface: 'pptx-viewer',
  present,
})

registerSidecar({
  key: { transport: '*', mime: MIME.pptx, codec: 'pptx' },
  sidecar,
})
```

- **codec**：字节 ↔ `{ doc, anchorMap, outline? }`；
- **presenter**：决定用哪个页面 surface 呈现；
- **sidecar**：给 agent 的纯文本，不参与页面渲染。

格式身份是 `codec.id`，不是扩展名，也不是 `hints` / `docType`。

### 4.2 页面会话

页面只拿：

```ts
window.extension.ura.page.get<T>(sessionId)   // 页面数据
window.extension.ura.page.saveBytes(id, bytes) // 写回
window.extension.bus.on(...)                   // 主题 / 脏态 / 壳层事件
```

页面**不碰路径**，不直接写文件，不持有 Node 能力。

保存返回：

```ts
{ ok: true, version: string }
```

页面只负责生成正确的 OOXML 字节；写队列、版本和冲突归 WorkStudio 的 `PageSessions`。

### 4.3 页面生命周期

```
打开资源
  → codec.decode(bytes)
  → presenter.present({ bytes, name, ref, version, ... })
  → 页面加载 dist/page/index.html?session=…
  → 页面读取 session 数据
  → 安装 host adapter
  → 渲染编辑器
  → 用户编辑
  → saveBytes
  → 页面继续编辑
```

---

## 5. 我们怎样改造和封装

### 5.1 一条原则：不改 renderer 的交互，不改 engine 的语义

```
GenOffice renderer  → fork，尽量原样保留
GenOffice engine    → vendor，原样复用
GenOffice main      → 不进入插件
GenOffice preload   → 不进入插件
宿主差异            → 全部收敛到 host adapter
```

host adapter 做三件事：

1. 安装 `window.desktopApi` / `window.slidesApi` 兼容对象；
2. 把页面会话数据构造成编辑器需要的 `WorkbookFile` / `OpenedPptx` / `RenderSlide`；
3. 把 editor 的保存请求翻译成字节，再走 `ura.page.saveBytes`。

### 5.2 目录结构

```
packages/
├─ docx/
│  └─ src/
│     ├─ ura.ts                     codec / presenter / sidecar 注册
│     ├─ page/                      WorkStudio 页面入口与 host 适配
│     └─ editor/                    GenOffice Docs renderer fork
│
├─ xlsx/
│  └─ src/
│     ├─ ura.ts
│     ├─ shared/                    desktopApi 类型、IPC 类型、保存 schema
│     └─ editor/
│        ├─ App.tsx / ExcelShell.tsx
│        ├─ workstudio-host.ts      desktopApi adapter
│        ├─ workstudio-save.ts      WorkbookSaveRequest → xlsx-gateway
│        ├─ main.tsx                页面入口
│        └─ shims/                  node:crypto / fs / path 的浏览器替代
│
└─ pptx/
   └─ src/
      ├─ ura.ts
      ├─ shared/                    slidesApi 类型、编辑操作类型
      └─ editor/
         ├─ App.tsx / Ribbon / SlideCanvas
         ├─ workstudio-host.ts      slidesApi adapter + pptx-engine/ops
         ├─ main.tsx                页面入口
         ├─ redaction.ts            人工脱敏标签
         └─ shims/                  node:crypto / zlib 的浏览器替代
```

### 5.3 构建与 vendor

GenOffice 通过 `scripts/setup-genoffice.mjs` sparse checkout 到 `vendor/genoffice`。

Vite 侧：

- `genofficeSource()` 把 `@genoffice/*` 解析到 vendor 源码；
- alias 处理 CSS / 字体 / 子路径；
- 页面 bundle 内含 engine，不依赖 WorkStudio 新 RPC；
- `buffer-global.ts` 必须在页面入口第一行 import，原因见 §5.5。

tsconfig 侧：

- 为 `@genoffice/*` 显式配置 `paths`；
- `src/editor` / `src/shared` 已纳入 typecheck；
- CSS / 图片 / 字体 / `?raw` / `?url` 有模块声明。

### 5.4 保存链路

| 格式 | 页面请求 | 页面内处理 | 最终写入 |
| --- | --- | --- | --- |
| DOCX | 编辑器序列化 | 重建 DOCX 包 | `saveBytes` |
| XLSX | `WorkbookSaveRequest` | `planCellEditsToXlsx` + `assembleWithJsZip` | `saveBytes` |
| PPTX | 编辑操作累积 | `runTxn` + `savePptx` | `saveBytes` |

三条链路共享一个不变量：

> 页面只产生字节；路径、版本、冲突和串行写入归宿主。

### 5.5 运行时坑

这些不是架构问题，但会直接影响“能不能跑”：

- **JSZip 的 `Buffer` 快照**：必须在模块求值前注入 `buffer-global`，否则同一 bundle 内 `typeof Buffer` 为 `undefined`，ZIP 读写失败。
- **node:crypto**：xlsx-gateway 需要同步 `sha256`，页面用 `@noble/hashes` 实现。
- **node:zlib**：pptx-engine 需要 `deflateSync`，页面用 `fflate`。
- **node:fs / path**：xlsx-gateway 有一部分 Node API，用内存 shim 兜底；真正的文件读写仍只走 `saveBytes`。
- **React / react-konva**：Konva 版本必须和 React 版本匹配，否则运行时 canvas 报错。
- **元素 id**：PPTX 保存/重排/图表编辑后元素 id 会重解析；任何“保存后继续编辑”的路径都必须用新的 `RenderSlide`，不能保留旧 `sourceId`。

### 5.6 AI 的删除方式

AI 不是用 CSS 隐藏，而是从页面层删掉：

- 不挂载 AI 面板；
- 不注册 AI Ribbon / 右键 / 快捷键；
- 不 import `agent-core` / `ai-provider` / `ai-search`；
- 不保留 AI IPC 和 host stub；
- XLSX / PPTX 的 `ai/`、AI i18n、AI host 方法已删除；
- DOCX 页面入口不暴露 AI，但源码中的 AI 模块尚未完全清理。

---

## 5.7 合体版：`docx-pro`（完整编辑器 + 计划型写）

`packages/docx` 与 `packages/office-docx` 各占一半：前者是完整编辑器（字节权威
在 ProseMirror 序列化器），后者是计划型写协议（字节权威在块级 XML 切片）。
宿主把写形态做成**每插件单值**（`page-sessions.commit` 用同一个 `writeModeOf`
裁决 `bytes` 与 `plan`），所以合体不能靠“装两个插件”，只能让页面也交计划。

`docx-pro` 的做法：

```text
                     ┌─ 有页面：PM 事务 → 编辑器 save() → serialize
plan { kind: 'bytes' } ┘                        ↘
                                                  exec → 权威层 CAS → 落盘
plan { kind: 'save' }  ── 无页面：块级 XML 切片 ──↗
```

- `write: 'plan'` + `engine.{tools,plan,exec,read,runtime}`；
- 编辑器源码从 `packages/docx/src/editor` **复制**进 `src/editor`（插件之间不互相
  import），只换宿主适配层的保存函数：
  `saveBytes` → `commit({ payload: { type: 'plan', plan: { kind: 'bytes', bytes } } })`；
- 无页面时 `engine.plan` 用与 `office-docx` 同一套块命令产 `{ kind: 'save' }`，
  `engine.exec` 只重写 `word/document.xml`，其余 zip 条目逐字节拷贝；
- 读类工具只能在执行点（宿主不把 `readOnly` 工具交给页面代理），因此
  `get_document_context` / `read_blocks` / 读批注与修订都从字节侧回答——代价是
  拿不到“当前选区”，SKILL 里明确要求用块索引定位；
- 工具面只有三个：`docx_guide`（目录）/ `docx_read`（读）/ `docx_apply`（命令批），
  目录里 41 条命令（blocks 8 · editor 23 · actions 10）。需要页面的 op 由页面代理执行，
  无页面时回 `command.page-unavailable`，不回退 headless 写第二份模型。

三者的取舍：

| 插件 | 写形态 | 页面 | 适合 |
| --- | --- | --- | --- |
| `docx` | `bytes` | 完整编辑器 | 纯人工编辑 |
| `office-docx` | `plan` | 块列表 | 纯后台/批处理 |
| `docx-pro` | `plan` | 完整编辑器 | 人在场 + agent 同一份文档 |

三者声明同一 MIME：**同一工作区只能启用一个**。

---

## 6. 不变量

只要下面这些成立，改造就不会退回“手写精简 Office”：

1. **文件字节只走 `ura.page.saveBytes`**；页面不写路径。
2. **渲染层不直接读文件系统**；所有数据来自页面会话。
3. **模型修改必须走操作层**：PPTX `runTxn`，XLSX gateway / Univer mutation，DOCX transaction。
4. **保存映射必须完整**：XLSX 使用完整 `planCellEditsToXlsx` 参数，不能调用丢字段的旧包装。
5. **一次手势一个撤销单元**：预览合并、保存后重开、id 刷新都要围绕这个原则。
6. **AI 循环不进入页面**：模型调用与工具编排由 WorkStudio 助手承担；页面只提供
   领域工具的执行体（`docx-pro` 的 3 个工具 / 41 条命令 + 页面代理），不挂聊天面板、
   不 import agent/provider 包。
7. **renderer 的差异集中在 host adapter**：不要在每个组件里写 WorkStudio 分支。
8. **typecheck 覆盖 renderer**：`src/editor` / `src/shared` 不能靠 `tsconfig` 排除来“通过”。

---

## 7. 当前能力边界

| 格式 | 已有 | 边界 |
| --- | --- | --- |
| DOCX | 打开、编辑、分页、Ribbon、审阅、保存 | 上游同步与部分 WorkStudio 系统命令仍待补 |
| DOCX（`docx-pro`） | 同上 + 3 个工具 / 41 条命令 + 无页面块命令 | 页面路径 `maxCommitBytes` 16 MB；无页面无选区；批注/页脚/页面设置等 15 个动作未接线 |
| XLSX | Ribbon、网格、公式、选择、格式、查找、筛选、透视、图表、保存 | 大文件/公式流式仍按页面内加载策略处理 |
| PPTX | Ribbon、画布、原位文本、变换、幻灯片/节、表格/图表、图片/媒体、动画、录音、母版、脱敏、放映、PDF、保存 | 观众窗口是 Document PiP；自动第二屏全屏/交换显示器受浏览器权限限制 |

观众窗口说明：

- 演示者视图会打开一个独立的 Document PiP 窗口；
- 窗口可以和主窗口分屏或拖到第二显示器；
- Chromium 不向页面暴露屏幕枚举/PiP 全屏权限，所以“自动检测第二屏并全屏”和“交换显示器”需要宿主提供 screen API。

---

## 8. 新增一个功能时怎么做

按层定位，不要从 UI 直接改字节：

1. **确认属于哪一层**
   - 格式/解析 → engine / gateway；
   - 模型修改 → ops / transaction；
   - 屏幕表现 → render tree；
   - 用户入口 → Ribbon / 对话框 / 快捷键。
2. **优先找 GenOffice 现有实现**
   - 有 op：直接映射；
   - 有 gateway 参数：接到保存请求；
   - 有 UI：原样 fork，不改交互。
3. **补 host adapter**
   - 输入：页面会话 / 编辑器状态；
   - 输出：`RenderSlide` / `ApplyOutcome` / bytes。
4. **补保存字段**
   - XLSX：确认 `WorkbookSaveRequest` → `planCellEditsToXlsx` 的完整参数；
   - PPTX：确认 op 已写回 `opened`，`savePptx` 能序列化。
5. **验证**
   - `pnpm typecheck`；
   - 页面打开、操作、保存、重新打开；
   - ZIP 部件与关键 XML 检查。

---

## 9. 验证与安装

```bash
pnpm setup          # sparse checkout GenOffice + pnpm install
pnpm typecheck      # docx / xlsx / pptx，包含 renderer
pnpm build          # codec + 页面 bundle
pnpm verify         # URA codec / presenter / sidecar 契约
```

打包与安装：

```bash
pnpm pack:plugins   # release/<id>-<version>.zip + sha256
pnpm install:plugins # <userData>/plugins/installed/<id>（开发）
pnpm install:builtin # WorkStudio 源码内置插件目录（仅开发）
```

WorkStudio 的插件装载是新机制：

- 插件是自包含目录，`package.json` 的 `workstudio` 声明
  `id / label / icon / page / apiVersion`，`exports["./ura"]` 是 ESM 注册入口；
- 远程只接受 zip，一个 zip 只装一个插件；WorkStudio 负责下载、校验 `sha256`
  与安装；
- 本地目录原地加载，不复制，适合开发；
- 安装后重启 WorkStudio 才装载；安装时不执行 npm install。

发布由 `.github/workflows/release.yml` 完成：推送 `v*` tag 后，在
GitHub Actions 中构建、打包、生成 sha256，创建 GitHub Release。普通分支/PR 由 `.github/workflows/ci.yml` 构建并上传 artifact。

运行态最低验收：

```text
xlsx: ribbon=true  ai=false  canvases=3
pptx: ribbon=true  ai=false  canvases=7
```

---

## 10. 最终判断

GenOffice 的价值不在“它有一个现成的 Office 页面”，而在于它已经把 Office 拆成了：

```
格式 → 模型 → 操作 → 渲染 → 交互 → 宿主
```

我们做的事情是：

- 保留中间的格式、模型、操作、渲染、交互；
- 替换最上层的应用壳和最下层的文件宿主；
- 删掉不属于 Office 格式能力的 AI；
- 用 WorkStudio 的页面会话和 `saveBytes` 把两端接起来。

这就是本仓库的结构，也是后续任何 Office 功能改造的判断标准：

> **功能应该落在它所属的层里；跨层的地方，只允许有一个 host adapter。**
