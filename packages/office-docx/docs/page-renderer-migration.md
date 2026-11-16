# 页面层迁移方案：docs 渲染器 + 照抄工具栏 → office-docx（v2 · plan 写盘）

> 目标：`office-docx` 的页面从「块列表 + 自绘精简工具栏」换成 **GenOffice Docs 渲染器（React + TipTap + 分页）**，
> 工具栏直接复用 `packages/docx/src/editor/components/Ribbon.tsx` 及其 5 个 tab，人编辑走 PM（单编辑宿主），
> 写盘仍走本包的 `plan → /commit → exec`（唯一字节产生者不变）。

---

## 0. 当前状态（P0–P6）

| 阶段 | 状态 | 结果 |
| --- | --- | --- |
| P0 渲染器副本 + 切分探针 + 构建骨架 | ✅ | `src/page-v2/genoffice/`；`check:slicing`；`dist/page-v2` |
| P1 只读挂载 | ✅ | 渲染器 + Ribbon + 分页跑起来，保存返回明确错误 |
| P2 plan 写盘 | ✅ | 字节 diff → `SavePlan` → `/commit`；能力闸门；自动保存恢复 |
| P2b agent 代理 | ✅ | context 投影 / apply（PM 事务 + 渲染器 ops）→ 保存 → 回执 |
| P3 切分同源 | ⛔ 取消 | 计划改成**字节 diff** 后不再依赖 docx-engine 的 `docxIndex`，切分口径天然一致 |
| P4 工具栏按组开放 | ✅ | 布局（sectPr）/ 表格已验证落盘；sectPr 保序硬校验 |
| P5 整包 part 通道 | ✅ | 图片 / rels / 页眉页脚 / 脚注 / 批注 / 编号 / 主题 / 文档属性 |
| P6 收尾 | 部分 | README/迁移文档与清理说明已完成；**代码分割（4.5 MB 主 chunk）与真机回归待做** |

默认页已切到 `dist/page-v2/index.html`；回退只需把 `package.json` 的 `page` 改回
`dist/page/index.html`（旧块列表页仍保留在 `src/page/`）。

## 0.1 一句话结论

**「照抄工具栏」= 「照抄编辑器」。** Ribbon 的每个按钮都是
`editor.chain().focus().setBold().run()` / `editor.isActive('bold')` / `editor.can()...`，
依赖 TipTap 实例与 PM state。搬 UI 而不搬编辑器，等于把 `Ribbon.tsx`（4144 行）+ 5 个 tab（合计 ~5000 行）
的每个 handler 重写成我们的块命令，**并凭空丢掉表格 / 图片 / 列表 / 分页 / 样式面板**——比重写更贵，且最终还是要换编辑器。

所以路线是：**页面层整体替换为 docs 渲染器**，把接缝收窄到 5 个点（§3）。

---

## 1. 现状核对（已读代码）

| 事实 | 位置 |
| --- | --- |
| v1 `packages/docx` 页面 = GenOffice Docs 渲染器 fork + 宿主适配层 | `packages/docx/src/editor/App.tsx`（7439 行）、`src/page/host.ts`（319 行） |
| 工具栏就是 `Ribbon.tsx` + `ribbon-tabs / insert / layout / references / design / hf` | `packages/docx/src/editor/components/` |
| 保存流水线：`PM doc → pmDocToSavePlan → saveBlocks → saveDocx → bytes → 写盘 → 重解析` | `src/editor/file-actions.ts`（`buildDocBytes`，line ~754-860） |
| **`pmDocToSavePlan(inputDoc, originalBlocks)` 已存在**，产出 `saveBlocks: {kind:'original'\|'xml'}`，形状与我们 `SavePlan` 基本一致 | `src/editor/convert.ts:2006`，类型 `:1841` |
| 页面注入 `window.desktop.saveDocx(path, buffer)`；v1 走 `ura.page.saveBytes` | `src/page/host.ts:103-160` |
| 渲染器依赖：`@tiptap/*` 3.31、React 19、`@genoffice/docx-engine|font-metrics|ui|i18n|zip-gate`、vite + `genofficeSource` 插件 | `packages/docx/{package.json,vite.config.ts}` |
| 分页是「连续流 + 测量 + gap decoration」，大文档有 tier + 分阶段装载 | `src/editor/pagination-*`（~11k 行）、`large-document.ts`、`phased-content.ts` |

我们这边已有的、**保持不变**的部分：`engine/{plan,exec,read,runtime,zip}`、`ura.ts`、`shared/*`（含 run 级格式与
`blocks-edit`）、`catalog.ts`、工具面、`sync.manifest`、`tests/`（56 项）。

---

## 2. 目标形态：目录与依赖

```
packages/
├─ docx/                          ← v1 **零改动**（git status 干净）
└─ office-docx/
   ├─ src/page-v2/                ← 渲染器版页面（vite + React，P1 只读）
   │  ├─ index.html / main.tsx    ← get/bytes → blobUrl → 启动渲染器
   │  ├─ host.ts                  ← 从 v1 移植的 window.desktop 适配层（P2 改保存缝）
   │  ├─ session.ts               ← （P2）commit / 回执 / manifest rebase / agent-request
   │  ├─ plan.ts                  ← （P2）pmDocToSavePlan 输出 → 我们的 SavePlan + 能力闸门
   │  └─ genoffice/               ← **复制**进来的渲染器 fork（不跨包引用）
   │     ├─ editor/               ← = packages/docx/src/editor 的整份拷贝
   │     └─ shared/               ← = packages/docx/src/shared（渲染器要的 ipc 类型）
   ├─ src/shared/…                ← 保留（engine/agent 路径与单测仍用）
   └─ src/page/                   ← 现有块列表页（P2 切换后删除）
```

**决策 A（已按此实施）：复制，不抽包。** 理由：两个 docx 插件最终只会保留一个，
不存在"长期双份 drift"问题；复制换来 `packages/docx` **零改动**（v1 可随时对照/回退）。
代价与约束：
- 渲染器补丁（P3 切分对齐、P5 part 通道）**只在 `office-docx/src/page-v2/genoffice/` 里打**，不回灌 v1；
- 日后删掉 v1 时，直接删 `packages/docx`，不涉及搬家；
- **副本上游的改动要手动同步**：真机调试时如果改了 `packages/docx/src/editor`（CSS / 交互），
  记得同步到 `src/page-v2/genoffice/editor`。目前已知的唯一一次同步：抑制 WorkStudio
  `appearance.css` 的 `:focus-visible { outline: … !important }` 打在纸张上的蓝框
  （`styles.css` 末尾 15 行，v1 与 v2 都需要）。
- 目录必须保持 `genoffice/{editor,shared}` 的兄弟关系：渲染器源码里有 `../shared/ipc` 与
  `../../shared/ipc` 两种相对引用（7 + 12 处），层级一错就解析不到。

依赖与构建：`scripts/build.mjs` 只负责 `ura/engine/lib`（esbuild），页面改由 vite 构建
（`root: 'src/page-v2'`，`outDir: dist/page-v2`，`genofficeSource()` 解析 `@genoffice/*` → `vendor/genoffice/packages/*/src`）。
`vendor/genoffice` 用仓库根已有的 `scripts/setup-genoffice.mjs`（sparse checkout 已列 `docx-engine` / `font-metrics` / `ui` 等）。

---

## 3. 五个接缝（全部改动集中在这里）

### 3.1 进：内容自取
```
ura.page.get(session) → info{ref,version,epoch,size}
ura.page.bytes(session,{start,end}) × N → Uint8Array      （现有 fetchAllBytes 逻辑直接搬）
→ @genoffice/docx-engine.parseDocx(bytes) → parsed.blocks  → blocksToPmDoc() → editor.setContent
```
比 v1 多一步 `size` 分块（我们已是 `content:'pulled'`），顺带把 v1 的 `blobUrl` 双通道去掉。

### 3.2 出：PM doc → 我们的 SavePlan（**核心**）
```ts
const gen = pmDocToSavePlan(editor.getJSON(), parsed.blocks);   // 渲染器已有
const plan: SavePlan = {
  kind: 'save',
  saveBlocks: gen.saveBlocks,                                    // 形状一致
  saveBlockIndexByDocx: Object.fromEntries(gen.saveBlockIndexByDocx),
  changedCount: gen.changedCount,
  deletedCount: gen.deletedCount,
  totalOriginal: gen.totalOriginal,
  applied: /* 本次提交的编辑数 */,
  trailingSectPrChange: false,
  ...extraParts(gen, ctx),                                       // §3.5 能力闸门
};
await ura.page.commit({ txnId, baseVersion: version, expect:{docRev, epoch}, payload:{type:'plan', plan}, origin:'page' });
```
`saveDocx(...)` 与 `window.desktop.saveDocx/…` 全部不走。

### 3.3 回执：manifest rebase
`receipt.sync.kind === 'manifest'` → 用 `applyManifest`（`shared/sync.ts`）把新 `docxIndex`/`hash` 写回
`parsed.blocks[].{docxIndex,originalXml}` **和** PM 顶层节点的 `docxIndex` attr（渲染器已有该 attr，
`extensions.ts` 的 `anchorAttrs`）。
`mismatched` → 重新 `bytes` + `parseDocx`（渲染器本来就会重解析，成本可接受）并重置 `epoch`。

### 3.4 宿主 API（`window.desktop`）
从 `packages/docx/src/page/host.ts` 整体移植，**只改 4 类**：

| v1 | v2 |
| --- | --- |
| `saveDocx/saveDocxAs/saveDocxNew(path,bytes)` | `commitPlan(plan)` → `/commit`；「另存副本」= 宿主能力或禁用 |
| `openPath / openFiles` | 禁用（一个页面只对应一个 ref） |
| `exportPdf / exportHtml / print*` | 保留：走页面内渲染 + 宿主保存对话框（不经过 plan） |
| `writeRecoveryCopy / dirty / theme / closeGuard` | 保留：`dirty → bus`、`action:save → flush`、`theme ← bus` |

### 3.5 能力闸门（**必须做**，否则用户点「插入图片」会保存失败）
`saveDocx` 的 options 覆盖 body 之外的 part（`newPictures` / `partXml` / `partBinary` / comments / notes /
numbering / styleUpserts / sectPr / watermark / inks / chartPatches）。我们的 `exec` 目前只重写
`word/document.xml`。因此：

- **P2 闸门（保守）**：`plan` 里出现 `newPictures` / `chartPatches` / `partXml` / `partBinary` / `inks` / 非空
  `comments`/`notes`/`numbering`/`styleUpserts` → 拒绝提交并提示「本版暂不支持，请在 Word 中完成」；
  同时用特性开关把这些 Ribbon 按钮置灰（§5 表）。
- **P5 打开通道**：给 `SavePlan` 加可选 `parts: Array<{path, xml?} | {path, base64?} | {path, delete:true}>`，
  `engine/zip.ts` 已有条目级重写能力，扩展到多 part 很直接；难的是**新增** part（图片：`word/media/*` +
  `document.xml.rels` + `[Content_Types].xml`），这一段单列一阶段。

---

## 4. 头号风险：**块切分必须与 exec 同源**

`pmDocToSavePlan` 里的 `docxIndex` 来自 `@genoffice/docx-engine` 的 `parsed.blocks`，而 `exec` 用我们自己的
`parseDocxBytes`。两者**在以下文档上不等价**（已核源码）：

| 结构 | docx-engine | 我们的 `parseDocxBytes` |
| --- | --- | --- |
| `w:sdt` 内容控件 | `splitSdtParts()`：一个 sdt 拆成 **N 个 block / N 个 element**（各自 `originalXml` 是片段，`sdtShell` 记壳） | 整个 `<w:sdt>` = **1 个块** |
| `w:altChunk` | 装到转换器时展开成 M 个 block（只有第一个带字节，其余 range 为空） | 1 个 `other` 块 |
| 域代码折叠 / 浮表归一 | `foldFieldCodeParagraphs` / `bandWrappedBoxesBeforeTables` 可能改动顺序 | 无 |

一旦编号错位，`{kind:'original', docxIndex}` 就会回写**错误的基线切片**——这是静默写坏文件，必须靠工程手段堵死。

**对策（按顺序）**：
1. **P0 对齐校验脚本**（`scripts/check-slicing.mjs`）：对 `tests/fixtures` + 一批真实 docx，比较
   `docx-engine.parseDocx().blocks` 与 `parseDocxBytes().blocks` 的 `docxIndex → originalXml` 一一对应率；
   输出不一致清单。
2. **P3 单一实现**：把 docx-engine 的 body 扫描结果（`scan.elements` + `splitSdtParts` + altChunk 锚点）
   抽成不依赖 DOM 的 `bodySlices(documentXml)`，放进 `office-docx/src/shared/`，**exec 与页面共用**；
   `parseDocxBytes` 的 `blocks[]` 改为按 `bodySlices` 产出（`docxIndex = element 下标`），
   `baselineSlices` / `buildManifest` / `rewriteZipPart` 不用改。
3. **加载门槛（P0 就生效）**：`bodySlices` 未统一前，文档若命中 `w:sdt` / `w:altChunk` / 修订标记，
   **以只读打开**（`tier: 'readOnly'` + 提示），绝不产生 plan。宁可只读，也不写错块。

---

## 5. 分阶段计划

每阶段都有独立可验收的产物；★ 标的是「用户可见」。

### P0 · 结构与校验（2~3 天）—— ✅ 已完成
- [x] 渲染器**复制**进 `office-docx/src/page-v2/genoffice/{editor,shared}`（363 文件 / 18.4 MB）；
      `packages/docx` **零改动**（`git status` 干净，tsc + vite build 仍通过）。
      渲染器内部相对引用（`../shared` 7 处、`../../shared` 12 处）按 `genoffice/` 兄弟目录布局解析。
- [x] `scripts/check-slicing.mjs`：esbuild 现场打包 `@genoffice/docx-engine` + 本包 `parseDocxBytes`，
      对比 `docxIndex → 切片 hash`。用法：`pnpm check:slicing [a.docx …]`。
- [x] `office-docx` vite 页面骨架（`src/page-v2/` → `dist/page-v2`，不影响现有 `dist/page`）。

**P0 实测（7 个合成夹具）**：

| 夹具 | ours | engine | 结论 |
| --- | --- | --- | --- |
| plain / runs-and-table | 3 / 4 | 3 / 4 | ✅ 对齐 |
| `w:sdt` 包**单**个块（含嵌套） | 4 / 3 | 4 / 3 | ✅ 对齐（单子节点不拆片） |
| `w:altChunk`（无转换器） | 4 | 4 | ✅ 对齐（passthrough） |
| `w:sdt` 包**多**个块（两段 / 表格+段） | 4 / 2 | **5 / 3** | ❌ 分歧（`splitSdtParts` 拆片） |

→ **门槛规则因此收紧**：不是「含 `w:sdt` 就只读」，而是「`w:sdt` 含 ≥2 个顶层块子节点」才只读；
单子节点 sdt、无转换器的 altChunk 都可以直接编辑。P3 的 `body-slices.ts` 按这个口径对齐即可。

### P1 · 页面壳与只读打开（3~4 天）—— ✅ 骨架已跑通（待真机验收）
- [x] `src/page-v2/`：`get` + `bytes`（1 MB 分块）→ blobUrl → `installDocxHost` → `import('@workstudio/docx-renderer/main')` → `bus.loaded()`。
- [x] `src/page-v2/host.ts`：从 `packages/docx/src/page/host.ts` 移植 `window.desktop` 全表；
      保存改成明确错误「只读预览：写盘通道在 P2 接入」，并**关掉自动保存**避免反复弹错。
- [x] 构建：`pnpm build:page-v2` → `dist/page-v2`（**不覆盖** `dist/page`）；`tsc --noEmit` 干净。
      `@genoffice/*` 与 `@workstudio/docx-renderer/*` 在两处 tsconfig / vite alias 里对齐。
- [x] 清单暂不改：`workstudio.page` 仍指 `dist/page/index.html`（块列表版）。切到渲染器只需一行，
      等 P2 的 plan 写盘就绪后一起切，避免用户拿到“不能保存”的编辑器。
- [ ] 真机验收：打开真实 docx（分页外观 / Ribbon / 只读），确认 worker、字体、主题、关闭守卫无异常。

### P2 · plan 写盘 + 自动保存（4~6 天）—— ✅ 代码完成（待真机验收）
- [x] `shared/plan-bytes.ts`：**块级 diff 还原计划**（关键取舍，见下）。
- [x] `page-v2/host.ts`：`saveDocxPlan` → `/commit`（`baseVersion` + `expect{docRev,epoch}`）→ 回执 → 刷新基线；
      `command.stale / stale-index / too-large / page-unavailable` 映射成人话；成功后清 dirty；无改动不产生空提交。
- [x] `page-v2/genoffice/editor/file-actions.ts`：`saveOnce` 加 plan 分支（宿主有 `saveDocxPlan` 时走它；提交后
      `readDocxBytes()` 回读，后段「重解析 + 更新 docxIndex / 分节 / 页眉页脚」逻辑原样复用）。
- [x] 自动保存恢复（`getAutoSaveDefault → on`，节奏仍由渲染器控制）。
- [x] 能力闸门 `shared/zip-parts.ts`：只读 zip 中央目录比对部件指纹；`word/document.xml` 之外任何部件变化
      （media / rels / styles / comments …）**直接拒绝并报出部件名**；`docProps/{core,app}.xml` 例外
      （引擎每次真实保存都会写修改时间 / 修订号 / 页数缓存）。
- [x] **P2b**：`agent-request` 接到渲染器（`page-v2/agent.ts`）：
      `context` = live PM 投影（与 `shared/project.ts` 同格式：`[docxIndex] label text`）；
      `apply` = 先整批校验（任一非法整批不执行）→ 文本 / 结构走 PM 事务，格式 / 选择器走渲染器
      自带 ops 引擎（`ai/ops.ts` 的 `setFont` / `setHeadingLevel` / `findReplace`）→ `__aidocs.save()`
      触发 plan 提交 → 回执 `{applied, version, epoch}`。epoch 不符 → `command.stale-index`；
      `dryRun` 只校验不派发；apply 前先把本地未保存改动 flush 掉（agent 的坐标基于磁盘基线）。
- [x] flip：`package.json` → `"page": "dist/page-v2/index.html"`（回退：改回 `dist/page/index.html`，两套产物都在）。

> **偏离原方案的地方（重要）：计划不是来自 `pmDocToSavePlan`，而是来自字节 diff。**
> 它返回的 `SaveBlock[]` 里改过的块是 `kind:'generated'`（结构化块），序列化成 OOXML 需要引擎内部的
> `GenerateContext`（样式表 / 编号 / rels 分配器），页面拿不到；只有 `original` 与 `xml` 能直接交给 `exec`。
> 因此：渲染器照旧跑它的 `saveDocx`（**它本来就在跑**），页面把「新 document.xml 顶层切片」与「基线切片」
> 按哈希顺序匹配（`shared/plan-bytes.ts`）：命中 → `{kind:'original'}`（exec 原字节回写）；未命中 → `{kind:'xml'}`。
> 未命中的块再按「身前还有没被消耗的基线块」区分**改写**与**新增**，剩余基线块才是**删除**。
>
> 代价：页面多跑一趟 `saveDocx`（只为拿块的最终 XML）；收益：`exec` 零改动、零依赖引擎内部上下文，
> 且**字节等价性被单测钉住**（`tests/plan-bytes.test.mjs` 端到端：计划 → exec → 重解析，块序 / 文本与引擎产出一致，
> 未改块逐字节等于基线切片）。P5 之后可改为 `generated → xml` 直出，去掉这趟重复序列化。

### P3 · 切分同源（3~5 天）—— ⛔ 已取消（P2 换了计划来源后不再需要）
> 原风险：页面用 docx-engine 的 `docxIndex`，与 `exec` 的切片编号可能错位（多块 `w:sdt` 拆片、
> `altChunk` 展开）。P2 改成「基线切片 ↔ 产出切片」的**字节 diff** 之后，两侧都用本包的
> `parseDocxBytes`，`{kind:'original', docxIndex}` 指的就是同一套切片 —— 错位不可能发生。
> `scripts/check-slicing.mjs` 保留作回归探针（现状：5 对齐 + 2 已知分歧，均不影响写盘）。
>
> 原计划（留档）：
- [ ] `shared/body-slices.ts`：sdt 拆片 / altChunk 锚点 / 域折叠的对齐实现；`parseDocxBytes` 改用它。
- [ ] 解锁含内容控件的文档（读写）；`check-slicing` 全绿。
- 验收：含 sdt / 多节 / 修订的文档可编辑保存，未改块逐字节不变。

### P4 · 工具栏按组开放（3~5 天）—— ✅ 布局/表格/域已验，其余按部件拒绝并解释
- [x] **布局组**（纸张 / 页边距 / 分栏 / 分节符 / 页码格式）：尾部 `w:sectPr` 就在
      `word/document.xml` 里，走**块级 xml**，执行点落盘即生效 —— 已用「计划 → exec → 重解析」
      钉住（纸张改 12240 / 加 `w:type`+`w:cols` 两个用例），其余块逐字节不变。
- [x] **表格组**：表格结构改动同样是 body 块（加行用例已钉住），执行点不需要任何扩展。
- [x] **字体 / 段落 / 样式（用现有样式）/ 剪贴板 / 撤销 / 域 / 公式 / 符号**：P2 起即可落盘。
- [x] **保存被拒时给功能名**：`shared/zip-parts.ts` 的 `describePart` 把部件翻成
      「页眉 / 页脚 / 图片 / 媒体 / 编号 / 列表定义 / 样式表 / 批注 / 主题 / 文档设置 …」，
      错误文案形如：
      > 本次修改涉及正文以外的内容：页眉（word/header1.xml）、图片 / 媒体（word/media/image1.png）。
      > 本版只写 word/document.xml，请在 Word 中完成这类修改（未写入磁盘）。
- [x] **新增执行点硬校验**：计划最后一块必须是尾部 sectPr（`original(sectPr)` 或 `xml` 且根名
      `sectPr`），否则 `command.invalid` —— 防止页面 / agent 送来的计划把分节信息挪到中间或丢掉。
- [x] P5 已把这些全部打通（见下）：`numbering / comments / footnotes / theme / settings / media` 都作为 part 补丁落盘。
- [ ] 可选优化：保存被拒后一键「放弃本地修改并重新载入」（渲染器已有 `onOpenDocx` 通道，
      页面可推一份磁盘字节让它整篇重载）。

#### 原计划的 P4 分组表（现状）
- 逐组打开 Ribbon 功能，每组都要求「保存能落盘」：

| Ribbon 组 | 依赖 | 状态 |
| --- | --- | --- |
| 字体 / 段落 / 样式 / 剪贴板 / 撤销 | body blocks + PM marks | **P2 即可用** |
| 插入：表格 / 分隔符 / 页码域 / 公式 / 符号 | body blocks | ✅ 已验（表格加行） |
| 布局：页边距 / 纸张 / 分栏 / 分节 | `sectPr`（在 body 尾块内） | ✅ 已验（P4） |
| 插入：图片 / 形状 / 文本框 | `newPictures` + `partXml` + rels | ✅ P5（整包通道） |
| 引用：目录 / 题注（域在 body） | body blocks | ✅ 预期可用 |
| 引用：脚注 / 尾注 | notes parts | ✅ P5（整包通道） |
| 审阅：修订（w:ins/w:del 在 document.xml） | body blocks | ✅ 预期可用 |
| 审阅：批注 | comments parts | ✅ P5（整包通道） |
| 设计：主题 / 水印 / 页面颜色 | theme/fonts/settings parts | ✅ P5（整包通道） |
| 图表 / 墨迹 / 邮件合并 / AI | 各自 part + 外部能力 | **P6 或不做** |

### P5 · exec 的整包 part 通道（5~10 天）—— ✅ 已完成（图片 / 页眉页脚 / 编号 / 批注一次打通）
- [x] `SavePlan.parts`：`{path, xml} | {path, base64} | {path, delete:true}`（JSON 安全：XML 走文本，二进制走 base64）。
- [x] `engine/zip.ts`：`rewriteZipParts(bytes, patches)` —— 替换 / 新增 / 删除多条目；
      **未涉及的条目仍然逐字节拷贝**（local record 与 CD record 只改 `localOffset`），
      新增条目按 path 排序追加，EOCD 条目数按最终结果写。`rewriteZipPart` 保持「必须已存在」的严格语义。
- [x] `engine/exec.ts`：正文块 + parts 一次写完；`audit.changedParts` 报数。
- [x] `shared/plan.ts`：`assertPartsShape` —— 路径安全（拒绝绝对路径 / `..` / 反斜杠）、
      唯一性、`xml/base64/delete` 三选一、`word/document.xml` 不许走 parts；
      预算 `MAX_PART_BYTES = MAX_COMMIT_BYTES - 6MB`、`MAX_PARTS_BYTES = MAX_COMMIT_BYTES - 4MB`（提交上限 16 MB）；
      仍要求最后一块是尾部 sectPr。
- [x] `shared/zip-parts.ts`：`diffZipParts`（replace / add / remove）+ `readZipEntries`（只解压需要的条目）+ `isTextPart`。
- [x] `shared/plan-bytes.ts`：**P4 的「拒绝」改成「携带」** —— 产出里变化过的条目直接进 `parts`，
      图片 / rels / 页眉 / 编号 / 批注 / 主题 / 文档属性全部走同一条路。
- 结果：**渲染器能正确序列化的东西，我们都能落盘**（因为我们拷贝的就是引擎自己产出的那几段字节）。
  剩下的限制只有：单次提交 16 MB 上限、zip64 / 加密文档拒绝。

#### 原计划（分段）
- [ ] `SavePlan.parts`：XML part 替换（sectPr 之外的 numbering/styles/settings/theme/comments/notes）。
- [ ] 新增 part（图片）：`word/media/**` + `word/_rels/document.xml.rels` + `[Content_Types].xml` 三处一致；
      沿用 `zip.ts` 的「未改条目逐字节拷贝」哲学，新条目单独写入。
- [ ] `inks` / `chartPatches` / `partBinary`（xlsx 图表数据）——最后做，或明确列为不做。

### P6 · 收尾（3~5 天）
- [ ] 体积：按需 `import()` 对话框；关掉 `pdf-lib/pdfjs/opentype` 等只在导出用的依赖（vite 分块）。
- [ ] i18n / 主题 / 快捷键表 / 右键菜单 / 查找替换 回归。
- [ ] 文档：README §1/§3/§7/§8 重写（页面层不再是「块列表」），`catalog` 与工具面按新能力增补。
- [ ] 清理：删 `src/page/editor.ts`、`styles.css`、`index.html`、`shared/blocks-edit.ts`（**保留**：无页面 agent 路径仍用）。

总工期（单人）：**P0–P4 ≈ 3 周**（此时「工具栏照抄 + 多段落 + run 级格式 + 分页 + 表格」全部可用，图片/脚注/批注按钮置灰）；
**P5 ≈ +1~2 周**（图片、脚注、批注、主题落地）。

---

## 6. 与现有实现的取舍

| 现有 | 迁移后 |
| --- | --- |
| `src/page/editor.ts`（652 行：块列表 + 分帧挂载 + run 上色） | 删除（渲染交给渲染器；那套分帧/增量 patch 由 `phased-content.ts` + PM 接管） |
| `shared/blocks-edit.ts`（合并 / 选区替换 / 粘贴分段） | **保留**：`engine.plan`（无页面 agent 路径）与单测仍用 |
| `shared/xml-ops.ts` 的 run 级函数（`runSpans`/`applyRunStyle`/`replaceRunText`） | **保留**：agent 的区间命令仍走它；页面路径由 PM marks 覆盖 |
| `shared/{plan,sync,ops}` | 不变（`plan` 的 `SavePlan` 形状被 `plan.ts` 复用） |
| `page/main.ts` 的提交循环（flush/冲突/回执） | 搬到 `page/session.ts`，逻辑不变 |
| `tests/*`（56 项） | 全部保留；新增 `page-plan.test.mjs`（`pmDocToSavePlan` 输出 → 我们的 `exec` 能跑通）与 `check-slicing` |

---

## 7. 验收清单（迁移完成定义）

1. 打开一个真实 docx：分页外观、页眉页脚、表格、图片、目录、批注按 docs 行为呈现。
2. 工具栏与 docs 完全一致（含 tab、组、按钮启用态）；不出现「点了保存失败」的按钮。
3. 编辑 → 保存：`exec` 产出的字节中，**未改块切片逐字节不变**（对比 v1 保存结果）。
4. 保存后 `manifest` rebase 生效：连续两次提交不会写错块（沿用 `roundtrip.test.mjs` 的断言方式在页面路径复测）。
5. agent `context/apply` 可用，且 `stale / stale-index / stale-index(epoch) / syncing` 语义与现在一致。
6. 大文档：`tier` 分级 + 分阶段装载生效；2000 页文档只读不崩。
7. 全仓 `pnpm test` / `pnpm typecheck` / `verify` 绿。
