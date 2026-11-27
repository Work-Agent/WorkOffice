# xlsx 改造方案（对齐 WorkStudio 0.4 / SDK 0.3 插件框架）

对象：`packages/xlsx-pro`（URA apiVersion 2 + `plan` 写）→ `packages/xlsx`（`@work-studio/sdk` 插件）。

结论先说：

- **身份换掉，能力不换**：4 个 agent 命令、56 条 op、完整 Sheets 编辑器、无页面/有页面两条写路径全部保留。
- **GenOffice 原生代码不进仓库副本**：`vendor/genoffice` submodule 直取，包名直接 `import '@genoffice/...'`（不再包一层 barrel）；未改造的页面文件直接从 submodule 编译。
- **只有我们要改的页面文件才 fork**：fork 用一个可验证的清单（`editor.manifest.json`）锁死，其余一律回落到 vendor。
- **通讯从"会话 + 总线 + 计划"改成"命令 + 事件租约 + 哈希 CAS"**：`ura.page.*`、`window.extension.bus`、`agent-request`、`plan {kind:'save'|'bytes'}` 全部删除。
- **范围是单个文件**：不枚举目录、不做跨文件合并、不做批量转换（见 §0）。
- **宿主零改动**：WorkStudio 0.4.0 原样使用。页面命令用「单命令 + 租约令牌」自保，脏态用自动保存兜住（见 §10）。

---

## 0. 范围：单个文件

插件的输入面是**一个 `.xlsx` 文件**，不是目录、不是工作簿集合：

| 项 | 规则 |
| --- | --- |
| ref | 必须是 `file://` 且指向**普通文件**；`http(s)://`、`plugin://`、目录一律拒绝 |
| 扩展名 | `.xlsx` / `.xlsm` / `.xltx`（与 `claims.extensions` 同集） |
| `resources.xlsx.ls/find` | 恒返回 `[]`（不枚举目录，不进资源树） |
| `resources.xlsx.read` | 单条 `Item`：`{locator:'xlsx:file', ref, name, size, hash, mtime, tier}` |
| `resources.xlsx.grep` | 恒返回 `[]`；内容检索走 `xlsx.read {what:'find'}` |
| 写 | 单文件整份写；没有多文件批，没有目录级操作 |
| 新建 / 另存 | 单文件；必须落在当前工作区内，父目录必须已存在 |

因此以下能力**不做**：打开文件夹、跨文件合并（`merge-workbooks`）、批量转换、多选导入、目录监听。
页面里对应的 Ribbon 入口与 `desktopApi.selectWorkbook(s)` 在 fork 中删掉或返回 `null`。

---

## 1. 契约映射（旧 → 新）

| 旧（xlsx-pro，URA apiVersion 2）                      | 新（xlsx，SDK 0.3）                                    | 说明 |
| ----------------------------------------------------- | ----------------------------------------------------------- | ---- |
| `workstudio.id` / `label` / `icon`                    | 同名                                                        | 不变 |
| `apiVersion` / `write: 'plan'` / `content` / `model`  | **删除**                                                    | 新清单没有这些字段 |
| `workstudio.page: 'dist/page/index.html'`             | `workstudio.pages.content`                                  | 字段名变化 |
| `workstudio.engine.{tools,plan,exec,read,runtime}`    | `exports['./workstudio'] → dist/engine.js`（`definePlugin`）| 五个模块塌缩成一个入口 |
| `workstudio.skills: [{name, dir}]`                    | `workstudio.skills: ['skills/xlsx']`                        | 名称/描述改由 SKILL.md frontmatter 定义 |
| （无）                                                | `workstudio.claims: [{extensions:['xlsx','xlsm','xltx'], priority:100}]` | 取代 codec 的 MIME 认领 |
| `registerCodec`（MIME → 快照 + anchorMap）            | `resources.xlsx.read()/grep()` → `Item{locator}`            | codec 概念消失；读投影即资源适配器 |
| `registerSidecar`（agent 文本）                       | 同上 `read/grep` + `xlsx.read` 命令                     | 文本由 `Item.text` 承载，locator 保留可寻址 |
| `registerStore`（store.query）                        | **删除**                                                    | 需要表格化查询时用 `read` 命令参数，不建第二套模型 |
| `registerPresenter`（page + sessionId + data）        | `pages.content` + `?ref=`                                   | 页面不再收"会话数据包" |
| `content:'pulled'` + `ura.page.bytes` 分块            | `xlsx.page {op:'chunk'}`                                | 字节仍分块，但走命令通道 |
| `ura.page.commit({type:'plan', plan})`                | `xlsx.page {op:'save'}`（`bytes` + `baseHash`）         | 写=命令；平台层不再有"计划" |
| `ura.page.invoke(id, 'xlsx_read', …)`（页面结构化读） | `xlsx.page {op:'meta'|'range'}`                         | 页面白名单自动放行自己的域 |
| `ura.page.agentResult` / `extension.bus agent-request`| `ctx.events.emit('request')` + `pending`/`ack` + 租约        | 见 §5 |
| `extension.bus.send({type:'dirty'|'theme'})`          | `report` 命令 / `subscribe('shell','appearance')`             | 见 §5.1 |
| `EngineCtx.{budget, sessionDir, tempDir, importEpoch}`| `PluginContext` + `PluginScope.tempDir`                      | 预算改插件自管常量 |
| `codec.encode()` 抛 `command.forbidden`               | **不存在**                                                  | 平台层没有整份编码通道 |

---

## 2. 依赖策略：submodule 直取，只 fork 改过的页面文件

### 2.1 三层来源

```text
vendor/genoffice                        ← git submodule，钉 commit
  packages/xlsx-gateway/ …              ← 引擎 / 网关 / DSL 与基础件（不复制）
  apps/sheets/src/{renderer,shared}/    ← 页面层原版（不复制，作基线 + 未改造文件的来源）
  apps/sheets/native/xlsx-engine/       ← wasm 引擎源码（不复制，构建时 overlay ABI）

packages/xlsx/build/sheets/src/    ← 构建期物化视图（gitignore）
  renderer/  shared/                    ← vendor 全量 + 下面两层覆盖

packages/xlsx/src/editor/          ← 只放「我们改过的」renderer 文件（125 个）
packages/xlsx/src/editor-shared/   ← 与渲染层同版本的 shared 类型（desktop-api / ipc-channels）
```

解析规则：

```text
@sheets/<path>            →  build/sheets/src/renderer/<path>      （物化视图，相对 import 可用）
@xlsx/<path>              →  src/<path>                            （fork 文件引用插件自有模块）
@genoffice/<pkg>[/<sub>]  →  vendor/genoffice/packages/<pkg>/src/<sub>
```

**为什么物化，而不是逐文件回落**：renderer 内部全是相对 import（`./App`、`../shared/desktop-api`）。
只把裸说明符 `@sheets/*` 指向 fork，会让 fork 的 `main.tsx` 找不到没有 fork 的 `./App`。
物化一棵完整目录树后，fork 只需提供「改过的文件」，其余一律来自上游。
代价是每次构建多一次 7 MB 目录复制（约 100 ms）。

- 物化：`scripts/sync-renderer.mjs`（`build/` 在 `.gitignore`）；
- 解析：`scripts/genoffice-resolve.mjs`，esbuild / vite / tsconfig 三处共用；
- sparse-checkout 需要 `/apps/sheets/src/renderer/`、`/apps/sheets/src/shared/`、`/apps/sheets/native/xlsx-engine/`，已写进 `vendor.manifest.json`。

### 2.2 fork 清单（`editor.manifest.json`）

`src/editor` 的每个文件都必须和上游「有差异」，否则就是陈旧 fork：

```bash
node scripts/editor-diff.mjs           # 重写清单
node scripts/editor-diff.mjs --check   # CI：清单外有差异 / 清单内文件与上游一致 → 失败
```

规则：

1. `src/editor`、`src/editor-shared` 的每个文件都必须在 `forked[]` 里（shared 用 `shared/` 前缀）；
2. 与上游逐字节相同的文件必须删除，回落上游，否则 `--check` 失败；
3. fork 文件不许出现宿主协议细节：`window.workStudio` 只在 `src/page/**`，
   渲染层只碰 `window.__xlsxPlus` 与 `window.__xlsxProBoot`；
4. `verify.mjs` 扫描 `src/**`，禁止旧 URA 面（`window.extension` / `ura.page` / `agent-request` / `registerCodec`）。

当前 125 个 fork 文件的分组：`App.tsx`（去 AI + 宿主接缝）、`workstudio-host.ts` /
`workstudio-save.ts`（页面 I/O，全部改到 `__xlsxPlus`）、`shims/`（浏览器侧 node 兼容）、
`i18n/*`（上游新增键未跟）。其余 245 个渲染层文件、全部 `shared/` 类型与所有
`@genoffice/*` 包都直接来自 vendor。

#### 为什么不继续削减 fork（实测）

已核对：`packages/xlsx/src` 里 **没有任何**非 editor 文件与 vendor 逐字节相同（全量 hash 比对），
`@genoffice/*` 的 import 只出现在具体实现文件里（不再包一层 barrel）—— 也就是说
「复制上游代码」这条已经清零，剩下的只有「改过的页面文件」。

又试过一次激进削减：把 125 个 fork 里 87 个「正文没有宿主接缝字样」的文件移回 vendor，
让上游版本接管。结果 48 个类型错误，全部来自上游漂移而不是接缝本身：

- 上游 `ExcelShell`/`App` 重新引入 AI 面板，fork 已删；
- `DesktopApi` 的方法被上游改名/移除（`listExternalLinks`、`armWorkbookPick`…）；
- `SaveTickState` / `PivotActionContext` 结构变了；
- i18n 键被上游改名（`appComboNoSelectData`、`appTableFilterNoEdit`、`appSaveErr*`）。

所以 **125 就是这个 vendor pin（f6c2358c）下的最小安全集**。要真正缩小，只有一条路：
一次性把 fork 基线升到当前上游（AI 移除、i18n、desktop-api 同一批做掉），再跑
`fork-diff --write`；逐文件回落会被上游漂移拖垮。

另外 `scripts/setup-genoffice.mjs` 的 sparse 列表已补上
`apps/sheets/src/renderer` 与 `apps/sheets/src/shared`：前者是物化视图的来源，
后者是渲染层 `../shared/*` 类型的来源，缺一个 `pnpm setup` 后都构建不了。

### 2.3 wasm 引擎

保持现有做法（已经是"引用 + overlay"，不是复制）：

- `rust/xlsx-engine/wasm.rs` 是**接入层 ABI 补丁**，始终 overlay 到 vendor 源码副本上编译；
- 源码路径只认 `vendor/genoffice/apps/sheets/native/xlsx-engine`（`--source` 仅作应急覆盖）；
- 产物 `dist/xlsx-engine.wasm` + `.json`（sha256/体积/导出面），`verify.mjs` 校验导出面与体积。

页面侧不再有"页面内 gateway 产字节"这条路径，因此 `workstudio-save.ts` 里的 `assembleWithJsZip` 调用删除；页面只保留 `planWorkbookSave`（请求映射），字节只由引擎产生。

---

## 3. 包结构与清单

```text
packages/xlsx/
├─ package.json                  workstudio 清单（新字段）
├─ commands.json / api.json / api.d.ts   ← workstudio-plugin 生成，不手写
├─ vendor.manifest.json          vendor commit / sparse / wasm overlay
├─ editor.manifest.json          editor fork 白名单（生成 + review）
├─ skills/xlsx/SKILL.md
├─ scripts/
│  ├─ genoffice-resolve.mjs      解析规则（唯一）
│  ├─ fork-diff.mjs              fork 清单生成 / 校验
│  ├─ build.mjs                  esbuild(engine) + vite(page) + buildPlugin()
│  ├─ build-wasm.mjs             现脚本，改读 vendor.manifest.json
│  └─ verify.mjs                 清单/产物/往返自检
├─ src/
│  ├─ editor/                    fork 的页面文件（125 个）
│  ├─ editor-shared/             与渲染层同版本的 shared 类型
│  ├─ engine/                    Node 侧（definePlugin）
│  │  ├─ index.ts                workstudio = definePlugin({...})
│  │  ├─ resource.ts             resources.xlsx（ls/find/read/grep/open）
│  │  ├─ ref.ts                  file:// + 扩展名 + 体积守卫
│  │  ├─ state.ts                租约 + pending + docRev + dirty
│  │  ├─ lock.ts                 同 ref 写串行化
│  │  ├─ gateway.ts              planCellEditsToXlsx / recalc 封装
│  │  ├─ read.ts                 元数据 / 范围 / 公式 / 查找 / 审计（字节侧）
│  │  ├─ wasm.ts                 node:wasi reactor（几乎原样搬）
│  │  ├─ sessions.ts             (ref, hash) → wasm session，LRU 2
│  │  └─ commands/               page.ts（页面通道单入口）/ guide.ts / read.ts /
│  │                             create.ts / apply.ts / resource.ts
│  ├─ shared/                    引擎与页面共用契约（types / op-tiers / chunks）
│  ├─ tools/catalog.ts           56 条 op 目录（从 xlsx-pro 原样搬，无宿主依赖）
│  ├─ genoffice/                 中转 barrel（engine/page/ui/i18n）
│  ├─ fork/                      「必须改」的页面文件
│  └─ page/                      页面专属：boot / desktop-api / agent / load
└─ tests/                        node --test：命令、往返、令牌、分块、wasm
```

`package.json`：

```jsonc
{
  "name": "@workstudio/plugin-xlsx",
  "version": "0.2.0",
  "type": "module",
  "exports": { "./workstudio": "./dist/engine.js" },
  "workstudio": {
    "id": "xlsx",
    "label": "Excel 表格",
    "icon": "table",
    "pages": { "content": "dist/page/index.html" },
    "claims": [{ "extensions": ["xlsx", "xlsm", "xltx"], "priority": 100 }],
    "skills": ["skills/xlsx"]
  }
}
```

> id 统一为 `xlsx`：新框架按 `claims.priority` 竞争，不比 MIME。旧 `xlsx-pro` 已删除，不再保留回退路径。

---

## 4. 命令面

### 4.1 agent 可见（`expose: { agent: true }`，共 4 条，与旧工具一一对应）

| 命令 | 对应旧工具 | 说明 |
| ---- | ---------- | ---- |
| `xlsx.guide` | `xlsx_guide` | 命令目录（无参/分组/单条），静态，不占页面往返 |
| `xlsx.read` | `xlsx_read` | `context / sheets / range / formulas / find / audit`；有租约读页面模型，无租约读磁盘 |
| `xlsx.create` | `xlsx_create` | 目标 ref 必须不存在；必须是工作区内单文件，**父目录必须已存在**（不建目录） |
| `xlsx.apply` | `xlsx_apply` | ≤200 条；有租约 → 事件 + `pending/ack` 交页面执行；无租约 → 网关展开后写盘；`objects` 需页面 |
| `xlsx.create` | `xlsx_create` | `sheets/rows/csv`；目标必须不存在、父目录已存在 |

`args` 规范（`commands.json` 即由此生成）：

```ts
'xlsx.apply': {
  args: {
    ref: { type: 'string', required: true },
    commands: { type: 'string', required: true },   // JSON 字符串，与 docx 一致
    baseHash: { type: 'string' },                    // 磁盘 CAS
    docRev: { type: 'number' },                      // L1 读内容型
    structureEpoch: { type: 'number' },              // L2 绝对坐标
    dryRun: { type: 'boolean' },
  },
  expose: { agent: true },
}
```

### 4.2 页面通道（`agent: false`，单条命令 + 租约令牌）

页面命令收成**一条** `xlsx.page`：助手命令表只多一行，授权只有一条规则，
所有 op 共用同一套令牌校验 —— 这是「不改宿主」的关键（见 §10）。

| `op` | 入参 | 返回 | 替代旧接口 |
| --- | --- | --- | --- |
| `status` | — | `{size, hash, mtime, tier}` | presenter data |
| `open` | — | `{size, hash, tier}`（不传字节） | `ura.page.get` |
| `chunk` | `{hash, offset, length}` | `{bytes, next, eof}` | `ura.page.bytes` |
| `meta` | `{baseHash}` | 引擎 `open` 的元数据 | `invoke(what:'meta')` |
| `range` | `{baseHash, sheetId, range}` | wasm `RangeResult` 原样 | `invoke(what:'cells')` |
| `save` / `saveBegin` / `saveChunk` / `saveCommit` | `{bytes}` / `{transferId, seq, bytes}` / `{transferId, size, sha256}` | `SaveReceipt` | `ura.page.commit(plan)` |
| `saveAs` | `{bytes, token}` | `SaveAsResult` | `saveBytes` |
| `attach` / `detach` | `{pageId, takeover?}` | `{token, hash, docRev, pending[]}` | `agent-request` 生命周期 |
| `pending` / `ack` | `{id}` / `{id, ok, payload}` | 请求正文 / 结算结果 | `agent-request` / `agentResult` |
| `report` | `{hash?, docRev?, dirty?, selection?}` | 租约续期 | 无 |
| `pending` / `ack` | `{id}` / `{id, ok, payload}` | agent 请求正文 / 页面回执 | `agent-request` / `agentResult` |
| `draft` / `draftStatus` / `draftLoad` / `draftClear` | 恢复副本读写 | 崩溃恢复 | 无 |

```ts
'xlsx.page': {
  summary: '页面通道（内部）：status/open/chunk/meta/range/save/saveAs/attach/pending/ack/report/detach',
  args: {
    ref: { type: 'string', required: true },
    pageId: { type: 'string', required: true },
    token: { type: 'string' },   // attach 之后所有 op 必填
    op: { enum: ['status','open','chunk','meta','range','save','saveBegin','saveChunk','saveCommit','saveAs','attach','pending','ack','report','detach'] },
    input: { type: 'string' },   // JSON
  },
  expose: { agent: false },
}
```

守卫（插件侧替代 `ctx.origin`）：

1. 只读 op（`status` / `open` / `chunk` / `meta` / `range`）不要求令牌 —— 页面在 `attach` 之前就要用它们建模型；
2. `attach` 发放随机 `token`；已有未过期租约且 `pageId` 不同 → `lease.busy`（页面 reload 用 `takeover:true` 才替换）；
3. 写与结算 op（`save` / `saveAs` / `saveBegin` / `saveChunk` / `saveCommit` / `report` / `ack` / `pending` / `detach`）必带 `token`，不匹配 → `denied`；
4. `save` 额外校验 ZIP 魔数 / 必需 part / 体积 / `baseHash`，不合法一律拒；
5. `pending` / `ack` 的 `id` 必须属于当前 `pageId`。

额度沿用：`≤200` 命令、`≤2,000` 展开格、页面提交 `≤16 MB`（超出走分块 save）、wasm 文件 `≤150 MB`、TS 回退 `≤64 MB`。

---

## 5. 通讯（改造重点）

### 5.1 通道总表

| 方向 | 通道 | 用途 |
| ---- | ---- | ---- |
| 页面 → 引擎 | `command('xlsx','xlsx','page',{ref,pageId,token,op,input})` | 所有页面动作（status/open/chunk/meta/range/save/attach/report…） |
| 引擎 → 页面 | `subscribe('xlsx','events')` + `ctx.events.emit('request',{id,type})` | 把 agent 请求交给页面执行 |
| 引擎 → 页面 | `page {op:'pending', input:{id}}` | 页面取请求正文（事件只带 id/type，不塞大 payload） |
| 页面 → 引擎 | `page {op:'ack', input:{id,ok,payload}}` | 页面回执（`applied` / `receipt` / `data` / `error`） |
| 页面 → 引擎 | `page {op:'report', input:{hash,docRev,dirty,selection}}` | 租约续期 + 页面状态（15 s 心跳 / visibility / focus） |
| 页面 → 宿主 | `subscribe('shell','appearance')` + `onUpdate` | 主题 / 缩放（白名单已开） |
| 页面 → 宿主 | `command('shell','native','saveFile',{name})` | 另存为对话框 → `save:` token |
| 引擎 → 宿主 | `ctx.changed(ref,'content')` | 落盘后广播给 `workspace/<ref>` 订阅者（资源栏 + 页面） |
| 删除 | `window.extension` / `ura.page.*` / `ws:bus` / `agent-request` | — |

页面能调什么由宿主鉴权决定，`xlsx` 只需记住三条：

1. 自己的域全放行（`xlsx/*`）；
2. 宿主白名单只有 `shell/native.saveFile`、`shell/native.openExternal`；
3. 能订阅 `xlsx/*`、`workspace/<自己 ref>`、`shell/appearance`。

宿主没有的通道（脏态、关窗确认、语言）一律不依赖，改用插件侧机制兜住（见 §10）。

### 5.2 打开（页面 = 唯一编辑器）

```text
宿主 → plugin://xlsx/content?ref=file:///w/预算.xlsx
页面 pageId = crypto.randomUUID()
  ├─ page {op:'status'}        → {size, hash, tier}
  ├─ subscribe shell/appearance → {theme…}
  ├─ subscribe xlsx/events → snapshot
  ├─ open 循环：
  │    page {op:'open'}                        → {size, hash, tier}
  │    page {op:'meta',  input:{baseHash}}     → sheets/styles/…（几 KB）
  │    page {op:'range', input:{baseHash, sheetId, range}} × N（可视区/已加载区）
  ├─ 用 meta + range 重建 file 模型（不再解析整包）
  ├─ page {op:'attach', input:{pageId}}        → {token, hash, docRev, pending[]}
  └─ 对 pending 逐条走 §5.4
```

- 页面**不读整包字节**：编辑器只需要 `meta + range`；只有"另存为/崩溃恢复/内联媒体"这类需要原字节的场景才 `open/chunk`。
- 分块约定：`open` 返回 `size/hash`；`chunk` 必须带 `hash`，引擎按 `(ref,hash)` 定位已打开文件，`offset` 递增，`eof` 结束；哈希不一致 → `stale`。
- 大文件降级：`size > 64 MB` 时 `meta` 只返回 sheet 清单 + used range，编辑器进 `readOnlyOversized`（沿用现有 tier 逻辑）。

### 5.3 人保存（编辑器 → 磁盘）

```text
编辑器 saveWorkbookEdits(request)
  → planWorkbookSave(request, file)              // sheetId→sheetName，字段全映射
  → page {op:'save', input:{ref, bytes?, baseHash: currentHash, type:'page'}}
        └─ bytes ≤ 16 MB：一次带 bytes
        └─ bytes > 16 MB：saveBegin → saveChunk×N → saveCommit
  → 引擎 withLock(ref)：
       read(ref) → sha256 → 与 baseHash 比对（不一致 → stale）
       write(ref, bytes) → changed(ref,'content')
       → SaveReceipt {ref, hash, size, at, type}
  → 页面：currentHash = receipt.hash；report({dirty:false, hash})
         → 重取 meta（文件可能结构变了）；range 缓存清空
```

- **只有一条写路径**：页面不再产出/提交"计划"，也不再自己 `assembleWithJsZip`；`shouldUseSavePlan`/`save-plan-budget.ts` 删除。
- `applyCellEditsStaging`（save-edits 分块）保留：它是编辑器大编辑批的上行通道，只是终点从 `commit` 改为 `saveChunk`。

### 5.4 agent 写（有页面 / 无页面）

有页面（租约在）：

```text
command xlsx.apply {ref, commands, baseHash, docRev, structureEpoch}
  ├─ leaseOf(ref) 命中
  ├─ baseHash ≠ lease.hash → stale（先 read 再写）
  ├─ 命令合法性与批上限校验（引擎侧，不信任调用方）
  ├─ enqueue → ctx.events.emit('request', {id, type:'apply'})
  └─ 等 ack（30 s 超时）
       页面：pending{id} → 校验 docRev/structureEpoch（L1/L2，op 级）
             → planFromOps + applyChangePlan（一个撤销项）
             → save（同 §5.3）→ ack{applied, receipt{hash,size,applied}}
             → report{dirty:false, hash, docRev}
  → 回 {applied, mutated, receipt}
```

无页面（无租约）：

```text
command xlsx.apply {ref, commands, baseHash, dryRun}
  ├─ 无租约 → 引擎自己展开：buildSaveRequest(view, ops) → planCellEditsToXlsx
  ├─ objects 组 → command.page-unavailable（不回退第二套模型）
  ├─ dryRun → 只回 touched entries，不写
  └─ 写：wasm save_archive（缺失则 JSZip）→ 校验未涉及条目不变
        → withLock + baseHash CAS → ctx.fs.write → changed(ref,'content')
```

**两套新鲜度分工，不要混**：

| 层 | 令牌 | 解决什么 |
| --- | --- | --- |
| 磁盘 | `baseHash`（sha256）+ `withLock(ref)` | 并发写、外部改动、重放 |
| 页面模型 | `docRev`（事务计数）/ `structureEpoch`（结构修订） | L1 读内容型 / L2 绝对坐标（沿用 `shared/op-tiers.ts`） |

### 5.5 agent 读

```text
command xlsx.read {ref, what}
  ├─ 有租约 → enqueue{type:'read'} → 页面 pending → ack{data}
  │    页面只实现 context/meta/sheets/range/formulas/find/audit
  │    其它 what 回 command.unsupported → 引擎落回磁盘/wasm
  └─ 无租约 → 磁盘解析（≥16 MB 走 wasm 结构化读，否则 TS 解析）
```

读结果首行仍回令牌（`hash=… docRev=… structureEpoch=…`），否则 agent 拿不到写所需的新鲜度。

### 5.6 租约与失败语义

| 项 | 值 | 说明 |
| --- | --- | --- |
| 租约 TTL | 5 min | 心跳 `report` 15 s；后台标签 timer 被节流不误判 |
| 页面令牌 | `attach` 发放，租约内唯一 | 其余 page op 必须匹配；助手拿不到令牌，调不动写 |
| 令牌冲突 | 已有租约 + 不同 `pageId` → `lease.busy` | 只有页面 reload 用 `takeover:true`；助手抢不到 |
| pending 超时 | 30 s | 超时即拒绝，页面回 `ack` 时回 `not_found`（忽略） |
| 页面替换 | 同 ref 新 `pageId` attach | 旧 pending 全部作废，报 `页面已更换，请求作废` |
| 保存失败 | 页面回滚本批（撤销栈深度守卫） | 与旧 `applyOnPage` 同策略，避免"页面已脏、磁盘未变" |
| 并发写 | 引擎锁 + `baseHash` | 用户保存与 agent 写互斥；后者哈希过期要求重读 |
| 错误码 | `command.invalid` / `stale` / `page-unavailable` / `stale-index` / `too-large` | 保持旧码，SKILL.md 的处置说明不变 |

---

## 6. 页面改造

### 6.1 fork 文件与改法

| fork 文件 | 改什么 | 不改什么 |
| --- | --- | --- |
| `main.tsx` | boot：`?ref=`、`status`、`subscribe`、`installDesktop`；删 `__xlsxProBoot`/`__xlsxProHost` | 字体加载、主题应用、ReactDOM 渲染 |
| `App.tsx` | `window.__xlsxPro` 桥改 `__xlsxPlus`（context/readCells/applyOps/save/docRev/structureEpoch/undo）；删 MCP 桥 | Univer 初始化、Ribbon 布局 |
| `ExcelShell.tsx` | 标题/菜单/快捷键入口挂到 `shell` 命令（打印、导出、另存） | 面板布局 |
| `ribbon-actions.ts` / `ribbon-icons.tsx` | 依赖 `desktopApi` 的入口改由 `src/page/host` 提供 | 图标与分组顺序 |
| `host/**`（新文件） | `window.desktopApi` 兼容层：把 40 余个调用面映射到 `xlsx.*` 命令 | — |
| `page/agent.ts`（页面侧，新文件） | 租约 + pending/ack + 令牌校验 + 保存 + 回滚 | — |
| `page/boot.ts`（新文件） | `open/meta/range` 与 `save` 的实现，替代 `workstudio-host.ts` 的 I/O 部分 | 模型重建、样式映射 |

保留的 GenOffice 页面文件（**不 fork**，走 `@sheets/*`）：`*.tsx` 对话框、`print-*`、`formula-*`、`i18n/**`、`styles.css`、`op-executor.ts`（除非要改）。

### 6.2 `window.desktopApi` 兼容层

页面渲染器有 ~40 处 `desktopApi.*`。沿用 docx 插件的 `Proxy` 方案，一次性把**未实现的**表面收成默认值，只显式实现真正用到的：

```ts
// src/page/host/desktop-api.ts
window.desktopApi = new Proxy(api, {
  get(target, key) {
    if (key in target) return target[key]
    if (typeof key === 'string' && /^on[A-Z]/.test(key)) return () => () => {}
    return async () => undefined
  },
})
```

必须实现：`readWorkbookRange`、`readWorkbookFormulas`、`readWorkbookMedia`、`readWorkbookRowOutline`、`findWorkbookCells`、`recalcWorkbook`、`saveWorkbookEdits`、`reopenWorkbook`、`closeWorkbook`、`exportPdf`、`printWorkbook`、`exportCsv`/`confirmCsvSave`、`readLocalImage`/`fetchImage`、`getUserDisplayName`/`onUserDisplayNameChanged`、`onLanguageChanged`、`listPrinters`、`openExternal`。

单文件范围内直接置空：`selectWorkbook` → `null`、`selectWorkbooksForMerge` → `null`、`openWorkbooksForMerge` → `[]`、`listDocsTabs` → `[]`、`getRecentFiles` → `[]`、`writeRecoveryCopy` → `ok`。

### 6.3 页面内不再需要的东西

- `session-bytes.ts`（`fetchSessionBytes`）→ `page/load.ts`（`open/chunk`）；
- `workstudio-save.ts` 的 `saveWorkbookEdits`（gateway + JSZip）→ 删除，只留 `planWorkbookSave`；
- `shims/node-fs*.ts` 若仅服务于 `assembleWithJsZip` 则删除；gateway `domain/*` 仍需要的保留（`node:path`/`node:crypto` 兜底）；
- `src/shared/xlsx-parse.ts` 从页面移除（页面用 `meta+range`），只在引擎侧保留。

---

## 7. 引擎改造

```text
src/engine/index.ts
  export const workstudio = definePlugin({
    events: true,
    resources: { xlsx: xlsxResource },
    commands: { …§4 },
    lifecycle: { activate: setScope, deactivate: clearScope },
  })
```

- `wasm.ts` / `sessions.ts` 基本原样移植；`EngineCtx.sessionDir` → `getScope().tempDir`（插件私有、随卸载清理），LRU 2 与版本失效逻辑不变。
- `budget`：宿主不再给，改为常量（`MAX_FILE_BYTES` 64/150 MB、`MAX_COMMIT_BYTES` 16 MB、`MAX_READ_CELLS` 2,000）；`AbortSignal` 用 `ctx` 外的定时器替代（单线程 wasm 本来也不支持取消，`degraded:['cancel-unavailable']` 保留在 `status` 诊断里）。
- 公式重算顺序**不得改动**：`recalc_cells` 先拿 stale 值 → 再和本次编辑一起交给 `planCellEditsToXlsx`；新写公式用 1×1 `read` 回填。
- `read.ts` 的 `audit`（包结构诊断）保留；`workbook.ts` 的解析/投影保留（引擎侧唯一解析器）。
- `state.ts` 里 `docRev`/`structureEpoch` 由**页面**上报（`report`），引擎只做比较，不自己推演页面模型。

---

## 8. 阶段与验收

> 当前进度：**P0–P4 已完成**（含 `save_archive` 提速、旧包弃用、CI 接入）。剩下只有注册表发布。

| 阶段 | 内容 | 验收 |
| --- | --- | --- |
| **P0 契约** | 新包骨架：清单/claims/skills、`src/genoffice/*` barrel、resolver、`page` 通道（status/open/chunk/meta/range）、页面只读网格 | `workstudio-plugin` 生成 `commands.json`；打开 1 MB 工作簿、滚动可视区、主题切换正常；单文件守卫用例通过；`verify` 通过 |
| **P1 fork 收敛** | `editor-diff` → `editor.manifest.json`；`src/editor` 只留改过的文件，其余回落 vendor；shims 收敛 | fork 文件数 ≤ 40；`--check` 无未登记差异；页面 bundle 体积不增长 |
| **P2 流式与写通路** | wasm 接入（`open` / `read_range` / `recalc_cells`）；`meta + range` 流式打开；`saveAs` / 分块 save；自动保存；保存后只重建元数据 | 人工编辑保存往返一致（`sha256` 稳定）；保存后 reopen 不回落旧值；分块边界与空白区域（`<dimension ref="A1"/>`）各测一次 |
| **P3 agent 面** | `apply` / `create` + `pending/ack` + op 级令牌；恢复副本；页面读投影 | 无页面 cells/structure/sheets 全通过；`objects` 无页面回 page-unavailable；dryRun 不落盘；stale 分支有单测 |
| **P4 清理与发布** | `save_archive` 提速、旧 `xlsx-pro` 弃用并退出默认打包、CI 装 wasm target、smoke / 真机清单 | `typecheck + build + test + verify + smoke`；zip 安装到 0.4.0 实机按 `docs/xlsx-e2e.md` 走一遍 |

每个阶段都要求：`tests/engine.test.mjs` 等价用例迁移（命令级），新增 `tests/lease.test.mjs`、`tests/save-parity.test.mjs`（页面请求 vs 引擎字节一致）、`tests/chunks.test.mjs`（分块哈希）。

---

## 9. 删除清单（不要带进新包）

```text
src/sdk.ts                      → @work-studio/sdk
src/ura-types.ts                → SDK 契约类型
src/ura.ts                      → definePlugin + commands
src/engine/{abi,tools,plan,exec,read,runtime}.ts → commands/* + gateway.ts
src/page/{main,agent,session-bytes}.ts           → 重写（boot/agent/load）
src/editor/workstudio-host.ts   → 拆进 src/page/host/*
src/editor/workstudio-save.ts 的 saveWorkbookEdits / assembleWithJsZip
src/editor/save-plan-budget.ts
src/editor/mcp-bridge.ts / mcp-sheet-refs.ts（若仅服务 AI 桥）
src/editor/merge-workbooks.ts（单文件范围不做跨文件合并）
src/editor/shims/node-fs-promises.ts（若只剩 page 内组包用）
scripts/verify.mjs 中的 codec/presenter/sidecar/apiVersion 检查
package.json 的 workstudio.engine / apiVersion / write / content / model
```

保留并改写：`tools/catalog.ts`（56 op）、`shared/op-tiers.ts`、`shared/headless.ts`、`shared/workbook.ts`、`shared/zip-raw.ts`、`engine/wasm.ts`、`engine/sessions.ts`、`rust/xlsx-engine/wasm.rs`。

`packages/xlsx-pro` 不删源码：标记弃用、移出默认 `pack:plugins` / `install:plugins` 列表，留给旧 URA 宿主回退。

---

## 10. 宿主零改动（决策）

不改 WorkStudio。四个缺口全部用插件侧机制补，代价写在右列：

| 缺口 | 插件侧替代 | 残留代价 |
| --- | --- | --- |
| 技能激活后，页面命令对助手**域级放行** | 页面命令收成单条 `page`；`attach` 令牌 + `pageId` 归属；`save` 包结构校验；租约忙时拒绝 | 助手命令表多一行 `page`；鉴权强度依赖令牌而非来源 |
| `PluginContext` 没有 `origin` | 令牌 + `lease.busy` + `ack/pending` 的 id 归属 | 助手若先 `attach` 会占租约；页面心跳失败即 `takeover:true` 自愈 |
| 脏态 no-op / 关窗不确认 | 空闲自动保存（800 ms）+ `pagehide` 尽力 flush + `scope.dataDir` 恢复副本 | 标题栏无脏点；关窗不弹确认，最多丢最后一个防抖窗口 |
| 无语言通道 | `navigator.language` → `zh` 兜底 | 与 WorkStudio 的语言设置可能不一致 |
| 无字节通道 / 非原子写 / 另存越界 | `page` 分块 op；`baseHash` + 锁；先写恢复副本；另存仅工作区内 | 大文件保存峰值内存 2×；跨目录另存不支持 |
| 多窗口写根 | 不处理：写入最近激活工作区的根 | 两个工作区同时打开同一文件时，以最近激活者为准 |
| 卸载清理 | `lifecycle.deactivate` 里自己清 wasm 实例与会话缓存 | 无（插件可控） |

触发「回头改宿主」的条件（任一条成立就单独提宿主改动，插件契约不变）：

1. 要标题栏脏点或关窗确认；
2. 要给助手强隔离页面命令，不再依赖令牌；
3. 要另存到工作区之外；
4. 单文件 >150 MB 仍要保持可编辑。

其余风险（沿用旧结论）：

| 项 | 方案 |
| --- | --- |
| 大文件 IPC | `page {op:'chunk'}`，1 MiB/块 + `hash` 绑定；>150 MB 明确拒绝 |
| 原子写 | 先写 `scope.tempDir` 校验、再 `ctx.fs.write`；`baseHash` + `withLock` 兜并发；崩溃靠恢复副本 |
| 上游漂移 | `fork-diff --check` 进 CI；升级 vendor 先看 fork 清单，不整包 diff |
| 许可 | `NOTICE` 写明 vendor commit 与 Apache-2.0 来源；zip 内附 LICENSE |
