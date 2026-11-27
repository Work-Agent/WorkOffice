# pptx 整改方案（WorkStudio 新插件框架 · 就地改造）

> 依据：`docs/docx-pro-remediation.md` 的同一套做法，以及 `@work-studio/sdk`（当前宿主 0.3.3 / SDK 0.2.0）。
>
> **只改 `packages/pptx`，不改 WorkStudio 宿主、不改 SDK。**
> 包 id 仍是 `pptx`，清单从 `apiVersion: 2` 换成 `pages + claims + skills`，
> 目录结构对齐 `packages/docx-pro`。

---

## 1. 插件到底做什么

| 归属 | 职责 |
| --- | --- |
| 宿主 shell | 标签页、打开/关闭视图、dirty 标记、文件列表、重命名、工作区刷新 |
| pptx 插件 | 接收 `ref` → 打开页面 → `load` 字节 → 编辑 → `save` 当前 `ref`；助手在页面内读写 |
| 宿主 / SDK | IPC 通道、命令鉴权、页面生命周期、`shell/native.saveFile` |

所以插件**不实现**标签页、文件列表、目录遍历、dirty 上报；`ref` 永远是单个文件。
**新建是唯一由引擎产字节的场景**（`create`），其余写入全部发生在已打开的页面里。

---

## 2. 旧实现与新框架的差距

| 旧能力 | 旧位置 | 新框架状态 |
| --- | --- | --- |
| URA codec / presenter / sidecar / store / transport | `src/ura.ts`、`src/sdk.ts`、`src/ura-types.ts` | 不存在；改用 `ResourceAdapter` + `claims` |
| `apiVersion: 2` 清单（`write/content/model/engine`） | `package.json` | 不存在；改用 `workstudio.pages` + `exports['./workstudio']` |
| 页面会话 `ura.page.get/bytes/commit/agentResult` | `src/page/{main,agent}.ts`、`src/editor/workstudio-host.ts` | 不存在；页面只有 `window.workStudio` |
| `plan` → `exec` → CAS（引擎产计划） | `src/engine/{plan,exec,state,abi,runtime}.ts` | 不存在；改为"引擎只入队，页面产字节" |
| 四个逐工具声明 | `src/engine/tools.ts` | 不存在；改为 5 个助手命令 + 8 个页面命令 |
| `agent-request` 页面代理 | `src/page/agent.ts` | 不存在；用 `events` + `pending` + `ack` 自建 |
| `version` / `epoch` | 页面回执 | 不存在；改用 `baseHash` / `docRev` |
| 1 MiB `bytes` 分块拉取 | `page/main.ts:36`、`workstudio-host.ts:379` | 不存在；一次 `load` 返回整份字节 |

### 2.1 传输事实

| 页面 API | 通道 |
| --- | --- |
| `command` | `ipcRenderer.invoke('workstudio:command')` → `Dispatcher.call` |
| `subscribe` / `onUpdate` | `invoke('workstudio:subscribe')` / `webContents.send('workstudio:update')` |
| `notify` | `ipcRenderer.send('workstudio:notify')` |

- 传输支持 `Uint8Array`（structured clone）；`ArgSpec` 没有 `bytes` 只影响描述，不影响运行时。
- 命令地址是 `<domain>/<resource>/<action>`，插件固定为 `pptx/pptx/<action>`；
  引擎里用 `'pptx.<action>'` 作 key（宿主按第一个 `.` 切出 action，与 docx-pro 一致）。
- 根目录 `commands.json` 必填，且与引擎命令 key **双向**对齐（`mergeStaticCommands` 会报错）。
- 插件页只能调自己域 + `shell/native.saveFile|openExternal`；`shell/<ws>/views`、`workspace/*` 不可订。
- 助手来源能否调页面命令只由 `expose.agent` 决定；`save` 必须 `expose.agent: false`。

---

## 3. 目标形态

### 3.1 清单

```jsonc
{
  "name": "@workstudio/plugin-pptx",
  "version": "1.0.0",
  "type": "module",
  "exports": { "./workstudio": "./dist/engine.js" },
  "workstudio": {
    "id": "pptx",
    "label": "PowerPoint 演示文稿",
    "icon": "list",
    "pages": { "content": "dist/page/index.html" },
    "claims": [{ "extensions": ["pptx", "pptm", "potx"], "priority": 100 }],
    "skills": ["skills/pptx"]
  }
}
```

- 没有 `apiVersion` / `write` / `content` / `model` / `engine` / `page`（旧的 `page` 单字段也不再合法，必须是 `pages.content`）。
- `icon` 只认宿主 `icons.ts` 里的 18 个名字（`file-text`/`list`/`eye`…）；`presentation`、`table` 不在表里，会退化成 `Package` 图标。
- 迁移期与旧 `pptx` 冲突，`claims.priority: 100` + 卸载旧包。

### 3.2 目录结构（对齐 docx-pro，已落盘）

```text
packages/pptx/
├─ package.json                 id/pptx · pages.content · claims · skills · exports['./workstudio']
├─ commands.json / api.json / api.d.ts   gen-commands 生成（提交）
├─ assets/seed.pptx             create 的模板
├─ skills/pptx/{SKILL.md,design.md,spec.md}
├─ src/
│  ├─ engine/
│  │  ├─ index.ts               definePlugin：resources + 13 命令 + events + lifecycle
│  │  ├─ read.ts / check.ts / create.ts / build.ts / images.ts / abi.ts
│  ├─ page/
│  │  ├─ main.ts                ?ref= → load → 装 host → 挂编辑器 → attach → 队列
│  │  ├─ host.ts                GenOffice Slides 适配层（原 editor/workstudio-host.ts）
│  │  ├─ channel.ts             window.workStudio 的类型化薄封装
│  │  ├─ lease.ts               events → pending → 执行 → save → ack + 15s 心跳
│  │  ├─ env.d.ts / index.html
│  ├─ shared/                   19 个纯逻辑模块（deck/read/check/spec/ops/design/…）
│  ├─ editor/                   GenOffice Slides 整包（226 文件，fork 常驻）
│  └─ tools/{catalog.ts,op-catalog.generated.ts,types.ts}
├─ scripts/{build,gen-commands,gen-op-catalog,gen-seed,verify}.mjs
├─ tests/{engine.test.mjs,harness.mjs}
└─ vite.config.ts / tsconfig.json
```

删除的旧文件：`src/ura.ts`、`src/ura-types.ts`、`src/sdk.ts`、`src/types/raw.d.ts`、
`src/engine/{plan,exec,state,runtime,tools}.ts`、`src/page/agent.ts`。

与 docx-pro 的对应关系：

| docx-pro | pptx |
| --- | --- |
| `src/engine/index.ts`（600 行，单文件） | 同形，约 700 行 |
| `src/engine/read.ts` | 保留 + 复用 `shared/read.ts` |
| `src/page/{main,host,lease,command-batch}.ts` | 同四件套 |
| `src/editor/`（GenOffice Docs 整包） | `src/editor/`（GenOffice Slides 整包） |
| `src/tools/{catalog,commands.generated,op-catalog.generated}.ts` | 同形 |
| `scripts/{build,gen-commands,gen-seed,verify}.mjs` | 同形 |
| `assets/seed.docx` / `skills/docx/SKILL.md` | `assets/seed.pptx` / `skills/pptx/SKILL.md` |

### 3.3 运行结构

```text
plugin://pptx/content?ref=file:///w/a.pptx
        ├─ pptx.load    → { name, size, hash, bytes }
        ├─ pptx.attach  → { hash, pending[] }；report 续期；detach 释放
        ├─ pptx.save    → 当前 ref 唯一落盘点
        └─ subscribe('pptx','events') → 助手请求

主进程引擎（dist/engine.js）
        ├─ ResourceAdapter：pptx（read / grep；ls/find 恒空）
        ├─ 助手命令：guide / read / check / create / apply（apply 只入队）
        ├─ 页面命令：load / save / saveAs / attach / pending / report / detach / ack
        └─ PluginScope：租约、pending、events
```

### 3.4 写不变量

1. **没有页面就没有写**：`apply` 无有效租约直接 `command.page-unavailable`；不存在后台回退。
2. **助手只入队**：`apply` 只做参数校验 + 入队 + 等 `ack`；主进程不跑 `runTxn`、不写盘。
3. **唯一落盘点**：页面的 `save`（`expose.agent: false`），且必须持有该 ref 的租约。
4. **`create` 例外**：新建时没有页面，引擎 `seed.pptx + buildCreateOps + runTxn + savePptx` 写一次。
5. **只读不受限**：无页面时 `guide/read/check` 仍读磁盘。
6. **一次手势一个事务**：一批 op = 一个 `runTxn` + 一条 undo；失败整批回滚。

### 3.5 两条写路径

| 路径 | 入口 | 写入者 | 落盘 |
| --- | --- | --- | --- |
| 人工编辑 | 编辑器事务 → `savePptx(opened)` | 页面 | `pptx.save` |
| 助手编辑 | `apply` → 页面队列 → `runTxn` | 页面 | 同上 |

没有"助手 + 无页面"。

---

## 4. 契约设计

### 4.1 ResourceAdapter：`pptx`

```ts
{
  read: async (ref, options) => [{
    locator: `${ref}#outline`,
    title, slideCount, size, hash,
    text: '…引导语 + 每页标题…',
  }],
  grep: async (ref, query) => [{ locator, text }],
  open: async (ref, locator) => ({ ref, locator }),
}
```

`ls` / `find` 恒 `[]`（不认领目录）。`read` 只在助手侧走磁盘；页面里用 `shared/read.ts` 直接读 live 模型。

### 4.2 命令面（13 条）

| 命令 | 参数 | `expose.agent` | 语义 |
| --- | --- | --- | --- |
| `pptx.guide` | `ref`；`group` enum；`op` | true | 命令目录 |
| `pptx.read` | `ref`；`what` enum；`slide`；`query`；`limit` | true | 有租约读 live，无租约读磁盘 |
| `pptx.check` | `ref`；`outline`/`pages`/`theme`（JSON 文本）；`audit`/`skeleton`/`slide` | true | 写前检查 / 几何审计 |
| `pptx.create` | `ref`；`pages`/`slides`/`theme`（JSON）；`content`/`style`/`size` | true | 新建（目标必须不存在） |
| `pptx.apply` | `ref`；`commands`（JSON 文本）；`baseHash`；`docRev`；`dryRun` | true | 只入队 |
| `pptx.load` | `ref` | false | `{ name, size, hash, bytes }` |
| `pptx.save` | `ref`；`bytes`；`baseHash` | false | 锁 + CAS + 校验 + 写 |
| `pptx.saveAs` | `ref`；`bytes`；`token` | false | P3；token 来自 `shell/native.saveFile` |
| `pptx.attach` | `ref`；`pageId` | false | 发租约 + 未完成请求 |
| `pptx.pending` | `id` | false | 取请求正文 |
| `pptx.report` | `ref`；`pageId`；`hash`；`docRev`；`dirty` | false | 续期 + 状态 |
| `pptx.detach` | `ref`；`pageId` | false | 释放 + 清 pending |
| `pptx.ack` | `id`；`ok`；`payload`（JSON） | false | 结算 |

### 4.3 类型

```ts
type SaveReceipt = { ref: string; hash: string; size: number; at: number }
type PageRequest =
  | { type: 'apply'; id: string; commands: unknown[]; baseHash?: string; docRev?: number; dryRun?: boolean }
  | { type: 'read'; id: string; what: string; params: Record<string, unknown> }
type AckPayload = { receipt?: SaveReceipt; data?: unknown; error?: { code: string; message: string; index?: number } }
```

### 4.4 字节通道

- `load` 一次返回整份 `Uint8Array`；上限 **64 MB**（docx-pro 16 MB；pptx 带媒体，包更大）。
- `save` 一次提交整份字节；上限同样 64 MB；超限 `command.too-large`。
- **不分块**：`saveBegin/saveChunk/saveCommit` 与 `open/chunk` 全部不做（上一版抄 xlsx-plus 的分块是多余复杂度）。
- 校验：ZIP 魔数 `PK\x03\x04`、包内必须有 `[Content_Types].xml`、`ppt/presentation.xml`、体积匹配。

### 4.5 租约与锁

```ts
type Lease = { ref: string; pageId: string; hash: string; docRev: number; dirty: boolean; updatedAt: number }
```

- 每 ref 一份租约；新 `pageId` `attach` 覆盖旧的（reload 自然工作），旧 pending 作废。
- TTL 5 min，页面 15 s `report` 续期；`pending` 30 s 超时。
- 写锁只在 `save` / `saveAs` / `create` 取；`apply` 不取锁。
- 同 ref 的 `apply` 串行：引擎等上一条 `ack` 再放行下一条。

### 4.6 页面代理

```text
apply → 校验 ref/baseHash/docRev/commands → enqueue
      → scope.events.emit('request', { id, type:'apply' })
page  → onUpdate('pptx','events') → pending(id) → 正文
      → docRev 校验 → flushActiveEdit → history.push(snapshot)
      → runTxn(opened,{ ops, isolation:'atomic' }) → buildSlides
      → savePptx(opened) → save { bytes, baseHash } → report
      → ack { ok, payload:{ receipt } }
engine→ settlePending → 工具结果
```

失败：页面不可用 → `command.page-unavailable`；页面 `docRev` 过期 → `command.stale-index`；
`runTxn` 被拒 → 整批回滚，`ack.ok=false` 带 `index` 与 guided error。

---

## 5. 代码改造

### 5.1 新增

| 文件 | 作用 |
| --- | --- |
| `src/engine/index.ts` | `definePlugin` 入口：adapter、13 命令、events、lifecycle |
| `src/page/lease.ts` | pending 处理、15 s 上报、detach |
| `src/page/channel.ts` | `window.workStudio` 薄封装（load/save/saveAs/attach/…/theme/saveFile） |
| `scripts/gen-commands.mjs` | 从 `dist/engine.js` 生成 `commands.json` / `api.json` / `api.d.ts` |
| `tests/{engine,page}.test.mjs` | 假页面 harness + 磁盘 harness |
| `src/shared/sha256.ts` | 已由 `@noble/hashes` 覆盖；抽一层给引擎/页面共用 |

### 5.2 重写

| 文件 | 改造 |
| --- | --- |
| `package.json` | 新清单、`exports`、去掉 `ura` 入口、版本 `1.0.0` |
| `src/page/main.ts` | `?ref=`；`load`；挂 `editor/main`；`attach`；订阅 `events`/`appearance` |
| `src/editor/workstudio-host.ts` → `src/page/host.ts` | 2 处 `bytes()`、1 处 `commit()` 换成 `load`/`save`；`bus.send({dirty})` → `report` |
| `src/engine/{read,check,create,build,images,abi}.ts` | 去掉 URA 依赖，纯函数化 |
| `scripts/build.mjs` | 引擎 bundle + 页面 vite + 生成描述文件 |
| `scripts/verify.mjs` | 清单 / 命令对齐 / 旧面扫描 / 一次往返 |
| `skills/pptx/SKILL.md` | 5 个助手命令；`apply` 前必须页面已开 |

### 5.3 删除

```text
src/ura.ts
src/ura-types.ts
src/sdk.ts
src/types/raw.d.ts
src/engine/{plan,exec,state,runtime,tools}.ts
src/page/agent.ts
```

`src/shared/ipc.ts` 保留（编辑器 fork 的 `window.desktop` 类型还在用），文档面型换成 `workStudio`。

### 5.4 编辑器 fork 最小改动（`src/editor`，9 处）

| 文件 | 改动 |
| --- | --- |
| `App.tsx` | 删 AI 面板挂载；`__pptxProUi` 只留 2 个入口 |
| `components/RibbonHomeTab.tsx` | 删 Genspark AI 组 |
| `file-actions.ts` | 删 AI 命名 / 视频导出 |
| `components/PrintDialog.tsx` | 宿主无 `listPrinters` → 走浏览器打印 |
| `MasterView.tsx` | 宿主无 `masterAddElement` → 隐藏 |
| `components/SlideShowView.tsx` | 全屏 IPC 600 ms 兜底 |
| `components/PresenterView.tsx` | Document PiP → `window.open` 兜底 |
| `konva-adapter.ts` | JPEG 容差 16→4、CJK 字体栈 |
| `redact-actions.ts` / `keyboard-actions.ts` / `action-context.ts` | 去掉 AI 依赖 |

`Ribbon.tsx`（3504 行）对宿主调用为 0，**不 fork 不改**。

### 5.5 技能

```text
助手 5 个命令（command 工具 / bash）：
  pptx.guide   命令目录
  pptx.read    读 outline/pages/theme/styles/audit/skeleton/notes
  pptx.check   写前检查（outline JSON 断言 + 几何审计）
  pptx.create  新建（目标不存在才可）
  pptx.apply   commands 是 JSON 字符串；要求页面已打开

页面没开 → command.page-unavailable → 调宿主 open(ref) 再重发。
写 .pptx 禁止 write / edit；不走 workspace 目录命令。
```

---

## 6. 分阶段

| 阶段 | 内容 | 人日 |
| --- | --- | --- |
| P0 | 清单 + `dist/engine.js` + `commands.json`；`load` 打通；页面能渲染 | 1–2 |
| P1 | `guide/read/check`；adapter `read/grep`；删 URA 三件套 | 2–3 |
| P2 | `save` + `attach/report/detach` + `page/host.ts` 改造；CAS/stale | 3–4 |
| P3 | `apply` 队列 + `pending/ack` + 页面事务 + `docRev`；无页面拒绝 | 3–4 |
| P4 | `create` / `saveAs` / 自动保存 / tests / verify / 打包 | 3–4 |

合计 **12–17 人日**（docx-pro 当时估 14–19）。

**P0 退出**：打开 40 页 deck 可翻页；`commands.json` 与引擎对齐；无 `ura.*`。
**P2 退出**：改文字/拖元素/增删页 → 保存 → 重开一致；外部改动后保存报 `command.stale`。
**P3 退出**：助手一批命令一个 undo 步；页面关掉后所有写失败且零落盘。

---

## 7. 测试与验收

| 用例 | 断言 |
| --- | --- |
| 清单 | `pages.content` / `claims` / `skills`；无 `apiVersion`；`icon` 在宿主表内 |
| 命令对齐 | 动态 13 条 ↔ `commands.json` 双向一致；`save*` 的 `agent=false` |
| 旧面扫描 | `src` 内无 `window.extension` / `ura.` / `agent-request` / `registerCodec` |
| 无页面 apply | `command.page-unavailable`；`ctx.fs.write` 零调用 |
| 页面 apply | 假页面 attach → pending → `runTxn` → ack；`dryRun` 不落盘；整批回滚 |
| save | CAS 拒绝（stale）；非 pptx / 空字节 / 超 64 MB / 缺 `[Content_Types].xml` 拒绝 |
| create | 目标已存在 → `resource.exists`；成功后解包含 seed 页 |
| 并发 | 两个页面 attach 后者覆盖；旧 pending 作废 |

---

## 8. 不改宿主的边界

| 能力 | 插件 | 说明 |
| --- | --- | --- |
| 打开/编辑/保存当前 ref | ✅ | `load` / `save` |
| 助手 guide/read/check/apply | ✅ | `apply` 需页面 |
| 新建空演示文稿 | ✅ | `create`（引擎产字节） |
| 另存为 | ⚠️ P3 | 页面调 `shell/native.saveFile` 拿 token 再 `saveAs` |
| 目录/文件列表 | ❌ | 宿主 shell 职责 |
| 标签 / dirty | ❌ | 宿主视图命令；插件只 `report` 自用状态 |
| 外部修改实时刷新 | ⚠️ | 不做 watcher；靠 `baseHash` 在保存时发现 |

结论：不改宿主即可交付；只有"自动刷新"是产品取舍，不是阻塞。

---

## 9. 提交切分

1. `chore(pptx): sdk manifest, drop ura entry, commands.json`
2. `refactor(pptx): engine single file — adapter + guide/read/check`
3. `refactor(pptx): delete plan/exec/state/runtime/tools and page agent bridge`
4. `feat(pptx): page load/save over workStudio commands`
5. `feat(pptx): lease, events, pending/ack, apply queue`
6. `feat(pptx): create from seed; saveAs via saveFile token`
7. `test(pptx): fake page harness, no-page rejection, CAS, batch rollback`
8. `chore(pptx): verify, skills, packaging`

---

## 10. 剩余问题（实现验证 / 产品取舍）

| 项 | 现状 | 处理 |
| --- | --- | --- |
| `savePptx` 在页面里可用 | ✅ 已确认：`vendor/.../pptx-engine/src/index.ts:746`；`node:crypto`/`node:zlib` 已由 `src/editor/shims` 别名替换 | P0 实测 40 页 deck 保存往返 |
| `runTxn` 回滚 | ✅ 已确认：`pptx-ops/src/ops/executor.ts:160`，atomic 默认，内部快照回滚 | 页面侧再包一层 `history.push` 保证一个 undo 步 |
| 上传/下载上限 | docx 16 MB；pptx 需更大 | 定 64 MB；超限拒绝而非分块 |
| 另存为 | docx-pro 判为宿主职责 | pptx 保留 `saveAs`（页面已有字节、宿主已有 `saveFile`），放 P3 |
| 大图/媒体插入 | `engine/images.ts` 走引擎；页面侧需 `media-insert` | P2 检查 `node:zlib` shim 是否覆盖 `deflateSync` 路径 |
| 页面忙时 apply | 拖动中 | `flushActiveEdit()` 后提交；P3 实测 |
| `powerpoint` 之外的扩展名 | pptm/potx | claims 声明三个；`create` 只产 `.pptx` |

---

## 11. 实现状态（本轮已落地）

已完成并通过 `typecheck` / `test`（11 个用例）/ `verify` / 全量 `build`：

- **清单**：`@workstudio/plugin-pptx@1.0.0`，`id: pptx`、`pages.content`、`claims(pptx/pptm/potx, 100)`、
  `skills: ['skills/pptx']`、`exports['./workstudio'] → dist/engine.js`；`apiVersion/write/content/model/engine` 全部移除。
- **引擎**（`src/engine/index.ts`，单文件）：`resources.pptx{read,grep,open,describe}` +
  13 条命令（`guide/read/check/create/apply` + `load/save/saveAs/attach/pending/report/detach/ack`）+ `events` + 租约/pending/超时。
- **页面**：`main.ts` 一次 `load`；`host.ts` 用 `savePptx(opened)` 产字节、`commitSaved` 同步基线；
  `lease.ts` 走 events → pending → `runTxn` → save → ack；`channel.ts` 收全部 IPC。
- **删除**：URA 三件套、plan/exec/state/runtime/tools、page/agent、raw types、`tests/ura.test.mjs`。
- **脚本**：`build.mjs`（engine + seed + gen-commands + vite 页面）、`gen-commands.mjs`、`verify.mjs`。
- **文档**：README / SKILL.md 改成 13 命令 + 页面优先；`pptx_check`→`pptx.check` 等命名统一。
- **打包**：`scripts/pack.mjs` / `scripts/install.mjs` 改成自动发现（扫 `packages/*` 里声明了
  `workstudio` 的包，`--only` 会校验包存在），产物 `release/pptx-1.0.0.zip`（3.2 MB）+
  `.sha256`；zip 内只有精简清单、`dist/`、`commands.json`、`api.*`、`skills/`。
  已实测：解包 → 安装到临时目录 → `import` 安装后的 `dist/engine.js` → `pptx.load` 跑通。

与 docx-pro 的偏差（都写在 §10）：

| 点 | docx-pro | 本轮 pptx | 原因 |
| --- | --- | --- | --- |
| `saveAs` | 不做（判给宿主） | **做了**（`saveFile` token → 写工作区内） | 页面已有整份字节，成本只有 8 行；跨盘仍 `denied` |
| `check` | 折进 `read what:'audit'` | **独立命令** | pptx 的 check 要吃 outline/pages 做断言，参数形状与读不同 |
| 通道封装 | 直接调 `window.workStudio` | 多一个 `page/channel.ts` | 9 条页面命令 + 主题 + saveFile，收一处 |
| 事务粒度 | per-block / per-op | 一批 = 一个 `runTxn` | pptx 的 op 天然是批量原子 |

未做（明确留待后续）：

- 宿主 GUI e2e（本环境未启动 WorkStudio；页面已构建但没在真窗口里跑过）；
- 另存为跨盘（受 `ctx.fs` 工作区边界限制，需要宿主新能力）；
- 自动保存开关（壳层动作目前只有面板内 Ctrl+S）。

## 12. 结论：可以执行且已执行

三个前提都已在代码里落地并被 verify 覆盖：

1. **页面能产整份字节**：`pptx-engine` 的 `savePptx(opened)` 在页面 bundle 里可用
   （`node:crypto` / `node:zlib` 已由 `src/editor/shims` 别名替换）。
2. **批量 op 有原子事务**：`runTxn(opened,{ops})` 默认 atomic，失败内部回滚；
   页面在它外面再包一层 `history.push` 得到"一批 = 一个 undo 步"。
3. **旧面收口**：`ura.` 调用点 7 处全部替换，`verify` 会扫描并拒绝
   `window.extension` / `ura.page` / `apiVersion` / `registerCodec` / `agent-request` / `saveBegin`。
