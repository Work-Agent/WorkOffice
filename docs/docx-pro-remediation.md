# docx-pro 整改方案（WorkStudio 新插件框架 · 只做文档页）

> 依据：`@work-studio/sdk`（npm `0.1.1`）与当前 preload / dispatcher 实现。
>
> 本方案只改 `packages/docx-pro`，不改 WorkStudio 宿主、不改 SDK。
>
> **范围纠正：docx-pro 只负责“给定 ref 打开文档并读写内容”。标签页、文件列表、新建文件、另存为、dirty 标记、工作区文件变更通知都属于宿主 shell / 文件管理，不属于插件。**

---

## 1. 插件到底做什么

| 归属 | 职责 |
| --- | --- |
| 宿主 shell | 标签页、打开/关闭视图、dirty 标记、文件列表、新建文件、另存为、重命名、工作区刷新 |
| docx-pro 插件 | 接收 `ref` → 打开文档页 → 加载字节 → 编辑 → 保存当前 `ref`；Agent 在页面内读写 |
| 宿主 / SDK | IPC 通道、文件 token、命令鉴权、页面生命周期 |

因此：

- 插件**不需要**实现标签页，也不需要上报 dirty；
- 插件**不需要**做文件列表 / 新建 / 另存为；
- 插件**不需要**订阅 `workspace` 文件树；
- 新建文档是宿主行为：宿主创建文件后打开 `plugin://docx-pro/content?ref=...`，插件只看到已存在的 ref。

这样就不存在“新建要先有页面、页面又依赖文件”的循环：**新建不归插件**。

---

## 2. 旧实现与新框架的差距

| 旧能力 | 旧位置 | 新框架状态 |
| --- | --- | --- |
| URA codec / presenter / sidecar / store | `src/ura.ts` | 不存在；改用 `ResourceAdapter` + claims |
| `apiVersion: 2` 清单 | `package.json` | 不存在；改用 `workstudio.pages` / `claims` / `skills` + `exports['./workstudio']` |
| 页面会话 `ura.page.get/bytes/commit/agentResult` | `src/page/*` | 不存在；页面只有 `window.workStudio` |
| `plan` → `exec` → CAS | `src/engine/plan.ts` / `exec.ts` | 不存在；本方案不要无页面写，整条删除 |
| `version` / `epoch` | 页面回执 | 不存在；插件用 `baseHash` / `docRev` 自管 |
| Agent 页面代理 `agent-request` | `src/page/agent.ts` | 不存在；用 `events` + `pending` + `ack` 自建 |
| 四个逐工具声明 | `src/engine/tools.ts` | 不存在；改为 4 个助手命令 + 页面命令 |
| legacy 工具面 | `--legacy-tools` | 删除 |

### 2.1 传输事实

`window.workStudio` 就是当前 preload 对 `ipcRenderer` 的封装：

| 页面 API | 通道 |
| --- | --- |
| `command` | `ipcRenderer.invoke('workstudio:command', ...)` → `ipcMain.handle` → `Dispatcher.call` |
| `notify` | `ipcRenderer.send('workstudio:notify', ...)` |
| `subscribe` | `ipcRenderer.invoke('workstudio:subscribe', ...)` |
| `onUpdate` | `webContents.send('workstudio:update')` → `ipcRenderer.on` |
| `fileToken` | `ipcRenderer.invoke('workstudio:file-token', ...)` |

所以：

- **传输层支持 `Uint8Array`**：Electron IPC 使用 structured clone。`ArgSpec` 没有 `bytes` 只影响 `commands.json` 描述，不影响运行时。
- **插件页拿不到裸 `ipcRenderer`**：preload 只暴露 `window.workStudio`；能发什么由命令表与 `origin` 鉴权决定。
- **`workspace` 订阅被拒是鉴权问题**：`subscriptionDenialOf` 只放行本插件域与 `shell/appearance`。但本插件不需要 workspace 文件树，所以不构成阻塞。
- **标签 dirty 不在插件范围**：`shell/<wsId>/view/<id>/dirty` 是宿主视图命令；插件不碰标签，也不需要这个通道。

---

## 3. 目标形态

### 3.1 清单

```jsonc
{
  "name": "@workstudio/plugin-docx-pro",
  "version": "1.0.0",
  "type": "module",
  "exports": { "./workstudio": "./dist/engine.js" },
  "workstudio": {
    "id": "docx-pro",
    "label": "Word 文档",
    "icon": "file-text",
    "pages": { "content": "dist/page/index.html" },
    "claims": [{ "extensions": ["docx", "docm", "dotx"], "priority": 100 }],
    "skills": ["skills/docx"]
  }
}
```

- 没有 `apiVersion` / `write` / `content` / `model` / `engine` / `page`。
- `skills` 是目录字符串数组，名称 / 描述在 `SKILL.md` frontmatter。
- 根目录必须有 `commands.json`；`api.json` / `api.d.ts` 同层生成。
- 迁移期建议卸载旧 `docx` / `office-docx`；`claims.priority: 100` 保证胜出。

### 3.2 运行结构

```text
plugin://docx-pro/content?ref=<file:///w/report.docx>&locator=docx:block:12
        │
        ├─ docx-pro.load            → 字节 + hash + tier
        ├─ docx-pro.attach/report/detach
        ├─ docx-pro.save            → 当前 ref 唯一落盘点
        └─ 订阅 docx-pro/events     → Agent 请求 / 保存回执

主进程引擎（dist/engine.js）
        ├─ ResourceAdapter：docx（read / grep / open）
        ├─ 助手命令：guide / read / create / apply（apply 只入队）
        ├─ 页面命令：load / save / attach / pending / report / detach / ack
        └─ PluginScope：租约、pending、事件
```

### 3.3 写不变量

1. **没有页面就没有写**：`docx-pro.apply` 没有有效页面租约时直接 `command.page-unavailable`。
2. **Agent 只入队**：`apply` 只解析参数、入队、等待页面 `ack`；主进程不执行块命令、不拼 XML、不写文件。
3. **唯一落盘点**：只有页面调用的 `docx-pro.save` 写当前 `ref`，且必须持有该 ref 的页面租约；`expose.agent: false`。
4. **不碰文件管理**：只有 `docx-pro.create` 创建空白 seed（唯一创建例外）；没有 `saveAs` / `status` / watcher；正文仍只由页面 `save` 写。
5. **只读不受限**：没有页面时 `guide` / `read` 仍可读磁盘；这是读，不是写。

### 3.4 两条写路径

| 路径 | 入口 | 写入者 | 结果 |
| --- | --- | --- | --- |
| 人工编辑 | 编辑器 PM 事务 → 编辑器 `save()` | 页面 `docx-pro.save` | 写当前 ref |
| Agent + 页面 | Agent `command` → `docx-pro.apply` → 页面队列 | 页面 PM 事务 + `docx-pro.save` | 写当前 ref |

没有“Agent + 无页面”；保存回执只对当前页面。

---

## 4. 契约设计

### 4.1 ResourceAdapter：`docx`

```ts
type DocxItem = Item & {
  index?: number
  type?: 'doc' | 'p' | 'tbl' | 'sdt' | 'other'
  heading?: number
  text?: string
  hash?: string
}

resources: {
  docx: {
    read: async (ref, options) => readItems(ref, options),
    grep: async (ref, query) => grepBlocks(ref, query),
    open: async (ref, locator) => ({ ref, ...(locator ? { locator } : {}) }),
  },
}
```

- 第一条是文档元信息：`locator: 'docx:doc'`，带 `hash` / `size` / `name` / `blockCount` / `tier`。
- 后续块：`locator: 'docx:block:<docxIndex>'`，带 `index` / `type` / `heading` / `text` / `hash`。
- `read` 自己按 `offset/limit` 切片，默认最多 500 条；`grep` 返回命中块。
- 适配器没有 `PluginContext`，用 `node:fs/promises` + `fileURLToPath(ref)`；只接受 `file://` 与 docx 扩展名。

### 4.2 命令面

**助手可见（4 个）**

| 命令 | 参数 | 语义 |
| --- | --- | --- |
| `docx-pro.guide` | `ref`、`group?`、`op?` | 命令目录；不读字节 |
| `docx-pro.read` | `ref`、`what?`、`startBlockIndex?`、`endBlockIndex?`、`offset?` | 页面 live 读；无页面读磁盘 |
| `docx-pro.create` | `ref` | 创建空白 seed 文件；目标不存在；不写正文 |
| `docx-pro.apply` | `ref`、`commands: string`、`baseHash?`、`docRev?`、`dryRun?` | `commands` 是 JSON 字符串；只入队；无页面拒绝 |

**页面可见（7 个，均 `expose.agent: false`）**

| 命令 | 参数 | 语义 |
| --- | --- | --- |
| `docx-pro.load` | `ref`、`base64?` | 默认返回 `{ bytes, hash, name, size, tier }`；`base64:true` 返回 `{ data, ... }` |
| `docx-pro.save` | `ref`、`data?: base64`、`bytes?`（快路径）、`baseHash?`、`type: 'page' \| 'agent'` | 优先 typed array，否则 base64；校验后写当前 ref |
| `docx-pro.attach` | `ref`、`pageId` | 登记页面租约；返回 `{ hash, pending }` |
| `docx-pro.pending` | `id` | 取回待处理请求正文（事件只带 id） |
| `docx-pro.report` | `ref`、`pageId`、`docRev`、`hash` | 页面节流上报 live 版本 |
| `docx-pro.detach` | `ref`、`pageId` | 撤销租约 |
| `docx-pro.ack` | `id`、`ok`、`payload: string`（JSON） | 页面回执，结算 pending |

没有 `saveAs` / `status` / `watch`；`create` 只写空白 seed。

`commands` 用 JSON 字符串：现有 `ArgSpec` 没有对象数组；字符串能进入 `commands.json` / `api.d.ts`，宿主校验也能通过。技能里明确 `JSON.stringify`，引擎 `JSON.parse` 后校验数组。

### 4.3 类型

```ts
export interface LoadResult {
  ref: string
  name: string
  size: number
  hash: string
  tier: 'normal' | 'lite' | 'readOnly' | 'refused'
  bytes?: Uint8Array // invoke 快路径
  data?: string      // base64 兜底
}

export interface SaveReceipt {
  ref: string
  hash: string
  size: number
  at: number
  applied?: number
  warnings?: string[]
}

export interface PageLease {
  ref: string
  pageId: string
  docRev: number
  hash: string
  updatedAt: number
}

export type PageRequest =
  | { type: 'apply'; id: string; commands: unknown[]; baseHash?: string; docRev?: number; dryRun?: boolean }
  | { type: 'read'; id: string; what: string; params: Record<string, unknown> }
  | { type: 'flush'; id: string }

export interface PageAckPayload {
  data?: unknown
  receipt?: SaveReceipt
  error?: { code: string; message: string }
}
```

`baseHash` / `docRev` 是命令参数，不是 locator。

### 4.4 字节通道

`window.workStudio.command` 底层是 `ipcRenderer.invoke`，structured clone 支持 typed array：

```ts
const loaded = await window.workStudio.command('docx-pro', 'docx-pro', 'load', { ref })
const bytes = loaded.ok && loaded.data.bytes instanceof Uint8Array
  ? loaded.data.bytes
  : base64ToBytes((loaded as { data: { data: string } }).data.data)

await window.workStudio.command('docx-pro', 'docx-pro', 'save', {
  ref,
  bytes,       // 契约外快路径；校验忽略未声明字段
  baseHash,
  type: 'page',
})
```

命令声明保留可选 `data: string`（base64）作为契约 / CLI 兜底；页面优先传 `bytes`。`commands.json` 仍可读，页面不承担 16 MB → 21 MB 的 base64 膨胀。

- 页面单次可编辑上限 **16 MB**；`load` 默认返回 `bytes`，失败时可带 `base64: true` 重取。
- 超过 16 MB：页面只读；Agent `apply` 因无写页面拒绝。
- 引擎 `load` 读取 64 MB 以内；`save` 拒绝空字节 / 非 docx / 超 16 MB。

### 4.5 并发、锁与租约

1. **写前 CAS**：`load` / `read` 返回 `hash`；`save` 带 `baseHash`，锁内重读文件计算 hash，不一致返回 `command.stale`。
2. **页面租约**：`attach` 登记；页面每 15 秒或每次事务后 `report` 更新 lastSeen。租约**不按时间失效**：页面关闭由新 `attach` 接管，客户端退出随进程清除。
3. **锁范围**：每个 `ref` 一把写锁，只有 `save` 取锁；`apply` 只入队、不取锁，避免 apply 等 ack、save 等 apply 锁的死锁。`read` 不取写锁。
4. **docRev**：`report` 更新；`apply` 带 `docRev`，页面比较不一致返回 `command.stale-index`；`baseHash` 在保存时做磁盘 CAS。

### 4.6 页面代理

```text
Agent → docx-pro.apply(commands: JSON 字符串)
  │
  ├─ 无页面 → command.page-unavailable（不写盘）
  │
  └─ 页面存在
        ├─ 引擎 JSON.parse + 校验，生成 requestId，挂起 Promise，放入 scope.pending
        ├─ scope.events.emit('request', { id, type })      ← 事件只带 id/type
        ├─ 页面订阅 docx-pro/events，收到后调 docx-pro.pending({ id })
        ├─ type=apply：执行 applyCommandBatch（PM 事务，整批回滚）
        │    └─ 落盘 → docx-pro.save({ bytes, type:'agent', baseHash })
        ├─ type=read：序列化 live 文档 → readMany
        ├─ type=flush：编辑器 save() → docx-pro.save
        └─ 页面 docx-pro.ack({ id, ok, payload: JSON.stringify({ data, receipt, error }) })
  │
  └─ 引擎结算 Promise；30s 超时返回 command.page-unavailable
```

- 每个 `ref` 只允许一个活跃页面；第二个 `attach` 返回 `already-open`。
- `dryRun`：页面执行前抓 PM JSON 快照，执行后恢复，不写盘。
- 超时 / `detach` / 页面崩溃：清理 pending，所有挂起请求失败，不留半批写入。
- 无 watcher、无 workspace 订阅：外部修改由下一次 `save` 的 `baseHash` 冲突发现，返回 `command.stale`。

---

## 5. 代码改造

### 5.1 新增

| 文件 | 作用 |
| --- | --- |
| `src/engine/index.ts` | `definePlugin` 入口；adapter、命令、`events`、`lifecycle` |
| `src/engine/resource.ts` | `docx` adapter：`read` / `grep` / `open` |
| `src/engine/commands/guide.ts` | `docx-pro.guide` |
| `src/engine/commands/read.ts` | `docx-pro.read` |
| `src/engine/commands/apply.ts` | `docx-pro.apply`；JSON 解析、入队；不写盘 |
| `src/engine/commands/save.ts` | `load` / `save`；typed array / base64、CAS |
| `src/engine/state.ts` | 租约、pending、超时、事件 |
| `src/engine/base64.ts` | base64 兜底编解码 |
| `src/page/lease.ts` | attach / pending / report / detach / ack |
| `src/page/bytes.ts` | typed array 快路径 + base64 兜底 + 16 MB 限制 |

### 5.2 重写

| 文件 | 改造 |
| --- | --- |
| `package.json` | 新清单、`exports['./workstudio']`、版本 `1.0.0`、去掉 `./ura` |
| `scripts/build.mjs` | esbuild 引擎 → `dist/engine.js`；Vite 页面；`workstudio-plugin .` 生成静态描述 |
| `scripts/verify.mjs` | 校验清单、exports、页面、commands.json 对齐、一次 load→save 往返 |
| `src/page/main.ts` | `?ref=` / `?locator=`；`docx-pro.load`；订阅 `docx-pro/events`；启动编辑器 |
| `src/page/host.ts` | `window.desktop` 适配：保存 → `docx-pro.save`；打开 → `load`；主题 → `shell/appearance` |
| `src/page/env.d.ts` | 只保留 `window.workStudio` 与页面内部类型 |
| `src/page/block-commands.ts` / `command-batch.ts` | 保留；补 `docRev`、`dryRun` 快照、`ack` |
| `src/engine/read.ts` | 去掉旧 ABI 依赖；纯函数化 |
| `src/tools/catalog.ts` | 保留 41 条目录；guide 加 `baseHash` / `docRev` |
| `skills/docx/SKILL.md` | 4 个助手命令、JSON 字符串、页面写 |

### 5.3 删除

- `src/ura.ts`
- `src/engine/{abi,plan,exec,runtime,tools,tools-legacy,create,zip}.ts`
- `src/shared/{ops,plan,sync,blocks-edit,runtime}.ts`
- `src/tools/declarations.generated.ts`
- `--legacy-tools` 构建与测试分支
- `assets/seed.docx` / `scripts/gen-seed.mjs`：保留为 `create` 的空白模板

`src/shared/index.ts` 裁剪导出；`src/shared/ipc.ts` 保留（编辑器 fork 的 `window.desktop` 类型仍用）。

### 5.4 编辑器 fork 最小改动

1. `__docxPro` 保留 `getCtx` / `getEditor` / `serialize`；新增 `reload(bytes)` / `replaceDocument(bytes)`（供页面重新加载用）。
2. `window.desktop.saveDocx*` → `docx-pro.save`；不再接收 `version/epoch`。
3. `consumePendingOpenDocx` 返回 `{ path: ref, name, dataUrl, hash }`；`consumeNewBlankDoc` 返回 `false`；AI 面板关闭。
4. 主题订阅 `shell/appearance`。
5. 打印 / 导出保持页面内浏览器能力。

### 5.5 技能

```text
助手只有 4 个命令（通过 command / bash）：
  docx-pro.guide   → 命令目录
  docx-pro.read    → 读 context / blocks / comments / notes / revisions / styles / audit
  docx-pro.create  → 创建空白 seed；不写正文
  docx-pro.apply   → 命令批；commands 是 JSON 字符串；要求文档页面已打开

没有 saveAs：另存为属于宿主文件管理。
没有无页面写：页面没打开就拒绝，不要尝试后台改文件。
写 .docx 仍然禁止 write / edit。
```

| 旧 | 新 |
| --- | --- |
| `docx_guide` | `command({ domain:'docx-pro', action:'guide', input:{ ref, group?, op? } })` |
| `docx_read` | 块列表用通用 `read`；其余用 `docx-pro.read` |
| `docx_create` | `docx-pro.create`：只创建空白 seed，然后打开页面再 apply |
| `docx_apply` | `command({ domain:'docx-pro', action:'apply', input:{ ref, commands: JSON.stringify([...]), baseHash?, docRev?, dryRun? } })` |
| `version` / `epoch` | `baseHash` / `docRev` |
| `command.page-unavailable` | 页面未打开；让用户打开文档 |
| `command.stale` | 文件已变；重新 `read` 后带新 `baseHash` 重发 |
| `command.stale-index` | 页面 `docRev` 已变；重新 `read` 后再发 |

---

## 6. 分阶段

### P0 — 约束确认（1 人日）

- 确认 typed array 在页面 ↔ 引擎的 invoke 往返；
- 确认 `commands` JSON 字符串在 `commands.json` / `command` 工具中可用；
- 确认 `docx-pro/events` 页面订阅可用；
- 确认插件不碰标签 / 文件管理。

### P1 — 引擎只读（2–3 人日）

- 新清单、`dist/engine.js`、`commands.json`；
- adapter `read/grep/open`；`guide/read`；
- 删除 URA / plan / exec / runtime / tools / create / zip。

**退出**：Agent `read` 返回带 locator 的块；`command_list` 只有 4 个助手命令。

### P2 — 页面宿主与人工保存（4–5 人日）

- `page/main.ts`：`ref` / `locator` / `load` / 编辑器；
- `page/host.ts`：`saveDocx*` → `docx-pro.save`；
- `attach` / `detach`；typed array 快路径 + base64 兜底；
- 16 MB 上限、只读 tier。

**退出**：打开→编辑→保存→重开正确；无页面零写盘；页面无 `ura.*`。

### P3 — 页面代理与 Agent 写（4–6 人日）

- `events`；`pending` / `report` / `ack`；30s 超时；
- `apply` JSON 解析；页面 blocks / editor / actions；`dryRun` 回滚；`read` live；`flush`；
- `docRev` stale-index；`baseHash` stale；无页面拒绝。

**退出**：页面打开时 41 条命令 parity；页面关闭时所有写失败且不落盘。

### P4 — 清理、测试、打包（3–4 人日）

- 删除 legacy / generated declarations；测试迁移；
- 安装 / 打包 / verify / CI / README / SKILL；
- e2e：安装、打开、编辑、Agent 写、关闭写拒绝、卸载；
- 版本 `1.0.0`；zip + sha256。

总估算：**14–19 人日**；两人并行约 2 周。

---

## 7. 测试与验收

| 用例 | 断言 |
| --- | --- |
| commands.json 对齐 | 动态命令都有静态描述；4 个助手命令；无 `saveAs/status` |
| guide | 三组 op / schema / example |
| read | 7 种 `what`；blocks 范围 / offset；locator |
| 无页面 apply | `command.page-unavailable`；`ctx.fs.write` 零调用 |
| 页面 apply | JSON 解析；入队 → ack；dryRun 不落盘；整批回滚 |
| load/save | typed array 快路径 + base64 兜底；非 docx / 空字节 / 超 16 MB / hash 冲突拒绝 |
| save 权限 | 助手来源被拒；页面来源可调用；无租约拒绝 |
| 页面队列 | attach → apply → ack；超时；detach 拒绝 pending |
| 文件管理 | 插件命令表无 create/saveAs；不订阅 workspace；不碰 dirty |

**验收**：

| 场景 | 通过标准 |
| --- | --- |
| 打开已存在文档 | 页面显示字节；编辑保存正确 |
| Agent 写 | 页面打开时生效；页面关闭时失败且不落盘 |
| 外观 | 标签 / dirty / 文件列表仍由宿主负责，插件不干预 |
| 并发 | 旧 hash 保存 `command.stale` |
| 打包 | zip 自包含；干净环境安装成功 |

---

## 8. 不改宿主的边界

| 能力 | 插件 | 说明 |
| --- | --- | --- |
| 打开已有 `.docx` | ✅ | 宿主给 ref，插件开页面 |
| 编辑 / 保存当前 ref | ✅ | 页面 `save` |
| Agent guide/read/apply | ✅ | apply 仅页面打开时 |
| 字节传输 | ✅ | invoke structured clone；base64 兜底 |
| 标签页 / dirty / 文件列表 | ❌ | **宿主 shell 职责，不是插件缺口** |
| 新建 / 另存为 / 重命名 | ❌ | **宿主文件管理职责** |
| workspace 树订阅 | ❌ | 插件不需要；外部修改由 `baseHash` 冲突发现 |
| 宿主 CAS/version | ✅（插件等价） | `baseHash` / `docRev` |

结论：不改宿主时，docx-pro 作为**文档页插件**是完整可交付的；不需要标签页、文件管理或新建能力。

---

## 9. 提交切分

1. `chore(docx-pro): add sdk manifest and remove ura entry`
2. `refactor(docx-pro): extract docx read adapter and guide/read commands`
3. `refactor(docx-pro): remove headless plan/exec and page-session bridge`
4. `feat(docx-pro): page load/save over workStudio commands`
5. `feat(docx-pro): page lease and event-driven agent proxy`
6. `test(docx-pro): no-page rejection, typed-array roundtrip, save CAS, queue`
7. `chore(docx-pro): packaging, skills and docs for SDK framework`

---

## 10. 开工前剩余问题（均为实现验证 / 产品取舍，不是宿主阻塞）

| 项 | 现状 | 处理 |
| --- | --- | --- |
| typed array 的 invoke 往返 | 设计上可用，未实测 | P0 用最小页面 + `docx-pro.load/save` 验证；失败则 base64，功能不变 |
| `commands` JSON 字符串 | 可进 `commands.json`，但 Agent 必须 `JSON.stringify` | 技能写示例；`guide` 返回 schema；不改成宿主未支持的对象参数 |
| 外部修改 | 不做 watcher；下一次 `save` 用 `baseHash` 发现冲突 | 页面提示冲突；若产品要求自动刷新，再加插件级 watcher，仍是插件范围 |
| 16 MB 上限 | 页面编辑上限，超过只读 | 保留编辑器 tier；不引入分块 |
| 新建 / 另存为 | 不属于插件 | 宿主创建并打开页面；插件只保存当前 ref |
| 标签 dirty / 文件列表 | 不属于插件 | 宿主 shell 负责；插件不订阅、不上报 |
| 旧插件冲突 | `docx` / `office-docx` 同扩展名 | 迁移期卸载旧插件；`claims.priority: 100` |
| 编辑器 fork 的 `window.desktop` 面 | 需要逐个方法接到新命令 | P2 对照 `App.tsx` 调用清单补齐；不改编辑器内部 |
| 页面队列超时 | 30s 超时 | 超时返回 `command.page-unavailable`；不留半批写入 |
| 打包 | 自包含 zip，无 node_modules | 引擎 bundle SDK + fflate；页面 Vite；commands.json 根目录 |

结论：**没有阻塞性宿主问题，可以开始整改。** 建议顺序：P0 验证 typed array + events → P1 引擎只读 → P2 页面 load/save → P3 Agent 页面队列 → P4 清理与打包。

---

## 11. 实现状态

已完成：

- `package.json` 改为 SDK 清单：`exports['./workstudio']`、`pages.content`、`claims`、`skills`；
- `src/engine/index.ts`：`definePlugin` + `docx` ResourceAdapter + `guide/read/create/apply/load/save/attach/pending/report/detach/ack`；
- 页面：`load` typed array / base64、`save` CAS、`attach` 租约、`events` → `pending` → `ack` 页面队列；
- 删除 URA / `plan` / `exec` / 旧 tools / legacy 声明 / 无页面写路径；
- `scripts/build.mjs`、`gen-commands.mjs`、`verify.mjs`、`install.mjs`、`pack.mjs` 适配 SDK 插件；
- `tests/sdk-engine.test.mjs` 覆盖 guide/read/load、无页面拒绝、页面队列、save CAS；
- 产物：`release/docx-pro-1.1.0.zip`（12.9 MB）。

未做（明确属于宿主）：

- 标签页、dirty、文件列表、新建、另存为；
- 宿主 GUI e2e（本环境未启动 WorkStudio）；
- typed array 页面实机验证（已有 base64 兜底）。
