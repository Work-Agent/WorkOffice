# docx-plus 设计方案

> 目标：在 WorkStudio 新插件框架（`@work-studio/sdk`）下，做一个**只含 Office DOC 能力、不含 GenOffice AI 智能体**的 DOCX 插件。
> 依据：
> - WorkStudio `0.3.3` 宿主源码（`src/preload/work-studio.ts`、`src/main/ipc/*`、`src/main/plugin/*`、`src/main/registry/*`）
> - GenOffice `vendor/genoffice`（commit `f6c2358c`）
> - 本仓库审计：[`docx-pro-genoffice-audit.md`](./docx-pro-genoffice-audit.md)、[`docx-pro-office-doc-scope.md`](./docx-pro-office-doc-scope.md)

---

## 0. 一页结论

docx-plus 不是“把 docx-pro 改个清单”，而是**保留编辑器 fork 和 DOCX 命令层，重写插件外壳与宿主接缝**：

```text
GenOffice DOCX 能力
├─ docx-engine（submodule，直接 git）
├─ editor fork（App/Ribbon/ProseMirror/分页，整包 fork）
├─ 文档命令层（ai/ops、ai/tools 等 19 个文件，fork）
└─ i18n/字体（数据直接 git + 少量 overlay）

WorkStudio 插件
├─ engine：definePlugin + ResourceAdapter + 命令 + 租约/pending
├─ page：window.workStudio（invoke/send/events）+ DocsHost
└─ 写模型：页面唯一落盘，无页面不写
```

| 决策 | 结论 |
| --- | --- |
| UI | **整包 fork**，App/Ribbon 成对；不 submodule、不混合路径 |
| 后端复用 | `src/genoffice/engine.ts` + `page.ts` 双 barrel，唯一白名单边界 |
| 通讯 | `workStudio.command`（invoke）+ `notify`（send）+ `events`（主→页） |
| 字节 | `Uint8Array` 快路径已实测可用；契约用 base64 兜底 |
| 写 | 只有页面 `save` / `saveAs` 写；Agent `apply` 只入队 |
| AI | GenOffice AI 面板/聊天/技能/provider 全部不取 |
| 范围 | **单文件**（`.docx/.docm/.dotx`）；不做文件夹 / 资源树 / 批量文件操作 |
| 宿主改动 | **零** |

---

## 1. 目标与非目标

### 1.1 目标

1. 打开、编辑、保存 `.docx`，保留 GenOffice 的 ProseMirror 编辑器、分页、Ribbon、审阅、打印/导出。
2. 页面与引擎只通过 `window.workStudio` 通讯，不出现 `ura.*`、`window.extension`、`window.desktop`。
3. 文档命令层（41 条目录、PM 事务、HTML 片段、图表、修订）与 docx-pro 对齐。
4. Agent 只在页面在场时写；没有页面就拒绝，不后台写、不第二份模型。
5. 页面壳里没有 GenOffice AI 入口、AI provider、AI 文案、AI 设置。
6. 复用 `vendor/genoffice` 的 DOCX 引擎、i18n、字体、非 AI UI 基础件。

### 1.2 非目标

- 不把 GenOffice 的 AI 面板/聊天/agent loop 迁进来。
- 不做无页面写、后台批处理、定时生成。
- 不改 WorkStudio 宿主、不改 SDK。
- 不重新实现编辑器；UI 是 fork，不是重写。
- **不做文件夹操作**：没有 `ls` / `find` / `mkdir` / `rm` / `mv` / `cp` / `openFolder` / `listFiles`，没有 `resourceGroups` / 资源树 / 目录递归 / 批量导入导出。

### 1.3 单文件范围（硬约束）

| 项 | 约束 |
| --- | --- |
| `ref` | 只能是 `file://` 指向的**单个 .docx/.docm/.dotx 文件** |
| 拒绝 | 目录、非文件、非 docx 扩展名、`http(s)://`、`plugin://`、工作区根 |
| 命令 | 只有 `load/status/save/saveAs/attach/pending/report/detach/ack` + `guide/read/apply` |
| 没有 | `ls/find/mkdir/rm/mv/cp/openFolder/listFiles/importFolder/exportFolder` |
| resources | 只有单文件 `docx`；`ls/find` 返回空；没有 `resourceGroups` / 目录树贡献 |
| `grep` | 只在**当前 docx 的块树**里检索，不是跨文件/文件夹 grep |
| `read` | 只读当前文件的块；`offset/limit` 是块条数，不是文件列表 |
| `saveAs` | 只写一个文件；`saveFile` 返回单文件 token；目标扩展名必须是 docx/docm/dotx；不创建目录 |
| watcher | 只 watch 当前文件；不递归、不 watch 目录 |
| 页面 | 一个页面 = 一个文件；`attach` 按文件 `ref` 建租约 |

> 文件夹由宿主的通用工作区域处理，**不属于 docx-plus**；docx-plus 不认领目录，也不返回目录条目。
>
> Agent 找文件应使用工作区的 `ls` / `find`（`workspace/<ws>/search`），拿到 `file://` ref 后再 `open` / `docx-plus.read` / `docx-plus.apply`。`resources.docx.ls/find` 返回空是正确语义：`.docx` 是单文件，不是目录。
>
> DOCX 内部是 zip 包（多个 part），但插件始终把它当作**一个文件**；zip 读写不是文件夹操作。`dataDir` / `tempDir` 是插件私有目录，只用于缓存/临时文件，不暴露给页面，也不做用户文件管理。

---

## 2. 宿主可行性结论（已验证）

| 能力 | 宿主实现 | 结论 |
| --- | --- | --- |
| 页面 → 引擎请求/响应 | `src/preload/work-studio.ts`：`command` = `ipcRenderer.invoke('workstudio:command', ...)`；`src/main/ipc/channel.ts` 用 `ipcMain.handle` 接 | ✅ invoke |
| 页面 → 引擎单向通知 | `notify` = `ipcRenderer.send('workstudio:notify', ...)` | ✅ send |
| 引擎 → 页面事件 | `events: true` → `<pluginId>/events`；`scope.events.emit()` → `resource.event()` → `webContents.send` | ✅ |
| 插件页调用自己的命令 | `commandDenialOf` 放行 `entry.address.domain === pluginId` | ✅ |
| 插件页调 `saveFile` | 放行 `shell/native/saveFile`；插件来源拿到 `save:` token，不拿路径 | ✅ |
| 页面订阅自己的事件 | `subscriptionDenialOf` 放行 `${pluginId}/...` | ✅ |
| 页面订阅 workspace | 插件页被拒绝 | ❌ 用 `docx-plus/events` + `status` |
| 页面上报 dirty/loaded | `shell/<ws>/view/<id>` 命令对插件页不开放；`setViewDirty` 还是空实现 | ❌ 页面内提示 |
| 页面打开不存在的 ref | 插件页不能调 `shell/pages/open`；但可 `location.href` 到同插件 ref | ⚠️ Save As 后导航可行 |
| typed array | WorkStudio 的 Electron 44.7、`sandbox:true` + `contextIsolation:true` 实测：`contextBridge` + `ipcRenderer.invoke` 双向传 `Uint8Array` 成功 | ✅ 快路径 |
| `commands.json` 参数 | `ArgSpec` 只有 `string/number/boolean/locator/enum`；未声明参数不校验 | ✅ `data` 可选 + `bytes` 快路径 |
| 插件 scope | `PluginScope` 有 `dataDir/tempDir/settings/secrets/events/files/log` | ✅ |

**关键修正：**

1. `window.workStudio.command` 返回 `CallResult<T>`，页面必须先判 `ok`。
2. `save` / `saveAs` 的 `data`（base64）必须是**可选**参数，否则 typed array 快路径会被判缺参。
3. 不依赖 `workspace` 订阅和 `dirty/loaded`；用 `events` + `status` + 页面内提示。
4. `ctx.fs.write` 只写当前工作区；`saveAs` 出工作区时用 `node:fs`（插件主进程有权限）。

---

## 3. 总体结构

```text
packages/docx-plus/
├─ package.json                  workstudio 清单 + exports['./workstudio']
├─ commands.json / api.json / api.d.ts   （workstudio-plugin 生成）
├─ skills/docx/SKILL.md
├─ src/
│  ├─ genoffice/                 唯一允许 import @genoffice 的地方
│  │  ├─ engine.ts               Node barrel
│  │  ├─ page.ts                 Browser barrel
│  │  └─ i18n.ts                 vendor 语言包 + 8 键 overlay
│  ├─ editor/                    GenOffice Docs renderer fork
│  │  ├─ App.tsx
│  │  ├─ components/
│  │  ├─ editor/
│  │  └─ host.ts                 DocsHost 接口
│  ├─ page/                      WorkStudio 页面适配
│  │  ├─ main.ts
│  │  ├─ host.ts                 DocsHost → window.workStudio
│  │  ├─ lease.ts
│  │  ├─ events.ts
│  │  ├─ bytes.ts
│  │  ├─ changes.ts
│  │  └─ save.ts
│  ├─ engine/                    definePlugin
│  │  ├─ index.ts
│  │  ├─ resource.ts
│  │  ├─ state.ts
│  │  ├─ watch.ts
│  │  ├─ base64.ts
│  │  └─ commands/
│  │     ├─ guide.ts
│  │     ├─ read.ts
│  │     ├─ apply.ts
│  │     ├─ load.ts
│  │     ├─ save.ts
│  │     └─ lease.ts
│  ├─ shared/                    纯类型 + 块命令/XML 切片
│  └─ tools/                     命令目录（41 条）
└─ dist/
   ├─ engine.js                  esbuild（Node ESM）
   └─ page/{index.html,assets/**} Vite
```

### 3.1 三层职责

| 层 | 职责 | 禁止 |
| --- | --- | --- |
| `src/genoffice/` | 唯一 submodule 边界；白名单导出 | 不出现 AI provider / agent loop；不做业务 |
| `src/editor/` | UI fork；只认 `DocsHost` | 不出现 `window.desktop` / `workStudio` / AI |
| `src/page/` + `src/engine/` | 新框架适配：命令、租约、事件、保存 | 不直接 import 编辑器内部状态 |

---

## 4. 清单

```jsonc
{
  "name": "@workstudio/plugin-docx-plus",
  "version": "1.0.0",
  "type": "module",
  "exports": { "./workstudio": "./dist/engine.js" },
  "workstudio": {
    "id": "docx-plus",
    "label": "Word 文档",
    "icon": "file-text",
    "pages": { "content": "dist/page/index.html" },
    "claims": [{ "extensions": ["docx", "docm", "dotx"], "priority": 100 }],
    "skills": ["skills/docx"]
  }
}
```

- 插件 id：`docx-plus`；
- 页面：`plugin://docx-plus/content?ref=<file:///...>&locator=...`；
- 迁移期卸载 `docx` / `docx-pro` / `office-docx`，避免 claim 冲突；`priority: 100` 保证 docx-plus 胜出。

---

## 5. 命令面

地址：`docx-plus/docx-plus/<action>`；页面调用：

```ts
await window.workStudio.command('docx-plus', 'docx-plus', 'load', { ref })
```

### 5.1 助手命令（3 个，`expose.agent: true`）

| 命令 | 参数（ArgSpec） | 语义 |
| --- | --- | --- |
| `docx-plus.guide` | `ref` string required；`group` enum(`blocks/editor/actions`)；`op` string | 命令目录；不读字节、不需要页面 |
| `docx-plus.read` | `ref` required；`what` enum；`startBlockIndex`/`endBlockIndex`/`offset` number | 页面在场读 live；无页面读磁盘 |
| `docx-plus.apply` | `ref` required；`commands` string required（JSON）；`baseHash` string；`docRev` number；`dryRun` boolean | 只校验 + 入队；无页面 `command.page-unavailable` |

### 5.2 页面命令（9 个，`expose.agent: false`）

| 命令 | 参数 | 语义 |
| --- | --- | --- |
| `load` | `ref` required；`base64` boolean | 返回 `{ ref, name, size, hash, tier, bytes?, data? }` |
| `status` | `ref` required | 返回 `{ hash, size, mtime, tier }` |
| `save` | `ref` required；`data` string **可选**；`baseHash` string；`type` enum(`page/agent`) | 页面落盘；优先 `bytes` 快路径，否则 base64 |
| `saveAs` | `ref` required；`data` string **可选**；`token` string required | 用 `save:` token 写新文件；返回新 ref |
| `attach` | `ref` required；`pageId` string required | 登记租约；返回 `{ hash, pending }` |
| `pending` | `id` string required | 取回请求正文（事件只带 id/type） |
| `report` | `ref`、`pageId`、`docRev`、`dirty`、`hash`、`selection` | 节流上报 live 状态 |
| `detach` | `ref`、`pageId` | 撤销租约 |
| `ack` | `id` required；`ok` boolean required；`payload` string required（JSON） | 结算 pending |

> `expose.agent` 默认 false；页面命令不显式打开。`apply` 必须显式 `expose: { agent: true }`。

**没有文件夹命令。** 本插件不提供 `ls` / `find` / `mkdir` / `rm` / `mv` / `cp` / `openFolder` / `listFiles` / `importFolder` / `exportFolder`；也不声明 `resourceGroups`。所有命令的 `ref` 都必须是单个 DOCX 文件。

### 5.3 错误码

| 错误码 | 触发 |
| --- | --- |
| `command.page-unavailable` | 无页面 / 租约过期 / 页面只读 |
| `command.stale` | `baseHash` 与磁盘 hash 不一致 |
| `command.stale-index` | `docRev` 与页面当前 revision 不一致 |
| `command.too-large` | 页面字节 > 16 MB |
| `command.invalid` | JSON / 参数 / 包结构非法 |
| `command.not-found` | ref 不存在 / 非 docx |
| `command.conflict` | 同一 ref 第二个页面 attach |

---

## 6. 写模型

### 6.1 不变量

1. **没有页面就没有写**：`apply` 无租约直接拒绝；不存在后台回退。
2. **页面唯一落盘**：只有 `save` / `saveAs` 写；两者都校验租约。
3. **Agent 只入队**：`apply` 不解析 XML、不写文件，只发事件等 `ack`。
4. **一次手势一个事务**：页面执行一批命令用一个 PM 事务；失败整批回滚。
5. **CAS 由插件自己做**：`baseHash` + `docRev`；宿主没有 version/epoch。
6. **draft 不落盘**：页面内 New 的 `draft` 租约禁止 `save`，只能 `saveAs`。

### 6.2 打开 / 保存

```text
打开
  host 打开 plugin://docx-plus/content?ref=...
  → page: command('load', { ref })
  → engine: 读字节 / hash / tier；返回 bytes 快路径或 base64
  → page: 启动编辑器 → command('attach', { ref, pageId })
  → engine: 建租约；有 pending 则一并返回

人工保存
  page: 编辑器 save() → serialize → bytes
  → command('save', { ref, bytes, baseHash, type:'page' })
  → engine: 取 per-ref 写锁 → 重新读盘算 hash → CAS → 写 → changed
  → page: 收到 changed，按已知 hash 忽略自身保存
```

### 6.3 Agent 写入

```text
Agent → command('apply', { ref, commands: JSON.stringify([...]), baseHash, docRev })
  → engine: 校验 JSON / 租约 / hash / docRev
  → pending.set(id, request)
  → scope.events.emit('request', { id, type:'apply' })
  → page: onUpdate 收到 request → command('pending', { id })
  → page: 解析命令批 → 一个 PM 事务执行（dryRun 则快照回滚）
  → page: command('save', { ref, bytes, type:'agent', baseHash })
  → page: command('ack', { id, ok, payload: JSON.stringify({ receipt }) })
  → engine: 结算 Promise
```

超时 30s → `command.page-unavailable`；页面 `detach` / 崩溃 → 清理 pending，全部失败。

### 6.4 外部修改

```text
engine attach 时 fs.watch(fileURLToPath(ref))，debounce 200ms
  → 算新 hash → scope.events.emit('external', { hash })
page 每 2s command('status') 兜底
  → 干净：自动 load 重载
  → 脏：提示保留本地 / 放弃本地 / 另存为
```

### 6.5 新建 / 另存为

```text
页面内 New（draft）→ 编辑 → Save As
  → command('shell', 'native', 'saveFile', { name })
  → 宿主返回 { token, name }（插件只拿 token）
  → command('saveAs', { ref, bytes, token })
  → engine: scope.files.resolveSave(token) → 绝对路径
  → engine: 写文件（工作区内用 ctx.fs，工作区外用 node:fs）
  → 返回 { ref: file:///... }
  → page: location.href = plugin://docx-plus/content?ref=<新 ref>
```

约束：

- 插件页面不能调 `shell/pages/open`；只能自己 `location.href` 到同插件 id 的地址；
- 目标在工作区外：可以写，但不导航；插件不维护资源树。
- 没有任何文档页可打开时，插件无法启动“新建”——这是宿主入口缺失，不是写路径。

---

## 7. 宿主接缝：DocsHost

编辑器 fork 里现在有 90 处 `window.desktop`。docx-plus 把它收敛为一个接口：

```ts
// src/editor/host.ts
export interface DocsHost {
  load(ref: string): Promise<{ bytes: Uint8Array; name: string; hash: string }>
  save(bytes: Uint8Array, opts: { baseHash?: string; type: 'page' | 'agent' }): Promise<SaveReceipt>
  saveAs(bytes: Uint8Array, token: string): Promise<SaveAsResult>
  status(ref: string): Promise<{ hash: string; size: number; mtime: number }>
  subscribe(cb: (event: DocEvent) => void): () => void
  theme(): 'light' | 'dark' | 'system'
  print?(): Promise<void>
  export?(kind: 'pdf' | 'html' | 'md'): Promise<void>
  requestSaveAs?(): Promise<{ token: string; name: string } | null>
}
```

- `App` / `Ribbon` 只通过 props/context 拿 `host`；
- `page/host.ts` 用 `window.workStudio` 实现；
- 迁移期保留 `desktop-shim.ts`，把旧 `window.desktop` 映射到 `DocsHost`，逐文件替换；
- fork 内不再出现 `window.desktop` / `window.workStudio` / `ura.*`。

---

## 8. UI fork 策略

### 8.1 整包 fork，不混合

- fork `apps/docs/src/renderer` 整包；
- 不把 App/Ribbon 走 fork、其他文件走 submodule；相对 import 太密，混合会带错版本；
- App + Ribbon 成对维护；Ribbon 不做跨格式组件库；
- 跨格式只共享 `@genoffice/ui` 基础件和 i18n 词表。

### 8.2 AI 剥离

保留：

```
ai/ops.ts  ai/tools.ts  ai/protocol.ts  ai/table-ops.ts  ai/field-ops.ts
ai/style-ops.ts  ai/note-ops.ts  ai/floating-ops.ts  ai/page-setup.ts
ai/revision-ops.ts  ai/comment-ops.ts  ai/doc-utils.ts  ai/doc-nav.ts
ai/doc-writer.ts  ai/after-text.ts  ai/lengths.ts  ai/inherit-formatting.ts
ai/tab-stops.ts  ai/pending-sections.ts  ai/redact.ts  ai/redact-view.ts
```

删除：

```
ai/AiPanel.tsx  ai/AiAskPopover.tsx  ai/EditQueueCard.tsx  ai/MindMapCard.tsx
ai/edit-queue.ts  ai/chat-retention.ts  ai/continuation.ts  ai/transport.ts
ai/docs-skill.ts  ai/files-skill.ts  ai/skills-skill.ts  ai/mind-map.ts
ai/revision-author.ts
i18n/ai/**  i18n/strings-ai.ts
```

`ai/tools.ts` 保留文档动作，删除 agent-only 分支：

```
web_search / image_search / analyze_media / generate_image
generate_mindmap / create_document
```

删除 App/main 的 AI 残留：

- `AiSettings`、`AI_PROVIDERS`、`getAiSettings()`；
- 6 个 `ai-*.css`、`applyAiPanelPrefs` / `getAiPanelPrefs`；
- Ribbon 的 `showAi` / `onToggleAi` / `onAiPreset` / `onNameWithAi`；
- 右键菜单 `onAiPreset`。

可选：把 `src/editor/ai/` 改名 `src/editor/doc-commands/`，消除“ai = 智能体”的误解。

### 8.3 i18n

- `i18n/app/*`、`i18n/ribbon/*`、`i18n/strings-*` 删本地副本，import 上游；
- 保留 `i18n/locale.tsx` + `i18n/strings.ts` 薄层；
- 8 个本地键 overlay：
  `ribbonStylesShow`、`ribbonMoreStyles`、`ribbonWatermarkCustomTitle`、`ribbonSelectTip`、`ribbonPenTip`、`ribbonHighlighterTip`、`ribbonPenStyle`、`ribbonHighlighterStyle`。
- 不 spread `aiStrings`。

### 8.4 字体 / 图标

- 27 个字体 + `LICENSE-OFL.txt` + 7 张图标：与上游逐字节一致，直接从 vendor 取；
- 7 个重切 CJK/KR 子集 + `fonts.css`：保留 fork，与本地 `line-metrics`/分页绑定；
- 9 张 `assets/file-*.png`：无引用，删除。

---

## 9. genoffice barrel（唯一 submodule 边界）

### 9.1 `src/genoffice/engine.ts`（Node）

```ts
export {
  parseDocx,
  saveDocx,
  // ... docx-engine 根 index 的浏览器安全导出
} from '@genoffice/docx-engine'
export { assertZipWithinLimits } from '@genoffice/zip-gate'
export {
  checkPlaceholders,
  placeholderSource,
  // ...
} from '@genoffice/agent-core/redact-core'
export type { WithheldNode, WithheldMark } from '@genoffice/agent-core/redact-range'
export type { AgentToolDef, ToolDisplay } from '@genoffice/agent-core'
```

约束：

- 不从 `@genoffice/docx-engine/zip-splice` 取（`node:fs` / `node:zlib`）；
- 不 import `@genoffice/agent-core` 根（会带 agent loop / skills / transport）；
- 不 import `@genoffice/ai-provider`。

### 9.2 `src/genoffice/page.ts`（Browser）

```ts
export { setAltChunkHtmlConverter } from '@genoffice/docx-engine'
export { createI18n, htmlLang, htmlDir, ribbonCommonEn } from '@genoffice/i18n'
export { Dropdown } from '@genoffice/ui/dropdown'
export { ColorPicker } from '@genoffice/ui/color-picker'
export { ShapePreview, WORDART_PRESETS, wordArtStrokePx } from '@genoffice/ui/shape-gallery'
export { ImageViewer } from '@genoffice/ui/image-viewer'
```

约束：

- 不使用 `@genoffice/ui` 根 index（会 re-export AI 组件 + `agent-core` + `ai-provider`）；
- 页面 bundle 不允许出现 `node:*`；
- `docx-engine` 只用根 index 的浏览器安全部分。

### 9.3 `src/genoffice/i18n.ts`

- 从 `@docs/i18n/strings-{app,ribbon,table,editor,zotero}` 取上游语言包；
- 合并 8 键 overlay；
- 不导入 `strings-ai`。

### 9.4 边界 gate

```bash
# 只有 src/genoffice/** 可以 import @genoffice
! grep -rn "from '@genoffice/" src --exclude-dir=genoffice

# 禁止 AI 根导入
! grep -rn "from '@genoffice/ui'\|from '@genoffice/agent-core'\|from '@genoffice/ai-provider" src
```

`src/editor` 的 fork 内部相对 import 不经过 barrel。

### 9.5 ResourceAdapter（`docx`）

```ts
const docx: ResourceAdapter = {
  // 单文件不是容器：即使宿主调用 ls/find，也不返回目录/子项
  ls: async () => [],
  find: async () => [],
  read: async (ref, options) => readItems(assertDocxFile(ref), options),
  grep: async (ref, query) => grepBlocks(assertDocxFile(ref), query),
  open: async (ref, locator) => ({ ref: assertDocxFile(ref), ...(locator ? { locator } : {}) }),
}
```

- **`ls` / `find` 显式返回空**：即使宿主把 docx-plus 当资源域调用，也不会列出目录；对文件夹的枚举由 workspace 通用域处理，不经过本插件；
- **不定义 `resourceGroups`**：不向宿主贡献任何目录/资源树；
- 第一条是文档元信息：`locator: 'docx:doc'`，带 `hash/size/name/blockCount/tier`；
- 块条目：`locator: 'docx:block:<docxIndex>'`，带 `index/type/heading/text/hash`；
- `read` 的 `offset/limit` 自己切片，默认最多 500 条；
- `grep` 只回当前文件的命中块，带 `snippet`，不回整篇，也不跨文件；
- adapter 没有 `PluginContext`：用 `node:fs/promises` + `fileURLToPath(ref)`；
- `assertDocxFile(ref)`：只接受 `file://`、扩展名 `docx/docm/dotx`、`stat.isFile()`；目录/非文件/非 docx 一律拒绝；
- 解析结果按 `ref + mtime/hash` 缓存，写后失效。

> 只有 `docx-plus.read` 需要 adapter；`guide` 不读字节。

---

## 10. 引擎状态机

```ts
interface Lease {
  ref: string
  pageId: string
  docRev: number
  dirty: boolean
  draft?: boolean
  hash: string
  selection?: string
  updatedAt: number
}

interface PendingRequest {
  id: string
  type: 'apply' | 'read' | 'flush'
  request: unknown
  createdAt: number
  timer: NodeJS.Timeout
  resolve(value: unknown): void
  reject(error: Error): void
}
```

规则：

- `PluginScope` 在 `lifecycle.activate(scope)` 时捕获，命令通过闭包使用 `scope.events` / `scope.files` / `scope.log`；`PluginContext` 本身不带 scope；
- 每个 `ref` 只允许一个活跃页面；第二个 `attach` 返回 `command.conflict`；
- 租约 TTL 30s；`report` 续期；
- `apply` 前校验 `baseHash` / `docRev`；
- `pending` 30s 超时；
- `detach` / 页面销毁 → 清理该 ref 的 pending；
- 写锁只在 `save` / `saveAs` 取；`apply` 不取锁，避免死锁。

---

## 11. 页面状态机

```text
boot
  ref = query.ref
  res = command('load', { ref })   // 先判 ok
  if bytes/data → start editor
  attach = command('attach', { ref, pageId })
  订阅 events → 处理 attach.pending

editing
  保存: editor.save() → command('save')
  外部: events.external / status 轮询 → 提示或重载
  Agent: events.request → pending → 执行 → save → ack

new/draft
  New → draft=true
  Save As → shell/native/saveFile → saveAs → location.href
```

页面内提示：

- 脏态：标题栏 / 状态栏 / beforeunload；
- 外部变化：保留本地 / 放弃本地 / 另存为；
- 无租约 / 只读：禁用保存。

---

## 12. 构建与打包

### 12.1 engine

```bash
esbuild src/engine/index.ts \
  --bundle --format=esm --platform=node --target=node22 \
  --outfile=dist/engine.js
```

- external：`node:*`；
- bundle：`@work-studio/sdk`、`fflate`；
- 入口只导出 `workstudio`（`definePlugin`），不在 import 时执行副作用。

### 12.2 page

```bash
vite build   # root = src/page，alias @genoffice/@docs → vendor
```

### 12.3 静态描述

```bash
pnpm exec workstudio-plugin .
# → commands.json / api.json / api.d.ts
```

### 12.4 zip

只含：

```
package.json
commands.json / api.json / api.d.ts
dist/engine.js
dist/page/**
skills/**
```

不含 `src`、`vendor`、`node_modules`。

---

## 13. 测试

### 13.1 单元

- `commands.json` 与 `definePlugin` 对齐；并断言命令表里**没有** `ls/find/mkdir/rm/mv/cp/openFolder/listFiles`；
- `resources` 只有 `docx`；`ls/find` 返回空；没有 `resourceGroups`；
- `ref` 守卫：目录、非文件、非 `docx/docm/dotx`、`http(s)://`、`plugin://` 全部拒绝；
- `apply` JSON 解析 / 参数校验 / 无页面拒绝；
- `baseHash` / `docRev` 冲突；
- `data` 可选 + `bytes` 快路径 + base64 兜底；
- `save` / `saveAs` 权限：页面可调、agent 被拒；
- `pending` 超时 / `detach` 清理；
- 路径守卫：只接受 `file://` + docx 扩展名；
- i18n overlay 8 键。

### 13.2 集成 / e2e

1. 安装 → 打开 `.docx` → 编辑 → 保存 → 重开字节正确；
2. Agent `apply` → 页面执行 → 保存 → 回执；
3. 页面关闭 → `apply` 立即失败且文件不变；
4. 外部修改 → `events.external` / `status` 发现 → 提示；
5. New + Save As → 新文件可重开；
6. 插件页调 `dirty/loaded/workspace subscribe` 被拒绝；
7. AI-free gate：页面源码无 `AiPanel` / `ai-provider` / `AI_PROVIDERS`；
8. 16 MB 边界：只读打开；超限保存拒绝；
9. 对文件夹 `ref` 调 `load` / `read` / `status` → `invalid`；docx-plus 不出现在目录条目里；
10. `grep` 只在当前文件块内命中，不跨文件。

---

## 14. 从 docx-pro 迁移

| docx-pro | docx-plus | 处理 |
| --- | --- | --- |
| `src/ura.ts` | — | 删除 |
| `src/engine/plan.ts` / `exec.ts` / `tools.ts` / `create.ts` / `zip.ts` / `runtime.ts` | — | 删除；逻辑拆进 commands/save |
| `src/page/main.ts` | `src/page/main.ts` | 重写：`ref` + load + attach |
| `src/page/host.ts` | `src/page/host.ts` | 重写：DocsHost → workStudio |
| `src/page/agent.ts` | `src/page/events.ts` | 重写：events + pending + ack |
| `src/shared/*` | `src/shared/*` | 保留块命令/XML；删除 plan/sync/runtime |
| `src/editor/*` | `src/editor/*` | 复用；AI 剥离；`window.desktop` → `DocsHost` |
| `src/tools/catalog.ts` | `src/tools/catalog.ts` | 保留 41 条目录；guide 改 baseHash/docRev |
| `src/engine/tools.ts` 四工具 | 3 个助手命令 | `guide/read/apply` |
| `docx_create` | — | 页面 New + Save As |
| `apiVersion: 2` / `write: plan` / `content: pulled` | — | 删除 |
| 任何 `resourceGroups` / 目录树 / 文件夹导入导出 | — | **不迁移**；docx-plus 只做单文件 |

---

## 15. 分阶段实施

| 阶段 | 内容 | 出口标准 | 人日 |
| --- | --- | --- | ---: |
| **P0** | 最小插件：`load`/`save` + typed array + events + saveFile token | 4 个 spike 全过 | 2–3 |
| **P1** | 引擎只读：adapter / guide / read | Agent 读到带 locator 的块 | 2–3 |
| **P2** | 页面 + 人工保存：load/attach/save/status/watch | 打开→编辑→保存→重开正确 | 5–6 |
| **P3** | Agent 写：apply/pending/ack/dryRun/stale | 41 条命令 parity；无页面写失败 | 4–6 |
| **P4** | 清理 / i18n / AI-free / 测试 / 打包 | 安装包 + e2e + gate 全过 | 3–4 |
| **合计** | | | **16–22** |

P0 spike 必须先验证：

1. `command` 传 `{ bytes: Uint8Array }` → 引擎 `instanceof Uint8Array`；
2. `events: true` + `scope.events.emit` → 页面 `onUpdate` 收到；
3. `saveFile` token → `saveAs` → 新文件可 `load`；
4. 插件页调 `shell/.../dirty` 被拒绝。

---

## 16. 风险与取舍

| 风险 | 影响 | 处理 |
| --- | --- | --- |
| 16 MB 页面上限 | 大文档不能编辑 | typed array 快路径；超限只读；后续要更大需宿主分块 |
| 无页面不写 | 不能后台批处理 | 产品接受；否则改宿主 |
| 无标签脏点 | 用户可能误关 | 页面内标题 / beforeunload / 保存提示 |
| 空工作区不能 New | 首次使用受限 | 从欢迎页/模板打开一个 docx 后 New |
| fork 漂移 | 上游 92 个 patch | pinned base + patch 清单；只做 feature port，不整包 re-fork |
| `@genoffice/ui` 根带 AI | 页面 bundle 混入 AI | 子路径 import + lint gate |
| `docx-engine` 有 Node 子路径 | 浏览器构建失败 | engine/page 双 barrel；页面不取 `zip-splice` |
| `saveAs` 工作区外 | `ctx.fs` 拒绝 | 用 `node:fs` 写 token 路径；不导航 |
| `command` 返回 `CallResult` | 示例代码误用 | 统一 `invoke<T>()` 包装，先判 `ok` |
| `data` 必填会破快路径 | typed array 被拒 | `data` 声明可选 |

---

## 17. 验收清单

- [ ] 插件安装后 `commands.json` 与 `definePlugin` 对齐；
- [ ] 命令表没有文件夹命令；`resources` 的 `ls/find` 返回空；没有 `resourceGroups`；
- [ ] 目录 / 非 docx / 非 file:// 的 ref 全部拒绝；
- [ ] 打开 / 编辑 / 保存 / 重开字节正确；
- [ ] 无页面 `apply` 失败且不写盘；
- [ ] 有页面 `apply` 41 条命令可用；
- [ ] 外部修改能发现并提示；
- [ ] New + Save As 可创建并重开；
- [ ] 页面源码无 GenOffice AI 入口/依赖；
- [ ] 页面无 `window.desktop` / `ura.*` / `window.extension`；
- [ ] 页面 bundle 无 `node:*`；
- [ ] 16 MB 边界正确；
- [ ] 关闭页面 / 崩溃后 pending 清理，无半批写入。

---

## 18. 一句话总结

> **docx-plus = GenOffice DOCX 编辑器 fork + `docx-engine` submodule + WorkStudio 新框架的“页面唯一落盘”命令外壳。**
> UI 是自己的 fork，引擎是 submodule，通讯是 invoke/send/events，AI 一律不进页面。

---

## 19. 宿主依据（代码位置）

| 结论 | 位置 |
| --- | --- |
| `command` = invoke | `WorkStudio/src/preload/work-studio.ts`、`src/main/ipc/channel.ts` |
| `notify` = send | 同上 |
| `onUpdate` / `events` 投递 | `src/main/ipc/channel.ts` 的 `sendUpdate`、`src/main/plugin/manager.ts` 的 `scope.events.emit` |
| 插件命令注册 | `src/main/plugin/manager.ts`（`commandResource = domainName`） |
| 页面来源鉴权 | `src/main/registry/authorization.ts`（`commandDenialOf` / `subscriptionDenialOf`） |
| 插件页身份 | `src/main/domains/shell/native.ts`（从 `plugin://` URL 解析 `pluginId` 并登记 origin） |
| `saveFile` / `save:` token | `src/main/domains/shell/domains.ts`、`src/main/plugin/files.ts` |
| `events: true` | `packages/sdk/src/contribution.ts`、`src/main/plugin/manager.ts` |
| `PluginScope` | `packages/sdk/src/scope.ts`、`src/main/plugin/manager.ts` 的 `createScope` |
| `ArgSpec` / 校验 | `packages/sdk/src/command.ts`、`src/main/registry/command.ts` |
| 输入必须是对象 | `src/main/ipc/dispatcher.ts` |
| typed array | `contextBridge` + `ipcRenderer.invoke`（WorkStudio Electron 44.7 实测，`sandbox:true`） |
| 页面 URL | `src/shared/pages.ts`、`src/main/page-router.ts`、`src/main/protocols.ts` |
| `ctx.fs` | `packages/sdk/src/context.ts`、`src/main/transport/workspace-files.ts` |

---

## 20. 执行就绪度（当前仓库）

| 项 | 状态 | 说明 |
| --- | --- | --- |
| 设计方案 | ✅ | 本文档 |
| 宿主能力验证 | ✅ | invoke/send/events、saveFile、typed array 已核对/实测 |
| `@work-studio/sdk` 依赖 | ✅ | `file:` 链接本地 WorkStudio SDK `0.2.0` |
| `packages/docx-plus` 骨架 | ✅ | P0/P1/P2 已落地 |
| 编辑器 fork 拷贝 | ✅ | 已从 `HEAD:packages/docx-pro/src/editor` 解出 |
| AI 残留清理 | ✅ | 已删 AI 组件/文案/不可达文件；`i18n/ai` 仅保留 `aiSum*` 摘要 key |
| `DocsHost` 收敛 | ✅ | fork 内统一 `getDocsHost()`；只剩 main 一处 `setDocsHost(window.desktop)` |
| genoffice 双 barrel | ✅ | `engine.ts` + 本地 UI 白名单 barrel；`page.ts`/i18n 后续 |
| build / pack / install | ✅ | 新框架清单 + `commands.json` 已跑通 |
| i18n overlay | ✅ | 上游 app/ribbon/table/editor/zotero + 8 键 overlay |
| P0 / P1 / P2 / P3 / P4 | ✅ | 见 §21–§26 |

**结论：P0–P4 已执行；剩余 DocsHost 窄接口与全量 GUI e2e。**

### 20.1 P0 的具体范围

```text
packages/docx-plus/
├─ package.json              workstudio 清单 + exports['./workstudio']
├─ tsconfig.json
├─ scripts/build.mjs         esbuild engine + vite page + workstudio-plugin
├─ src/engine/index.ts       definePlugin（只注册 load/save/status + events）
├─ src/engine/commands/
│  ├─ load.ts
│  ├─ save.ts
│  └─ status.ts
├─ src/page/index.html
├─ src/page/main.ts          ?ref= → load → 显示字节 → save
└─ commands.json / api.json / api.d.ts
```

P0 只验证 4 件事：

1. `command('load')` 返回 `Uint8Array`，页面 `instanceof Uint8Array` 为 true；
2. `events: true` + `scope.events.emit('ping')` → 页面 `onUpdate` 收到；
3. `shell/native/saveFile` → `save:` token → `save` 写回；
4. 插件页调 `shell/<ws>/view/<id>/dirty` 被拒绝（确认脏点不可用）。

P0 通过后再做：

- P1：`docx` adapter / `guide` / `read`
- P2：完整 editor fork + `DocsHost` + AI 清理
- P3：`apply` / `pending` / `ack` / `dryRun`
- P4：i18n / 字体 / 打包 / e2e

### 20.2 两个必须先定的产品决策

1. **编辑器 fork 基线**：用 `docx-pro/src/editor` 现有的 AI-stripped 版本（推荐，省 2,000 行 App 剥离），还是从 vendor 重新 fork 拿上游新功能？
2. **`saveAs` 工作区外**：v1 是否只允许工作区内（推荐，`ctx.fs` 足够），还是允许 `node:fs` 写到任意路径？

---

## 21. P0 执行结果（已落地）

已创建 `packages/docx-plus` 并完成 P0 骨架：

| 项 | 结果 |
| --- | --- |
| 清单 | `workstudio.id = docx-plus`、`pages.content`、`claims`、`skills` |
| 命令 | `load` / `status` / `save` / `saveAs` / `ping` |
| 引擎 | `definePlugin` + `events: true` + `docx` ResourceAdapter（`ls/find` 返回空） |
| 页面 | 原生 DOM：显示 hash/size/tier，按钮 Save / Save As / Status / Ping |
| 构建 | `pnpm --filter @workstudio/plugin-docx-plus build` ✅ |
| 类型检查 | `pnpm --filter @workstudio/plugin-docx-plus typecheck` ✅ |
| 单测 | `node --test tests/p0.test.mjs` 5/5 ✅ |
| 打包 | `node scripts/pack.mjs --only docx-plus` ✅ |
| 安装 | `node scripts/install.mjs --only docx-plus --target <dir>` ✅ |

P0 单测覆盖：

- `load` 返回 `Uint8Array` + hash；
- `save` typed array / base64 / `baseHash` 冲突；
- 单文件守卫：目录 / 非 docx / 非 `file://` / 不存在；
- `saveAs` 用 `save:` token 写单文件；
- `ping` 发出 `docx-plus/events`；
- `commands.json` 中 `data` 可选、没有 `bytes` 声明、没有文件夹命令；
- `resources.docx.ls/find` 返回空。

尚未做（需要 WorkStudio GUI / 宿主 e2e）：

- 在真实 WorkStudio 里打开 `.docx`，点 Save / Save As / Ping；
- 确认插件页调 `shell/<ws>/view/<id>/dirty` 被拒绝；
- P1：接 `docx-engine`、`guide/read`、editor fork。

---

## 22. P1 执行结果（已落地）

在 P0 基础上接入 `docx-engine`，开放 `guide` / `read`：

| 项 | 结果 |
| --- | --- |
| GenOffice 边界 | `src/genoffice/engine.ts`：`parseDocx` + `Block/ParsedDoc/Run` 类型 |
| 构建解析 | `build.mjs` 增加 `@genoffice/*` → `vendor/genoffice/packages/*/src` resolver |
| 命令目录 | `src/tools/catalog.ts`（8 条 block ops）+ `guide.ts` |
| 块树读取 | `src/engine/commands/read.ts`：`meta` / `blocks` / `context`，支持 `start/endBlockIndex`、`offset/limit`（上限 500） |
| 助手命令 | `docx-plus.guide`、`docx-plus.read`（`expose.agent: true`） |
| 单文件守卫 | 仍只接受 `file://` 单个 `.docx/.docm/.dotx`；`resources.docx.ls/find` 为空 |
| 单测 | `node --test tests/p0.test.mjs tests/p1.test.mjs` 8/8 ✅ |
| 类型检查 | `tsc --noEmit` ✅ |
| 构建 | `dist/engine.js` ≈ 1.4 MB（含 docx-engine / pptx-engine / zip-gate） |

P1 新增测试覆盖：

- `guide` 无参 / `group:blocks` / `op:insertParagraph` / 未知 op；
- `read` 的 `meta` / `blocks` / `context` / `offset+limit`；
- `insertParagraph` 等 block op schema 可从 `guide` 取到；
- `read` 拒绝非 docx；
- `guide/read` 的 `expose.agent` 为 true。

下一步 P2：把 editor fork + `DocsHost` 接进页面；P3 再开放 `apply`（页面事务 + `pending/ack`）。

P2 的 fork 基线建议：

> 用 **`HEAD:packages/docx-pro/src/editor`**（已提交的 AI-stripped fork）作为快照，
> 不要直接复制当前工作区的 `packages/docx-pro/src/editor`——它正在被另一条整改修改，
> 复制到一半的迁移状态会把 P2 的构建和 AI 剥离都弄脏。

---

## 23. P2 执行结果（已落地）

P2 把 editor fork 接进页面，先用 `window.desktop` 兼容层验证打开/编辑/保存：

| 项 | 结果 |
| --- | --- |
| editor fork | 从 `HEAD:packages/docx-pro/src/editor` 解出 361 个文件到 `src/editor` |
| AI 清理（第一轮） | 删 App 的 `AiSettings` / `AI_PROVIDERS`；删 main 的 AI CSS / `applyAiPanelPrefs`；删 9 个不可达 AI 文件与 `generate_mindmap` |
| 页面宿主 | `src/page/desktop.ts`：`window.desktop` → `docx-plus.load/save/saveAs` + `shell/native/saveFile` |
| UI 白名单 | `src/genoffice/ui.ts`：只 re-export Dropdown/ColorPicker/ShapePreview/ImageViewer/font-list 等非 AI 子路径；Vite alias `@genoffice/ui` → 本地 barrel |
| 构建 | Vite 页面 4.4 MB JS + 字体；`@genoffice` 走 vendor resolver |
| 类型检查 | `tsc --noEmit` ✅ |
| 单测 | P0/P1 8/8 ✅ |
| 页面 smoke | Playwright + mock `window.workStudio`：`.app` 渲染、Ribbon/页码/字数可见；`load` 调用成功；`window.desktop.saveDocx` → `docx-plus/save` 成功 |
| 打包 | `docx-plus-0.1.0.zip` ≈ 13.1 MB；安装目录 ≈ 22.7 MB |

P2 页面 smoke 结果：

```text
APP 1
CALLS [ 'docx-plus/load', 'docx-plus/save' ]
SAVE { ok: true }
```

尚未做：

- `window.desktop` 仍是兼容层（90 处 call site 未收敛成 `DocsHost`）；P3 再重构。
- `i18n/ai` 仍保留（`ai/tools.ts` 的文档动作摘要仍用 `aiSum*` key）。
- page bundle 仍有 `appExpandAiPanel` 等 AI 文案 key（无 AI 组件/依赖）。
- P3：`apply` + `pending/ack` + 页面事务。

---

## 24. P3 执行结果（已落地）

P3 把页面代理接进编辑器：Agent `apply` → 事件 → `pending` → PM 事务 → `save` → `ack`。

| 项 | 结果 |
| --- | --- |
| 引擎状态 | `src/engine/state.ts`：租约、pending、30s 超时、attach/detach/report |
| 命令面 | 新增 `apply`（助手）、`attach` / `pending` / `report` / `detach` / `ack`（页面） |
| 页面代理 | `src/page/agent.ts`：订阅 `docx-plus/events`、取 `pending`、执行 `applyCommandBatch`、保存、回执；15s heartbeat 续租 |
| 命令批 | 从 HEAD 复制 `page/block-commands.ts` / `page/command-batch.ts`；guide 升到 41 op（blocks 8 / editor 23 / actions 10） |
| 页面 id | `src/page/page-id.ts`：desktop 保存回执与 agent 租约共用 |
| 保存回执 | 页面保存后 `report(hash)` + `ack({receipt})`；桌面手动保存也 report hash |
| 类型检查 | `tsc --noEmit` ✅ |
| 单测 | 11/11 ✅（P0 5 + P1 3 + P3 3） |
| 页面 smoke | `attach` → `pending` → `applyCommandBatch` → `save` → `ack`；DOM 出现 `Hello P3`；无 console error |
| P3.1 live read | `read` 有租约时 `pending` → 页面 `serialize()` → `buildReadResult` → `ack.data`；失败/无租约回退磁盘 |

P3 smoke 结果：

```text
ATTACHED true
CALLS [ load, attach, report, pending, save, status, ack, report ]
ACK {"applied":1,"mutated":true,"receipt":{...}}
HAS_HELLO true
ERRORS []
```

P3 单测覆盖：

- 无租约 `apply` 拒绝；
- `attach` 返回 hash / pending；
- `apply` 只入队，`pending` 取正文，`ack` 结算；
- `baseHash` 冲突拒绝；
- `detach` 后 `apply` 拒绝；
- `apply` 对助手暴露；
- 未知 pending id 拒绝。

尚未做：

- `window.desktop` 兼容层未收敛成 `DocsHost`。
- `i18n/ai` 与 AI 文案 key 未清；无 AI 组件/依赖。
- Save As 后新 ref 的 page 导航 / 租约迁移待 P4。

---

## 25. P4 执行结果（已落地）

| 项 | 结果 |
| --- | --- |
| live read | `read` 有页面租约时 `pending(type:read)` → 页面 `serialize()` → `buildReadResult` → `ack.data`；无页面/失败回退磁盘 |
| Save As | v1 只允许工作区内；`ctx.fs.write(ref)` 写新文件；成功后 `location.assign(plugin://...?ref=新ref)` |
| Save As 拒绝 | 工作区外 `ctx.fs.write` 抛 denied，页面不会留下“已写但无法重开”的状态 |
| AI 文案 | 删除本地 `appAiRail` / `appAiSettings` / `appExpandAiPanel` 与 `ribbonAi*`；页面 bundle 无 `anthropic` / `agent-core` / `ai-provider`（vendor i18n 仍可能带 `appExpandAiPanel` 等 key 名） |
| 页面 smoke | `.app` 渲染、`attach`、`pending`、`apply`、`save`、`ack`、`HAS_P4 true`、无 console error |
| 单测 | 12/12 ✅ |
| 打包 | `docx-plus-0.4.0.zip` ≈ 13.2 MB；安装目录 ≈ 22.8 MB |

P4 smoke：

```text
APP 1
CALLS [ load, attach, report, pending, save, status, ack, report ]
HAS_P4 true
ERRORS []
```

仍未做：

- `window.desktop` 兼容层未收敛成 `DocsHost`（90 处 call site）。
- i18n overlay（上游语言包 + 8 键）仍沿用 fork 内旧语言包。
- `i18n/ai` 仍保留（文档动作摘要 key），但不含 AI 组件/provider。
- 全量 WorkStudio GUI e2e 仍需人工跑一次。

---

## 26. P4.1 执行结果（DocsHost + i18n overlay）

| 项 | 结果 |
| --- | --- |
| DocsHost 入口 | 新增 `src/editor/host.ts`；fork 内 93 处 `getDocsHost()`，只剩 `main.tsx` 里 `setDocsHost(window.desktop)` 一处全局赋值 |
| i18n overlay | `src/editor/i18n/strings.ts` 改为 `@docs/i18n/strings-{app,ribbon,table,editor,zotero}` + 本地 `aiStrings` + 8 键 overlay |
| 本地 i18n 删除 | 删除本地 `i18n/app` / `i18n/ribbon` 与 5 个 `strings-*.ts`（42+ 文件）；保留 `i18n/ai`（文档动作摘要 key） |
| Vite | `@docs/*` → `vendor/genoffice/apps/docs/src/renderer/*` resolver |
| AI 文案 | page bundle 无 `anthropic` / `agent-core` / `ai-provider`；vendor i18n 仍可能带 AI 文案 key 名 |
| 页面 smoke | `APP 1`、`ERRORS []` |
| 单测 | 12/12 ✅ |

剩余：

- `DocsHost` 类型仍等于 `DesktopApi`，还没拆成更窄的接口（下一步可拆 read/write/theme/print 四组）。
- `i18n/ai` 仍保留（`aiSum*` 文档动作摘要）；不包含 AI 组件/provider。
- 全量 WorkStudio GUI e2e 仍需人工跑一次。
