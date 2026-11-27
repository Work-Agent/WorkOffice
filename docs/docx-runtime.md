# docx 当前运行框架与数据流

> 对应实现：`packages/docx`（WorkStudio SDK 插件，不改宿主）。
> `docx` 负责“打开已存在的 ref 并读写这个文档”；`docx.create` 只创建空白 seed（不写正文）。标签页、dirty、文件列表、另存为归宿主。

## 1. 分层

```text
┌─ WorkStudio 宿主 ───────────────────────────────────────────────┐
│ PageRouter：把 ref 变成 plugin://docx/content?ref=...      │
│ preload：window.workStudio → ipcRenderer.invoke / send         │
│ Dispatcher：origin 鉴权 → DomainRegistry → 插件命令            │
│ PluginManager：加载 dist/engine.js、合并 commands.json         │
│ TransportRouter：generic read/grep → 插件的 docx adapter       │
│ WorkspaceFiles：ctx.fs.read/write 的落盘实现                    │
└───────────────────────────────┬─────────────────────────────────┘
                                │ invoke / send
┌─ docx 页面（plugin://） ──▼─────────────────────────────────┐
│ src/page/main.ts      加载字节、启动编辑器、attach 租约         │
│ src/page/host.ts      window.desktop → docx.save            │
│ src/page/lease.ts     events → pending → PM 事务 → save → ack  │
│ src/editor/**         GenOffice Docs 编辑器 fork（页面内 live） │
└───────────────────────────────┬─────────────────────────────────┘
                                │ 命令 / 事件
┌─ docx 引擎（dist/engine.js）▼──────────────────────────────┐
│ definePlugin({ resources, commands, events, lifecycle })       │
│ docx adapter：read / grep / open（Node fs + DOCX 解析）        │
│ guide / read / create / apply（助手）                          │
│ load / save / attach / pending / report / detach / ack（页面） │
│ leases: ref → pageId / docRev / hash                           │
│ pending: requestId → 等待页面 ack 的 Agent 请求                │
└───────────────────────────────┬─────────────────────────────────┘
                                │ ctx.fs / ctx.changed
┌─ 数据层 ───────────────────────▼────────────────────────────────┐
│ 工作区里的 .docx 文件（唯一磁盘真相）                           │
└─────────────────────────────────────────────────────────────────┘
```

## 2. 组件职责

| 层 | 负责 | 不负责 |
| --- | --- | --- |
| 宿主 shell | 标签页、dirty、文件列表、新建、另存为、窗口 | 文档内容解析、编辑事务 |
| 宿主 transport | 按 claims 找插件 adapter；提供 `ctx.fs` | 解释 DOCX |
| docx 引擎 | 读投影、写前 CAS、页面租约、Agent 队列 | 常驻文档模型、后台写 |
| docx 页面 | 编辑器、PM 事务、序列化、保存当前 ref | 文件管理、标签页 |
| 磁盘 `.docx` | 唯一持久化真相 | — |

引擎没有 DOCX 常驻模型：`read` 解析后返回条目；`apply` 不执行块命令；只有页面序列化出的整份字节会被写回。

## 3. 打开文档

```text
宿主资源面板
  → PageRouter.resolve(ref) → plugin://docx/content?ref=…&locator=…
  → 页面 main.ts
      1. 取 ref / locator
      2. 订阅 shell/appearance（主题）
      3. command docx.load { ref }
           → 引擎 ctx.fs.read(ref)
           → 返回 { bytes|data, hash, name, size, tier }
      4. 生成 Blob URL
      5. installDocxHost：把 window.desktop 接到 docx.save
      6. installPageQueue：订阅 docx/events
      7. command docx.attach { ref, pageId }
           → 引擎读当前 hash，登记租约
           → 返回 pending 请求（如果有）
      8. import editor/main → 编辑器消费 consumePendingOpenDocx()
      9. handlePending(attach.pending)
```

页面不订阅 `workspace`；主题订阅 `shell/appearance`，Agent 请求订阅 `docx/events`。

## 4. 人工保存

```text
编辑器 Save
  → buildDocBytes(ctx)（GenOffice 序列化整份 docx）
  → window.desktop.saveDocx(buffer)
  → page/host.ts commit()
  → command docx.save {
        ref, pageId, bytes, baseHash, type: 'page'
     }
  → 引擎（每个 ref 一把写锁）
       1. requireLease(ref)
       2. bytes instanceof Uint8Array / data base64
       3. 空字节 / >16MB / 非 DOCX 拒绝
       4. ctx.fs.read(ref) 当前字节 → sha256
       5. 与 baseHash 比较：不一致 → command.stale
       6. ctx.fs.write(ref, bytes)
       7. ctx.changed(ref, 'content')  // 宿主的资源面板刷新
       8. 更新 lease.hash
       9. 返回 { ref, hash, size, at, applied? }
  → host 更新 write.hash
  → 编辑器刷新自己的基线
```

`baseHash` 是插件级 CAS；宿主没有 `version/epoch`。外部修改不会被 watcher 自动发现；下一次保存时由 hash 冲突暴露。

## 5. Agent 读

### 5.1 通用 `read` 工具

```text
Agent read(ref)
  → 宿主 TransportRouter
  → docx adapter.read(ref)
  → Node fs 读取 + parseDocxBytes
  → [doc 元信息, block 条目…]
       locator: docx:doc / docx:block:<index>
       item: index / type / heading / text / hash
```

这条路径不经过页面，用于资源地图 / 通用内容读取。

### 5.2 `docx.read`

```text
Agent command docx.read
  ├─ 页面租约存在
  │     → 引擎生成 requestId，挂起，emit request
  │     → 页面 onUpdate → pending({id})
  │     → 页面 buildDocBytes + readMany
  │     → 页面 report({docRev, hash})
  │     → 页面 ack({data:{text, hash, docRev}})
  │     → 引擎返回：
  │          hash=<当前> docRev=<n>
  │          <readMany 文本>
  └─ 没有页面
        → ctx.fs.read → readMany
        → 返回 hash=<磁盘> docRev=-
```

页面打开时读 live 编辑状态；没有页面时读磁盘。

## 5.3 Agent 创建空白文档

```text
Agent command docx.create { ref }
  → 检查目标不存在
  → ctx.fs.write(ref, dist/seed.docx 空白模板)
  → ctx.changed(ref, 'structure')
  → 返回 { ref, size, hash, created:true }
  → 宿主 / 用户打开新 ref 的页面
  → 后续正文仍走 docx.apply（页面内）
```

`create` 只写空白模板，不写正文；这是唯一不在页面内发生的写入。

## 6. Agent 写

```text
Agent command docx.apply {
  ref,
  commands: '[{"op":"setText","index":1,"text":"新"}]',
  baseHash?, docRev?, dryRun?
}
  → 引擎 JSON.parse + 校验数组 / ≤200
  → requireLease(ref)：没有页面 → command.page-unavailable
  → docRev 比较：不一致 → command.stale-index
  → 生成 requestId，放入 pending，emit request
  → 页面 onUpdate → pending({id})
  → 页面 handleApply：
       1. editor.getJSON() 快照
       2. applyCommandBatch(editor, commands, ctx)
          - blocks：PM 直接编辑
          - editor：executeOps（一个事务）
          - actions：executeTool
       3. dryRun：恢复快照，ack；不写盘
       4. 有变更：host.setOrigin('agent') → save(ctx,false,true)
            → 同人工保存的 docx.save
       5. ack({receipt})
  → 引擎 resolve → Agent 得到 applied / hash
```

失败时页面恢复快照并 ack 错误；整批不落盘。页面关闭 / 崩溃 / 30s 超时，引擎拒绝所有挂起请求。

## 7. 无页面时

| 命令 | 行为 |
| --- | --- |
| `docx.guide` | 可用（不读字节） |
| `docx.read` | 读磁盘 |
| `docx.apply` | **拒绝**：`command.page-unavailable` |
| `docx.save` / `load` / `attach` / … | 页面命令；助手不可见 |

不存在“无页面写”，也不存在后台块级 XML 写。

## 8. 事件与状态

```text
scope.events.emit('request', { id, type })
  → DomainRegistry.publish event
  → 只有订阅了 docx/events 的页面收到
  → webContents.send('workstudio:update')
  → window.workStudio.onUpdate

页面 ack
  → window.workStudio.command('docx','docx','ack',{id,ok,payload})
  → 引擎 pending.resolve/reject
```

| 状态 | 位置 | 生命周期 |
| --- | --- | --- |
| 磁盘字节 | 工作区 `.docx` | 持久 |
| live 文档 | 页面编辑器（PM） | 页面 |
| `lease` | 引擎内存 `leases` | attach → detach / 新页面接管 / 进程退出 |
| `pending` | 引擎内存 `pending` | apply/read → ack / timeout |
| 命令表 | `commands.json` + `definePlugin` | 插件激活 |

## 9. 当前边界

- **不做**：标签页、dirty、文件列表、另存为、重命名；这些归宿主。`create` 只创建空白 seed。
- **不订阅** `workspace`；外部修改只在保存时通过 `baseHash` 冲突暴露。
- **不常驻** DOCX 模型；引擎只读投影、校验、入队、落盘。
- **写只经页面**：页面打开才写；页面序列化是唯一字节来源。
- **字节**：优先 `Uint8Array`（invoke structured clone），base64 兜底；单次 16 MB。
- **命令批**：`commands` 是 JSON 字符串（现有 `ArgSpec` 没有对象数组）。
- **新建**：`docx.create` 创建空白 seed 后，由宿主/用户打开页面；正文仍只经页面 `apply`。插件不为不存在的 ref 开页面。

## 10. 为什么插件自己维护租约（宿主代码证据）

宿主其实有视图生命周期，但它没有暴露给插件：

- `src/main/domains/shell/domains.ts`：
  - `viewResourceName(windowId, viewId)` → `<windowId>/view/<viewId>`；
  - 该资源有 `loaded` / `dirty` 命令；
  - `views.close` 会删除 view 资源。
- `src/main/domains/shell/native.ts`：
  - `webContents.on('destroyed')` 里触发 `onViewClosed(windowId, id)`；
  - 这是宿主内部的关闭通知。
- `src/main/domains/shell/module.ts`：
  - `host.onViewClosed(...)` 只删除 shell 的 view 记录和资源。
- `src/main/registry/authorization.ts`：
  - 插件页面只能订阅本插件域与 `shell/appearance`；
  - 插件页面不能调用 `shell/<wsId>/view/<id>/dirty|loaded`。
- `packages/sdk/src/bridge.ts` / `scope.ts`：
  - `window.workStudio` 只有 `workspaceId`，没有 `viewId`；
  - `PluginScope` 没有 view / page session / close 回调。

所以：

```text
宿主知道 view 开/关
  ≠ 插件知道 page 开/关
```

租约不是宿主要求，也不是安全边界；它是插件在“宿主不暴露 page session”前提下的协调层。  
如果宿主以后提供 page session（或允许插件订阅自己的 view 资源），租约可以删掉，直接使用宿主的 session。
