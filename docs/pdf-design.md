# PDF 插件改造方案（WorkStudio SDK 代）

> 对象：`packages/pdf`（改造前 `@workstudio/plugin-pdf@0.1.1`，URA apiVersion 2；改造后 0.2.0）
>
> 目标：迁到 WorkStudio 新插件框架（`definePlugin` + `exports['./workstudio']` +
> `pages.content` + `claims` + 命令 / 资源适配器 / 页面 bridge），保留现有
> PDF.js 字节渲染与查看器交互，并补上 agent 可读面。
>
> 结论先说：**PDF 不需要页面租约。** 它是只读文档，字节永远来自磁盘，
> 没有“页面模型才是权威”的问题。因此不照搬 PPTX 的
> `attach / pending / ack / report / save` 五件套，只保留
> `load / chunk / saveCopy` 三个页面命令和 `pdf.read` 一个助手命令。
>
> **实施状态**：P0（页面 / load / chunk / saveCopy）与 P1（resources / pdf.read /
> 文本 worker / locator / 外部变更）已落地；P2 文档级操作未实施。

---

## 0. 一页结论

改造分三层，边界和现有 `docx-pro` / `xlsx` / `pptx` 完全一致：

```text
宿主主进程（engine，可 import node:）
  definePlugin
    ├─ resources.pdf      read / grep / open / describe
    ├─ pdf.read           助手读：context / outline / pages / text / find
    └─ pdf.load / chunk / saveCopy    页面专用命令

插件页面（page，只认 window.workStudio）
  plugin://pdf/content?ref=…&locator=…
    ├─ pdf.load  → 元数据 + hash（小文件直接带 bytes）
    ├─ pdf.chunk → 大文件按 4 MB 取，页面预分配一次
    ├─ PDF.js 渲染 / 搜索 / 大纲 / 缩略图 / 打印
    └─ workspace/<ref> 的 changed 事件 → “文件已更新”横幅

共享层（shared，纯函数）
  pdf-meta（头部 / trailer / xref 探测）
  chunk（分块拼装）
  print / render-policy / geometry / search
```

三条不变量：

1. **页面不碰路径**：只拿 `ref`，所有字节经 `pdf.load` / `pdf.chunk` 命令；
2. **只读插件不写 `ref`**：默认没有 `ctx.fs.write` 调用；唯一写操作是
   `pdf.saveCopy` 写宿主签发的 save token；
3. **不再有 codec / presenter / sidecar**：`decode` 的校验逻辑进
   `pdf.load` 与 `resources.read`，`presenter` 由 `claims + pages.content`
   取代，`sidecar` 由 `resources.read` 取代。

---

## 1. 现状与差距

### 1.1 当前实现

```text
package.json
  exports['./ura'] = dist/ura.js
  workstudio.apiVersion = 2
  workstudio.content = 'pulled'
  workstudio.page = dist/page/index.html

src/ura.ts
  registerCodec({ id:'pdf', decode: 头部校验 + probePdfMeta + enrichMeta })
  registerPresenter({ surface:'pdf-viewer', data: { name, size, … } })
  registerSidecar({ content: summaryOf(doc) })

src/page/main.ts
  ?session=… → ura.page.get(session) → ura.page.bytes(session, range) × N
  → window.__pdfBoot → viewer

src/viewer/*
  PDF.js 渲染、大纲、缩略图、搜索、链接、打印、另存副本
  window.extension.bus.loaded / send(dirty)
  window.extension.ura.onChange
  window.appBridge.peek('/theme')
```

### 1.2 新框架要求

| 维度 | URA（现在） | SDK（目标） |
| --- | --- | --- |
| 入口 | `exports['./ura']` | `exports['./workstudio']` |
| 清单 | `apiVersion` / `content` / `page` | `pages.content` / `claims` |
| 注册面 | codec / presenter / sidecar | `definePlugin({ resources, commands })` |
| 页面身份 | `?session=` + 页面会话 | `?ref=`（可选 `locator`） |
| 读字节 | `ura.page.bytes(id, range)` | 自己的命令 `pdf.load` / `pdf.chunk` |
| 写字节 | `ura.page.saveBytes` | 没有；只读插件不写 |
| 变更 | `ura.onChange` | `subscribe('workspace', ref)` + `changed` 事件 |
| 主题 | `appBridge.peek('/theme')` / bus | `subscribe('shell', 'appearance')` |
| agent 读 | sidecar | `resources.read` / `resources.grep` + `pdf.read` |
| 静态描述 | 无 | `commands.json` / `api.json` / `api.d.ts` |
| 校验 | `verify-contract.mjs` 查 URA | 包内 `scripts/verify.mjs` 查 SDK 契约 |

### 1.3 必须保留的能力

- PDF.js 真字节渲染：canvas + textLayer + linkLayer；
- 大纲、缩略图、跨页搜索、页码跳转、缩放 / 旋转 / 适合宽高；
- 密码保护 PDF 的对话框与取消路径；
- 打印（当前是 150–200 DPI 栅格打印）；
- 另存副本；
- pdf-lib 的 ≤ 8 MB 元数据增强（标题 / 作者 / 首页尺寸 / 页数）；
- 128 MB 上限与“超限拒绝、不静默截断”的行为。

### 1.4 应删除的旧面

- `src/ura.ts`、`src/sdk.ts`、`src/ura-types.ts`；
- `window.extension`、`window.appBridge`、`ura.page.*`、`bus.loaded` / `bus.send`；
- `workstudio.apiVersion` / `content` / `page` / `write`；
- `sidecar` 的 `summaryOf` 双份实现（并入 engine 的 `describe` / `resource.read`）；
- 根 `scripts/verify-contract.mjs` 里的 PDF URA 条目；
- `tests/ura.test.mjs`。

---

## 2. 目标结构

```text
packages/pdf/
├─ src/
│  ├─ engine/                 # 主进程：宿主 exports['./workstudio']
│  │  ├─ index.ts             # definePlugin + commands + resources
│  │  ├─ ref.ts               # file:// / .pdf / 大小守卫；pathOf / statOf
│  │  ├─ load.ts              # readBytes + hash + probePdfMeta + enrichMeta
│  │  ├─ meta.ts              # summaryOf / itemOf（原 shared/meta.ts 的 engine 侧）
│  │  ├─ text.ts              # PDF.js 文本抽取（esbuild 单独 chunk）
│  │  ├─ resource.ts          # resources.pdf
│  │  └─ cache.ts             # 字节 LRU（大文件 chunk 用）+ 文本 LRU
│  ├─ page/                   # 页面：?ref=，只认 window.workStudio
│  │  ├─ channel.ts           # command / theme / watch / saveFile 薄封装
│  │  ├─ main.ts              # boot：load → 拼字节 → 装 __pdfBoot → viewer
│  │  ├─ env.d.ts             # WorkStudioBridge / LoadResult / PdfBoot 类型
│  │  └─ index.html
│  ├─ shared/                 # 纯函数，engine 与 page 共用
│  │  ├─ pdf-meta.ts          # 原样保留
│  │  ├─ pdf-types.ts         # PdfDoc / PdfPageMeta / PdfBoot
│  │  ├─ chunk.ts             # 分块拼装（保留预分配）
│  │  ├─ print.ts
│  │  ├─ render-policy.ts
│  │  └─ format.ts
│  └─ viewer/                 # React + PDF.js，基本原样
├─ skills/pdf/SKILL.md        # 可选，告诉 agent 用 read/grep/pdf.read
├─ scripts/
│  ├─ build.mjs               # engine + text chunk + page + pdfjs assets
│  ├─ gen-commands.mjs        # definePlugin → commands/api
│  └─ verify.mjs              # 清单 / 产物 / 静态动态对齐 / 旧面残留
├─ commands.json              # 生成并提交
├─ api.json / api.d.ts        # 生成并提交
└─ tests/
   ├─ engine.test.mjs         # 命令 + 资源 + 往返
   ├─ text.test.mjs           # 文本抽取（pdf-lib 现场造 PDF）
   └─ chunk / geometry / search / print / pdf-meta  # 保留
```

分层规则与 `docs/architecture.md` 一致：

- `engine/` 可以 `node:fs` / `node:crypto`，**不得** import React / viewer；
- `page/` 不得 import `node:*`，只经 `window.workStudio`；
- `shared/` 只放无副作用的纯函数。

---

## 3. 清单

```jsonc
{
  "name": "@workstudio/plugin-pdf",
  "version": "0.2.0",
  "type": "module",
  "exports": { "./workstudio": "./dist/engine.js" },
  "workstudio": {
    "id": "pdf",
    "label": "PDF 文档",
    "icon": "eye",
    "pages": { "content": "dist/page/index.html" },
    "claims": [{ "extensions": ["pdf"], "priority": 100 }],
    "skills": ["skills/pdf"]
  }
}
```

要点：

- 不再声明 `apiVersion` / `content: 'pulled'` / `page` / `write`；
- `claims` 只认 `.pdf`，`priority: 100` 与其它 Office 插件一致；
- 一个 zip 只装一个插件，`pack.mjs` 已经按 SDK 清单分支处理，无需改；
- `skills` 是可选但建议：agent 需要知道“读全文用 `pdf.read`，不要 grep 二进制”。

---

## 4. 引擎：命令与资源

### 4.1 命令表

| 命令 | 调用方 | `expose.agent` | 作用 |
| --- | --- | --- | --- |
| `pdf.read` | 助手 / 页面 | ✅ | 读元数据、大纲、页尺寸、页文本、全文查找 |
| `pdf.load` | 页面 | ❌ | 元数据 + hash；≤ 16 MB 直接带字节 |
| `pdf.chunk` | 页面 | ❌ | 大文件按范围取字节（4 MB/块） |
| `pdf.saveCopy` | 页面 | ❌ | 用 `shell/native.saveFile` 的 token 写副本 |

> 不设 `pdf.apply`：一期只读。若后续要支持旋转 / 删页 / 重排 / 抽取 / 合并，
> 再加 `pdf.apply`（见 §9 P2），写路径是“读字节 → pdf-lib → `ctx.fs.write` →
> `ctx.changed`”，仍然由命令完成，不做页面事务。

### 4.2 `pdf.read`

```ts
'pdf.read': {
  summary: '读 PDF：元数据 / 大纲 / 页尺寸 / 页文本 / 全文查找',
  args: {
    ref: { type: 'string', required: true },
    what: { enum: ['context', 'outline', 'pages', 'text', 'find'] },
    page: { type: 'number' },
    query: { type: 'string' },
    limit: { type: 'number' },
  },
  expose: { agent: true },
  run: (input, ctx) => readPdf(input, ctx),
}
```

语义：

- `context`：`name / size / pageCount / hash / encrypted / title / author /
  metadataWarning`，首行带 `hash` 与 `size` 令牌，供后续定位；
- `outline`：PDF.js 的大纲树（标题 + 目标页），无大纲时回 `(no outline)`；
- `pages`：`[{ page, width, height, rotate }]`，`limit` 默认 100；
- `text`：需要 `page`（1-based），返回该页文本，`limit` 是字符上限（默认 20k）；
- `find`：需要 `query`，逐页扫描，返回 `[{ page, snippet }]`，`limit` 默认 50。

加密 / 损坏 / 扫描件：

- `encrypted: true` 时 `text/find` 回
  `command.page-unavailable: 该 PDF 受密码保护，文本抽取需要页面解锁`；
- 没有可抽取文本时返回 `text: '(no extractable text)'`，不抛错；
- 扫描件 OCR 不在本期范围（见 §11）。

### 4.3 `pdf.load` / `pdf.chunk`

```ts
'pdf.load': {
  summary: '页面加载：元数据 + hash；小文件直接带字节',
  args: { ref: { type: 'string', required: true }, base64: { type: 'boolean' } },
  expose: { agent: false },
  run: async (input, ctx) => {
    const ref = refOf(input)                 // file:// + .pdf + 大小上限
    const bytes = await ctx.fs.read(ref)
    guardBytes(bytes)                        // 非空 / ≤ 128 MB / %PDF- 头
    const hash = sha256Hex(bytes)
    const meta = await readMeta(bytes, nameOf(ref), hash) // probe + enrich
    const inline = bytes.byteLength <= INLINE_BYTES
    if (!inline) cache.put(ref, hash, bytes) // 小文件字节已发给页面，再缓存是浪费
    const base = { ref, name: nameOf(ref), size: bytes.byteLength, hash, meta, inline }
    if (input.base64 === true) return { ...base, data: Buffer.from(bytes).toString('base64') }
    return inline ? { ...base, bytes } : base
  },
}

'pdf.chunk': {
  summary: '页面分块取字节',
  args: {
    ref: { type: 'string', required: true },
    start: { type: 'number', required: true },
    length: { type: 'number', required: true },
    hash: { type: 'string' },
  },
  expose: { agent: false },
  run: async (input, ctx) => {
    const ref = refOf(input)
    const bytes = await cache.getOrRead(ref, input.hash, () => ctx.fs.read(ref))
    const start = Math.max(0, Math.min(input.start, bytes.byteLength))
    const end = Math.max(start, Math.min(start + input.length, bytes.byteLength))
    return { ref, start, end, size: bytes.byteLength, hash, bytes: bytes.subarray(start, end) }
  },
}
```

设计理由：

- 旧实现一次 `page.bytes` 只读 1 MB，但页面仍要预分配整份；新实现把
  “整份读取”放到主进程，页面只收块。峰值与旧版相同（页面一份、主进程一份），
  但避免了 128 MB 单条 IPC 消息；
- `pdf.load` 已经算过 hash（要读整份），所以大文件把它留在 `cache` 里，
  `pdf.chunk` 直接切片；缓存 1–2 份、上限 160 MB，避免重复读盘；
- `hash` 作为版本令牌：`chunk` 时如果缓存被换出，重新读盘并比对；
  不一致就回 `command.stale`，页面提示重新打开；
- `base64` 是 IPC 不保留 `Uint8Array` 时的退路（`docx-pro.load` 已有同样处理）；
  大文件按块 base64，单块膨胀可接受。

### 4.4 `pdf.saveCopy`

```ts
'pdf.saveCopy': {
  summary: '另存副本（宿主 save token + 页面当前字节）',
  args: {
    ref: { type: 'string', required: true },
    token: { type: 'string', required: true },
  },
  expose: { agent: false },
  run: async (input) => {
    refOf(input)                                         // 只接受 file:// + .pdf
    const target = scope.files.resolveSave(input.token) // save:<uuid> → 绝对路径
    const bytes = toBytes(input)                       // input.bytes（≤ 64 MB）
    if (bytes.byteLength === 0) fail('command.invalid', '副本没有内容')
    if (bytes.byteLength > MAX_SAVE_BYTES) fail('command.too-large', '副本超过 64 MB')
    await writeFile(target, bytes)                     // 目标在工作区外，不能用 ctx.fs.write
    return { ok: true, size: bytes.byteLength }
  },
}
```

页面把“当前看到的字节”（重新取一份，避免 PDF.js transfer 后的空 buffer）
交给命令，而不是让引擎再读一次磁盘，这样副本内容与屏幕一致。

> 这是 SDK 目前的一个缺口：`ctx.fs.write` 只允许写工作区内 `ref`，
> 而 save token 指向用户选择的任意路径。二选一：
>
> - 一期：直接 `node:fs/promises.writeFile`（插件主进程有完整权限，`resolveSave`
>   已经校验 token 归属）；
> - 或者推动 SDK 增加 `ctx.files.write(token, bytes)`，插件不碰绝对路径。
>
> 页面侧保留 anchor 下载作为降级：拿不到 `shell/native.saveFile` 时仍可下载副本。

### 4.5 资源适配器

```ts
const resources = {
  pdf: {
    ls: async () => [],
    find: async () => [],
    read: async (ref, options) => {
      const file = await assertPdfFile(ref)
      // ≤16 MB：整份读 + probe + enrich + hash；更大：stat + 头尾探测
      const meta = await readResourceMeta(file)
      return [{
        locator: 'pdf:doc',
        kind: 'pdf',
        ref,
        name: file.name,
        size: file.size,
        hash: meta.hash ?? null,
        version: meta.hash ?? `${file.size}:${file.mtimeMs}`,
        pageCount: meta.pageCount ?? null,
        encrypted: meta.encrypted === true,
        title: meta.title ?? null,
        author: meta.author ?? null,
        text: summaryOf(meta, file.name),
      }]
    },
    grep: async (ref, query) => grepPages(ref, query),
    open: async (ref, locator) => ({ ref, ...(locator ? { locator } : {}) }),
    describe: () => ({
      root: 'file://',
      label: 'PDF 文档',
      summary:
        '单个 .pdf 文件：read 取元数据，grep 查页文本，pdf.read 读大纲/单页/全文，open 交给查看器。没有目录能力。',
      examples: [
        'read file:///w/报告.pdf',
        'grep file:///w/报告.pdf 收入',
        'command pdf/read {ref, what:"text", page:3}',
      ],
    }),
  },
}
```

关键取舍：**`read` 只返回元数据，不做文本抽取。**

原因：页面打开时会 `subscribe('workspace', ref)` 监听外部变更，
而 `workspace/<ref>` 的 snapshot 就是 `resources.read(ref)`。
如果 `read` 默认抽取全文，打开一个 500 页 PDF 会先在主进程跑一遍 PDF.js，
用户还没看到页面就先卡住。文本走两条显式路径：

- `resources.grep`：agent 明确要搜文本时才抽取；
- `pdf.read {what:'text'|'find'|'outline'}`：agent 明确要读时才抽取。

`SKILL.md` 里写清这一点即可。

大文件还有一层优化：`read` 不要为了元数据整份读 128 MB。
≤ 16 MB 时读整份走 `probePdfMeta` + `enrichMeta`；更大时只 `stat` +
头 1 KB 做 headerVersion 探测（必要时再读尾部），`pageCount` / `title` 允许为空，
`version` 用 `size:mtimeMs`。整份读取与 sha256 留给页面真正打开时的
`pdf.load`，`grep` / `pdf.read` 也按需再读。

### 4.6 生命周期与状态

```ts
let scope: PluginScope | null = null

export const workstudio = definePlugin({
  resources,
  commands,
  // 只读插件没有页面队列，不开 events 通道
  lifecycle: {
    activate(value) {
      scope = value
      value.log.info('pdf: read-only, 4 commands')
    },
    deactivate() {
      scope = null
      cache.clear()
    },
  },
})
```

没有 `attach / pending / ack / report / detach`，也没有 `leases` / `pending` /
`locks`。PDF 的“写”只有 `saveCopy`，且不改变 `ref` 的字节。

---

## 5. 页面

### 5.1 通道 `src/page/channel.ts`

```ts
const DOMAIN = 'pdf'

export const channel = {
  load: (ref: string, base64 = false) =>
    call<LoadResult>('load', { ref, ...(base64 ? { base64 } : {}) }),

  chunk: (input: { ref: string; start: number; length: number; hash?: string }) =>
    call<ChunkResult>('chunk', input),

  saveCopy: (input: { ref: string; token: string; bytes: Uint8Array }) =>
    call<{ ok: boolean; size: number }>('saveCopy', input),

  async pickSaveFile(name: string): Promise<string | null> {
    const result = await window.workStudio.command<{
      cancelled?: boolean
      token?: string
    }>('shell', 'native', 'saveFile', { name })
    if (result.ok !== true || !result.data || result.data.cancelled) return null
    return result.data.token ?? null
  },

  watch(ref: string, handler: () => void): () => void {
    void window.workStudio.subscribe('workspace', ref).catch(() => undefined)
    return window.workStudio.onUpdate((update) => {
      if (update.type !== 'event') return
      if (update.domain !== 'workspace' || update.resource !== ref) return
      if (update.event !== 'changed') return
      handler()
    })
  },

  async theme(): Promise<'light' | 'dark'> { /* subscribe shell/appearance，2s 超时回系统偏好 */ },
  onTheme(handler: (theme: 'light' | 'dark') => void): () => void { /* 同上 */ },
}
```

`subscribe('workspace', ref)` 的 snapshot 是元数据 item，代价很小；
`changed` 事件由 `ctx.changed(ref, 'content')` 发出（P2 写命令时用；
外部工具改文件时由宿主的工作区监听发出）。

### 5.2 Boot

```ts
async function boot(): Promise<void> {
  const params = new URLSearchParams(location.search)
  const ref = params.get('ref') ?? ''
  const locator = params.get('locator') ?? ''
  if (!ref) return fail('缺少 ref：请从工作区打开 PDF')

  await applyTheme(await channel.theme())

  const loaded = await channel.load(ref)
  const bytes = loaded.bytes ?? (await fetchChunks(ref, loaded.size, loaded.hash))

  window.__pdfBoot = {
    ref,
    data: loaded.meta,
    bytes,
    version: loaded.hash,
    refetch: async () => (await loadBytes(ref)),
    reload: async () => {
      const next = await channel.load(ref)
      return {
        bytes: next.bytes ?? (await fetchChunks(ref, next.size, next.hash)),
        data: next.meta,
        version: next.hash,
      }
    },
    watchExternal: (handler) => channel.watch(ref, handler),
    saveCopy: async () => {
      const token = await channel.pickSaveFile(loaded.name)
      if (!token) return false
      const bytes = await loadBytes(ref)   // 重新取一份，避免 PDF.js transfer 后的空 buffer
      await channel.saveCopy({ ref, token, bytes })
      return true
    },
    locator,
  }

  await import('../viewer/main')
}

async function fetchChunks(ref: string, size: number, hash: string): Promise<Uint8Array> {
  const out = new Uint8Array(size)
  for (let offset = 0; offset < size; offset += CHUNK_BYTES) {
    const end = Math.min(offset + CHUNK_BYTES, size)
    const chunk = await channel.chunk({ ref, start: offset, length: end - offset, hash })
    out.set(chunk.bytes, offset)
  }
  return out
}
```

`PdfBoot` 的变化：

```ts
export interface PdfBoot {
  ref: string
  data: PdfPageMeta        // { kind, name, ref, size, version, pageCount, encrypted, title, author? }
  bytes: Uint8Array
  version: string          // hash
  refetch(): Promise<Uint8Array>
  reload(): Promise<{ bytes: Uint8Array; data: PdfPageMeta; version: string }>
  watchExternal(handler: () => void): () => void
  saveCopy(bytes: Uint8Array): Promise<boolean>
  locator: string
}
```

去掉：`sessionId`、`degraded`（旧宿主降级）、`window.extension`、
`window.appBridge`。

> PDF.js 可能把 `bytes` 的 `ArrayBuffer` transfer 给 worker，`boot.bytes`
> 之后可能变成 detached。查看器只在首次 `openDocument(boot.bytes)` 使用它，
> 另存 / 重载 / 重试都走 `refetch` / `reload`，不受影响。

### 5.3 查看器改动点

| 文件 | 改动 |
| --- | --- |
| `viewer/App.tsx` | 去掉 `window.extension.bus.loaded()` / `send(dirty)`；外部变更改 `boot.watchExternal`；另存改 `boot.saveCopy`（保留 anchor 降级）；`boot.degraded` 改为 `metadataWarning` |
| `viewer/main.tsx` | 不变 |
| `viewer/PdfPage.tsx` | 不变（PDF.js 渲染 / textLayer / 链接） |
| `viewer/components/*` | 不变 |
| `viewer/i18n.ts` | 增加“文本不可用”“已定位到第 N 页”“另存失败”等词条 |
| `viewer/print.ts` | 不变（栅格打印）；宿主将来提供 vector print 再接 |
| `viewer/outline.ts` / `search.ts` / `geometry.ts` | 不变 |
| `shared/chunk.ts` | `ChunkReader` 签名改为 `pdf.chunk` 的 `{start,length}`；保留预分配 |
| `page/index.html` | 不变（`./main.ts`） |

`locator` 处理放在 boot 之后：`pdf:page:3` / `#p=3` / `3` 都映射到
`goToPage(3)`；这样资源面板点击、agent `open` 都能直达页码。

### 5.4 资源基址（重要）

新页面地址是 `plugin://pdf/content?ref=…`，没有 `index.html` 路径。
`new URL('pdfjs/', document.baseURI)` 会解析成 `plugin://pdf/pdfjs/`，
宿主协议表不会把它映射到页面目录。

正确做法：

```ts
// dist/page/assets/pdfjs/** 由构建脚本复制
export const ASSET_BASE = new URL('assets/pdfjs/', document.baseURI).href
// → plugin://pdf/assets/pdfjs/ → 协议表 kind === 'assets' 分支
```

构建时把 PDF.js 的 `cmaps / standard_fonts / wasm / iccs` 复制到
`dist/page/assets/pdfjs/`（而不是 `dist/page/pdfjs/`）；dev 模式对应
`src/page/assets/pdfjs/`。宿主对 `plugin://<id>/assets/*` 不做扩展名白名单，
`.bcmap` / `.pfb` / `.wasm` 都能取到；走其它路径会被 `ASSET_EXT` 拒绝。

---

## 6. 字节与内存预算

| 文件大小 | `pdf.load` | 页面 |
| --- | --- | --- |
| ≤ 16 MB | 元数据 + `hash` + `bytes` | 一次 IPC，直接渲染 |
| 16–128 MB | 元数据 + `hash`（不带 bytes） | `pdf.chunk` × N（4 MB/块） |
| > 128 MB | `command.too-large` | 拒绝打开，不截断 |

- `INLINE_BYTES = 16 MB`，`CHUNK_BYTES = 4 MB`，`MAX_PDF_BYTES = 128 MB`；
- 页面预分配一次 `Uint8Array(size)`，按块直填，避免 collect + concat 的 2× 峰值
  （沿用现有 `assembleBytes` 的思路）；
- 主进程 `bytesCache` 只缓存大文件（小文件字节已经发给页面，再缓存是浪费），
  LRU 1–2 份、上限 160 MB；
- `pdf.load` 对 128 MB 文件做一次 sha256（~百毫秒级），换来可靠的外部变更检测；
- 若 IPC 不保留 `Uint8Array`，`pdf.load` / `pdf.chunk` 接受 `base64: true`，
  页面按块解码。

---

## 7. 文本抽取

### 7.1 为什么放在 engine

- agent 读 PDF 时页面通常没打开（资源面板扫描、批量问答），文本不能依赖页面；
- `resources.grep` 在 engine 里执行，必须能直接读磁盘；
- 页面已经用 PDF.js 渲染，但把文本再走一遍页面队列会给只读插件引入
  `attach / pending / ack` 一整套租约，收益不大。

### 7.2 实现

```ts
// src/engine/text.ts
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'

export interface TextResult {
  pageCount: number
  pages: Array<{ page: number; text: string }>
  truncated?: boolean
  warning?: string
}

export async function extract(
  path: string,
  options: { pages: number[] | 'all'; maxChars: number; maxPages: number },
): Promise<TextResult> { … }
```

- 使用 legacy build（与查看器一致），`isEvalSupported: false`、
  `useSystemFonts: false`、`disableFontFace: true`、`enableXfa: false`；
- 只做 `getTextContent()`，不渲染、不碰 canvas；
- `pages: 'all'` 用于 grep，按 `maxPages`（默认 500）与 `maxChars`（默认 200k）
  截断，并在结果里标 `truncated`；
- 单页 `text` 默认 20k 字符，超出标 `truncated`；
- 加密 / 损坏 / 扫描件：捕获异常，回 `warning`，不阻断元数据读取；
- 文本缓存键 `ref + size + mtimeMs`，LRU 2 份、上限 32 MB；
  `ctx.changed` 时按 ref 失效。

### 7.3 构建与 worker

`pdfjs-dist` 不能和 `engine.js` 打成一个巨型 bundle，否则插件激活就多背 3 MB。
用 esbuild 的 code splitting：

```text
dist/engine.js        definePlugin（~几十 KB）
dist/text.js          PDF.js 文本抽取（~1 MB）
dist/pdf.worker.mjs   pdfjs 的 fake worker 文件（~2.3 MB）
```

`engine.js` 里 `await import('./text.js')`；`text.js` 的 `import.meta.url`
指向 `dist/text.js`，PDF.js 会在同目录找 `pdf.worker.mjs`，所以构建脚本必须

```js
copyFile(
  resolve('node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs'),
  join(root, 'dist/pdf.worker.mjs'),
)
```

已验证：Node 22 下 `getDocument({ data })` 走 fake worker 可以正常抽文本；
缺 `standardFontDataUrl` 只会打一条 warning，不影响文本。

### 7.4 主进程阻塞

fake worker 跑在主进程线程。500 页 PDF 的文本抽取可能到秒级，会卡 UI。
一期用 `maxPages / maxChars / 15s` 预算兜底；二期把 `dist/text.js`
作为 `node:worker_threads` 的 worker 入口（`workerData` 传 path/pages，
`parentPort.postMessage` 回结果，超时 `terminate`），主进程只等消息。

> 不把 PDF.js 的真实 Web Worker 搬进主进程：Node 没有全局 `Worker`，
> `worker_threads` 与 PDF.js 的 `MessagePort` 兼容性要额外验证。
> 直接把“整个抽取脚本”放进 worker_thread 更简单。

---

## 8. 改造清单（按文件）

### 新增

| 文件 | 内容 |
| --- | --- |
| `src/engine/index.ts` | `definePlugin`、命令表、resources、lifecycle |
| `src/engine/ref.ts` | `file://` 守卫、`.pdf`、大小上限、`pathOf` / `statOf` |
| `src/engine/load.ts` | `readBytes` / `hash` / `probePdfMeta` / `enrichMeta` |
| `src/engine/meta.ts` | `summaryOf`、resource item 组装 |
| `src/engine/text.ts` | PDF.js 文本抽取（worker 入口） |
| `src/engine/resource.ts` | `resources.pdf` |
| `src/engine/cache.ts` | 字节 LRU + 文本 LRU |
| `src/page/channel.ts` | `window.workStudio` 薄封装 |
| `scripts/gen-commands.mjs` | 从 `dist/engine.js` 生成静态描述 |
| `scripts/verify.mjs` | SDK 契约自检 |
| `skills/pdf/SKILL.md` | agent 使用说明（读元数据 / grep / pdf.read / open） |
| `tests/engine.test.mjs` | 命令 + 资源 + 往返 |
| `tests/text.test.mjs` | 现场用 pdf-lib 造 PDF，验证文本抽取 |

### 修改

| 文件 | 改动 |
| --- | --- |
| `package.json` | SDK 清单、`exports`、依赖、脚本、版本 0.2.0 |
| `src/page/main.ts` | `?ref=` + `pdf.load` / `chunk`；不再有 session / ura / bus |
| `src/page/env.d.ts` | 换成 `WorkStudioBridge` / `LoadResult` / `PdfBoot` |
| `src/shared/pdf-types.ts` | `PdfBoot` 去掉 sessionId，加 `ref` / `watchExternal` / `saveCopy` |
| `src/shared/chunk.ts` | 适配 `{start,length}` 分块读取；保留预分配 |
| `src/viewer/App.tsx` | 去 bus / appBridge；外部变更 / 另存 / locator 跳转 |
| `src/viewer/view-config.ts` | `ASSET_BASE` 改 `assets/pdfjs/` |
| `scripts/build.mjs` | engine + text chunk + worker + assets 复制 + gen-commands |
| `tsconfig.json` | 保持 `src/**/*`；engine / page 同仓 typecheck |
| `NOTICE` | 不变（pdfjs / pdf-lib 已在） |

### 删除

```text
src/ura.ts
src/sdk.ts
src/ura-types.ts
src/shared/meta.ts        # 内容并入 engine/load + engine/meta
tests/ura.test.mjs
```

### 仓库级

| 文件 | 改动 |
| --- | --- |
| `scripts/verify-contract.mjs` | 从 URA 列表移除 `pdf`（PDF 不再有 `dist/ura.js`） |
| `package.json` | `build:pdf` 保留；加 `verify:pdf`；`test:pdf` 改为 `node scripts/build.mjs --engine-only && node --test tests/*.test.mjs`，本地不依赖全量 build |
| `.github/workflows/ci.yml` | 用 `pnpm --filter @workstudio/plugin-pdf verify` 替代 URA 校验里的 PDF |
| `docs/architecture.md` | 代际表加 `pdf`；§3 加 PDF 映射；§7 能力边界更新 |
| `README.md` | 插件表加 `pdf` 行（只读查看器 + agent 读） |

---

## 9. 分阶段实施

### P0：页面可用（对齐现有 0.1.1 的查看能力）

1. 清单 / 目录 / `definePlugin` 骨架；
2. `ref.ts` + `load.ts` + `pdf.load` / `pdf.chunk`；
3. `page/channel.ts` + `page/main.ts` + `PdfBoot`；
4. `ASSET_BASE` 与 PDF.js 资源复制修正；
5. 查看器去 `window.extension`；
6. `gen-commands` + `verify` + `tests/engine.test.mjs`（先只测 load / chunk / 非 PDF / 超限）。

验收：

- `pnpm build:pdf` 产出 `dist/engine.js`、`dist/page/index.html`、
  `dist/page/assets/pdfjs/{cmaps,standard_fonts,wasm,iccs}`、`commands.json`；
- 打开 1 页 / 1000 页 / 加密 PDF：渲染、翻页、缩放、搜索、大纲、缩略图、
  链接、打印、另存副本全部正常；
- 1000 页文件走 `pdf.chunk`，峰值内存与旧版相当；
- 不存在 `window.extension` / `apiVersion` / `ura.page` 残留。

### P1：agent 可读

1. `text.ts` 文本抽取 + 缓存；
2. `resources.grep`、`pdf.read` 的 `outline/pages/text/find`；
3. `skills/pdf/SKILL.md`；
4. `tests/text.test.mjs`；
5. 外部变更 `watchExternal` + “文件已更新”横幅；
6. locator 跳页（`pdf:page:N`）。

验收：

- 页面关闭时：`read` 返回元数据；`grep` 命中页文本；
  `pdf.read {what:'context'|'outline'|'text'|'find'}` 正常；
- 页面打开时：同样结果，且不需要页面参与；
- 加密 PDF：`pdf.read text/find` 回明确错误，不崩；
- 扫描件：`context` 仍返回元数据，`grep` / `pdf.read text` 回空或 `(no extractable text)`；
- 外部改文件：页面出现横幅，点“重载”后页码 / 缩放保留；
- 500 页 PDF 的 grep 在预算内返回（P1 可接受秒级；P2 再上 worker_threads）。

### P2（可选）：文档级操作

只读 → 可写的最小扩展，仍然“命令写、页面订阅”：

```text
pdf.apply { ref, ops, baseHash? }
  ops: [{ op:'rotate', pages:[1,2], angle:90 },
        { op:'delete', pages:[3] },
        { op:'reorder', order:[1,3,2] },
        { op:'extract', pages:[1-3], token? },
        { op:'merge', refs:[...] }]
  → readBytes → pdf-lib → ctx.fs.write(ref) → ctx.changed(ref,'content')
```

- `expose.agent: true`，同时页面可调用；
- `baseHash` 做 CAS，冲突回 `command.stale`；
- 页面收到 `changed` 后显示“文件已更新”，用户确认后 `reload`；
- 抽取副本走 `shell/native.saveFile` token（与 `saveCopy` 同一套）；
- 不引入页面事务 / undo：PDF 页面不是权威模型，撤销由用户重开或宿主文件历史负责。

验收：旋转 / 删页 / 重排 / 抽取 / 合并没有损坏 xref；页面重载后页码正确；
助手在页面关闭时也能完成操作。

### P3（远期，不在本期）

- 表单填写 / 注释：PDF.js 只读渲染，保存注释需要额外的写入层；
- OCR（扫描件）；
- 宿主 vector print（当前是栅格）；
- 标签标题显示文件名（需要宿主 `ViewRecord.title` 支持）。

---

## 10. 风险与宿主缺口

| 风险 / 缺口 | 影响 | 处理 |
| --- | --- | --- |
| SDK 没有 range read | `pdf.chunk` 只能整份读 + 内存切片 | 主进程缓存 1–2 份；建议 SDK 加 `PluginFs.stat / readRange` |
| save token 目标在工作区外 | `ctx.fs.write` 会被 workspace 守卫拒绝 | 一期用 `node:fs/promises.writeFile`；建议 SDK 加 `ctx.files.write(token, bytes)` |
| `resources.read` 会被 `workspace` snapshot 调用 | 默认抽全文会拖慢打开 | `read` 只返回元数据；文本走 `grep` / `pdf.read` |
| PDF.js fake worker 在主进程 | 大文档抽文本卡 UI | 预算 + P2 worker_threads |
| `plugin://` 页面没有 `index.html` 路径 | 相对路径 `pdfjs/` 会解析到插件根 | 资源放 `dist/page/assets/pdfjs/`，`ASSET_BASE = assets/pdfjs/` |
| 宿主协议 `ASSET_EXT` 不含 `.bcmap/.pfb/.wasm` | 走错路径会 404 | 必须走 `plugin://<id>/assets/*` 分支 |
| PDF.js 可能 transfer 字节 | `boot.bytes` 之后不可用 | 另存 / 重载 / 重试都重新取字节 |
| 标准字体数据在 Node 不可 fetch | 抽文本有 warning | 不影响文本；可忽略或后续内置字体度量 |
| 加密 PDF | 文本不可读 | 页面解锁；agent 读回明确错误 |
| 扫描件 | 无文本层 | `grep` 回空、`pdf.read text` 回 `(no extractable text)`；OCR 不在本期 |
| 标签标题固定为插件 label | 多个 PDF 标签都叫“PDF 文档” | 宿主缺口；不影响功能，后续可加页面改标题命令 |
| 打印是栅格 | 文字不可选、矢量细节损失 | 现有行为；vector print 需宿主能力 |
| 包体积 +4 MB 左右 | 安装包变大 | `text.js` + `pdf.worker.mjs` 只在需要时加载；可接受 |

---

## 11. 不做的事

- 不把 PDF 做成可编辑文本编辑器；不重写 PDF 解析 / 渲染；
- 不引入第二套模型：页面与 agent 读的是同一份字节；
- 不做页面租约（只读，没有页面权威模型）；
- 不在页面里读路径、不直接写文件；
- 不 import AI / agent 包，不挂聊天面板；
- 不为了“像 Office”而复制 `attach / pending / ack`。

---

## 12. 验收清单

```text
构建
  pnpm build:pdf
  pnpm --filter @workstudio/plugin-pdf typecheck
  pnpm --filter @workstudio/plugin-pdf test
  pnpm --filter @workstudio/plugin-pdf verify

清单 / 产物
  exports['./workstudio'] === './dist/engine.js'
  workstudio.pages.content / claims.pdf / priority 100
  不存在 apiVersion / content / page / write / engine / tools
  dist/engine.js + dist/text.js + dist/pdf.worker.mjs
  dist/page/index.html + dist/page/assets/pdfjs/{cmaps,standard_fonts,wasm,iccs}
  commands.json 与 definePlugin 双向对齐
  pdf.read 暴露给助手；load/chunk/saveCopy 不暴露

运行
  小文件（≤16 MB）：一次 load 即渲染
  大文件（>16 MB）：load + chunk 拼装，128 MB 上限
  加密 PDF：密码对话框、取消、重试
  1000 页：缩略图 / 搜索 / 大纲 / 缩放 / 打印
  外部改文件：changed → 横幅 → 重载保留页码
  另存副本：saveFile token → 文件落盘
  agent：read 元数据 / grep 命中 / pdf.read text|find|outline
  locator：pdf:page:12 直接跳到第 12 页

残留
  rg "window\.extension|ura\.page|apiVersion|registerCodec|appBridge|sessionId"
  仅测试 / 文档里的历史说明命中
```

---

## 附：与现有 SDK 插件的对照

| 能力 | `docx-pro` | `xlsx` | `pptx` | `pdf`（本方案） |
| --- | --- | --- | --- | --- |
| 页面租约 | 有 | 有 | 有 | 无（只读） |
| 页面队列 | `attach/pending/ack` | gateway / 页面事务 | `attach/pending/ack` | 无 |
| 写形态 | `plan` / 页面事务 | 命令 | 命令 + 页面事务 | 无（P2 才加 `pdf.apply`） |
| 助手命令 | 3 | 多个 | 5 | 1（`pdf.read`） |
| 资源读 | `read/grep/open` | `read/open` | `read/grep/open` | `read/grep/open` |
| 页面字节 | `load` | wasm / chunks | `load` | `load` + `chunk` |
| 主题 | `shell/appearance` | `shell/appearance` | `shell/appearance` | `shell/appearance` |
| 变更 | `ctx.changed` | gateway publish | `ctx.changed` | `ctx.changed`（P2） |

差异只在“只读”这一条：没有租约、没有页面事务、没有保存命令。
其余边界与三个 Office 插件一致。
