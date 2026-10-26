# WorkStudio 接入接缝

Office 插件不修改 WorkStudio 的插件模型，只要求四个通用接缝。它们都不是
Office 专用逻辑，任何二进制文档（PDF、OOXML、压缩包）都适用。

## 1. 格式分发：MIME，而不是 hints

新版 URA 不再使用 `hints` / `docType`。codec 直接声明自己能处理的
`(transport, mime)`：

```ts
registerCodec({
  id: 'docx',
  keys: [
    {
      transport: '*',
      mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    },
  ],
  extensions: ['docx', 'docm', 'dotx'],
  patchType: 'bytes',
  decode,
  locate,
  encode,
})
```

来源负责给出 MIME：

- `file` transport 按扩展名映射：`docx` / `docm` / `dotx` →
  `application/vnd.openxmlformats-officedocument.wordprocessingml.document`；
  `xlsx` / `xlsm` / `xltx` → spreadsheet MIME；`pptx` / `pptm` / `potx` →
  presentation MIME。
- `http(s)` transport 使用响应 `Content-Type`；服务端返回
  `application/octet-stream` 时不会命中专用 codec（需要来源侧补 MIME，
  或由 WorkStudio 在 `Content-Type` 缺失时按 URL 扩展名兜底）。

没有匹配 codec 时管道降级 raw；页面出口与 agent 文本出口分别由
presenter / sidecar 决定。

## 2. 页面会话的二进制写

`src/shared/ura.ts`

```ts
export interface CommitResult {
  ok: true
  version: string
}
```

`src/main/ura/facade.ts`

```ts
async readBytes(ref: string): Promise<{ bytes: Uint8Array; version: string }>
async write(ref: string, content: string | Uint8Array): Promise<{ version: string }>
```

`src/main/ura/page-sessions.ts`

- `PageSessions.save(id, content: string | Uint8Array)`
- `PageSessions.saveBytes(id, bytes: Uint8Array)`
- 提交返回 `{ ok: true, version }`
- 写入按资源键串行，后写覆盖先写；页面不用版本号做条件写，也不处理冲突稿

页面命令：

```
app:ura.page.saveBytes  { id, data: Uint8Array }  → CommitResult
```

preload 暴露：

```ts
window.extension.ura.page.saveBytes(id, data)
```

页面只拿会话 id、页面入口与 presenter 投递的数据；路径、写队列和版本
都在主进程。

## 3. `app://page` 的资源路径

插件页地址必须是“保留斜杠、逐段编码”的绝对路径：

```
app://page/Users/me/plugins/docx/dist/page/index.html?session=…
```

不能把整条绝对路径 `encodeURIComponent` 成单段，否则页面里的
`./assets/main.js` 会解析成 `app://page/assets/main.js`，落到协议根目录，
被 `containsPage` 拒绝。

`src/main/capabilities/webview/presets.ts` 提供 `pageUrlFor(file)`；协议处理器
解码后保留前导 `/`，并做路径归一化。

## 4. 外部插件的 SDK 注入

外部插件包无法相对引用 WorkStudio 源码。`PluginHost` 在装载插件前注入：

```ts
globalThis.__workstudioSdk = uraSdk
```

插件侧的 `src/sdk.ts` 只做类型化转发。新版注册面：

```ts
registerTransport(transport)
registerCodec(codec)
registerStore(store)
registerPresenter(presenter)
registerSidecar(provider)
registerCollection(provider)
registerViewAction(action)
```

Office 插件使用其中三个：

```ts
registerCodec(codec)
registerPresenter(editorPresenter)
registerSidecar(agentSidecar)
```

- `codec`：`{ id, keys, extensions, patchType, decode, locate, encode }`
- `presenter`：`{ key, type: 'page', surface, present }`
- `sidecar`：`{ key, sidecar }`

`presenter` / `sidecar` 的键是 `(transport, mime, codec)`：

```ts
{ transport: '*', mime: MIME.docx, codec: 'docx' }
```

`present()` 返回的 `page` 由宿主按插件清单解析，插件入口里留空串即可。
`sidecar()` 只负责 agent 读到的文本；页面渲染不经过它。

装载插件仍然是主进程权限模型：安装即信任，页面仍是无 Node 的沙箱。

## 5. 页面层归属

`packages/docx/src/editor` 是从 GenOffice Docs 页面层 fork 下来的源码，
由本仓库持有：

- `src/page` = WorkStudio 页面入口（会话数据、主题、启动）；
- `src/editor` = 文档编辑器（Ribbon / 分页 / 选区 / 保存序列化）；
- `src/editor/host` 边界 = 页面会话 ↔ 编辑器运行时。

AI 入口已在 fork 内移除，而不是靠宿主 CSS / 开关隐藏。后续要接
WorkStudio 的 AI 能力时，走宿主助手或显式注入的 action，不恢复
GenOffice 的 AI 面板与提示词。
