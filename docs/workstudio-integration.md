# WorkStudio 接入接缝

Office 插件不修改 WorkStudio 的插件模型，只要求四个通用接缝。它们都不是
Office 专用逻辑，任何二进制文档（PDF、OOXML、压缩包）都适用。

## 1. `file` transport 的格式提示

`src/main/ura/transports/fs.ts` 为 OOXML 扩展名给出具体提示，越具体越靠前：

```ts
docx: ['ooxml-docx', 'ooxml', 'binary']
xlsx: ['ooxml-xlsx', 'ooxml', 'binary']
pptx: ['ooxml-pptx', 'ooxml', 'binary']
pdf:  ['pdf', 'binary']
```

`http(s)` transport 按 MIME 给出同样的提示。否则所有二进制文件只会命中
`text` 兜底格式，文档会以文本编辑器打开。

## 2. 页面会话的二进制写

`src/shared/ura.ts`

```ts
export interface WriteConflict {
  base: string
  current: string
  mine: string
  encoding?: 'text' | 'base64'
}
```

`src/main/ura/facade.ts`

```ts
async readBytes(ref: string): Promise<{ bytes: Uint8Array; version: string }>
```

`src/main/ura/page-sessions.ts`

- `PageSessions.save(id, content: string | Uint8Array)`
- `PageSessions.saveBytes(id, bytes: Uint8Array)`
- 冲突稿的 `base / current / mine` 使用 base64，并带 `encoding: 'base64'`。

页面命令：

```
app:ura.page.saveBytes  { id, data: Uint8Array }  → CommitResult
```

preload 暴露：

```ts
window.extension.ura.page.saveBytes(id, data)
```

页面会话仍持有唯一版本基线；页面不接触版本号。保存是严格写，冲突不覆盖。

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

插件侧的 `src/sdk.ts` 只做类型化转发：

```ts
registerCodec(codec)
registerConsumer(consumer)
```

内置插件可以继续直接 import `src/main/ura/sdk`；外部插件必须走注入面。
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
