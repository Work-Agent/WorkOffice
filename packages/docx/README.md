# docx-pro（Word · `.docx`）· WorkStudio SDK 文档页插件

`docx-pro` 只做一件事：**给定一个已存在的 `.docx` ref，打开文档页并读写这个文档**。

- 不管理标签页、文件列表、dirty、新建、另存为；这些属于宿主 shell / 文件管理。
- 页面打开才允许写；没有页面时 `docx-pro.apply` 直接拒绝，不做后台写。
- Agent 的写通过页面队列执行：`apply` → 事件 → 页面 PM 事务 → `save`。

## 清单与产物

```jsonc
{
  "exports": { "./workstudio": "./dist/engine.js" },
  "workstudio": {
    "id": "docx-pro",
    "pages": { "content": "dist/page/index.html" },
    "claims": [{ "extensions": ["docx", "docm", "dotx"], "priority": 100 }],
    "skills": ["skills/docx"]
  }
}
```

```text
dist/engine.js                definePlugin：ResourceAdapter + 命令 + 页面队列
dist/page/**                  GenOffice Docs 编辑器页面
commands.json / api.json      workstudio-plugin 等价的静态描述
skills/docx/SKILL.md          4 个助手命令与写入前提
```

## 命令

| 命令 | 可见 | 用途 |
| --- | --- | --- |
| `docx-pro.guide` | 助手 | 41 条 op 的三级目录 |
| `docx-pro.read` | 助手 | `context/blocks/comments/notes/revisions/styles/audit` |
| `docx-pro.create` | 助手 | 创建空白 seed 文件；不写正文；目标必须不存在 |
| `docx-pro.apply` | 助手 | 命令批；`commands` 是 JSON 字符串；无页面拒绝 |
| `docx-pro.load` | 页面 | 加载当前 ref 字节 |
| `docx-pro.save` | 页面 | 保存当前 ref（页面唯一落盘点） |
| `docx-pro.attach` / `report` / `detach` / `pending` / `ack` | 页面 | 页面租约与 Agent 队列 |

## 构建

```bash
pnpm --filter @workstudio/plugin-docx-pro build
pnpm --filter @workstudio/plugin-docx-pro typecheck
pnpm --filter @workstudio/plugin-docx-pro verify
```

页面构建依赖 `vendor/genoffice`；引擎 bundle 自带 SDK 等价入口与 `fflate`，插件 zip 不依赖 `node_modules`。

## 边界

- 标签页、dirty、文件列表、新建、另存为：宿主职责。
- 新建文档：`docx-pro.create` 只创建空白 seed，随后由宿主打开页面；正文仍经 `apply` 页面写。
- 外部修改：不订阅 workspace；保存时 `baseHash` 不一致返回 `command.stale`。
- 字节：页面 ↔ 引擎走 `window.workStudio.command`（invoke structured clone），`Uint8Array` 优先，base64 兜底。
