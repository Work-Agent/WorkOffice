# pptx（PowerPoint · `.pptx`）

WorkStudio SDK 文档页插件：完整 GenOffice Slides 编辑器跑在 `plugin://pptx/content?ref=…` 页面里，
引擎只做"校验 + 落盘"，助手写一律经页面队列。

- **打开**：页面一次 `pptx.load` 拿整份字节（不分块、不流式）
- **保存**：页面 `savePptx(opened)` 产整份字节 → `pptx.save`（`baseHash` CAS + 包校验 + 写盘）
- **助手写**：`pptx.apply` 只入队 → 页面 `runTxn`（一批一个 undo 步）→ 保存 → `ack`
- **无页面**：`apply` 直接 `command.page-unavailable`；唯一例外是 `pptx.create`

## 命令（13 条）

| 命令 | 助手 | 作用 |
| --- | --- | --- |
| `pptx.guide` | ✅ | 三级命令目录（7 组 / 75 op） |
| `pptx.read` | ✅ | context / slide / notes / comments / charts / tables / theme / audit / find；有页面读 live |
| `pptx.check` | ✅ | 风格清单 / outline 校验 / pages 对账 / 几何审计 / 版面骨架 |
| `pptx.create` | ✅ | 新建（seed + `buildCreateOps`，或设计稿 `pages` → 字节） |
| `pptx.apply` | ✅ | 只入队，等页面 `ack` |
| `pptx.load` | — | `{ name, size, hash, bytes }` |
| `pptx.save` | — | 页面唯一落盘点 |
| `pptx.saveAs` | — | `shell/native.saveFile` 的 token → 写工作区内目标 |
| `pptx.attach` / `pending` / `report` / `detach` / `ack` | — | 租约 5 min、心跳 15 s、请求超时 30 s |

地址固定 `pptx/pptx/<action>`；根目录 `commands.json` 由 `scripts/gen-commands.mjs` 生成，必须与引擎命令双向对齐。

## 目录

```text
src/
├─ engine/index.ts     definePlugin：adapter + 13 命令 + events + lifecycle（单文件，同 docx-pro）
├─ engine/{read,check,create,build,images,abi}.ts
├─ page/{main,host,channel,lease,env.d.ts,index.html}
├─ shared/             与页面共用的纯逻辑（deck / read / check / spec / ops / design / …）
├─ editor/             GenOffice Slides 整包（fork 常驻，改动只在接缝处）
└─ tools/              op 目录（op-catalog.generated.ts 由 gen-op-catalog 生成）
scripts/{build,gen-commands,gen-op-catalog,gen-seed,verify}.mjs
tests/{engine.test.mjs,harness.mjs}
assets/seed.pptx       create 的空白模板
skills/pptx/           SKILL.md / design.md / spec.md
```

## 构建 / 测试

```bash
pnpm --filter @workstudio/plugin-pptx build      # op 目录 → engine.js + commands.json → 页面
pnpm --filter @workstudio/plugin-pptx typecheck
pnpm --filter @workstudio/plugin-pptx test       # 11 个用例：只读 / 无页面拒绝 / 页面队列 / CAS
pnpm --filter @workstudio/plugin-pptx verify     # 清单 / 对齐 / 旧面扫描 / 一次往返
```

产物：

```text
dist/
├─ engine.js        exports['./workstudio']
├─ seed.pptx
└─ page/            ?ref= 页面
```

## 打包 / 安装

```bash
pnpm pack:pptx                              # release/pptx-0.2.0.zip (+ .sha256)
pnpm install:plugins --only pptx            # 复制到宿主插件目录（重启生效）
```

zip 里只有 `package.json`（精简清单）+ `dist/`（engine.js / seed.pptx / page）+ `commands.json` +
`api.*` + `skills/`；不含源码、node_modules、测试。`dist/engine.js` 与 `dist/seed.pptx` 必须同级
（`create` 用 `new URL('./seed.pptx', import.meta.url)` 找模板）。

## 限制

- 单文件 ≤ 64 MB（`load` / `save` 上限；超限 `command.too-large`，不分块）
- 外部修改不做 watcher：保存时用 `baseHash` 发现冲突 → `command.stale`
- 另存为只写当前工作区内（`ctx.fs` 的边界），跨盘另存会回 `denied`
