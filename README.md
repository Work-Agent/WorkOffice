# WorkStudio Office 插件

四个 WorkStudio SDK 文档页插件：`docx` / `xlsx` / `pptx` / `pdf`。页面层复用 GenOffice
的 renderer、engine 与操作层，宿主侧只接 `window.workStudio.command` / `subscribe` /
`onUpdate`：写一律是命令，页面不在场就不写。

整体架构、GenOffice 源码分层、Office 功能解构与封装方式：

→ [`docs/architecture.md`](./docs/architecture.md)

## 插件

| 包 | id | 页面 | 说明 |
| --- | --- | --- | --- |
| `@workstudio/plugin-docx` | `docx` | GenOffice Docs 编辑器 fork | 完整编辑器 + 块命令引擎：11 个命令（`guide` / `read` / `create` / `apply` / `load` / `save` + 租约 `attach` / `pending` / `report` / `detach` / `ack`）、41 条 op；有页面走编辑器事务，无页面走块级 XML 切片 |
| `@workstudio/plugin-xlsx` | `xlsx` | GenOffice Sheets 页面层 + `__xlsxPlus` 桥 | 单工作簿、命令 + 租约令牌；wasm 流式打开，保存/另存/自动保存/恢复副本，agent 读与写 |
| `@workstudio/plugin-pptx` | `pptx` | GenOffice Slides 编辑器 | 13 个命令；打开一次 `load`、保存一次 `save`，助手写经页面队列（无页面即拒） |
| `@workstudio/plugin-pdf` | `pdf` | PDF.js 查看器 | 只读：`pdf.read` / `pdf.load` / `pdf.chunk` / `pdf.saveCopy`；渲染、大纲、搜索、打印、另存副本 |

> 四个插件都用 `claims` 认领扩展名并设 `priority: 100`：`docx/docm/dotx`、`xlsx/xlsm/xltx`、
> `pptx/pptm/potx`、`pdf`。同一扩展名只能有一个插件生效。

> 插件入口统一是 `definePlugin`（`export const workstudio = ...`），不注册 codec / presenter /
> sidecar，也没有 `ura.page` / `window.extension`。旧 URA 代际与历史 `*-pro` 包已删除，
> 历史文档已随代际删除。

## 仓库结构

```text
packages/
├─ docx/                      # id: docx
│  └─ src/{engine,page,plugin,shared,editor,tools}
├─ xlsx/                      # id: xlsx
│  └─ src/{engine,page,shared,editor,editor-shared,tools}
├─ pptx/                      # id: pptx
│  └─ src/{engine,page,shared,editor,tools}
├─ pdf/                       # id: pdf
│  └─ src/{engine,page,shared,viewer}
├─ vendor/genoffice/          # submodule：sparse checkout 的 GenOffice 源码
scripts/
├─ setup-genoffice.mjs        # 拉取 / 稀疏化 GenOffice
├─ install.mjs                # 装到 WorkStudio 插件目录
├─ pack.mjs                   # 产出 release/<id>-<version>.zip + sha256
└─ write-registry-entry.mjs   # 写 PluginRegistry 的 plugins/<id>.json
docs/architecture.md          # 唯一架构文档
```

`workstudio.id` 与目录名一致（`docx` / `xlsx` / `pptx` / `pdf`）；每个包的
`skills/<id>/SKILL.md` 随 zip 分发。XLSX / PPTX 的 `src/editor` 是从 GenOffice 对应
renderer fork 的页面层，`workstudio-host.ts` 是唯一的宿主适配层。

## 开发

```bash
pnpm setup          # sparse checkout GenOffice + 安装依赖
pnpm typecheck      # docx / xlsx / pptx / pdf（含 renderer 与 engine）
pnpm build          # 四个插件的 engine + 页面 bundle
pnpm test           # 各包 tests/
pnpm verify         # 清单 / 产物 / 静态命令 / 旧面残留
```

单独构建、验证与测试：

```bash
pnpm build:docx && pnpm verify:docx && pnpm test:docx
pnpm build:xlsx && pnpm verify:xlsx && pnpm test:xlsx
pnpm build:pptx && pnpm verify:pptx && pnpm test:pptx
pnpm build:pdf  && pnpm verify:pdf  && pnpm test:pdf
```

打包 / 安装：

```bash
# 生成四个可安装 zip（release/<id>-<version>.zip + .sha256）
pnpm pack:plugins
pnpm pack:docx                     # 只打一个（--only docx）

# 开发用：直接安装到当前用户的插件目录
pnpm install:plugins

# 随 WorkStudio 源码构建的内置插件目录（仅开发/内置分发）
pnpm install:builtin
```

WorkStudio 新插件机制：

- **zip**：一个 zip 只装一个插件；WorkStudio 负责下载、校验 sha256 并安装。
- **本地目录**：在 WorkStudio 设置 → 插件中填绝对路径，原地加载，不复制；
  适合开发。
- zip / 本地目录安装后都需要**重启 WorkStudio**才会装载。
- zip 自包含，不执行 `npm install`；构建产物只包含 `package.json` 与 `dist/`。
- `pnpm install:plugins` 直接复制到 `installed/` 是开发快捷方式；需要来源、卸载和安装状态时，用 WorkStudio 设置 → 插件添加本地目录或 zip。

## 发布

四个插件（`docx` / `xlsx` / `pptx` / `pdf`）共用**同一个版本**和**同一个 tag**：
所有 `packages/*/package.json` 的 `version` 必须与根 `package.json` 一致，也要与 tag 一致
（`v0.2.0` 或 `0.2.0`，短 tag `0.2` 也接受）。

```bash
git tag v0.2.0
git push origin v0.2.0
```

`.github/workflows/release.yml` 会：

1. 校验 tag 与所有插件版本一致；
2. 拉取 `vendor/genoffice`，`pnpm install --frozen-lockfile`，装 `wasm32-wasip1`；
3. typecheck / build / test / verify；
4. 生成 `release/<id>-<version>.zip` 与 `.sha256`，逐个校验 zip 内容（入口、页面、技能、最小清单）；
5. 把 zip 与 sha256 挂到 GitHub Release；
6. 用 PluginRegistry App token 写入 `plugins/{docx,xlsx,pptx,pdf}.json`，
   `npm run validate` 后开 PR（`release/office-plugins-<version>`）。

需要的仓库配置：`REGISTRY_APP_PRIVATE_KEY`（secret）与 `REGISTRY_APP_CLIENT_ID`
（secret 或 variable），GitHub App 需对 `Work-Agent/PluginRegistry` 有 contents / pull_requests 写权限。

本地生成注册表条目（预演）：

```bash
REGISTRY_DIR=../PluginRegistry \
REPO_URL=https://github.com/Work-Agent/WorkOffice \
RELEASE_BASE_URL=https://github.com/Work-Agent/WorkOffice/releases/download/v0.2.0 \
node scripts/write-registry-entry.mjs            # 可加 --only docx,xlsx
```

安装入口归 WorkStudio 管理；本仓库只负责产出自包含的插件 zip、校验值与注册表条目。

仓库只有一条 workflow（`.github/workflows/release.yml`）：推 tag 才跑，普通分支/PR 不跑 CI。
本地提交前请自行跑 `pnpm typecheck && pnpm test && pnpm verify && pnpm pack:plugins`。

## 运行时约束

- 页面字节只由页面命令写回（`<id>.save`）；宿主与 agent 都不直接写字节；
- 模型修改必须走 GenOffice 的操作层（PPTX `runTxn`、XLSX gateway/Univer mutation、DOCX transaction）；
- XLSX 保存使用完整 `planCellEditsToXlsx` 参数，不能调用会丢字段的旧包装；
- PPTX 保存后重新 `openPptx(bytes)`，刷新元素 id 与字节锚点；
- 页面不依赖 AI 包；XLSX / PPTX 的 AI 代码已删除。
- 插件清单用 `exports['./workstudio']` + `workstudio.id` / `pages.content` / `claims`；
  一个 zip 只包含一个 `package.json`，且不携带依赖。

## 已知边界

- DOCX 仍需定期同步 GenOffice Docs 上游，并补齐部分 WorkStudio 系统命令；
- PPTX 观众窗口是独立的 Document PiP 窗口，可手动拖到第二显示器；自动全屏/交换显示器需要宿主提供 screen API；
- XLSX：`xlsx` 走 wasm 流式打开（`meta + range`），无 wasm 时回退整份字节（≤64 MB）；页面关闭不弹保存确认，依赖自动保存与恢复副本。

## 许可

插件代码 Apache-2.0；GenOffice 源码与引擎遵循其 Apache-2.0 许可，见
`NOTICE` 与 `vendor/genoffice/LICENSE`。
