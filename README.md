# WorkStudio Office 插件

DOCX / XLSX / PPTX 三种 WorkStudio 资源页。页面层复用 GenOffice 的
renderer、engine 与操作层，宿主侧只接 WorkStudio 的 URA 页面会话与
`saveBytes`。

整体架构、GenOffice 源码分层、Office 功能解构与封装方式：

→ [`docs/architecture.md`](./docs/architecture.md)

## 插件

| 包 | 页面 | 说明 |
| --- | --- | --- |
| `@workstudio/plugin-docx` | GenOffice Docs 编辑器 fork | 文本、分页、Ribbon、审阅、打印/导出 |
| `@workstudio/plugin-xlsx` | GenOffice Sheets renderer | Ribbon、Univer 网格、公式栏、格式、查找、筛选、透视、图表 |
| `@workstudio/plugin-pptx` | GenOffice Slides renderer | Konva 画布、原位文本、变换、幻灯片/节、表格/图表、媒体、动画、母版、放映、PDF |

三个插件都注册：

```ts
registerCodec(codec)
registerPresenter(presenter)
registerSidecar(sidecar)
```

页面只拿页面会话 id 和 presenter 投递的字节；保存统一走
`ura.page.saveBytes`，页面不碰路径，也没有 Node 能力。

## 仓库结构

```text
packages/
├─ docx/
│  └─ src/{page,editor,shared,ura.ts}
├─ xlsx/
│  └─ src/{editor,shared,ura.ts,ura-types.ts}
├─ pptx/
│  └─ src/{editor,shared,ura.ts}
├─ vendor/genoffice/          # sparse checkout 的 GenOffice 源码
├─ scripts/
│  ├─ setup-genoffice.mjs
│  ├─ install.mjs
│  └─ verify-contract.mjs
└─ docs/architecture.md       # 唯一架构文档
```

XLSX / PPTX 的 `src/editor` 是从 GenOffice 对应 renderer fork 的页面层；
`workstudio-host.ts` 是唯一的宿主适配层；`workstudio-save.ts`（XLSX）
负责保存请求到 `xlsx-gateway` 的映射。

## 开发

```bash
pnpm setup          # sparse checkout GenOffice + 安装依赖
pnpm typecheck      # docx / xlsx / pptx，包含 renderer
pnpm build          # codec + 页面 bundle
pnpm verify         # 校验 codec / presenter / sidecar 契约
```

单独构建：

```bash
pnpm build:docx
pnpm build:codec    # 只构建 ura.js
```

打包 / 安装：

```bash
# 生成三个可安装 zip（release/<id>-<version>.zip + sha256）
pnpm pack:plugins

# 开发用：直接安装到当前用户的插件目录
#   macOS: ~/Library/Application Support/WorkStudio/plugins/installed/<id>
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

GitHub Actions 会在推送 `v*` tag 时构建并创建 Release：

```bash
git tag v0.1.0
git push origin v0.1.0
```

`.github/workflows/release.yml` 会：

1. 拉取 `vendor/genoffice`；
2. `pnpm install --frozen-lockfile`；
3. typecheck / build / verify；
4. 生成 `release/<id>-<version>.zip` 与 `.sha256`；
5. 把 zip 与 sha256 挂到 GitHub Release。

安装入口和可安装清单归 WorkStudio 管理；本仓库只负责产出自包含的插件 zip
与校验值。

普通分支/PR 走 `.github/workflows/ci.yml`，只构建并上传 artifact，不创建 Release。

## 运行时约束

- 页面字节只通过 `ura.page.saveBytes` 写回；
- 模型修改必须走 GenOffice 的操作层（PPTX `runTxn`、XLSX gateway/Univer mutation、DOCX transaction）；
- XLSX 保存使用完整 `planCellEditsToXlsx` 参数，不能调用会丢字段的旧包装；
- PPTX 保存后重新 `openPptx(bytes)`，刷新元素 id 与字节锚点；
- 页面不依赖 AI 包；XLSX / PPTX 的 AI 代码已删除。
- 插件清单声明 `workstudio.apiVersion: 1`；一个 zip 只包含一个 `package.json`。

## 已知边界

- DOCX 仍需定期同步 GenOffice Docs 上游，并补齐部分 WorkStudio 系统命令；
- PPTX 观众窗口是独立的 Document PiP 窗口，可手动拖到第二显示器；自动全屏/交换显示器需要宿主提供 screen API；
- XLSX 大文件/公式流式加载仍按页面内策略处理。

## 许可

插件代码 Apache-2.0；GenOffice 源码与引擎遵循其 Apache-2.0 许可，见
`NOTICE` 与 `vendor/genoffice/LICENSE`。
