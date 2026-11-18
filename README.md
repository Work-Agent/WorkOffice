# WorkStudio Office 插件

DOCX / XLSX / PPTX 三种 WorkStudio 资源页。页面层复用 GenOffice 的
renderer、engine 与操作层，宿主侧只接 WorkStudio 的 URA 页面会话与写通道
（整份 `saveBytes` 或计划型 `commit`）。

整体架构、GenOffice 源码分层、Office 功能解构与封装方式：

→ [`docs/architecture.md`](./docs/architecture.md)

## 插件

| 包 | 页面 | 写形态 | 说明 |
| --- | --- | --- | --- |
| `@workstudio/plugin-docx` | GenOffice Docs 编辑器 fork | `bytes` | 文本、分页、Ribbon、审阅、打印/导出 |
| `@workstudio/plugin-xlsx` | GenOffice Sheets renderer | `bytes` | Ribbon、Univer 网格、公式栏、格式、查找、筛选、透视、图表 |
| `@workstudio/plugin-xlsx-pro` | 同上 renderer + gateway 计划引擎 | `plan` | **合体版**：完整 Sheets 编辑器 + 4 个工具 / 56 条命令；有页面交整份字节，无页面交 gateway 保存请求 |
| `@workstudio/plugin-pptx` | GenOffice Slides renderer | `bytes` | Konva 画布、原位文本、变换、幻灯片/节、表格/图表、媒体、动画、母版、放映、PDF |
| `@workstudio/plugin-docx-pro` | 同上编辑器 + 块命令引擎 | `plan` | **合体版**：完整编辑器 + 3 个工具 / 41 条命令；有页面走编辑器事务，无页面走块级 XML 切片 |
| `@workstudio/plugin-pptx-pro` | 同上 Slides 编辑器 + 常驻 `OpenedPptx` | `plan` | **合体版**：完整 Slides 编辑器 + 4 个工具 / 63 条命令；页面交 `TxnRequest`，无页面走 `runTxn + savePptxToFile` |
| `@workstudio/plugin-office-docx` | 块列表 | `plan` | 纯后台/批处理的最小实现（无分页渲染） |

> `docx` / `docx-pro` / `office-docx` 声明同一个 MIME，`xlsx` / `xlsx-pro` 声明同一个 MIME，`pptx` / `pptx-pro` 声明同一个 MIME，**同一工作区只能启用一个**；
> 后注册者会在注册表里 override 前者的 codec / presenter。

插件都注册：

```ts
registerCodec(codec)
registerPresenter(presenter)
registerSidecar(sidecar)
```

页面只拿页面会话 id 和 presenter 投递的元数据；字节走 `bytes` 分块通道，
写回只有两条合法路径：

- `write: 'bytes'` → `ura.page.saveBytes(id, bytes)`（整份提交）
- `write: 'plan'` → `ura.page.commit(id, { payload: { type: 'plan', plan } })`
  （宿主锁内 `exec` 产字节）

页面不碰路径，也没有 Node 能力。

## 仓库结构

```text
packages/
├─ docx/
│  └─ src/{page,editor,shared,ura.ts}
├─ docx-pro/                  # 合体：完整编辑器 + 计划型写
│  ├─ src/{page,engine,shared,tools,ura.ts}
│  ├─ skills/docx/SKILL.md
│  └─ tests/                  # 引擎 / 声明 / 往返
├─ xlsx/
│  └─ src/{editor,shared,ura.ts,ura-types.ts}
├─ xlsx-pro/                  # 合体：完整 Sheets 编辑器 + 计划型写
│  ├─ src/{page,engine,shared,tools,ura.ts}
│  ├─ skills/xlsx/SKILL.md
│  └─ tests/
├─ pptx/
│  └─ src/{editor,shared,ura.ts}
├─ pptx-pro/                 # 合体：完整 Slides 编辑器 + 常驻模型计划型写
│  ├─ src/{page,engine,shared,editor,tools,ura.ts}
│  ├─ skills/pptx/SKILL.md
│  └─ tests/
├─ vendor/genoffice/          # sparse checkout 的 GenOffice 源码
├─ scripts/
│  ├─ setup-genoffice.mjs
│  ├─ install.mjs
│  └─ verify-contract.mjs
└─ docs/architecture.md       # 唯一架构文档
```

XLSX / PPTX 的 `src/editor` 是从 GenOffice 对应 renderer fork 的页面层；
`workstudio-host.ts` 是唯一的宿主适配层；`workstudio-save.ts`（XLSX）
负责保存请求到 `xlsx-gateway` 的映射。`xlsx-pro` 复用同一份 renderer fork，
并把页面保存改成计划型 `commit`。

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
pnpm build:docx-pro # 完整编辑器 + 执行点 + 页面
pnpm build:xlsx-pro # Sheets 编辑器 + gateway 执行点 + 页面
pnpm build:pptx-pro # Slides 编辑器 + 常驻 OpenedPptx + 执行点
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

- 页面字节只通过 `ura.page.saveBytes`（整份）或 `ura.page.commit`（计划）写回；
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
