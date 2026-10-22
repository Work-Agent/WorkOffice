# WorkStudio Office 插件

以 WorkStudio 的 **transport / codec / projection / 页面会话** 为骨架。
codec 与 projection 由本仓库实现，只依赖通用 ZIP/XML 库；GenOffice
只用于需要富文本/分页的页面层。

当前进度：

| 插件 | 能力 | 状态 |
| --- | --- | --- |
| `@workstudio/plugin-docx` | DOCX 打开、渲染、编辑、字节保存 | 可用 |
| `@workstudio/plugin-xlsx` | 工作表/单元格浏览 · 基础值/公式编辑 | 可用 |
| `@workstudio/plugin-pptx` | 幻灯片浏览 · 文本节点编辑 | 可用 |

## 为什么不是“直接嵌 GenOffice 应用”

GenOffice 是完整 Electron 应用：文档身份是文件路径，打开/保存/菜单/原生对话框
都由它自己的主进程和 preload 承担。WorkStudio 的交互模型不同：

- 文档身份是 URA 统一标识，页面只拿页面会话 id 和投影数据；
- 页面没有 Node、没有裸 IPC，只能走白名单命令；
- 保存必须经过严格版本写（`ura.page.saveBytes`），冲突由页面会话返回；
- 标签、脏态、主题、侧栏、文件选择归 WorkStudio 壳层。

因此这里 **不引入 GenOffice 的 main / preload**，也不把它的 App 当黑盒：

1. **codec/projection 自有**：DOCX/XLSX 用 `fflate` / `fast-xml-parser` 读 OOXML，
   不依赖 `@genoffice/*`；
2. Docs 页面层 **fork 到 `packages/docx/src/editor`**，由我们持有并改造；
3. `packages/docx/src/page` 是 WorkStudio 页面入口与宿主适配层，只把
   `window.desktop` 语义映射到页面会话。

页面层已经按 Studio 的逻辑裁过：

- 移除 Home 的 Genspark AI 组（AI 总结 / 润色 / 排版 / 思维导图）；
- 移除选区浮出的“AI 修改”气泡；
- 移除右侧 AI 面板挂载点；
- 移除审阅里的 AI 改写 / AI 翻译 / AI 评论 / AI 修订入口；
- 移除右键菜单的 AI 同义 / AI 翻译，以及 F7 AI 校对入口；
- 移除 AI 面板状态、选区 AI 队列与对应宿主依赖。

后续页面改造（文件菜单、打开/保存、最近文件、导出、项目历史）都在
`src/editor` 内完成，不再依赖 GenOffice 的宿主代码。

## DOCX：复用 GenOffice 编辑器壳

```
WorkStudio file transport
        │  bytes + version + formatHints
        ▼
DOCX codec（自有 OOXML 解析）
        │  doc: raw bytes / text / outline
        ├── agent consumer ──► 纯文本投影（AI 读取）
        └── editor consumer ─► page data: { bytes, name, ref, version }
                                   │
                                   ▼
                         DOCX 页面（GenOffice Docs 渲染器）
                                   │  编辑后重建 DOCX 字节
                                   ▼
                     window.extension.ura.page.saveBytes(session, bytes)
                                   │
                                   ▼
                     PageSessions → Ura.write(strict version)
```

- **transport**：沿用 WorkStudio 的 `file://` 来源；插件不碰文件系统。
- **codec**：自有 OOXML 解析生成文本/大纲；`encode` 只接受二进制整篇写回。
- **projection**：agent 得到文本；editor 得到页面可消费的字节与元数据。
- **页面保存**：整篇 DOCX 字节走 `saveBytes`，版本冲突返回 base64 基线，
  页面提示重新打开。GenOffice 的“写路径”交互在这里不存在。

## XLSX：只复用解析引擎，页面自己画

XLSX 不走 GenOffice 的 Univer 外壳，也不依赖 GenOffice 的
`xlsx-gateway`：

```
transport bytes
  → xlsx codec（fflate + fast-xml-parser 读 workbook/sharedStrings/worksheet）
  → page data: { sheets: [{ id, name, cells: { A1: { value, formula } } }] }
  → WorkStudio 网格页（标签 / 选区 / 公式栏 / 滚动）
  → agent projection: 按 sheet/row 拼文本
```

页面支持选中单元格后在公式栏改值/公式；保存时由 `src/page/patch.ts`
直接改 worksheet 的 `<c>` 节点（inlineStr / 数字 / 布尔 / `<f>`），
其余 ZIP 部件原样保留，再走 `ura.page.saveBytes`。样式、合并、列宽、
行高、缓存公式值与数字格式已进入投影；图表、透视、条件格式等仍待补。

## PPTX：自有 codec + GenOffice 渲染树

PPTX 也分成两层：

```
transport bytes
  → pptx codec（fflate 解 ppt/slides/slideN.xml，抽文本/页数）
  → page data: { bytes, name, slideCount }
  → 页面层：@genoffice/pptx-engine openPptx
             @genoffice/pptx-render buildRenderSlide
             自有 SVG 渲染器（形状/图片/表格/图表/文本）
  → 幻灯片导航（上一页 / 下一页 / 缩略序号）
```

codec 不依赖 GenOffice；页面层使用 GenOffice 的解析与绘制列表，渲染器由
WorkStudio 自己实现。双击文本可编辑，保存走 `savePptx` + `ura.page.saveBytes`；
形状/图片/表格/图表编辑、动画、放映、备注等仍待补。

## GenOffice 页面层来源

官方 GenOffice 没有发布到 npm：`packages/*` 的 `package.json` 都是
`private: true`，`npm view @genoffice/docx-engine` 为 404。  
因此 codec 不依赖它；只有 DOCX 页面层 fork 需要官方源码。

仓库用 git submodule 固定官方源码，但只稀疏检出页面层用到的 11 个包：

```
vendor/genoffice   →  https://github.com/genspark-ai/genoffice
  只取 packages/{agent-core,ai-provider,docx-engine,electron-utils,font-metrics,
                i18n,pptx-engine,pptx-render,project-store,ui,zip-gate}
```

- `packages/docx/src/editor` 是从 `apps/docs` fork 下来的页面层，不再需要 `apps/*`；
- 构建时 Vite 插件把页面层里的 `@genoffice/*` 映射到 `vendor/genoffice/packages/*`；
- 发布产物 `dist/ura.js` 与 `dist/page` 都是自包含的，不把 `vendor/` 或任何
  `@genoffice/*` 包带进插件依赖；
- 更新 GenOffice：改 submodule 指向的 commit，再跑一次 `node scripts/setup-genoffice.mjs`。

如果以后要彻底去掉 `vendor/`，只有两条路：改用第三方已发布的 GenOffice
封装包，或我们发布自己的 `@workstudio/genoffice-*` 构建产物。

## 构建

```bash
# 首次：稀疏拉取官方 GenOffice 子模块（只取页面层用到的包）
node scripts/setup-genoffice.mjs

pnpm install
pnpm build          # codec + 页面
# 或只构建自有 codec（不需要 GenOffice 子模块）
pnpm build:codec
```

产物：

```
packages/docx/dist/
  ura.js              # 主进程注册入口（含 @genoffice/docx-engine）
  page/index.html     # 插件页面
  page/assets/…       # GenOffice 渲染器、字体、worker
```

## 接入 WorkStudio

WorkStudio 侧需要四个接缝（见 `docs/workstudio-integration.md`）：

1. `file` transport 为 OOXML 扩展名给出 `ooxml-*` 格式提示；
2. 页面会话支持 `saveBytes` / 二进制冲突稿；
3. `app://page` 只编码路径段、保留斜杠，插件页的相对资源才能解析；
4. 宿主装载外部插件前注入 `globalThis.__workstudioSdk` 注册面。

```bash
pnpm install:plugins   # 复制 docx / xlsx 到 ../WorkStudio/src/plugins
# 或安装到 userData：
node scripts/install.mjs --target "~/Library/Application Support/WorkStudio/plugins"
```

## 许可

插件代码 Apache-2.0；GenOffice 源码与引擎遵循其 Apache-2.0 许可，见
`NOTICE` 与 `vendor/genoffice/LICENSE`。

## 页面层改造清单（DOCX）

| 区域 | 现状 | 目标 |
| --- | --- | --- |
| AI 入口 | 已从 `src/editor` 删除 Home AI 组、选区气泡、AI 面板、审阅 AI 动作、右键同义/翻译、F7 校对 | 只保留 WorkStudio 助手 |
| 打开 | 页面会话投递字节；`openDocx` 不再弹 GenOffice 文件框 | 打开入口归资源面板 |
| 插入图片 | 页面内文件选择器，读为 base64 后插入 | 保持 |
| 导出 HTML / Markdown | 已恢复“转换”菜单，走浏览器文件保存 | PDF / 图片导出待补 |
| 保存 | `saveDocx` → `ura.page.saveBytes`，严格版本写 | 保持；补冲突提示与关闭前 flush |
| 另存为 / 新建 | 适配层暂按“保存当前文档”处理 | 改走 WorkStudio 的页面/工作区命令 |
| 最近文件 | 返回空 | 删除入口，或改用资源索引 |
| 导出 / 打印 / PDF | 多数为 no-op | 明确能力位；未实现则不暴露入口 |
| 项目历史 / 附件 / Zotero | 适配层 no-op | 从页面层删除 |
| 原生菜单 | 不承载 | 保留 Ribbon，菜单命令由壳层决定 |
| 脏态 | 输入事件上报 + 保存后清零 | 关闭标签前的保存策略由壳层接管 |

这些改造都在 `packages/docx/src/editor` 内完成，不再通过运行时
`window.desktop` 兜底；`src/page/host.ts` 只保留页面会话必须的几个入口。
