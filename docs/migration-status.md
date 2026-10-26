# GenOffice renderer 迁移状态

> 目标：用 GenOffice renderer 替换 XLSX / PPTX 精简页，保留除 AI 外的功能。
> 原则：renderer 尽量原样复用；差异集中在 host adapter；AI 全部移除。

## 总览

| 项目 | 状态 |
| --- | --- |
| GenOffice 组件复用分析 | ✅ `docs/genoffice-reuse-map.md` |
| vendor：`xlsx-gateway` / `pptx-ops` | ✅ 已加入 sparse checkout |
| XLSX renderer fork | ✅ `packages/xlsx/src/editor`（285 文件） |
| PPTX renderer fork | ✅ `packages/pptx/src/editor`（222 文件） |
| XLSX 接入 WorkStudio | ✅ GenOffice renderer + xlsx-gateway 保存 |
| PPTX 接入 WorkStudio | ✅ GenOffice renderer + pptx-engine/pptx-ops 保存；AI 代码已删除 |
| DOCX 上游同步 | ⏳ 待开始 |

XLSX / PPTX 页面入口均已切到 `src/editor`，旧 `src/page` 已删除。
`tsconfig` 仍暂时排除 `src/editor` / `src/shared`，待类型清理后放开。

## Phase 1：XLSX

- [x] 复制 `apps/sheets/src/renderer` → `packages/xlsx/src/editor`
- [x] 复制 `apps/sheets/src/shared` → `packages/xlsx/src/shared`
- [x] 复制 `create-univer` / Univer 集成 / Ribbon / 对话框 / 保存链路源码
- [x] 隐藏 AI 面板与 AI Ribbon 组
- [x] 配置 `@genoffice/*` / Univer presets / React 依赖与 Vite alias
- [x] 实现 `window.desktopApi` WorkStudio host adapter
- [x] 从页面会话构造 `WorkbookFile` 与 range 读取
- [x] 接通 `saveWorkbookEdits` → `xlsx-gateway` / `saveBytes`
- [x] 页面入口改为渲染 `src/editor/App`
- [x] 运行时验证：Ribbon / 网格 / 双击编辑 / 保存
- [x] 样式写回验证（`xl/styles.xml` + cell `s=`）
- [ ] AI 代码彻底删除（`ai/`、App AI state）

## Phase 2：PPTX

- [x] 复制 `apps/slides/src/renderer` → `packages/pptx/src/editor`
- [x] 复制 `apps/slides/src/shared` → `packages/pptx/src/shared`
- [x] 隐藏 AI 面板与 AI Ribbon 组
- [x] 配置 Konva / `pptx-ops` / UI 依赖
- [x] 实现 `window.slidesApi` WorkStudio host adapter
- [x] 页面内接 `pptx-engine` / `pptx-ops`
- [x] 页面入口改为渲染 `src/editor/App`
- [x] 运行时验证：画布 / 缩略图 / 原位文本编辑 / 变换 / 撤销 / 保存
- [x] 幻灯片增删 / 复制 / 排序 / 隐藏
- [x] 按版式新建幻灯片
- [x] 节的新增 / 重命名 / 移动 / 删除（含连同幻灯片删除）
- [x] 文本 / 字体 / 段落 / 文本框插入
- [x] 表格 / 图表 / SmartArt / 文本框插入
- [x] 表格单元格 / 合并 / 行列尺寸 / 表格样式
- [x] 图表数据读取 / 改类型 / 改数据 / 改配色
- [x] 图片裁剪 / 图片填充形状 / 图片透明度
- [x] 组合 / 层级 / 填充 / 描边 / 背景
- [x] 图片 / 媒体 / 3D 插入（页面文件选择器）
- [x] 主题 / 版式 / 查找替换 / 转场 / 页脚 / 页面尺寸 / 链接
- [x] 动画 / 节 / 放映（F5 已验证）
- [x] PDF 导出 / 图片导出 / 打印（页面兜底）
- [x] 元素复制 / 粘贴 / 复制副本
- [x] 系统剪贴板文本 / 图片跨窗口
- [x] 演示者视图（单屏模式已验证）
- [ ] 第二显示器外部观众窗口
- [x] AI 代码彻底删除（`ai/`、AI i18n、App AI state、host AI stub）
- [ ] 放开 `tsconfig` 排除并清理类型

## Phase 3：DOCX

- [ ] 同步 GenOffice Docs 上游
- [ ] 保持 AI 移除
- [ ] 补齐 WorkStudio 打开/保存/打印/导出命令

## 当前可运行状态

- XLSX：`packages/xlsx/dist/page`，GenOffice Sheets renderer + Univer
- PPTX：`packages/pptx/dist/page`，GenOffice Slides renderer + Konva
- DOCX：GenOffice Docs fork

```bash
pnpm typecheck
pnpm build
pnpm verify
```

均通过；安装到 WorkStudio 后运行态验证：

```text
xlsx: ribbon=true  ai=false  canvases=3
pptx: ribbon=true  ai=false  canvases=7
```
