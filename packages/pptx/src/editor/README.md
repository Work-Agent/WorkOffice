# GenOffice Slides renderer fork

来源：`vendor/genoffice` 的 `apps/slides/src/renderer`。

这是 **GenOffice Slides 页面层的原样 fork**，用于替代先前手写的 SVG 渲染 + 模态文本编辑页。
当前状态：

- 已接入 WorkStudio host adapter（`page/host.ts`）；
- 打开 / 渲染 / 文本编辑 / 变换 / 撤销 / 保存走页面内 `pptx-engine` / `pptx-ops`；
- AI 面板 / AI Ribbon 组已移除，AI 代码清理待后续；
- 插入元素 / 表格 / 图表 / 组合 / 动画 / 放映 / 导出等高级操作仍待接入；
- `tsconfig` 暂时排除本目录，等待 AI 类型清理。

不要在这个目录里手写 WorkStudio 专用逻辑；差异应放在 `src/page/host-adapter.ts`。
