# GenOffice Sheets renderer fork

来源：`vendor/genoffice` 的 `apps/sheets/src/renderer`。

这是 **GenOffice Sheets 页面层的原样 fork**，用于替代先前手写的 Univer 精简页。
当前状态：

- 已接入 WorkStudio host adapter（`workstudio-host.ts`）；
- 保存通过页面内 `xlsx-gateway` 写回，再走 `ura.page.saveBytes`；
- AI 面板 / AI Ribbon 组已移除，AI 代码清理待后续；
- `tsconfig` 暂时排除本目录，等待 AI 类型清理。

不要在这个目录里手写 WorkStudio 专用逻辑；差异应放在 `src/page/host-adapter.ts`。
