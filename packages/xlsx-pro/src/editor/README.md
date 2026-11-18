# GenOffice Sheets renderer fork

来源：`vendor/genoffice` 的 `apps/sheets/src/renderer`（经 `packages/xlsx` fork）。

这是 **GenOffice Sheets 页面层的原样 fork**，用于替代先前手写的 Univer 精简页。
当前状态：

- 已接入 WorkStudio host adapter（`workstudio-host.ts`）；
- 保存通过页面内 `xlsx-gateway` 产字节，再走计划型 `ura.page.commit({ kind:'bytes' })`；
- AI 面板 / AI Ribbon 组已移除，AI 代码清理待后续；
- `src/page` 是 WorkStudio 专用适配层；`App.tsx` 暴露 `__xlsxPro` 句柄给 agent 桥。

不要在这个目录里手写 WorkStudio 专用逻辑；差异应放在 `src/page/*` 与 `workstudio-host.ts` 的宿主适配层。
