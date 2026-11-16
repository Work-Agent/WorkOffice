// ─── engine.runtime：探活 → 能力报告 ───────────────────────────────────
//
// docx 不需要 wasm / 字体探活：exec 是进程内 JS，唯一的外部依赖是 fflate
// （随 bundle 打进来）。探活只做一次解析器自检，失败即如实降级
// （契约 §4.3 / P7：不得静默降级）。

import { parseDocumentXml } from '../shared/document';
import { runtimeReport } from '../shared/runtime';
import type { EngineCtx } from './abi';

export async function runtime(_ctx?: EngineCtx) {
  const report = runtimeReport();
  const at = Date.now();
  let ok = true;
  try {
    parseDocumentXml(
      '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>probe</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
    );
  } catch {
    ok = false;
  }
  report.probe = { at, ok };
  if (!ok) report.degraded = ['engine-init-failed'];
  return report;
}
