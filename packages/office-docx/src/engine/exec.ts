// ─── engine.exec：SavePlan → DOCX 字节（唯一字节产生者） ────────────────
//
// 契约 §6.2：exec 只写 `ctx.tempDir`，目标路径由权威层 rename；返回体不带
// `version`（由宿主算出）。docx 没有常驻写侧模型，因此没有 modelToken。
//
// 字节保真：只重写 `word/document.xml`；其余 zip 条目由 `zip.ts` 逐字节拷贝。

import { strToU8 } from 'fflate';
import {
  DOCUMENT_PART,
  DocxParseError,
  parseDocumentXml,
  parseDocxBytes,
} from '../shared/document';
import { PlanShapeError, assertPlanShape, baselineSlices } from '../shared/plan';
import { buildManifest, spliceBody } from '../shared/sync';
import type { SavePartPatch } from '../shared/types';
import { runtimeReport } from '../shared/runtime';
import type { ExecOutcome, ExecRequest, EngineCtx } from './abi';
import { coded, messageOf } from './abi';
import { ZipShapeError, rewriteZipParts, type ZipPartPatch } from './zip';
import { sha256Hex } from '../shared/sha256';

export async function exec(req: ExecRequest, ctx: EngineCtx): Promise<ExecOutcome> {
  if (req.command === 'undo' || req.command === 'redo')
    throw coded(
      'command.unsupported',
      'docx 的撤销是页面模型历史（I16）；无页面路径不提供 undo',
    );

  const baselineBytes = req.bytes ?? (await ctx.readBytes()).bytes;
  let parsed;
  try {
    parsed = parseDocxBytes(baselineBytes);
  } catch (error) {
    throw coded('command.invalid', messageOf(error));
  }
  const slices = baselineSlices(parsed);

  let plan;
  try {
    plan = assertPlanShape(req.plan, slices);
  } catch (error) {
    if (error instanceof PlanShapeError) throw coded('command.invalid', error.message);
    throw error;
  }

  if (plan.kind === 'reject')
    return {
      result: {
        applied: 0,
        rejected: plan.rejected ?? [],
        ...(plan.warnings ? { warnings: plan.warnings } : {}),
      },
    };

  if (plan.kind === 'check')
    return {
      result: { applied: 0, ...(plan.audit !== undefined ? { audit: plan.audit } : {}) },
    };

  // 拼接 → 重新解析校验（良构 / sectPr 保序），再重写 zip 条目。
  let documentXml: string;
  let nextParsed;
  try {
    documentXml = spliceBody(parsed, plan);
    nextParsed = parseDocumentXml(documentXml);
  } catch (error) {
    if (error instanceof DocxParseError)
      throw coded('command.invalid', `计划产出的 document.xml 非法：${messageOf(error)}`);
    throw error;
  }
  const manifest = buildManifest(parsed, plan);
  const warnings = [...(plan.warnings ?? [])];
  if (nextParsed.blocks.every((block) => block.hidden))
    warnings.push('document-empty');

  if (req.dryRun || plan.dryRun)
    return {
      result: {
        applied: plan.applied ?? plan.changedCount,
        ...(warnings.length > 0 ? { warnings } : {}),
        audit: {
          changedCount: plan.changedCount,
          deletedCount: plan.deletedCount,
          changedParts: plan.parts?.length ?? 0,
        },
        sync: { kind: 'manifest', blocks: manifest.blocks },
      },
    };

  let bytes: Uint8Array;
  try {
    // 正文块 + 整包补丁一次写完：未涉及的 zip 条目仍逐字节原样拷贝。
    bytes = rewriteZipParts(baselineBytes, [
      { path: DOCUMENT_PART, data: strToU8(documentXml) },
      ...decodePartPatches(plan.parts),
    ]);
  } catch (error) {
    if (error instanceof ZipShapeError) throw coded('command.unsupported', messageOf(error));
    throw error;
  }
  if (bytes.byteLength > ctx.budget.maxBytes)
    throw coded(
      'command.too-large',
      `产出超过执行点上限：${bytes.byteLength} > ${ctx.budget.maxBytes}`,
    );
  if (bytes.byteLength > runtimeReport().limits.maxCommitBytes)
    throw coded('command.too-large', '产出超过插件提交上限');

  return {
    bytes,
    result: {
      applied: plan.applied ?? plan.changedCount,
      ...(warnings.length > 0 ? { warnings } : {}),
      sync: {
        kind: 'manifest',
        blocks: manifest.blocks,
        ...(manifest.hidden ? { hidden: manifest.hidden } : {}),
      },
      audit: {
        changedCount: plan.changedCount,
        deletedCount: plan.deletedCount,
        changedParts: plan.parts?.length ?? 0,
        resultDocHash: sha256Hex(bytes),
      },
    },
  };
}

function decodeBase64(text: string): Uint8Array {
  const binary = atob(text);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

/** 计划里的 part 补丁 → zip 层补丁（xml 用 UTF-8，二进制用 base64）。 */
function decodePartPatches(parts: SavePartPatch[] | undefined): ZipPartPatch[] {
  if (!parts || parts.length === 0) return [];
  return parts.map((patch) => {
    if ('delete' in patch) return { path: patch.path, delete: true };
    if ('xml' in patch) return { path: patch.path, data: strToU8(patch.xml) };
    return { path: patch.path, data: decodeBase64(patch.base64) };
  });
}

// docx 无常驻写侧模型：两阶段提交为空实现（宿主只在 modelToken 存在时调用）。
export async function commitModel(): Promise<void> {}
export async function abortModel(): Promise<void> {}
