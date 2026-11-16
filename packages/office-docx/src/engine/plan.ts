// ─── engine.plan：命令批 → SavePlan（执行点 / 无页面路径） ─────────────
//
// 与页面路径共用 `applyCommands` + `planFromModel`（契约 §14.1 验收 2）。
// 只依赖基线字节与参数，不依赖进程内隐藏状态（契约 §6.2 第 3 条）。

import {
  DocxParseError,
  modelFromParsed,
  parseDocxBytes,
} from '../shared/document';
import { buildGuide } from '../shared/catalog';
import { applyCommands } from '../shared/ops';
import {
  baselineSlices as slicesOf,
  checkPlan,
  planFromModel,
  rejectedPlan,
} from '../shared/plan';
import { projectParsed } from '../shared/project';
import { runtimeReport } from '../shared/runtime';
import { CommandRejected, type DocxCommand } from '../shared/types';
import type { EngineCtx, PlanOutcome, PlanRequest } from './abi';
import { coded, messageOf } from './abi';

const MAX_COMMANDS = 200;

export async function plan(req: PlanRequest, ctx: EngineCtx): Promise<PlanOutcome> {
  const params = (req.params ?? {}) as Record<string, unknown>;
  const tool = String(req.tool ?? '');

  if (tool === 'docx.guide' || tool === 'docx_guide') return guideOutcome(req, ctx, params);

  // 无页面写：`docx_apply` 或宿主直接给 commands（契约 §8.2）。
  const rawCommands = Array.isArray(req.commands)
    ? req.commands
    : Array.isArray(params['commands'])
      ? params['commands']
      : null;
  if (!rawCommands)
    throw coded('command.invalid', '缺少 commands 数组');

  const bytes = req.bytes ?? (await ctx.readBytes()).bytes;
  let parsed;
  try {
    parsed = parseDocxBytes(bytes);
  } catch (error) {
    if (error instanceof DocxParseError) throw coded('command.invalid', messageOf(error));
    throw error;
  }

  if (rawCommands.length > MAX_COMMANDS)
    return {
      plan: rejectedPlan(
        new CommandRejected(0, 'command.too-large', `批上限 ${MAX_COMMANDS} 条（收到 ${rawCommands.length}）`),
        parsed,
      ),
      summary: `rejected: 批上限 ${MAX_COMMANDS} 条`,
    };

  const slices = slicesOf(parsed);
  try {
    const commands = rawCommands as DocxCommand[];
    const checks = commands.filter((command) => command.op === 'check');
    if (checks.length > 0) {
      if (commands.length !== 1)
        return {
          plan: rejectedPlan(
            new CommandRejected(0, 'command.invalid', 'check 必须单独成批（不与其他命令混排）'),
            parsed,
          ),
          summary: 'rejected: check 必须单独成批',
        };
      const checked = checkPlan(parsed);
      if (params['dryRun'] === true) checked.dryRun = true;
      return { plan: checked, summary: checkSummary(checked.audit) };
    }
    if (commands.length === 0) {
      const checked = checkPlan(parsed);
      return { plan: checked, summary: checkSummary(checked.audit) };
    }
    const model = modelFromParsed(parsed);
    const result = applyCommands(model, commands);
    const next = planFromModel(result.model, slices, result.applied);
    if (params['dryRun'] === true) next.dryRun = true;
    return {
      plan: next,
      summary: next.kind === 'check'
        ? checkSummary(next.audit)
        : `plan: ${result.applied} 条命令 · 改 ${next.changedCount} 块 · 删 ${next.deletedCount} 块 · ${shortVersion(req.baseVersion)}`,
    };
  } catch (error) {
    if (error instanceof CommandRejected) {
      const rejected = rejectedPlan(error, parsed);
      return {
        plan: rejected,
        summary: `rejected[${error.index}] ${error.code}: ${error.message}`,
      };
    }
    throw error;
  }

}

async function guideOutcome(
  req: PlanRequest,
  ctx: EngineCtx,
  params: Record<string, unknown>,
): Promise<PlanOutcome> {
  const { version } = await ctx.readBytes();
  const guide = buildGuide({
    version: version || req.baseVersion,
    group: typeof params['group'] === 'string' ? params['group'] : undefined,
    op: typeof params['op'] === 'string' ? params['op'] : undefined,
    epoch: typeof params['epoch'] === 'number' ? params['epoch'] : undefined,
  });
  const report = runtimeReport();
  const data = {
    ...(guide.data as Record<string, unknown>),
    runtime: { exec: report.exec, features: report.features, limits: report.limits },
  };
  return { plan: data, summary: guide.summary };
}

function checkSummary(audit: unknown): string {
  if (!audit || typeof audit !== 'object') return 'check: ok';
  const value = audit as Record<string, unknown>;
  return `check: blocks=${String(value['blocks'])} paragraphs=${String(value['paragraphs'])} tables=${String(value['tables'])} tier=${String(value['tier'])}`;
}

function shortVersion(version?: string): string {
  if (!version) return 'version=-';
  return `version=${version.slice(0, 12)}`;
}

/** 只读投影入口（宿主在 `engine.plan` 缺失时才用 read.project）。 */
export async function project(req: PlanRequest, ctx: EngineCtx): Promise<string> {
  const bytes = req.bytes ?? (await ctx.readBytes()).bytes;
  const parsed = parseDocxBytes(bytes);
  const params = (req.params ?? {}) as Record<string, unknown>;
  return projectParsed(parsed, {
    name: typeof params['name'] === 'string' ? params['name'] : undefined,
    version: req.baseVersion,
  });
}
