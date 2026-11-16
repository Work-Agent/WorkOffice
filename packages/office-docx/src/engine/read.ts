// ─── engine.read：只读投影 / 结构化查询 ─────────────────────────────────

import { parseDocxBytes } from '../shared/document';
import { projectParsed, queryModel, type QueryInput } from '../shared/project';
import type { EngineCtx } from './abi';
import { coded, messageOf } from './abi';

interface ReadRequest {
  ref: string;
  baseVersion?: string;
  tool?: string;
  params?: unknown;
  query?: unknown;
}

async function parsedOf(req: ReadRequest, ctx: EngineCtx) {
  const { bytes } = await ctx.readBytes();
  try {
    return parseDocxBytes(bytes);
  } catch (error) {
    throw coded('command.invalid', messageOf(error));
  }
}

export async function project(req: ReadRequest, ctx: EngineCtx): Promise<string> {
  const parsed = await parsedOf(req, ctx);
  const params = (req.params ?? {}) as Record<string, unknown>;
  return projectParsed(parsed, {
    name: typeof params['name'] === 'string' ? params['name'] : undefined,
    version: req.baseVersion,
    tier: typeof params['tier'] === 'string' ? params['tier'] : undefined,
  });
}

export async function query(req: ReadRequest, ctx: EngineCtx): Promise<unknown> {
  const parsed = await parsedOf(req, ctx);
  const input = (req.query ?? req.params ?? {}) as QueryInput;
  return queryModel(parsed.blocks, input);
}
