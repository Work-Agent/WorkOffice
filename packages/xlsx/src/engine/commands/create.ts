import type { PluginContext } from '@work-studio/sdk'

import { create, type CreateResult } from '../create'

export async function createWorkbook(
  input: Record<string, unknown>,
  ctx: PluginContext,
): Promise<CreateResult> {
  return create(input, ctx)
}
