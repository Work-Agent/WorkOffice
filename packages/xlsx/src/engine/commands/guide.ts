import { buildGuide } from '../../tools/catalog'

/** 命令目录：静态数据，不读文件、不占页面往返。 */
export async function guide(input: Record<string, unknown>): Promise<{ text: string; data: unknown }> {
  const result = buildGuide({
    ...(typeof input.group === 'string' ? { group: input.group } : {}),
    ...(typeof input.op === 'string' ? { op: input.op } : {}),
  })
  return { text: result.summary, data: result.data }
}
