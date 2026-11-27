/** 同一 ref 的写串行化；写只发生在命令里，页面不持锁。 */
const chains = new Map<string, Promise<unknown>>()

export async function withLock<T>(key: string, run: () => Promise<T>): Promise<T> {
  const previous = chains.get(key) ?? Promise.resolve()
  const next = previous.then(run, run)
  const settled = next.then(
    () => undefined,
    () => undefined,
  )
  chains.set(key, settled)
  try {
    return await next
  } finally {
    if (chains.get(key) === settled) chains.delete(key)
  }
}
