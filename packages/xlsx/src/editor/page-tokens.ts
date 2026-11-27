// ─── 页面令牌：docRev / structureEpoch（设计稿 §6.3） ──────────────────
//
// 两个计数器都属于**页面会话**，不进 (ref, version) 缓存、不持久化：
//
//   docRev          页面模型事务计数 → L1（读内容型）op 的新鲜度
//   structureEpoch  结构修订号       → L2（绝对坐标）op 的新鲜度
//
// 二者都**不随保存重置**：保存只是把内存模型与磁盘对齐，页面坐标没有变。
// 一旦重置，每次自动保存（400 ms 防抖）都会把 agent 的令牌打废。
//
// 与 `epoch` 的区别：`epoch` 是权威层回执里的提交代数（+1 只在提交时），
// 页面内未保存的结构改动不会推进它，因此不能拿它判 A1 坐标。

let docRev = 0
let structureEpoch = 0

export function bumpDocRev(): void {
  docRev += 1
}

export function bumpStructureEpoch(): void {
  structureEpoch += 1
}

export function pageTokens(): { docRev: number; structureEpoch: number } {
  return { docRev, structureEpoch }
}

export function resetPageTokens(): void {
  docRev = 0
  structureEpoch = 0
}
