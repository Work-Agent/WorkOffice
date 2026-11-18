// ─── 常驻模型缓存 / 快照 / 两阶段提交 ─────────────────────────────────
//
// 键：(ref, version)。宿主写成功后 commitModel 拿不到新 version，因此模型先
// 标记 lastArtifactSha256；下一次请求带新 version + baseline bytes 时按 hash
// 认领并重键。认领失败就丢弃重建，避免外部变更后复用旧模型。

import { createHash, randomUUID } from 'node:crypto'
import {
  commitSaved,
  openPptx,
  type OpenedPptx,
  type Slide,
  type SlideSize,
} from '@genoffice/pptx-engine'
import { MAX_CACHE_BYTES, MAX_CACHE_MODELS } from '../shared/runtime'

export interface Snapshot {
  slides: Slide[]
  entries: Map<string, Uint8Array>
  size: SlideSize
}

interface Pending {
  token: string
  entry: Entry
  snapshot: Snapshot
  undoDepth: number
  redoDepth: number
  mode: 'apply' | 'undo' | 'redo'
  artifactSha256?: string
}

export interface Entry {
  ref: string
  version: string
  opened: OpenedPptx
  bytes: number
  undo: Snapshot[]
  redo: Snapshot[]
  lastArtifactSha256?: string
}

const entries = new Map<string, Entry>()
const pending = new Map<string, Pending>()

function keyOf(ref: string, version: string): string {
  return `${ref}\u0000${version}`
}

export function sha256Hex(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

export function takeSnapshot(opened: OpenedPptx): Snapshot {
  return {
    slides: structuredClone(opened.deck.slides),
    entries: new Map(opened.archive.entries),
    size: { ...opened.deck.size },
  }
}

export function restoreSnapshot(opened: OpenedPptx, snapshot: Snapshot): void {
  opened.deck.slides = snapshot.slides
  opened.deck.size = snapshot.size
  const target = opened.archive.entries as Map<string, Uint8Array>
  target.clear()
  for (const [path, bytes] of snapshot.entries) target.set(path, bytes)
}

function estimateEntry(entry: Entry): number {
  let total = 0
  for (const bytes of entry.opened.archive.entries.values()) total += bytes.byteLength
  total += entry.opened.deck.slides.reduce((sum, slide) => sum + slide.originalXml.length, 0)
  return total
}

function trimCache(): void {
  while (entries.size > MAX_CACHE_MODELS) {
    const first = entries.keys().next().value as string | undefined
    if (!first) break
    entries.delete(first)
  }
  let total = 0
  for (const entry of entries.values()) total += estimateEntry(entry)
  while (total > MAX_CACHE_BYTES && entries.size > 1) {
    const first = entries.keys().next().value as string | undefined
    if (!first) break
    const entry = entries.get(first)
    if (entry) total -= estimateEntry(entry)
    entries.delete(first)
  }
}

/** 按 hash 认领刚写出的模型；失败返回 null。 */
function adoptByHash(ref: string, version: string, bytes: Uint8Array): Entry | null {
  const hash = sha256Hex(bytes)
  for (const [key, entry] of entries) {
    if (entry.ref !== ref || entry.lastArtifactSha256 !== hash) continue
    entries.delete(key)
    entry.version = version
    entry.lastArtifactSha256 = undefined
    entries.set(keyOf(ref, version), entry)
    return entry
  }
  return null
}

export async function getOrOpen(
  ref: string,
  version: string,
  bytes: Uint8Array,
): Promise<Entry> {
  const key = keyOf(ref, version)
  const hit = entries.get(key)
  if (hit) {
    // LRU：重新插入到队尾
    entries.delete(key)
    entries.set(key, hit)
    return hit
  }
  const adopted = adoptByHash(ref, version, bytes)
  if (adopted) return adopted

  const opened = await openPptx(bytes)
  const entry: Entry = {
    ref,
    version,
    opened,
    bytes: bytes.byteLength,
    undo: [],
    redo: [],
  }
  entries.set(key, entry)
  trimCache()
  return entry
}

export function beginApply(
  entry: Entry,
  mode: 'apply' | 'undo' | 'redo' = 'apply',
): { token: string; snapshot: Snapshot } {
  const token = `pptx-pro-${randomUUID()}`
  const snapshot = takeSnapshot(entry.opened)
  pending.set(token, {
    token,
    entry,
    snapshot,
    undoDepth: entry.undo.length,
    redoDepth: entry.redo.length,
    mode,
  })
  return { token, snapshot }
}

export function markArtifact(token: string, sha256: string): void {
  const item = pending.get(token)
  if (item) item.artifactSha256 = sha256
}

export function commit(token: string): void {
  const item = pending.get(token)
  if (!item) return
  const { entry, snapshot, mode } = item
  commitSaved(entry.opened)
  if (mode === 'apply') {
    entry.undo.push(snapshot)
    if (entry.undo.length > 20) entry.undo.shift()
    entry.redo.length = 0
  }
  entry.lastArtifactSha256 = item.artifactSha256
  pending.delete(token)
}

export function abort(token: string): void {
  const item = pending.get(token)
  if (!item) return
  const { entry, snapshot, undoDepth, redoDepth } = item
  restoreSnapshot(entry.opened, snapshot)
  entry.undo.length = undoDepth
  entry.redo.length = redoDepth
  pending.delete(token)
}

export function undo(entry: Entry): boolean {
  const snapshot = entry.undo.pop()
  if (!snapshot) return false
  entry.redo.push(takeSnapshot(entry.opened))
  restoreSnapshot(entry.opened, snapshot)
  return true
}

export function redo(entry: Entry): boolean {
  const snapshot = entry.redo.pop()
  if (!snapshot) return false
  entry.undo.push(takeSnapshot(entry.opened))
  restoreSnapshot(entry.opened, snapshot)
  return true
}

export function evictRef(ref: string): void {
  for (const [key, entry] of entries) {
    if (entry.ref === ref) entries.delete(key)
  }
}

export function clear(): void {
  entries.clear()
  pending.clear()
}

export function cacheSize(): { models: number; pending: number } {
  return { models: entries.size, pending: pending.size }
}
