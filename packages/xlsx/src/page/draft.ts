/**
 * 崩溃恢复：打开时若插件数据目录里有草稿，先问人。
 *
 * 恢复 = 用草稿字节当启动数据（走「整份字节」那条路），保存仍写原文件；
 * 丢弃 = 删草稿后照常打开磁盘版本。
 */

import type { Channel } from './channel'

interface DraftStatus {
  exists: boolean
  at?: number
  size?: number
}

/** 返回草稿字节（恢复）或 null（丢弃/无草稿）。 */
export async function askDraft(channel: Channel): Promise<Uint8Array | null> {
  const status = await channel
    .read<DraftStatus>('draftStatus')
    .catch(() => ({ exists: false }) as DraftStatus)
  if (!status.exists) return null

  const when = status.at ? new Date(status.at) : null
  return new Promise<Uint8Array | null>((resolve) => {
    const banner = document.createElement('div')
    banner.className = 'draft-banner'
    const text = document.createElement('span')
    text.textContent =
      `发现未保存的恢复副本${when ? `（${when.toLocaleString()}）` : ''}` +
      `${status.size ? ` · ${Math.round(status.size / 1024)} KB` : ''}`
    const restore = document.createElement('button')
    restore.type = 'button'
    restore.textContent = '恢复'
    const discard = document.createElement('button')
    discard.type = 'button'
    discard.textContent = '丢弃'
    discard.className = 'ghost'
    banner.append(text, restore, discard)
    document.body.append(banner)

    restore.addEventListener('click', () => {
      banner.remove()
      void channel
        .read<{ bytes?: Uint8Array }>('draftLoad')
        .then((loaded) => resolve(loaded.bytes instanceof Uint8Array ? loaded.bytes : null))
        .catch(() => resolve(null))
    })
    discard.addEventListener('click', () => {
      banner.remove()
      void channel.call('draftClear', {}).catch(() => undefined)
      resolve(null)
    })
  })
}
