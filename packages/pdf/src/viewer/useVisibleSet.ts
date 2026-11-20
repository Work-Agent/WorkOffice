import { useEffect, useRef, useState } from 'react'
import type { RefObject } from 'react'

/**
 * 维护可见项集合：页面/缩略图懒渲染的共享基础。
 * `enabled` 翻转时重建 observer（侧栏切换会 unmount/remount 根节点）。
 */
export function useVisibleSet(
  rootRef: RefObject<HTMLElement | null>,
  count: number,
  rootMargin: string,
  enabled = true,
): { visible: Set<number>; setItemRef: (index: number) => (el: HTMLElement | null) => void } {
  const [visible, setVisible] = useState<Set<number>>(new Set())
  const itemRefs = useRef<(HTMLElement | null)[]>([])
  useEffect(() => {
    const root = rootRef.current
    if (!enabled || !root || count === 0) return
    const io = new IntersectionObserver(
      (entries) => {
        setVisible((prev) => {
          const next = new Set(prev)
          for (const entry of entries) {
            const index = Number((entry.target as HTMLElement).dataset.idx)
            if (!Number.isFinite(index)) continue
            if (entry.isIntersecting) next.add(index)
            else next.delete(index)
          }
          return next
        })
      },
      { root, rootMargin },
    )
    for (const el of itemRefs.current) if (el) io.observe(el)
    return () => io.disconnect()
  }, [rootRef, count, rootMargin, enabled])
  return {
    visible,
    setItemRef: (index) => (el) => {
      itemRefs.current[index] = el
    },
  }
}
