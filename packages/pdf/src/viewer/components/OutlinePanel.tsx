import type { OutlineNode } from '../outline'
import { flattenOutline } from '../outline'

export function OutlinePanel({
  nodes,
  loading,
  error,
  onSelect,
  t,
}: {
  nodes: OutlineNode[]
  loading: boolean
  error: string | null
  onSelect: (page: number) => void
  t: (key: string, vars?: Record<string, string | number>) => string
}) {
  if (loading) return <div className="pdf-sidebar-empty">{t('loading')}</div>
  if (error) return <div className="pdf-sidebar-empty">{error}</div>
  if (nodes.length === 0) return <div className="pdf-sidebar-empty">{t('noOutline')}</div>

  return (
    <ul className="pdf-outline">
      {flattenOutline(nodes).map((node) => (
        <li key={node.id}>
          <button
            type="button"
            className="pdf-outline-item"
            style={{ paddingLeft: 10 + Math.min(node.depth, 8) * 14 }}
            disabled={node.page === null}
            onClick={() => {
              if (node.page !== null) onSelect(node.page)
            }}
          >
            {node.title}
          </button>
        </li>
      ))}
    </ul>
  )
}
