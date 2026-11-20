export function EmptyState({
  title,
  hint,
  detail,
  retryLabel,
  onRetry,
}: {
  title: string
  hint: string
  detail?: string
  retryLabel?: string
  onRetry?: () => void
}) {
  return (
    <div className="pdf-empty" role="alert">
      <div className="pdf-empty-mark" aria-hidden="true">
        PDF
      </div>
      <h2 className="pdf-empty-title">{title}</h2>
      <p className="pdf-empty-hint">{hint}</p>
      {detail && <pre className="pdf-empty-detail">{detail}</pre>}
      {onRetry && retryLabel && (
        <button type="button" className="pdf-btn pdf-btn-primary" onClick={onRetry}>
          {retryLabel}
        </button>
      )}
    </div>
  )
}
