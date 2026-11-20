import { useEffect, useRef, useState } from 'react'

export function PasswordDialog({
  wrong,
  hint,
  onSubmit,
  onCancel,
  t,
}: {
  wrong: boolean
  hint: string
  onSubmit: (password: string) => void
  onCancel: () => void
  t: (key: string, vars?: Record<string, string | number>) => string
}) {
  const [value, setValue] = useState('')
  const inputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    inputRef.current?.focus()
  }, [wrong])

  return (
    <div className="pdf-modal-backdrop" role="presentation">
      <div
        className="pdf-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('passwordRequired')}
        onKeyDown={(event) => {
          if (event.key === 'Escape') onCancel()
        }}
      >
        <h2 className="pdf-modal-title">
          {wrong ? t('passwordWrong') : t('passwordRequired')}
        </h2>
        <p className="pdf-modal-hint">{hint}</p>
        <form
          onSubmit={(event) => {
            event.preventDefault()
            if (value) onSubmit(value)
          }}
        >
          <input
            ref={inputRef}
            type="password"
            className="pdf-modal-input"
            value={value}
            placeholder={t('passwordPlaceholder')}
            onChange={(event) => setValue(event.target.value)}
            aria-label={t('passwordPlaceholder')}
            autoComplete="off"
          />
          <div className="pdf-modal-actions">
            <button type="button" className="pdf-btn pdf-btn-text" onClick={onCancel}>
              {t('passwordCancel')}
            </button>
            <button type="submit" className="pdf-btn pdf-btn-primary" disabled={!value}>
              {t('passwordSubmit')}
            </button>
          </div>
        </form>
      </div>
    </div>
  )
}
