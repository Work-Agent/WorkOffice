import { useEffect, useRef } from 'react'
import type { SearchMatch } from '../search'
import { CloseIcon, SearchIcon } from './icons'

export function SearchBar({
  query,
  onQuery,
  matchCase,
  wholeWord,
  onToggleCase,
  onToggleWord,
  indexing,
  progress,
  results,
  index,
  total,
  contextOf,
  onNext,
  onPrev,
  onSelect,
  onClose,
  t,
}: {
  query: string
  onQuery: (value: string) => void
  matchCase: boolean
  wholeWord: boolean
  onToggleCase: () => void
  onToggleWord: () => void
  indexing: boolean
  progress: number
  results: SearchMatch[]
  index: number
  total: number
  contextOf: (match: SearchMatch) => string
  onNext: () => void
  onPrev: () => void
  onSelect: (match: SearchMatch, index: number) => void
  onClose: () => void
  t: (key: string, vars?: Record<string, string | number>) => string
}) {
  const inputRef = useRef<HTMLInputElement>(null)
  useEffect(() => {
    inputRef.current?.focus()
    inputRef.current?.select()
  }, [])

  return (
    <div className="pdf-search" role="search">
      <div className="pdf-search-row">
        <SearchIcon className="pdf-search-icon" />
        <input
          ref={inputRef}
          className="pdf-search-input"
          value={query}
          placeholder={t('searchPlaceholder')}
          onChange={(event) => onQuery(event.target.value)}
          aria-label={t('searchPlaceholder')}
        />
        <button
          type="button"
          className={`pdf-chip${matchCase ? ' is-active' : ''}`}
          onClick={onToggleCase}
          aria-pressed={matchCase}
          title={t('matchCase')}
        >
          Aa
        </button>
        <button
          type="button"
          className={`pdf-chip${wholeWord ? ' is-active' : ''}`}
          onClick={onToggleWord}
          aria-pressed={wholeWord}
          title={t('wholeWord')}
        >
          Ab
        </button>
        <span className="pdf-search-count">
          {indexing
            ? t('searching', { percent: progress })
            : query
              ? total > 0
                ? `${index + 1} / ${total}`
                : t('noResults')
              : ''}
        </span>
        <button
          type="button"
          className="pdf-btn"
          onClick={onPrev}
          disabled={total === 0}
          aria-label={t('prevMatch')}
          title={t('prevMatch')}
        >
          ↑
        </button>
        <button
          type="button"
          className="pdf-btn"
          onClick={onNext}
          disabled={total === 0}
          aria-label={t('nextMatch')}
          title={t('nextMatch')}
        >
          ↓
        </button>
        <button
          type="button"
          className="pdf-btn"
          onClick={onClose}
          aria-label={t('dismiss')}
          title={t('dismiss')}
        >
          <CloseIcon />
        </button>
      </div>
      {query && results.length > 0 && (
        <ol className="pdf-search-results">
          {results.slice(0, 200).map((match, matchIndex) => (
            <li key={`${match.pageIndex}-${matchIndex}`}>
              <button
                type="button"
                className={`pdf-search-result${matchIndex === index ? ' is-active' : ''}`}
                onClick={() => onSelect(match, matchIndex)}
              >
                <span className="pdf-search-page">p.{match.pageIndex + 1}</span>
                <span className="pdf-search-context">{contextOf(match)}</span>
              </button>
            </li>
          ))}
        </ol>
      )}
    </div>
  )
}
