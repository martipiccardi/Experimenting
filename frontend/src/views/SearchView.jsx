import { useState, useEffect, useCallback } from 'react'
import { useSearchParams } from 'react-router-dom'
import Sidebar from '../components/Sidebar'
import RelatedTerms from '../components/RelatedTerms'
import ResultsTable from '../components/ResultsTable'
import Pagination from '../components/Pagination'
import { fetchSearch, downloadResults } from '../api/client'

const DEFAULT_FILTERS = {
  wave: '',
  questionNumber: '',
  periodFrom: '',
  periodTo: '',
  textSearch: '',
  searchScope: 'both',
  semanticOn: true,
  perPage: 100,
}

const SESSION_KEY = 'enes_search_state'

function loadSession() {
  try { return JSON.parse(sessionStorage.getItem(SESSION_KEY) || '{}') } catch { return {} }
}

export default function SearchView() {
  const [, setSearchParams] = useSearchParams()
  const [filters, setFilters] = useState(() => { const s = loadSession(); return s.filters || DEFAULT_FILTERS })
  const [page, setPage] = useState(() => { const s = loadSession(); return s.page || 1 })
  const [activeTerm, setActiveTerm] = useState(() => { const s = loadSession(); return s.activeTerm || null })
  const [committedText, setCommittedText] = useState(() => { const s = loadSession(); return s.committedText ?? (s.filters?.textSearch || '') })

  const [result, setResult] = useState(null)
  const [loading, setLoading] = useState(false)
  const [modelReady, setModelReady] = useState(false)

  useEffect(() => {
    if (modelReady) return
    const check = async () => {
      try {
        const r = await fetch('/api/model-ready')
        const d = await r.json()
        if (d.ready) setModelReady(true)
      } catch {}
    }
    check()
    const id = setInterval(check, 3000)
    return () => clearInterval(id)
  }, [modelReady])

  useEffect(() => {
    try { sessionStorage.setItem(SESSION_KEY, JSON.stringify({ filters, page, activeTerm, committedText })) } catch {}
  }, [filters, page, activeTerm, committedText])

  const doSearch = useCallback(async (f, p, term, text) => {
    setLoading(true)
    try {
      const data = await fetchSearch({
        semantic: f.semanticOn,
        wave: f.wave,
        question_number: f.questionNumber,
        period_from: f.periodFrom,
        period_to: f.periodTo,
        text_contains: text,
        search_scope: f.searchScope,
        sem_filter: term || '',
        page: p,
        per_page: f.perPage,
      })
      setResult(data)
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => {
    const t = setTimeout(() => doSearch(filters, page, activeTerm, committedText), 300)
    return () => clearTimeout(t)
  }, [filters.wave, filters.questionNumber, filters.periodFrom, filters.periodTo,
      filters.searchScope, filters.semanticOn, filters.perPage,
      committedText, page, activeTerm, doSearch])

  const handleFiltersChange = (newFilters) => {
    setFilters(newFilters)
    setPage(1)
    setActiveTerm(null)
  }

  const handleSearch = useCallback(() => {
    setCommittedText(filters.textSearch)
    setPage(1)
    setActiveTerm(null)
  }, [filters.textSearch])

  const handleTermClick = (term) => {
    setActiveTerm(term)
    setPage(1)
  }

  const handleShowWave = (wave, rowHash) => {
    setSearchParams({ show_wave: `${wave}___${rowHash}` })
  }

  const handleShowQWaves = (question, mnemo) => {
    setSearchParams({ show_q_waves: question, show_q_mnemo: mnemo || '' })
  }

  const handleShowVolA = (wave, question) => {
    window.open(`?show_vol_a=${encodeURIComponent(`${wave}___${question}`)}`, '_blank')
  }

  const handleDownload = async (fmt) => {
    await downloadResults({
      semantic: filters.semanticOn,
      wave: filters.wave,
      question_number: filters.questionNumber,
      period_from: filters.periodFrom,
      period_to: filters.periodTo,
      text_contains: committedText,
      search_scope: filters.searchScope,
      sem_filter: activeTerm || '',
    }, fmt)
  }

  const hasText = committedText.trim()
  const relatedTerms = result?.related_terms || []
  const allRelated = relatedTerms.map(t => t.term.toLowerCase())
  const expandedQueryTerms = result?.expanded_query_terms || []

  let expandedTerms = []
  if (activeTerm) {
    expandedTerms = [activeTerm.toLowerCase()]
  } else if (filters.semanticOn && hasText) {
    expandedTerms = [...new Set([...allRelated, ...expandedQueryTerms])]
  }
  const exactTerms = hasText ? [hasText.toLowerCase().trim()] : []

  const inQ = filters.searchScope === 'both' || filters.searchScope === 'q'
  const inA = filters.searchScope === 'both' || filters.searchScope === 'a'

  const qExact = inQ ? exactTerms : []
  const aExact = inA ? exactTerms : []
  const qExpanded = inQ ? expandedTerms : []
  const aExpanded = inA ? expandedTerms : []

  const total = result?.total ?? 0
  const rows = result?.rows ?? []
  const textPending = filters.textSearch !== committedText

  return (
    <div className="app-layout">
      <Sidebar filters={filters} onChange={handleFiltersChange} onSearch={handleSearch} textPending={textPending} />
      <main className="main-content">
        <h1>QUESTION BANK - SEARCH TOOL</h1>

        {filters.semanticOn && !modelReady && (
          <p style={{
            background: '#fff8e1', border: '1px solid #f9a825', borderRadius: 6,
            padding: '6px 12px', fontSize: 13, color: '#6d4c00', margin: '0 0 8px'
          }}>
            Semantic model loading… first search may take 1–2 min. Subsequent searches are instant.
          </p>
        )}

        {result?.semantic_count > 0 && (
          <p className="caption">Semantic search: {result.semantic_count} related results found</p>
        )}

        {filters.semanticOn && hasText && (
          <RelatedTerms
            terms={relatedTerms}
            activeTerm={activeTerm}
            onTermClick={handleTermClick}
          />
        )}

        {result?.waves_in_period?.length > 0 && (
          <>
            <p className="caption">Waves in this period ({result.waves_in_period.length}):</p>
            <div className="wave-links-container">
              {result.waves_in_period.map(w => (
                <button key={w} className="wave-link" onClick={() => handleShowWave(w, '')}>
                  {w}
                </button>
              ))}
            </div>
          </>
        )}

        <Pagination page={page} total={total} perPage={filters.perPage} onPageChange={setPage} />
        <div className="results-header">Results: {total.toLocaleString()}</div>

        {loading ? (
          <div className="loading">Loading…</div>
        ) : (
          <ResultsTable
            rows={rows}
            qExact={qExact}
            qExpanded={qExpanded}
            aExact={aExact}
            aExpanded={aExpanded}
            onShowWave={handleShowWave}
            onShowQWaves={handleShowQWaves}
            onShowVolA={handleShowVolA}
          />
        )}

        <div className="download-row">
          <button className="download-btn" onClick={() => handleDownload('csv')}>
            Download CSV ({total.toLocaleString()} results)
          </button>
          <button className="download-btn" onClick={() => handleDownload('xlsx')}>
            Download Excel ({total.toLocaleString()} results)
          </button>
        </div>
      </main>
    </div>
  )
}
