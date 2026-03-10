import { useEffect, useState } from 'react'
import { useSearchParams } from 'react-router-dom'
import ResultsTable from '../components/ResultsTable'
import { fetchWave, downloadResults } from '../api/client'

export default function WaveView() {
  const [searchParams, setSearchParams] = useSearchParams()
  const showWaveRaw = searchParams.get('show_wave') || ''
  const hlQParam = searchParams.get('hl_q') || ''
  const hlMnemoParam = searchParams.get('hl_q_mnemo') || ''

  const [wave, setWave] = useState('')
  const [hlId, setHlId] = useState('')
  const [rows, setRows] = useState([])
  const [total, setTotal] = useState(0)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    let wavePart = showWaveRaw
    let rowHash = ''
    if (showWaveRaw.includes('___')) {
      const idx = showWaveRaw.lastIndexOf('___')
      wavePart = showWaveRaw.slice(0, idx)
      rowHash = showWaveRaw.slice(idx + 3)
    }
    setWave(wavePart)
    setHlId(rowHash)

    if (!wavePart) return
    setLoading(true)
    fetchWave(wavePart).then(data => {
      setRows(data.rows)
      setTotal(data.total)
    }).finally(() => setLoading(false))
  }, [showWaveRaw])

  const handleShowVolA = (wave, question) => {
    window.open(`?show_vol_a=${encodeURIComponent(`${wave}___${question}`)}`, '_blank')
  }

  const handleBack = () => {
    if (hlQParam) {
      setSearchParams({ show_q_waves: hlQParam, show_q_mnemo: hlMnemoParam })
    } else {
      setSearchParams({})
    }
  }

  return (
    <div style={{ padding: '1rem 1.5rem' }}>
      <button className="back-link" onClick={handleBack}>
        {hlQParam ? 'Back to waves list' : 'Back to search'}
      </button>

      <h1>Complete wave: {wave}</h1>
      <p className="results-header">{total.toLocaleString()} questions</p>

      {loading ? (
        <div className="loading">Loading…</div>
      ) : (
        <ResultsTable
          rows={rows}
          qExact={[]}
          qExpanded={hlQParam ? [hlQParam.toLowerCase()] : []}
          aExact={[]}
          aExpanded={[]}
          highlightId={hlId}
          highlightQuestion={hlQParam}
          highlightMnemo={hlMnemoParam}
          onShowWave={(w, hash) => setSearchParams({ show_wave: `${w}___${hash}` })}
          onShowQWaves={(q, mnemo) => setSearchParams({ show_q_waves: q, show_q_mnemo: mnemo || '' })}
          onShowVolA={handleShowVolA}
        />
      )}

      <div className="download-row">
        <button className="download-btn" onClick={() => downloadResults({ wave }, 'csv')}>
          Download CSV ({total.toLocaleString()} rows)
        </button>
        <button className="download-btn" onClick={() => downloadResults({ wave }, 'xlsx')}>
          Download Excel ({total.toLocaleString()} rows)
        </button>
      </div>
    </div>
  )
}
