import { useEffect, useState } from 'react'
import { useSearchParams } from 'react-router-dom'
import { fetchWavesForQuestion } from '../api/client'
import { stripCaretCodes } from '../utils/highlight'

export default function WavesForQuestion() {
  const [searchParams, setSearchParams] = useSearchParams()
  const question = searchParams.get('show_q_waves') || ''
  const mnemo = searchParams.get('show_q_mnemo') || ''

  const [waves, setWaves] = useState([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    if (!question) return
    setLoading(true)
    fetchWavesForQuestion(question, mnemo).then(setWaves).finally(() => setLoading(false))
  }, [question, mnemo])

  const goBack = () => setSearchParams({})

  const goToWave = (wave) => {
    setSearchParams({ show_wave: wave, hl_q: question, hl_q_mnemo: mnemo })
  }

  return (
    <div style={{ padding: '1rem 1.5rem' }}>
      <button className="back-link" onClick={goBack}>Back to search</button>

      <h1>Waves containing this question</h1>
      <div className="question-highlight-box">{stripCaretCodes(question)}</div>

      {loading ? (
        <div className="loading">Loading…</div>
      ) : (
        <>
          <h3 style={{ marginBottom: '0.75rem' }}>{waves.length} wave(s)</h3>
          {waves.length > 0 ? (
            <div className="wave-links-container">
              {waves.map(w => (
                <button key={w} className="wave-link" onClick={() => goToWave(w)}>
                  {w}
                </button>
              ))}
            </div>
          ) : (
            <p style={{ color: '#888' }}>No waves found for this question.</p>
          )}
        </>
      )}
    </div>
  )
}
