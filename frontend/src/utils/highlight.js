export function stripCaretCodes(text) {
  if (!text) return text
  return text.replace(/\^[^^]*\^/g, '')
}

function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function escapeRegex(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

export function highlightText(text, exactTerms, expandedTerms) {
  text = stripCaretCodes(text)
  if (!text || (!exactTerms?.length && !expandedTerms?.length)) {
    return escapeHtml(text || '')
  }

  const escaped = escapeHtml(text)
  const greenSpans = []
  const yellowSpans = []

  for (const term of exactTerms || []) {
    if (!term) continue
    const re = new RegExp(escapeRegex(escapeHtml(term)), 'gi')
    let m
    while ((m = re.exec(escaped)) !== null) {
      greenSpans.push([m.index, m.index + m[0].length])
    }
  }

  for (const term of expandedTerms || []) {
    if (!term) continue
    const escapedTerm = escapeRegex(escapeHtml(term))
    // Use word boundaries only for short terms (single words/phrases);
    // for full question strings, skip \b to avoid mismatch on trailing '?'
    const pattern = term.length > 30
      ? escapedTerm
      : '\\b' + escapedTerm + '\\b'
    const re = new RegExp(pattern, 'gi')
    let m
    while ((m = re.exec(escaped)) !== null) {
      yellowSpans.push([m.index, m.index + m[0].length])
    }
  }

  if (!greenSpans.length && !yellowSpans.length) return escaped

  // Yellow takes priority over green
  const finalSpans = yellowSpans.map(([s, e]) => [s, e, '#FFFF99'])

  for (const [gs, ge] of greenSpans) {
    let remaining = [[gs, ge]]
    for (const [ys, ye] of yellowSpans) {
      const newRemaining = []
      for (const [rs, rend] of remaining) {
        if (ye <= rs || ys >= rend) {
          newRemaining.push([rs, rend])
        } else {
          if (rs < ys) newRemaining.push([rs, ys])
          if (rend > ye) newRemaining.push([ye, rend])
        }
      }
      remaining = newRemaining
    }
    for (const [rs, rend] of remaining) {
      finalSpans.push([rs, rend, '#90EE90'])
    }
  }

  finalSpans.sort((a, b) => a[0] - b[0])

  let result = ''
  let last = 0
  for (let [start, end, color] of finalSpans) {
    if (start < last) start = last
    if (start >= end) continue
    result += escaped.slice(last, start)
    result += `<mark style="background:${color};padding:1px 2px;border-radius:2px">${escaped.slice(start, end)}</mark>`
    last = end
  }
  result += escaped.slice(last)
  return result
}
