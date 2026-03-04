/** Thin API wrapper — all calls go through Vite's /api proxy to localhost:8000 */

const BASE = '/api'

export async function fetchPreview(payload) {
  const res = await fetch(`${BASE}/preview`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
  if (!res.ok) throw new Error('Preview failed')
  return res.json()
}

export async function generatePptx(payload) {
  const res = await fetch(`${BASE}/generate`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
  if (!res.ok) throw new Error('Generation failed')
  return res.blob()
}

// ── Song / OCR endpoints ────────────────────────────────────────

export async function extractSongLyrics(file) {
  const form = new FormData()
  form.append('file', file)
  const res = await fetch(`${BASE}/song/extract`, { method: 'POST', body: form })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.error || 'OCR extraction failed')
  }
  return res.json()  // { text: string }
}

export async function fetchSongPreview(payload) {
  const res = await fetch(`${BASE}/song/preview`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
  if (!res.ok) throw new Error('Song preview failed')
  return res.json()  // { slides: [...], total: number }
}

export async function generateSongPptx(payload) {
  const res = await fetch(`${BASE}/song/generate`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
  if (!res.ok) throw new Error('Song generation failed')
  return res.blob()
}

// ── Merge Slides endpoints ───────────────────────────────────────

/**
 * Count the number of slides in a single PPTX file.
 * @param {File} file
 * @returns {Promise<{slide_count: number}>}
 */
export async function countSlides(file) {
  const form = new FormData()
  form.append('file', file)
  const res = await fetch(`${BASE}/merge/count-slides`, { method: 'POST', body: form })
  if (!res.ok) throw new Error('Count slides failed')
  return res.json()
}

/**
 * Merge multiple PPTX files into one and return the blob.
 * @param {File[]} files - ordered list of PPTX files
 * @param {string} outputName - desired filename (without extension)
 * @returns {Promise<Blob>}
 */
export async function mergePptxFiles(files, outputName = 'merged_slides') {
  const form = new FormData()
  files.forEach(f => form.append('files', f))
  form.append('output_name', outputName)
  const res = await fetch(`${BASE}/merge/generate`, { method: 'POST', body: form })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || 'Merge failed')
  }
  return res.blob()
}
