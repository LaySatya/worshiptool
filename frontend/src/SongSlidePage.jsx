import { useState, useCallback, useEffect, useRef } from 'react'
import {
  Music,
  Upload,
  FileText,
  Download,
  Sliders,
  Loader2,
  CheckCircle2,
  AlertCircle,
  X,
  Layers,
  Sparkles,
  FileUp,
  Settings,
} from 'lucide-react'
import SongSettingsPanel from './components/SongSettingsPanel.jsx'
import SongPreviewPanel from './components/SongPreviewPanel.jsx'
import { extractSongLyrics, fetchSongPreview, generateSongPptx } from './api.js'
import { DEFAULT_SONG_SETTINGS } from './constants.js'

export const SAMPLE_HYMN_299 = `299 នៅពេញលោកិយទាំងមូល
온 세상 위하여 (통일 268장, 새 505장)
HYMN

១~នៅពេញលោកិយទាំងមូល យើងនឹងផ្សាយដំណឹងល្អ ដល់មនុស្សជាច្រើននៅគ្រប់ជាតិសាសន៍ ព្រះបន្ទូលព្រះអម្ចាស់ គេនឹងវង្វេងហើយស្លាប់ បើគ្មានព្រះបន្ទូលទ្រង់ គេគ្មានព្រះយេស៊ូវជួយសង្គ្រោះ ព្រោះមិនដែលធ្លាប់ឮឡើយ បន្ទរ៖ផ្សាយដំណឹង អធិស្ឋាន ធ្វើជាសាក្សី រៀងរាល់ថ្ងៃ មនុស្សរាប់លាននាក់ក្នុងលោកិយ នេះនឹងដឹងក្តីស្រឡាញ់ទ្រង់ ២.នៅពេញលោកិយទាំងមូល ព្រះបន្ទូលទ្រង់ ផ្សាយទៅ ដល់មនុស្សជាច្រើន ដែលស្លាប់ ក្នុងបាប ដើម្បីឲ្យគេប្រែចិត្ត ទ្រង់សុគត សង្គ្រោះគេ តែគេមិនបានដឹងឡើយ លុះយើងនាំអំណោយក្តីស្រឡាញ់~ និងដំណឹងល្អចេញទៅ បន្ទរ៖ផ្សាយដំណឹង អធិស្ឋាន ធ្វើជាសាក្សី រៀងរាល់ថ្ងៃ មនុស្សរាប់លាននាក់ក្នុងលោកិយ នេះនឹងដឹងក្តីស្រឡាញ់ទ្រង់ ៣.នៅពេញលោកិយទាំងមូល អ្នកនាំសារទ្រង់ប្រកាស ក្តីសង្គ្រោះដល់មនុស្សគ្រប់ទិសទី ក្នុងព្រះនាមព្រះយេស៊ូវ ឯយើងដែលមិនអាចផ្សាយ ដំណឹងល្អ ទៅទីឆ្ងាយ នឹងជួយអធិស្ឋាន ដល់ពួកគេ~ ហើយធ្វើជាសាក្សីទ្រង់ បន្ទរ៖ផ្សាយដំណឹង អធិស្ឋាន ធ្វើជាសាក្សី រៀងរាល់ថ្ងៃ មនុស្សរាប់លាននាក់ក្នុងលោកិយ នេះនឹងដឹងក្តីស្រឡាញ់ទ្រង់`

function useDebounce(value, delay) {
  const [d, setD] = useState(value)
  useEffect(() => {
    const t = setTimeout(() => setD(value), delay)
    return () => clearTimeout(t)
  }, [value, delay])
  return d
}

function Toast({ message, type = 'success', onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 4500)
    return () => clearTimeout(t)
  }, [onClose])
  return (
    <div className={`fixed bottom-5 right-5 z-50 flex items-center gap-3 px-4 py-3 rounded-xl shadow-panel border text-sm font-medium animate-fade-in
      ${type === 'success' ? 'bg-white border-green-200 text-green-700' : 'bg-white border-red-200 text-red-700'}`}>
      {type === 'success' ? <CheckCircle2 size={15} /> : <AlertCircle size={15} />}
      {message}
      <button onClick={onClose} className="ml-1 text-slate-400 hover:text-slate-600">
        <X size={13} />
      </button>
    </div>
  )
}

export default function SongSlidePage() {
  const [lyricsText, setLyricsText] = useState(SAMPLE_HYMN_299)
  const [songTitle, setSongTitle] = useState('299 នៅពេញលោកិយទាំងមូល')
  const [settings, setSettings] = useState(DEFAULT_SONG_SETTINGS)
  const [slides, setSlides] = useState([])
  const [currentSlide, setCurrentSlide] = useState(0)
  const [previewLoad, setPreviewLoad] = useState(false)
  const [generating, setGenerating] = useState(false)
  const [extracting, setExtracting] = useState(false)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [toast, setToast] = useState(null)

  const fileInputRef = useRef(null)

  const dLyrics   = useDebounce(lyricsText, 350)
  const dSettings = useDebounce(settings, 250)
  const dTitle    = useDebounce(songTitle, 250)

  // ── Auto-preview ────────────────────────────────────────────────
  useEffect(() => {
    if (!dLyrics.trim()) {
      setSlides([])
      setCurrentSlide(0)
      return
    }
    let cancelled = false
    setPreviewLoad(true)
    fetchSongPreview({ lyrics_text: dLyrics, settings: dSettings, song_title: dTitle })
      .then(data => {
        if (cancelled) return
        setSlides(data.slides ?? [])
        setCurrentSlide(0)
      })
      .catch(() => {
        if (!cancelled) setToast({ message: 'Preview error — check backend.', type: 'error' })
      })
      .finally(() => { if (!cancelled) setPreviewLoad(false) })
    return () => { cancelled = true }
  }, [dLyrics, dSettings, dTitle])

  // ── File Import (.txt directly or OCR for images/PDF) ───────────
  const handleFileImport = useCallback(async (file) => {
    if (!file) return

    // 1. Text file: read immediately into textarea
    if (file.name.toLowerCase().endsWith('.txt') || file.type.startsWith('text/')) {
      try {
        const text = await file.text()
        setLyricsText(text)
        const guessedTitle = file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ')
        setSongTitle(guessedTitle)
        setToast({ message: `Imported "${file.name}"! Slides ready.`, type: 'success' })
      } catch {
        setToast({ message: 'Could not read text file.', type: 'error' })
      }
      return
    }

    // 2. Music sheet image or PDF: run OCR
    if (file.type.startsWith('image/') || file.name.toLowerCase().endsWith('.pdf')) {
      setExtracting(true)
      try {
        const data = await extractSongLyrics(file)
        if (data.text) {
          setLyricsText(data.text)
          const guessedTitle = file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ')
          setSongTitle(guessedTitle)
          setToast({ message: 'Lyrics extracted from music sheet!', type: 'success' })
        } else {
          setToast({ message: 'No text found in music sheet.', type: 'error' })
        }
      } catch {
        setToast({ message: 'OCR failed — check backend.', type: 'error' })
      } finally {
        setExtracting(false)
      }
    }
  }, [])

  // ── Drag & drop handlers ─────────────────────────────────────────
  const onDrop = (e) => {
    e.preventDefault()
    const file = e.dataTransfer?.files?.[0]
    if (file) handleFileImport(file)
  }

  // ── Load Sample ──────────────────────────────────────────────────
  const handleLoadSample = () => {
    setLyricsText(SAMPLE_HYMN_299)
    setSongTitle('299 នៅពេញលោកិយទាំងមូល')
    setToast({ message: 'Loaded Hymn 299 sample!', type: 'success' })
  }

  // ── Generate PPTX ───────────────────────────────────────────────
  const handleGenerate = async () => {
    if (!lyricsText.trim()) {
      setToast({ message: 'Please enter or import song lyrics first.', type: 'error' })
      return
    }
    setGenerating(true)
    try {
      const blob = await generateSongPptx({
        lyrics_text: lyricsText,
        settings,
        song_title: songTitle,
      })
      const url = URL.createObjectURL(blob)
      const a   = document.createElement('a')
      a.href     = url
      a.download = `${songTitle || 'song'}_slides.pptx`
      a.click()
      URL.revokeObjectURL(url)
      setToast({ message: `${slides.length} slides downloaded!`, type: 'success' })
    } catch {
      setToast({ message: 'PPTX generation failed — check backend.', type: 'error' })
    } finally {
      setGenerating(false)
    }
  }

  return (
    <div className="flex-1 overflow-hidden flex flex-col h-full bg-slate-50">
      {/* Hidden file input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".txt,.pdf,.png,.jpg,.jpeg"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0]
          if (f) handleFileImport(f)
          e.target.value = ''
        }}
      />

      {/* Top action header for Song Slides */}
      <div className="bg-white border-b border-slate-200 px-6 py-3 flex items-center justify-between flex-shrink-0 z-10 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-brand-50 text-brand-600 flex items-center justify-center font-bold">
            <Music size={16} />
          </div>
          <div>
            <h2 className="text-sm font-bold text-slate-800">Khmer Song &amp; Hymn Slide Generator</h2>
            <p className="text-xs text-slate-400">Title slide, verses (១, ២...), and chorus with red R. marker</p>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            disabled={extracting}
            className="btn-secondary !py-2 !px-3 text-xs flex items-center gap-1.5"
            title="Import text file (.txt)"
          >
            {extracting ? <Loader2 size={13} className="animate-spin" /> : <FileUp size={13} className="text-brand-500" />}
            <span>Import .txt File</span>
          </button>

          <button
            type="button"
            onClick={handleLoadSample}
            className="btn-secondary !py-2 !px-3 text-xs flex items-center gap-1.5 hover:border-brand-300 hover:text-brand-600"
            title="Load Hymn 299 as sample"
          >
            <Sparkles size={13} className="text-amber-500" />
            <span>Load Sample Hymn 299</span>
          </button>

          <button
            type="button"
            onClick={() => setSettingsOpen(prev => !prev)}
            className={`btn-secondary !py-2 !px-3 text-xs flex items-center gap-1.5 ${settingsOpen ? 'bg-brand-50 text-brand-600 border-brand-200' : ''}`}
          >
            <Settings size={14} />
            <span>{settingsOpen ? 'Hide Settings' : 'Slide Settings'}</span>
          </button>

          <button
            type="button"
            onClick={handleGenerate}
            disabled={generating || !lyricsText.trim()}
            className="btn-primary !py-2 !px-4 text-xs flex items-center gap-1.5"
          >
            {generating ? (
              <><Loader2 size={14} className="animate-spin" />Generating…</>
            ) : (
              <><Download size={14} />Download PPTX</>
            )}
          </button>
        </div>
      </div>

      {/* Main Workspace (Side-by-Side: Lyrics input on left, Live preview on right) */}
      <div className="flex-1 overflow-hidden flex">
        {/* Left column: Text Input & Drop zone */}
        <div className="flex-1 overflow-y-auto p-5 min-w-0">
          <div className="max-w-[1500px] mx-auto grid grid-cols-1 lg:grid-cols-2 gap-5 h-full">

            {/* Input Card */}
            <div
              className="card p-5 flex flex-col h-full min-h-[500px]"
              onDragOver={(e) => e.preventDefault()}
              onDrop={onDrop}
            >
              <div className="flex items-center justify-between flex-shrink-0 pb-2 border-b border-slate-100">
                <div className="flex items-center gap-2">
                  <FileText size={16} className="text-brand-500" />
                  <h3 className="font-semibold text-slate-800 text-sm">Song Lyrics (.txt or paste)</h3>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => fileInputRef.current?.click()}
                    className="text-xs text-brand-600 hover:underline flex items-center gap-1 font-medium"
                  >
                    <Upload size={12} />
                    Browse .txt
                  </button>
                  {lyricsText && (
                    <button
                      onClick={() => { setLyricsText(''); setSongTitle(''); }}
                      className="text-xs text-slate-400 hover:text-red-500 transition font-medium ml-2"
                    >
                      Clear
                    </button>
                  )}
                </div>
              </div>

              {/* Textarea */}
              <div className="flex-1 flex flex-col mt-3 min-h-0">
                <textarea
                  className="input resize-none text-[13px] leading-relaxed flex-1 font-sans min-h-[300px]"
                  placeholder={`Example format:\n\n299 នៅពេញលោកិយទាំងមូល\n온 세상 위하여 (통일 268장, 새 505장)\nHYMN\n\n១~នៅពេញលោកិយទាំងមូល យើងនឹងផ្សាយដំណឹងល្អ... បន្ទរ៖ផ្សាយដំណឹង អធិស្ឋាន... ២.នៅពេញលោកិយទាំងមូល...`}
                  value={lyricsText}
                  onChange={(e) => setLyricsText(e.target.value)}
                />
              </div>

              {/* Quick instructions / tips */}
              <div className="mt-3 p-3 bg-slate-50 rounded-xl border border-slate-100 text-[11px] text-slate-500 space-y-1">
                <p className="font-semibold text-slate-700">💡 Song Slide Rules:</p>
                <ul className="list-disc list-inside space-y-0.5 text-slate-600">
                  <li><strong>Header lines:</strong> First lines before verse become the <strong>Title Slide</strong>.</li>
                  <li><strong>Verses:</strong> Marked with <code className="px-1 py-0.5 rounded bg-slate-200 text-slate-800">១~</code>, <code className="px-1 py-0.5 rounded bg-slate-200 text-slate-800">២.</code>, <code className="px-1 py-0.5 rounded bg-slate-200 text-slate-800">1.</code>, etc. (each verse = one slide).</li>
                  <li><strong>Chorus:</strong> Word <code className="px-1 py-0.5 rounded bg-slate-200 text-slate-800">បន្ទរ</code> is replaced by <strong className="text-red-500">R.</strong> (colored red) on its own slide.</li>
                </ul>
              </div>
            </div>

            {/* Right Card: Live Preview */}
            <div className="card p-5 flex flex-col h-full overflow-hidden">
              <SongPreviewPanel
                slides={slides}
                settings={settings}
                current={currentSlide}
                onNavigate={setCurrentSlide}
                isLoading={previewLoad}
              />
            </div>

          </div>
        </div>

        {/* Collapsible Slide Settings Drawer on the right */}
        <div
          className={`bg-white border-l border-slate-200 transition-all duration-300 overflow-hidden flex-shrink-0 z-20
            ${settingsOpen ? 'w-80' : 'w-0'}`}
        >
          {settingsOpen && (
            <div className="w-80 h-full overflow-y-auto">
              <div className="sticky top-0 bg-white border-b border-slate-100 px-5 py-4 z-10 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Sliders size={16} className="text-brand-500" />
                  <span className="font-semibold text-sm text-slate-800">Song Slide Settings</span>
                </div>
                <button
                  onClick={() => setSettingsOpen(false)}
                  className="p-1 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition"
                >
                  <X size={16} />
                </button>
              </div>
              <div className="p-5">
                <SongSettingsPanel settings={settings} onChange={setSettings} />
              </div>
            </div>
          )}
        </div>
      </div>

      {toast && (
        <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />
      )}
    </div>
  )
}
