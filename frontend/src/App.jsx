import { useState, useEffect } from 'react'
import {
  BookOpen,
  Download,
  Sliders,
  Loader2,
  CheckCircle2,
  AlertCircle,
  X,
  Layers,
  FileText,
  Music,
  Merge,
} from 'lucide-react'
import SettingsPanel from './components/SettingsPanel.jsx'
import PreviewPanel from './components/PreviewPanel.jsx'
import SongSlidePage from './SongSlidePage.jsx'
import MergeSlidesPage from './MergeSlidesPage.jsx'
import { fetchPreview, generatePptx } from './api.js'
import { DEFAULT_SETTINGS } from './constants.js'

function useDebounce(value, delay) {
  const [debounced, setDebounced] = useState(value)
  useEffect(() => {
    const t = setTimeout(() => setDebounced(value), delay)
    return () => clearTimeout(t)
  }, [value, delay])
  return debounced
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
      <button onClick={onClose} className="ml-1 text-slate-400 hover:text-slate-600 transition">
        <X size={13} />
      </button>
    </div>
  )
}

const PLACEHOLDER = `Example — paste Khmer numbered verses:

12 ដំក់ដិណាដាតិទាំទូរយេ្រីង គំជាពួកស្លោដែលមានក្រាប់គាម
13 នោះក័មានល្ហាចមានព្រីករេ្យីង ជាថ្ងៃទីរំ

Or English references:

John 3:16
For God so loved the world that he gave his one and only Son.

Romans 8:28
And we know that in all things God works for the good of those who love him.`

export default function App() {
  const [page, setPage]             = useState('bible')   // 'bible' | 'song' | 'merge'
  const [rawText, setRawText]       = useState('')
  const [reference, setReference]   = useState('')
  const [perSlide, setPerSlide]     = useState(1)
  const [settings, setSettings]     = useState(DEFAULT_SETTINGS)
  const [slides, setSlides]         = useState([])
  const [currentSlide, setCurrentSlide] = useState(0)
  const [previewLoading, setPreviewLoading] = useState(false)
  const [generating, setGenerating] = useState(false)
  const [toast, setToast]           = useState(null)
  const [detected, setDetected]     = useState(null)

  const dRawText   = useDebounce(rawText,   350)
  const dReference = useDebounce(reference, 300)
  const dPerSlide  = useDebounce(perSlide,  200)
  const dSettings  = useDebounce(settings,  200)

  useEffect(() => {
    if (!dRawText.trim()) {
      setSlides([])
      setCurrentSlide(0)
      setDetected(null)
      return
    }
    let cancelled = false
    setPreviewLoading(true)
    fetchPreview({ raw_text: dRawText, per_slide: dPerSlide, settings: dSettings, reference: dReference })
      .then(data => {
        if (cancelled) return
        setSlides(data.slides)
        setCurrentSlide(0)
        if (data.slides.some(s => s.verse_num)) setDetected('khmer')
        else if (data.slides.some(s => s.ref))  setDetected('english')
        else setDetected(null)
      })
      .catch(() => {
        if (!cancelled) setToast({ message: 'Preview error — is the backend running?', type: 'error' })
      })
      .finally(() => { if (!cancelled) setPreviewLoading(false) })
    return () => { cancelled = true }
  }, [dRawText, dReference, dPerSlide, dSettings])

  const handleGenerate = async () => {
    if (!rawText.trim()) {
      setToast({ message: 'Please paste some Bible verses first.', type: 'error' })
      return
    }
    setGenerating(true)
    try {
      const blob = await generatePptx({ raw_text: rawText, per_slide: perSlide, settings, reference })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'bible_slides.pptx'
      a.click()
      URL.revokeObjectURL(url)
      setToast({ message: `${slides.length} slides downloaded!`, type: 'success' })
    } catch {
      setToast({ message: 'Generation failed — check backend.', type: 'error' })
    } finally {
      setGenerating(false)
    }
  }

  return (
    <div className={`bg-slate-50 flex flex-col ${page === 'merge' ? 'h-screen overflow-hidden' : 'min-h-screen'}`}>

      {/* Header */}
      <header className="bg-white border-b border-slate-100 sticky top-0 z-40">
        <div className="max-w-[1600px] mx-auto px-6 h-14 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-lg bg-brand-500 flex items-center justify-center shadow-sm">
              <BookOpen size={15} className="text-white" />
            </div>
            <span className="font-bold text-slate-800 tracking-tight">ChurchTool</span>

            {/* Page tabs */}
            <div className="hidden sm:flex items-center gap-1 ml-3 bg-slate-100 rounded-lg p-1">
              <button
                onClick={() => setPage('bible')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition-all
                  ${page === 'bible'
                    ? 'bg-white text-brand-600 shadow-sm'
                    : 'text-slate-500 hover:text-slate-700'}`}
              >
                <BookOpen size={12} />
                Bible Slides
              </button>
              <button
                onClick={() => setPage('song')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition-all
                  ${page === 'song'
                    ? 'bg-white text-brand-600 shadow-sm'
                    : 'text-slate-500 hover:text-slate-700'}`}
              >
                <Music size={12} />
                Song Slides
              </button>
              <button
                onClick={() => setPage('merge')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition-all
                  ${page === 'merge'
                    ? 'bg-white text-brand-600 shadow-sm'
                    : 'text-slate-500 hover:text-slate-700'}`}
              >
                <Merge size={12} />
                Merge Slides
              </button>
            </div>
          </div>

          <div className="flex items-center gap-2.5">
            {page === 'bible' && slides.length > 0 && (
              <span className="hidden sm:inline-flex items-center gap-1.5 text-xs text-slate-500 bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1.5 font-medium">
                <Layers size={12} className="text-brand-500" />
                {slides.length} slide{slides.length !== 1 ? 's' : ''}
              </span>
            )}
            {page === 'bible' && (
              <button
                onClick={handleGenerate}
                disabled={generating || !rawText.trim()}
                className="btn-primary"
              >
                {generating
                  ? <><Loader2 size={14} className="animate-spin" />Generating…</>
                  : <><Download size={14} />Download PPTX</>}
              </button>
            )}
          </div>
        </div>

        {/* Mobile page tabs */}
        <div className="sm:hidden flex border-t border-slate-100">
          <button
            onClick={() => setPage('bible')}
            className={`flex-1 flex items-center justify-center gap-1.5 py-2.5 text-xs font-semibold border-b-2 transition
              ${page === 'bible' ? 'border-brand-500 text-brand-600' : 'border-transparent text-slate-400'}`}
          >
            <BookOpen size={12} /> Bible Slides
          </button>
          <button
            onClick={() => setPage('song')}
            className={`flex-1 flex items-center justify-center gap-1.5 py-2.5 text-xs font-semibold border-b-2 transition
              ${page === 'song' ? 'border-brand-500 text-brand-600' : 'border-transparent text-slate-400'}`}
          >
            <Music size={12} /> Song Slides
          </button>
          <button
            onClick={() => setPage('merge')}
            className={`flex-1 flex items-center justify-center gap-1.5 py-2.5 text-xs font-semibold border-b-2 transition
              ${page === 'merge' ? 'border-brand-500 text-brand-600' : 'border-transparent text-slate-400'}`}
          >
            <Merge size={12} /> Merge
          </button>
        </div>
      </header>

      {/* Page content */}
      {page === 'song' ? (
        <SongSlidePage />
      ) : page === 'merge' ? (
        <MergeSlidesPage />
      ) : (
        /* ── Bible Slides 3-column layout ── */
        <div className="flex-1 max-w-[1600px] mx-auto w-full px-4 py-5
                        grid grid-cols-1
                        lg:grid-cols-[360px_1fr_300px]
                        xl:grid-cols-[380px_1fr_320px]
                        gap-5 items-start">

          {/* LEFT: Input */}
          <div className="space-y-4 lg:sticky lg:top-20">
            <div className="card p-4 space-y-3">
              <div className="flex items-center justify-between">
                <h2 className="font-semibold text-slate-800 text-sm flex items-center gap-2">
                  <FileText size={14} className="text-brand-500" />
                  Bible Text
                </h2>
                {detected && (
                  <span className={`text-[10px] font-semibold px-2 py-0.5 rounded-full ${
                    detected === 'khmer'
                      ? 'bg-amber-50 text-amber-600 border border-amber-200'
                      : 'bg-blue-50 text-blue-600 border border-blue-200'
                  }`}>
                    {detected === 'khmer' ? '🇰🇭 Khmer verses' : '📖 English refs'}
                  </span>
                )}
              </div>
              <textarea
                rows={16}
                className="input resize-none text-[13px] leading-relaxed"
                placeholder={PLACEHOLDER}
                value={rawText}
                onChange={e => setRawText(e.target.value)}
              />
            </div>

            <div className="card p-4 space-y-3">
              <h3 className="text-[10px] font-bold uppercase tracking-widest text-slate-400">Options</h3>
              <div>
                <label className="label">Global Reference</label>
                <input
                  type="text"
                  className="input"
                  placeholder="e.g. Matthew 5:12-20"
                  value={reference}
                  onChange={e => setReference(e.target.value)}
                />
              </div>
              <div>
                <label className="label">Lines per Slide</label>
                <select
                  className="select"
                  value={perSlide}
                  onChange={e => setPerSlide(parseInt(e.target.value, 10))}
                >
                  {[1,2,3,4,5,6].map(n => (
                    <option key={n} value={n}>{n} line{n > 1 ? 's' : ''}</option>
                  ))}
                </select>
              </div>
              <button
                onClick={handleGenerate}
                disabled={generating || !rawText.trim()}
                className="btn-primary w-full"
              >
                {generating
                  ? <><Loader2 size={14} className="animate-spin" />Generating…</>
                  : <><Download size={14} />Generate &amp; Download PPTX</>}
              </button>
            </div>
          </div>

          {/* CENTRE: Preview */}
          <div className="card p-6 lg:sticky lg:top-20">
            <PreviewPanel
              slides={slides}
              settings={settings}
              current={currentSlide}
              onNavigate={setCurrentSlide}
              isLoading={previewLoading}
            />
          </div>

          {/* RIGHT: Settings */}
          <div className="card lg:sticky lg:top-20">
            <div className="flex items-center gap-2 px-4 pt-4 pb-3 border-b border-slate-100">
              <Sliders size={14} className="text-brand-500" />
              <span className="font-semibold text-sm text-slate-800">Slide Settings</span>
            </div>
            <div className="px-4 py-4 max-h-[80vh] overflow-y-auto">
              <SettingsPanel settings={settings} onChange={setSettings} />
            </div>
          </div>

        </div>
      )}

      {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
    </div>
  )
}
