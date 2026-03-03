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
  RefreshCw,
  Pencil,
  Eye,
  FileImage,
  Sparkles,
  ChevronRight,
  Clipboard,
  Image as ImageIcon,
} from 'lucide-react'
import SongSettingsPanel from './components/SongSettingsPanel.jsx'
import SongPreviewPanel from './components/SongPreviewPanel.jsx'
import { extractSongLyrics, fetchSongPreview, generateSongPptx } from './api.js'
import { DEFAULT_SONG_SETTINGS } from './constants.js'

// ── useDebounce ──────────────────────────────────────────────────
function useDebounce(value, delay) {
  const [d, setD] = useState(value)
  useEffect(() => {
    const t = setTimeout(() => setD(value), delay)
    return () => clearTimeout(t)
  }, [value, delay])
  return d
}

// ── Toast ────────────────────────────────────────────────────────
function Toast({ message, type = 'success', onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 5000)
    return () => clearTimeout(t)
  }, [onClose])
  return (
    <div className={`fixed bottom-5 right-5 z-50 flex items-center gap-3 px-4 py-3 rounded-xl shadow-panel border text-sm font-medium animate-fade-in
      ${type === 'success'
        ? 'bg-white border-green-200 text-green-700'
        : type === 'info'
          ? 'bg-white border-blue-200 text-blue-700'
          : 'bg-white border-red-200 text-red-700'}`}>
      {type === 'success' ? <CheckCircle2 size={15} />
        : type === 'info'  ? <Sparkles size={15} />
        : <AlertCircle size={15} />}
      {message}
      <button onClick={onClose} className="ml-1 text-slate-400 hover:text-slate-600">
        <X size={13} />
      </button>
    </div>
  )
}

// ── DropZone ─────────────────────────────────────────────────────
function DropZone({ onFile, extracting, fileName }) {
  const inputRef        = useRef(null)
  const zoneRef         = useRef(null)
  const [drag, setDrag] = useState(false)
  const [pasteHint, setPasteHint] = useState(false) // flash when paste lands

  // ── drag-drop ──────────────────────────────────────────────────
  const handleDrop = useCallback(e => {
    e.preventDefault()
    setDrag(false)
    const f = e.dataTransfer?.files?.[0]
    if (f) onFile(f)
  }, [onFile])

  const handleDragOver  = e => { e.preventDefault(); setDrag(true) }
  const handleDragLeave = ()  => setDrag(false)

  // ── clipboard paste (global Ctrl+V / ⌘+V) ────────────────────
  useEffect(() => {
    const onPaste = e => {
      if (extracting) return
      const items = Array.from(e.clipboardData?.items ?? [])

      // 1️⃣ Image blob pasted (e.g. screenshot or copied image)
      const imgItem = items.find(i => i.type.startsWith('image/'))
      if (imgItem) {
        const blob = imgItem.getAsFile()
        if (blob) {
          // Wrap in a proper File so handleFile can read it
          const ext  = blob.type === 'image/png' ? '.png' : '.jpg'
          const file = new File([blob], `pasted_image${ext}`, { type: blob.type })
          setPasteHint(true)
          setTimeout(() => setPasteHint(false), 1800)
          onFile(file)
        }
        return
      }
    }
    window.addEventListener('paste', onPaste)
    return () => window.removeEventListener('paste', onPaste)
  }, [onFile, extracting])

  // ── "Paste from clipboard" button (reads Clipboard API) ───────
  const handlePasteButton = async () => {
    if (extracting) return
    try {
      const clipItems = await navigator.clipboard.read()
      for (const item of clipItems) {
        const imgType = item.types.find(t => t.startsWith('image/'))
        if (imgType) {
          const blob = await item.getType(imgType)
          const ext  = imgType === 'image/png' ? '.png' : '.jpg'
          const file = new File([blob], `pasted_image${ext}`, { type: imgType })
          setPasteHint(true)
          setTimeout(() => setPasteHint(false), 1800)
          onFile(file)
          return
        }
      }
      // No image in clipboard
      alert('No image found in clipboard.\nCopy a screenshot or image first, then click Paste.')
    } catch {
      // Clipboard API blocked (user denied permission) → tell user to use Ctrl+V
      alert('Clipboard access was blocked.\nClick inside the drop zone first, then press Ctrl+V (or ⌘+V on Mac).')
    }
  }

  return (
    <div className="space-y-2">
      {/* Main drop target */}
      <div
        ref={zoneRef}
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onClick={() => !extracting && inputRef.current?.click()}
        className={`relative flex flex-col items-center justify-center gap-3 rounded-2xl border-2 border-dashed cursor-pointer transition-all p-8
          ${drag
            ? 'border-brand-400 bg-brand-50 scale-[1.01]'
            : pasteHint
              ? 'border-green-400 bg-green-50 scale-[1.01]'
              : extracting
                ? 'border-amber-300 bg-amber-50 cursor-default'
                : fileName
                  ? 'border-green-300 bg-green-50'
                  : 'border-slate-200 bg-slate-50 hover:border-brand-300 hover:bg-brand-50/40'}`}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".pdf,.png,.jpg,.jpeg"
          className="hidden"
          onChange={e => { const f = e.target.files?.[0]; if (f) onFile(f) }}
        />

        {extracting ? (
          <>
            <div className="w-12 h-12 rounded-full bg-amber-100 flex items-center justify-center">
              <Loader2 size={22} className="text-amber-500 animate-spin" />
            </div>
            <div className="text-center">
              <p className="font-semibold text-amber-700 text-sm">Extracting lyrics…</p>
              <p className="text-xs text-amber-500 mt-0.5">Running OCR — this may take a moment</p>
            </div>
          </>
        ) : pasteHint ? (
          <>
            <div className="w-12 h-12 rounded-full bg-green-100 flex items-center justify-center">
              <CheckCircle2 size={22} className="text-green-500" />
            </div>
            <div className="text-center">
              <p className="font-semibold text-green-700 text-sm">Image pasted!</p>
              <p className="text-xs text-green-500 mt-0.5">Starting OCR…</p>
            </div>
          </>
        ) : fileName ? (
          <>
            <div className="w-12 h-12 rounded-full bg-green-100 flex items-center justify-center">
              <CheckCircle2 size={22} className="text-green-500" />
            </div>
            <div className="text-center">
              <p className="font-semibold text-green-700 text-sm truncate max-w-[220px]">{fileName}</p>
              <p className="text-xs text-green-500 mt-0.5">Lyrics extracted — click to re-upload</p>
            </div>
          </>
        ) : (
          <>
            <div className="w-12 h-12 rounded-full bg-slate-100 flex items-center justify-center transition group-hover:bg-brand-100">
              <Upload size={22} className="text-slate-400" />
            </div>
            <div className="text-center">
              <p className="font-semibold text-slate-700 text-sm">Drop your music sheet here</p>
              <p className="text-xs text-slate-400 mt-0.5">or click to browse · PDF, JPG, PNG</p>
            </div>
            <div className="flex items-center gap-2 mt-1">
              {['PDF', 'JPG', 'PNG'].map(fmt => (
                <span key={fmt} className="text-[10px] font-bold px-2 py-0.5 rounded bg-white border border-slate-200 text-slate-500">
                  {fmt}
                </span>
              ))}
            </div>
          </>
        )}
      </div>

      {/* Paste button row */}
      {!extracting && !fileName && (
        <button
          type="button"
          onClick={handlePasteButton}
          className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl border border-dashed border-slate-200
                     bg-white hover:bg-slate-50 hover:border-brand-300 text-slate-500 hover:text-brand-600
                     text-xs font-medium transition-all"
        >
          <Clipboard size={13} />
          Paste image from clipboard
          <span className="ml-1 text-[10px] text-slate-300 font-normal">
            ({navigator.platform?.includes('Mac') ? '⌘V' : 'Ctrl+V'})
          </span>
        </button>
      )}
    </div>
  )
}

// ── ProcessingSteps ──────────────────────────────────────────────
function ProcessingStep({ step, label, active, done }) {
  return (
    <div className={`flex items-center gap-2 text-xs font-medium transition-colors
      ${done ? 'text-green-600' : active ? 'text-brand-600' : 'text-slate-300'}`}>
      <div className={`w-5 h-5 rounded-full border-2 flex items-center justify-center shrink-0
        ${done ? 'bg-green-500 border-green-500' : active ? 'border-brand-500' : 'border-slate-200'}`}>
        {done
          ? <CheckCircle2 size={11} className="text-white" />
          : active
            ? <Loader2 size={10} className="animate-spin text-brand-500" />
            : <span className="text-[9px] text-slate-400">{step}</span>}
      </div>
      {label}
    </div>
  )
}

// ── Main SongSlidePage ───────────────────────────────────────────
export default function SongSlidePage() {
  const [fileName,    setFileName]    = useState('')
  const [extracting,  setExtracting]  = useState(false)
  const [lyricsText,  setLyricsText]  = useState('')
  const [songTitle,   setSongTitle]   = useState('')
  const [settings,    setSettings]    = useState(DEFAULT_SONG_SETTINGS)
  const [slides,      setSlides]      = useState([])
  const [currentSlide,setCurrentSlide]= useState(0)
  const [previewLoad, setPreviewLoad] = useState(false)
  const [generating,  setGenerating]  = useState(false)
  const [toast,       setToast]       = useState(null)
  const [activeTab,   setActiveTab]   = useState('upload')  // 'upload' | 'edit'

  // OCR step tracker
  const [step, setStep] = useState(0)  // 0=idle, 1=uploading, 2=ocr, 3=done

  const dLyrics   = useDebounce(lyricsText, 400)
  const dSettings = useDebounce(settings,   300)
  const dTitle    = useDebounce(songTitle,   300)

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
        if (!cancelled) setToast({ message: 'Preview error — is the backend running?', type: 'error' })
      })
      .finally(() => { if (!cancelled) setPreviewLoad(false) })
    return () => { cancelled = true }
  }, [dLyrics, dSettings, dTitle])

  // ── File upload → OCR ───────────────────────────────────────────
  const handleFile = useCallback(async file => {
    setFileName(file.name)
    setExtracting(true)
    setStep(1)
    setActiveTab('upload')

    // Derive song title from filename (strip extension)
    const guessedTitle = file.name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ')
    setSongTitle(prev => prev || guessedTitle)

    try {
      setStep(2)
      const data = await extractSongLyrics(file)
      setLyricsText(data.text ?? '')
      setStep(3)
      setActiveTab('edit')
      setToast({
        message: 'Lyrics extracted! Review and correct the text below.',
        type: 'info',
      })
    } catch (err) {
      setToast({ message: err.message || 'OCR failed — check backend.', type: 'error' })
      setStep(0)
    } finally {
      setExtracting(false)
    }
  }, [])

  // ── Generate PPTX ───────────────────────────────────────────────
  const handleGenerate = async () => {
    if (!lyricsText.trim()) {
      setToast({ message: 'No lyrics to generate from.', type: 'error' })
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
      setToast({ message: `${slides.length} slide${slides.length !== 1 ? 's' : ''} downloaded!`, type: 'success' })
    } catch {
      setToast({ message: 'Generation failed — check backend.', type: 'error' })
    } finally {
      setGenerating(false)
    }
  }

  const canGenerate = lyricsText.trim().length > 0 && !generating

  return (
    <div className="flex-1 max-w-[1600px] mx-auto w-full px-4 py-5
                    grid grid-cols-1
                    lg:grid-cols-[400px_1fr_300px]
                    xl:grid-cols-[420px_1fr_320px]
                    gap-5 items-start">

      {/* ══ LEFT COLUMN ══════════════════════════════════════════ */}
      <div className="space-y-4 lg:sticky lg:top-20">

        {/* Upload + Edit tabs */}
        <div className="card overflow-hidden">
          {/* Tab bar */}
          <div className="flex border-b border-slate-100">
            {[
              { id: 'upload', icon: <Upload size={13} />,   label: 'Upload' },
              { id: 'edit',   icon: <Pencil size={13} />,   label: 'Edit Lyrics' },
            ].map(tab => (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`flex-1 flex items-center justify-center gap-1.5 py-3 text-xs font-semibold transition border-b-2
                  ${activeTab === tab.id
                    ? 'border-brand-500 text-brand-600 bg-brand-50/50'
                    : 'border-transparent text-slate-400 hover:text-slate-600 hover:bg-slate-50'}`}
              >
                {tab.icon}{tab.label}
              </button>
            ))}
          </div>

          <div className="p-4 space-y-4">
            {/* Upload tab */}
            {activeTab === 'upload' && (
              <>
                <DropZone onFile={handleFile} extracting={extracting} fileName={fileName} />

                {/* Processing steps */}
                <div className="space-y-2 pt-1">
                  <ProcessingStep step={1} label="File received"       active={step === 1} done={step > 1} />
                  <ProcessingStep step={2} label="Running OCR (Tesseract Khmer)" active={step === 2} done={step > 2} />
                  <ProcessingStep step={3} label="Lyrics extracted"    active={false}      done={step >= 3} />
                </div>

                {/* Tips */}
                <div className="rounded-xl bg-blue-50 border border-blue-100 px-3.5 py-3 space-y-1.5">
                  <p className="text-xs font-bold text-blue-700 flex items-center gap-1.5">
                    <Sparkles size={11} /> Tips for best OCR results
                  </p>
                  <ul className="text-[11px] text-blue-600 space-y-0.5 list-disc list-inside">
                    <li>Drag &amp; drop a PDF/image, click to browse, or paste a screenshot (⌘V / Ctrl+V)</li>
                    <li>Scan at 300 DPI or higher for sharper text</li>
                    <li>Ensure good contrast — avoid shadows or folds</li>
                    <li>Only Khmer text is extracted — chords and notation are ignored</li>
                    <li>Review &amp; correct the extracted text before generating</li>
                  </ul>
                </div>
              </>
            )}

            {/* Edit lyrics tab */}
            {activeTab === 'edit' && (
              <>
                <div>
                  <label className="label">Song Title (shown on last slide)</label>
                  <input
                    type="text"
                    className="input"
                    placeholder="e.g. ចំរៀងសរសើរព្រះ"
                    value={songTitle}
                    onChange={e => setSongTitle(e.target.value)}
                  />
                </div>

                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <label className="label mb-0">Khmer Lyrics</label>
                    {lyricsText && (
                      <button
                        onClick={() => { setLyricsText(''); setFileName(''); setStep(0) }}
                        className="text-[10px] text-slate-400 hover:text-red-500 flex items-center gap-0.5 transition"
                      >
                        <X size={10} /> Clear
                      </button>
                    )}
                  </div>
                  <textarea
                    rows={16}
                    className="input resize-none text-[13px] leading-relaxed font-[KhmerOSBattambang,sans-serif]"
                    placeholder={"Paste or type Khmer lyrics here…\n\nEach non-empty line becomes one lyric line.\nBlank lines separate verse groups.\n\nExample:\nដំណឹងល្អ\nព្រះជាម្ចាស់ស្រឡាញ់យើង"}
                    value={lyricsText}
                    onChange={e => setLyricsText(e.target.value)}
                  />
                  <p className="text-[10px] text-slate-400 mt-1.5">
                    {lyricsText.split('\n').filter(l => l.trim()).length} lines
                    &nbsp;·&nbsp; {slides.length} slide{slides.length !== 1 ? 's' : ''} will be generated
                  </p>
                </div>

                {/* Quick-group hint */}
                <div className="rounded-xl bg-slate-50 border border-slate-100 px-3 py-2.5">
                  <p className="text-[11px] text-slate-500">
                    <span className="font-semibold text-slate-600">Tip:</span> Use the
                    <span className="font-semibold text-brand-600"> "Lines per Slide"</span> setting
                    on the right to control how many lyric lines appear on each slide.
                    Blank lines in the textarea create natural verse breaks.
                  </p>
                </div>
              </>
            )}
          </div>
        </div>

        {/* Generate button */}
        <button
          onClick={handleGenerate}
          disabled={!canGenerate}
          className="btn-primary w-full !py-3 !text-base"
        >
          {generating
            ? <><Loader2 size={16} className="animate-spin" />Building PPTX…</>
            : <><Download size={16} />Generate &amp; Download PPTX</>}
        </button>

        {/* Stats bar */}
        {slides.length > 0 && (
          <div className="card px-4 py-3 flex items-center gap-3">
            <Layers size={14} className="text-brand-500 shrink-0" />
            <div className="flex-1 min-w-0">
              <p className="text-xs font-semibold text-slate-700">
                {slides.length} slide{slides.length !== 1 ? 's' : ''} ready
              </p>
              <p className="text-[10px] text-slate-400 truncate">
                {settings.lines_per_slide} line{settings.lines_per_slide > 1 ? 's' : ''} per slide
                &nbsp;·&nbsp; {settings.font_family}
                &nbsp;·&nbsp; {settings.font_size}pt
              </p>
            </div>
            <button
              onClick={handleGenerate}
              disabled={!canGenerate}
              className="btn-primary !py-1.5 !px-3 !text-xs"
            >
              <Download size={12} />
              Export
            </button>
          </div>
        )}
      </div>

      {/* ══ CENTRE — Preview ═══════════════════════════════════════ */}
      <div className="card p-6 lg:sticky lg:top-20">
        <SongPreviewPanel
          slides={slides}
          settings={settings}
          current={currentSlide}
          onNavigate={setCurrentSlide}
          isLoading={previewLoad}
        />
      </div>

      {/* ══ RIGHT — Settings ═══════════════════════════════════════ */}
      <div className="card lg:sticky lg:top-20">
        <div className="flex items-center gap-2 px-4 pt-4 pb-3 border-b border-slate-100">
          <Sliders size={14} className="text-brand-500" />
          <span className="font-semibold text-sm text-slate-800">Slide Settings</span>
        </div>
        <div className="px-4 py-4 max-h-[80vh] overflow-y-auto">
          <SongSettingsPanel settings={settings} onChange={setSettings} />
        </div>
      </div>

      {toast && (
        <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />
      )}
    </div>
  )
}
