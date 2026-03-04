import { useState, useCallback, useRef, useEffect } from 'react'
import {
  Layers,
  Upload,
  X,
  Download,
  Loader2,
  CheckCircle2,
  AlertCircle,
  GripVertical,
  ChevronLeft,
  ChevronRight,
  FileSliders,
  Trash2,
  Plus,
  Eye,
  Merge,
  ArrowUp,
  ArrowDown,
  Image as ImageIcon,
} from 'lucide-react'
import { mergePptxFiles, previewSlides } from './api.js'

// ── Toast ────────────────────────────────────────────────────────
function Toast({ message, type = 'success', onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 5000)
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

// ── Slide thumbnail (real PNG from backend) ───────────────────────
function SlideThumbnail({ thumb, label, active, onClick, loading }) {
  return (
    <button
      onClick={onClick}
      className={`relative group flex-shrink-0 rounded-lg overflow-hidden border-2 transition-all focus:outline-none
        ${active ? 'border-brand-500 shadow-lg scale-[1.03]' : 'border-slate-200 hover:border-brand-300 hover:shadow'}`}
      style={{ width: 120, aspectRatio: '16/9' }}
    >
      {loading ? (
        <div className="w-full h-full bg-slate-200 animate-pulse flex items-center justify-center">
          <Loader2 size={14} className="text-slate-400 animate-spin" />
        </div>
      ) : thumb ? (
        <img
          src={`data:image/png;base64,${thumb}`}
          alt={label}
          className="w-full h-full object-cover"
          draggable={false}
        />
      ) : (
        <div className="w-full h-full bg-slate-800 flex items-center justify-center">
          <ImageIcon size={16} className="text-slate-500" />
        </div>
      )}
      <div className="absolute bottom-0 left-0 right-0 bg-black/50 text-white/80 text-[9px] font-semibold px-1.5 py-0.5 truncate text-center">
        {label}
      </div>
      {active && (
        <div className="absolute inset-0 ring-2 ring-brand-500 ring-inset rounded-lg pointer-events-none" />
      )}
    </button>
  )
}

// ── Large slide viewer ────────────────────────────────────────────
function SlideViewer({ thumb, label, loading }) {
  return (
    <div
      className="relative w-full h-full rounded-xl overflow-hidden border border-slate-200 shadow select-none bg-white"
    >
      {loading ? (
        <div className="absolute inset-0 flex flex-col items-center justify-center gap-3 bg-slate-50">
          <Loader2 size={24} className="text-slate-400 animate-spin" />
          <p className="text-slate-400 text-xs">Rendering slide…</p>
        </div>
      ) : thumb ? (
        <img
          src={`data:image/png;base64,${thumb}`}
          alt={label}
          className="absolute inset-0 w-full h-full"
          style={{ objectFit: 'fill' }}
          draggable={false}
        />
      ) : (
        <div className="absolute inset-0 flex flex-col items-center justify-center gap-2 text-slate-400 bg-slate-50">
          <ImageIcon size={32} />
          <p className="text-xs">No preview available</p>
        </div>
      )}
      {label && !loading && thumb && (
        <div className="absolute top-2 left-2 bg-black/40 text-white/90 text-[10px] font-semibold px-2 py-0.5 rounded-full backdrop-blur-sm pointer-events-none">
          {label}
        </div>
      )}
    </div>
  )
}

// ── File card in the left panel ───────────────────────────────────
function FileCard({ entry, index, total, onRemove, onMoveUp, onMoveDown, isActive, onClick }) {
  const sizeKB = (entry.file.size / 1024).toFixed(1)
  const firstThumb = entry.thumbnails?.[0] ?? null

  return (
    <div
      onClick={onClick}
      className={`flex items-center gap-3 rounded-xl px-3 py-2.5 shadow-sm cursor-pointer transition-all border
        ${isActive
          ? 'bg-brand-50 border-brand-300'
          : 'bg-white border-slate-200 hover:border-brand-200'}`}
    >
      <GripVertical size={14} className="text-slate-300 shrink-0" />

      {/* mini thumbnail */}
      <div className="w-12 rounded overflow-hidden border border-slate-200 bg-slate-800 shrink-0" style={{ aspectRatio: '16/9' }}>
        {entry.loading ? (
          <div className="w-full h-full flex items-center justify-center">
            <Loader2 size={10} className="text-slate-400 animate-spin" />
          </div>
        ) : firstThumb ? (
          <img src={`data:image/png;base64,${firstThumb}`} className="w-full h-full object-cover" alt="" draggable={false} />
        ) : (
          <div className="w-full h-full flex items-center justify-center">
            <FileSliders size={10} className="text-slate-500" />
          </div>
        )}
      </div>

      <div className="flex-1 min-w-0">
        <p className="text-xs font-semibold text-slate-800 truncate">{entry.file.name}</p>
        <p className="text-[10px] text-slate-400 mt-0.5">
          {sizeKB} KB
          {entry.slideCount != null && (
            <span className="ml-1.5 text-brand-500 font-semibold">{entry.slideCount} slide{entry.slideCount !== 1 ? 's' : ''}</span>
          )}
        </p>
      </div>

      <span className="text-[10px] font-bold text-slate-400 tabular-nums w-4 text-center shrink-0">{index + 1}</span>

      <div className="flex items-center gap-0.5" onClick={e => e.stopPropagation()}>
        <div className="flex flex-col gap-0.5">
          <button onClick={onMoveUp} disabled={index === 0}
            className="p-0.5 rounded text-slate-300 hover:text-brand-500 disabled:opacity-20 transition">
            <ArrowUp size={10} />
          </button>
          <button onClick={onMoveDown} disabled={index === total - 1}
            className="p-0.5 rounded text-slate-300 hover:text-brand-500 disabled:opacity-20 transition">
            <ArrowDown size={10} />
          </button>
        </div>
        <button onClick={onRemove}
          className="p-1 rounded text-slate-300 hover:text-red-500 hover:bg-red-50 transition ml-1">
          <Trash2 size={11} />
        </button>
      </div>
    </div>
  )
}

// ── Horizontally scrollable strip with arrow buttons ─────────────
function ScrollStrip({ children, innerRef, className = '' }) {
  const localRef = useRef(null)
  const ref = innerRef ?? localRef
  const [canLeft, setCanLeft]   = useState(false)
  const [canRight, setCanRight] = useState(false)

  const update = () => {
    const el = ref.current
    if (!el) return
    setCanLeft(el.scrollLeft > 4)
    setCanRight(el.scrollLeft < el.scrollWidth - el.clientWidth - 4)
  }

  useEffect(() => {
    const el = ref.current
    if (!el) return
    update()
    el.addEventListener('scroll', update, { passive: true })
    const ro = new ResizeObserver(update)
    ro.observe(el)
    return () => { el.removeEventListener('scroll', update); ro.disconnect() }
  }, [children])

  const scroll = (dir) => {
    const el = ref.current
    if (!el) return
    el.scrollBy({ left: dir * Math.max(200, el.clientWidth * 0.6), behavior: 'smooth' })
  }

  return (
    <div className={`relative ${className}`}>
      {/* Left fade + arrow */}
      <div className={`absolute left-0 top-0 bottom-0 z-10 flex items-center transition-opacity duration-150 ${canLeft ? 'opacity-100' : 'opacity-0 pointer-events-none'}`}>
        <div className="absolute inset-y-0 left-0 w-10 bg-gradient-to-r from-white to-transparent pointer-events-none" />
        <button
          onClick={() => scroll(-1)}
          className="relative z-10 w-6 h-6 rounded-full bg-white border border-slate-200 shadow-sm flex items-center justify-center text-slate-500 hover:text-brand-500 hover:border-brand-300 transition -ml-0.5"
        >
          <ChevronLeft size={13} />
        </button>
      </div>

      {/* Right fade + arrow */}
      <div className={`absolute right-0 top-0 bottom-0 z-10 flex items-center transition-opacity duration-150 ${canRight ? 'opacity-100' : 'opacity-0 pointer-events-none'}`}>
        <div className="absolute inset-y-0 right-0 w-10 bg-gradient-to-l from-white to-transparent pointer-events-none" />
        <button
          onClick={() => scroll(1)}
          className="relative z-10 w-6 h-6 rounded-full bg-white border border-slate-200 shadow-sm flex items-center justify-center text-slate-500 hover:text-brand-500 hover:border-brand-300 transition -mr-0.5"
        >
          <ChevronRight size={13} />
        </button>
      </div>

      {/* Scrollable row — hide native scrollbar, use arrow buttons instead */}
      <div
        ref={ref}
        className="flex gap-2 overflow-x-auto py-1 px-1"
        style={{ scrollbarWidth: 'none', msOverflowStyle: 'none' }}
      >
        <style>{`.no-scrollbar::-webkit-scrollbar{display:none}`}</style>
        {children}
      </div>
    </div>
  )
}

// ── Drop Zone ─────────────────────────────────────────────────────
function DropZone({ onFiles, disabled }) {
  const inputRef = useRef(null)
  const [drag, setDrag] = useState(false)

  const handleDrop = useCallback(e => {
    e.preventDefault()
    setDrag(false)
    if (disabled) return
    const files = Array.from(e.dataTransfer?.files ?? []).filter(f =>
      f.name.toLowerCase().endsWith('.pptx'))
    if (files.length) onFiles(files)
  }, [onFiles, disabled])

  return (
    <div
      onDrop={handleDrop}
      onDragOver={e => { e.preventDefault(); setDrag(true) }}
      onDragLeave={() => setDrag(false)}
      onClick={() => !disabled && inputRef.current?.click()}
      className={`relative flex flex-col items-center justify-center gap-2.5 rounded-2xl border-2 border-dashed cursor-pointer transition-all py-6 px-4
        ${drag ? 'border-brand-400 bg-brand-50 scale-[1.01]' : 'border-slate-200 bg-slate-50 hover:border-brand-300 hover:bg-brand-50/40'}
        ${disabled ? 'opacity-60 cursor-not-allowed' : ''}`}
    >
      <input ref={inputRef} type="file" accept=".pptx" multiple className="hidden"
        onChange={e => {
          const files = Array.from(e.target.files ?? []).filter(f => f.name.toLowerCase().endsWith('.pptx'))
          if (files.length) onFiles(files)
          e.target.value = ''
        }} />
      <div className="w-10 h-10 rounded-xl bg-brand-100 flex items-center justify-center">
        <Upload size={18} className="text-brand-500" />
      </div>
      <div className="text-center">
        <p className="text-sm font-semibold text-slate-700">
          {drag ? 'Drop PPTX files here' : 'Click or drag & drop PPTX files'}
        </p>
        <p className="text-xs text-slate-400 mt-0.5">Slide previews are rendered automatically</p>
      </div>
    </div>
  )
}

// ── Main Page ─────────────────────────────────────────────────────
export default function MergeSlidesPage() {
  // entries: { id, file, slideCount, thumbnails: string[]|null, loading: bool }
  const [files, setFiles]               = useState([])
  const [activeFileIdx, setActiveFileIdx]   = useState(0)
  const [activeSlideIdx, setActiveSlideIdx] = useState(0)
  const [merging, setMerging]           = useState(false)
  const [outputName, setOutputName]     = useState('merged_slides')
  const [toast, setToast]               = useState(null)
  const stripRef = useRef(null)

  const totalSlides = files.reduce((s, f) => s + (f.slideCount ?? 0), 0)

  // Flat list of all slides across all files
  const allSlides = files.flatMap((entry, fi) =>
    (entry.thumbnails ?? Array.from({ length: entry.slideCount ?? 1 })).map((thumb, si) => ({
      fileIdx: fi,
      slideIdx: si,
      thumb: thumb ?? null,
      loading: entry.loading,
      label: `${entry.file.name.replace(/\.pptx$/i, '')} · ${si + 1}`,
    }))
  )

  const globalIdx = allSlides.findIndex(
    s => s.fileIdx === activeFileIdx && s.slideIdx === activeSlideIdx
  )
  const currentSlide = allSlides[Math.max(0, globalIdx)] ?? null

  const selectFile = (fi) => {
    setActiveFileIdx(fi)
    setActiveSlideIdx(0)
    setTimeout(() => {
      if (stripRef.current) {
        stripRef.current.scrollTo({ left: 0, behavior: 'smooth' })
      }
    }, 50)
  }

  const goTo = (flatIdx) => {
    const slide = allSlides[flatIdx]
    if (!slide) return
    setActiveFileIdx(slide.fileIdx)
    setActiveSlideIdx(slide.slideIdx)
    setTimeout(() => {
      if (!stripRef.current) return
      // find the nth visible button in the strip (only current file's slides shown)
      const btns = stripRef.current.querySelectorAll('button')
      const visibleIdx = allSlides
        .slice(0, flatIdx + 1)
        .filter(s => s.fileIdx === slide.fileIdx).length - 1
      btns[visibleIdx]?.scrollIntoView({ behavior: 'smooth', inline: 'nearest', block: 'nearest' })
    }, 30)
  }

  // ── Add files + auto-fetch thumbnails ────────────────────────
  const addFiles = useCallback(newFiles => {
    const entries = newFiles.map(f => ({
      id: crypto.randomUUID(),
      file: f,
      slideCount: null,
      thumbnails: null,
      loading: true,
    }))

    setFiles(prev => {
      const existing = new Set(prev.map(e => `${e.file.name}-${e.file.size}`))
      return [...prev, ...entries.filter(e => !existing.has(`${e.file.name}-${e.file.size}`))]
    })

    entries.forEach(async entry => {
      try {
        const data = await previewSlides(entry.file)
        setFiles(prev => prev.map(e =>
          e.id === entry.id
            ? { ...e, loading: false, slideCount: data.slide_count, thumbnails: data.thumbnails }
            : e
        ))
      } catch {
        try {
          const fd = new FormData()
          fd.append('file', entry.file)
          const res = await fetch('/api/merge/count-slides', { method: 'POST', body: fd })
          if (res.ok) {
            const { slide_count } = await res.json()
            setFiles(prev => prev.map(e =>
              e.id === entry.id
                ? { ...e, loading: false, slideCount: slide_count, thumbnails: Array(slide_count).fill(null) }
                : e
            ))
          } else {
            setFiles(prev => prev.map(e => e.id === entry.id ? { ...e, loading: false } : e))
          }
        } catch {
          setFiles(prev => prev.map(e => e.id === entry.id ? { ...e, loading: false } : e))
        }
      }
    })
  }, [])

  const removeFile = idx => {
    setFiles(prev => prev.filter((_, i) => i !== idx))
    setActiveFileIdx(i => Math.max(0, Math.min(i, files.length - 2)))
    setActiveSlideIdx(0)
  }

  const moveFile = (idx, dir) => {
    setFiles(prev => {
      const arr = [...prev]
      const target = idx + dir
      if (target < 0 || target >= arr.length) return arr
      ;[arr[idx], arr[target]] = [arr[target], arr[idx]]
      return arr
    })
  }

  // ── Merge & download ─────────────────────────────────────────
  const handleMerge = async () => {
    if (files.length < 2) {
      setToast({ message: 'Add at least 2 PPTX files to merge.', type: 'error' })
      return
    }
    setMerging(true)
    try {
      const blob = await mergePptxFiles(files.map(e => e.file), outputName.trim() || 'merged_slides')
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `${outputName.trim() || 'merged_slides'}.pptx`
      a.click()
      URL.revokeObjectURL(url)
      setToast({ message: `Merged ${totalSlides} slides from ${files.length} files!`, type: 'success' })
    } catch (err) {
      setToast({ message: `Merge failed — ${err.message}`, type: 'error' })
    } finally {
      setMerging(false)
    }
  }

  return (
    <div className="flex-1 overflow-hidden flex flex-col min-h-0">
      <div className="flex-1 min-h-0 max-w-[1500px] mx-auto w-full px-4 py-4
                      flex flex-col lg:flex-row gap-4 overflow-hidden">

        {/* ══════════════════ LEFT PANEL — scrollable ══════════════════ */}
        <div className="w-full lg:w-[300px] xl:w-[320px] shrink-0 flex flex-col gap-3 overflow-y-auto overflow-x-hidden min-h-0 pb-2"
             style={{ scrollbarWidth: 'thin' }}>

          <div className="card p-4 space-y-3 shrink-0">
            <div className="flex items-center gap-2">
              <Merge size={14} className="text-brand-500" />
              <h2 className="font-semibold text-sm text-slate-800">Attach PPTX Files</h2>
            </div>
            <DropZone onFiles={addFiles} disabled={merging} />
          </div>

          {files.length > 0 && (
            <div className="card p-4 space-y-2 shrink-0">
              <div className="flex items-center justify-between mb-1">
                <h3 className="text-[10px] font-bold uppercase tracking-widest text-slate-400">Files — merge order</h3>
                <span className="text-[11px] font-semibold text-brand-600 bg-brand-50 border border-brand-100 px-2 py-0.5 rounded-full">
                  {files.length} file{files.length !== 1 ? 's' : ''}{totalSlides > 0 && ` · ${totalSlides} slides`}
                </span>
              </div>
              <div className="space-y-1.5">
                {files.map((entry, idx) => (
                  <FileCard
                    key={entry.id}
                    entry={entry}
                    index={idx}
                    total={files.length}
                    isActive={idx === activeFileIdx}
                    onClick={() => selectFile(idx)}
                    onRemove={() => removeFile(idx)}
                    onMoveUp={() => moveFile(idx, -1)}
                    onMoveDown={() => moveFile(idx, 1)}
                  />
                ))}
              </div>
              <div className="flex items-center justify-between pt-1">
                <label className="flex items-center gap-1.5 text-xs text-brand-600 font-semibold cursor-pointer hover:text-brand-700 transition">
                  <input type="file" accept=".pptx" multiple className="hidden"
                    onChange={e => {
                      addFiles(Array.from(e.target.files ?? []).filter(f => f.name.toLowerCase().endsWith('.pptx')))
                      e.target.value = ''
                    }} />
                  <Plus size={12} /> Add more
                </label>
                <button onClick={() => { setFiles([]); setActiveFileIdx(0); setActiveSlideIdx(0) }}
                  className="text-[11px] text-slate-400 hover:text-red-500 flex items-center gap-1 transition">
                  <Trash2 size={11} /> Clear all
                </button>
              </div>
            </div>
          )}

          <div className="card p-4 space-y-3 shrink-0">
            <h3 className="text-[10px] font-bold uppercase tracking-widest text-slate-400">Output</h3>
            <div>
              <label className="label">File Name</label>
              <div className="flex gap-2 items-center">
                <input type="text" className="input flex-1" placeholder="merged_slides"
                  value={outputName} onChange={e => setOutputName(e.target.value)} />
                <span className="text-xs text-slate-400 font-mono shrink-0">.pptx</span>
              </div>
            </div>
            <button onClick={handleMerge} disabled={files.length < 2 || merging} className="btn-primary w-full">
              {merging
                ? <><Loader2 size={14} className="animate-spin" />Merging…</>
                : <><Download size={14} />Merge &amp; Download PPTX</>}
            </button>
            {files.length < 2 && (
              <p className="text-[11px] text-slate-400 text-center">Attach at least 2 PPTX files to merge.</p>
            )}
          </div>
        </div>

        {/* ══════════════════ RIGHT PANEL — flex column, no overflow ══════════════════ */}
        <div className="flex-1 min-w-0 min-h-0 flex flex-col gap-3 overflow-hidden">
          {files.length === 0 ? (
            <div className="card flex-1 p-6 flex items-center justify-center">
              <div className="w-full rounded-2xl border-2 border-dashed border-slate-200 flex flex-col items-center justify-center gap-4 bg-slate-50 h-full">
                <div className="w-16 h-16 rounded-2xl bg-slate-100 flex items-center justify-center">
                  <Layers size={28} className="text-slate-300" />
                </div>
                <div className="text-center">
                  <p className="text-sm font-semibold text-slate-500">No files attached yet</p>
                  <p className="text-xs text-slate-400 mt-1">Upload PPTX files — each slide will be previewed here</p>
                </div>
              </div>
            </div>
          ) : (
            <>
              {/* ── Slide viewer card — takes all remaining vertical space ── */}
              <div className="card p-3 flex flex-col gap-2 min-h-0 flex-1">
                {/* Header */}
                <div className="flex items-center gap-2 shrink-0">
                  <Eye size={14} className="text-brand-500" />
                  <span className="font-semibold text-sm text-slate-700">Slide Preview</span>
                  {allSlides.length > 0 && (
                    <span className="ml-auto text-xs font-semibold tabular-nums px-2.5 py-1 rounded-lg bg-brand-50 text-brand-600 border border-brand-100">
                      {Math.max(1, globalIdx + 1)} / {allSlides.length}
                    </span>
                  )}
                </div>

                {/* Viewer — fills remaining height, maintains 16:9 with max-width */}
                <div className="flex-1 min-h-0 flex items-center justify-center overflow-hidden">
                  <div className="w-full h-full flex items-center justify-center">
                    {/* Inner box: 16:9 constrained to both width AND height */}
                    <div className="relative w-full" style={{ aspectRatio: '16/9', maxHeight: '90%', maxWidth: '80%' }}>
                      <SlideViewer
                        thumb={currentSlide?.thumb ?? null}
                        label={currentSlide?.label ?? ''}
                        loading={currentSlide?.loading ?? false}
                      />
                    </div>
                  </div>
                </div>

                {/* Prev / Next */}
                {allSlides.length > 1 && (
                  <div className="flex items-center justify-center gap-3 shrink-0">
                    <button onClick={() => goTo(Math.max(0, globalIdx - 1))} disabled={globalIdx <= 0}
                      className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white border border-slate-200 text-slate-600 hover:border-brand-300 hover:text-brand-600 disabled:opacity-30 transition">
                      <ChevronLeft size={13} /> Prev
                    </button>
                    <span className="text-xs text-slate-400 tabular-nums">{Math.max(1, globalIdx + 1)} / {allSlides.length}</span>
                    <button onClick={() => goTo(Math.min(allSlides.length - 1, globalIdx + 1))} disabled={globalIdx >= allSlides.length - 1}
                      className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white border border-slate-200 text-slate-600 hover:border-brand-300 hover:text-brand-600 disabled:opacity-30 transition">
                      Next <ChevronRight size={13} />
                    </button>
                  </div>
                )}
              </div>

              {/* ── All-slides strip card — fixed height at bottom ── */}
              <div className="card p-3 shrink-0 space-y-2">
                {/* File tabs */}
                <ScrollStrip>
                  <span className="text-[10px] font-bold uppercase tracking-widest text-slate-400 self-center shrink-0 mr-1 whitespace-nowrap">All slides</span>
                  {files.map((entry, fi) => (
                    <button key={entry.id} onClick={() => selectFile(fi)}
                      className={`flex-shrink-0 flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-semibold border transition-all
                        ${fi === activeFileIdx
                          ? 'bg-brand-500 text-white border-brand-500 shadow'
                          : 'bg-white text-slate-600 border-slate-200 hover:border-brand-300 hover:text-brand-600'}`}>
                      <span className="opacity-70">#{fi + 1}</span>
                      <span className="max-w-[120px] truncate">{entry.file.name.replace(/\.pptx$/i, '')}</span>
                      {entry.slideCount != null && (
                        <span className={`text-[9px] rounded px-1 ${fi === activeFileIdx ? 'bg-white/20' : 'bg-slate-100 text-slate-400'}`}>
                          {entry.slideCount}s
                        </span>
                      )}
                    </button>
                  ))}
                </ScrollStrip>

                {/* Per-file thumbnail strip */}
                <ScrollStrip innerRef={stripRef}>
                  {allSlides.map((slide, flatIdx) => {
                    if (slide.fileIdx !== activeFileIdx) return null
                    return (
                      <SlideThumbnail
                        key={`${slide.fileIdx}-${slide.slideIdx}`}
                        thumb={slide.thumb}
                        label={`${slide.slideIdx + 1}`}
                        loading={slide.loading}
                        active={slide.fileIdx === activeFileIdx && slide.slideIdx === activeSlideIdx}
                        onClick={() => goTo(flatIdx)}
                      />
                    )
                  })}
                </ScrollStrip>

                {/* Merge-order strip — only when 2+ files */}
                {files.length >= 2 && (
                  <div className="border-t border-slate-100 pt-2">
                    <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-1.5">
                      Merge order · {totalSlides} slides total
                    </p>
                    <ScrollStrip>
                      {allSlides.map((slide, flatIdx) => {
                        const isActive = slide.fileIdx === activeFileIdx && slide.slideIdx === activeSlideIdx
                        const colors = ['border-blue-300','border-purple-300','border-green-300','border-orange-300','border-pink-300','border-teal-300']
                        const accent = colors[slide.fileIdx % colors.length]
                        return (
                          <button
                            key={`all-${slide.fileIdx}-${slide.slideIdx}`}
                            onClick={() => goTo(flatIdx)}
                            title={slide.label}
                            className={`relative flex-shrink-0 rounded overflow-hidden border-2 transition-all
                              ${isActive ? 'border-brand-500 shadow scale-105' : `${accent} hover:border-brand-400`}`}
                            style={{ width: 60, aspectRatio: '16/9' }}
                          >
                            {slide.loading ? (
                              <div className="w-full h-full bg-slate-200 animate-pulse" />
                            ) : slide.thumb ? (
                              <img src={`data:image/png;base64,${slide.thumb}`} className="w-full h-full" style={{ objectFit: 'fill' }} alt="" draggable={false} />
                            ) : (
                              <div className="w-full h-full bg-slate-100" />
                            )}
                            <div className="absolute bottom-0 right-0 bg-black/50 text-white text-[8px] px-0.5 font-mono leading-tight">
                              {flatIdx + 1}
                            </div>
                          </button>
                        )
                      })}
                      <div className="flex-shrink-0 flex flex-col items-center justify-center gap-0.5 bg-green-50 border-2 border-green-300 rounded px-2 text-green-700"
                        style={{ minWidth: 64, aspectRatio: '16/9' }}>
                        <CheckCircle2 size={10} />
                        <span className="text-[8px] font-bold text-center leading-tight truncate max-w-[56px]">{outputName || 'merged'}.pptx</span>
                        <span className="text-[8px] opacity-70">{totalSlides}s</span>
                      </div>
                    </ScrollStrip>
                  </div>
                )}
              </div>
            </>
          )}
        </div>
      </div>

      {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
    </div>
  )
}
