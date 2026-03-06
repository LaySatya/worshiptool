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
  RefreshCw,
  RotateCcw,
  Sparkles,
} from 'lucide-react'
import { mergePptxFiles, previewSlides } from './api.js'

const SESSION_KEY = 'mergeSlides_session'

// ── Persist session to sessionStorage (thumbnails + metadata only) ─
function saveSession(files, outputName, activeFileIdx) {
  try {
    const serialisable = files.map(e => ({
      id:         e.id,
      name:       e.file.name,
      size:       e.file.size,
      slideCount: e.slideCount,
      thumbnails: e.thumbnails,   // base64 strings — safe to store
      loading:    false,
      needsReattach: true,        // File binary is gone on refresh
    }))
    sessionStorage.setItem(SESSION_KEY, JSON.stringify({
      files: serialisable,
      outputName,
      activeFileIdx,
    }))
  } catch {
    // sessionStorage full (large thumbnails) — silently skip
  }
}

function loadSession() {
  try {
    const raw = sessionStorage.getItem(SESSION_KEY)
    if (!raw) return null
    return JSON.parse(raw)
  } catch {
    return null
  }
}

function clearSession() {
  sessionStorage.removeItem(SESSION_KEY)
}

// ── Toast ──────────────────────────────────────────────────────────────────
function Toast({ message, type = 'success', onClose }) {
  useEffect(() => { const t = setTimeout(onClose, 5000); return () => clearTimeout(t) }, [onClose])
  return (
    <div className={`fixed bottom-6 right-6 z-50 flex items-center gap-3 px-4 py-3 rounded-2xl
        shadow-[0_8px_32px_rgba(0,0,0,.14)] border text-sm font-medium animate-fade-in max-w-sm
        ${type === 'success' ? 'bg-white border-green-100 text-green-700' : 'bg-white border-red-100 text-red-600'}`}>
      {type === 'success'
        ? <CheckCircle2 size={16} className="shrink-0 text-green-500" />
        : <AlertCircle  size={16} className="shrink-0 text-red-500" />}
      <span className="flex-1">{message}</span>
      <button onClick={onClose} className="ml-1 text-slate-300 hover:text-slate-500 transition">
        <X size={14} />
      </button>
    </div>
  )
}

// ── Slide thumbnail ────────────────────────────────────────────────────────
function SlideThumbnail({ thumb, index, active, onClick, loading }) {
  return (
    <button
      onClick={onClick}
      className={`relative group flex-shrink-0 rounded-lg overflow-hidden transition-all duration-150 focus:outline-none
        ${active
          ? 'ring-2 ring-brand-500 ring-offset-1 shadow-lg scale-[1.04]'
          : 'ring-1 ring-slate-200 hover:ring-brand-300 hover:shadow-md'}`}
      style={{ width: 108, aspectRatio: '16/9' }}
    >
      {loading ? (
        <div className="w-full h-full bg-slate-100 animate-pulse flex items-center justify-center">
          <Loader2 size={12} className="text-slate-300 animate-spin" />
        </div>
      ) : thumb ? (
        <img src={`data:image/png;base64,${thumb}`} alt={`Slide ${index + 1}`}
          className="w-full h-full" style={{ objectFit: 'fill' }} draggable={false} />
      ) : (
        <div className="w-full h-full bg-slate-100 flex items-center justify-center">
          <ImageIcon size={14} className="text-slate-300" />
        </div>
      )}
      <div className="absolute bottom-1 right-1 bg-black/50 text-white text-[8px] font-bold
          px-1.5 py-0.5 rounded-full leading-none backdrop-blur-sm">
        {index + 1}
      </div>
      {active && <div className="absolute inset-0 bg-brand-500/5 pointer-events-none" />}
    </button>
  )
}

// ── Main slide viewer ──────────────────────────────────────────────────────
function SlideViewer({ thumb, label, loading }) {
  return (
    <div className="relative w-full h-full rounded-2xl overflow-hidden bg-slate-900 shadow-inner select-none">
      <div className="absolute inset-0 opacity-20"
        style={{ backgroundImage: 'radial-gradient(circle, #94a3b8 1px, transparent 1px)', backgroundSize: '24px 24px' }} />

      {loading ? (
        <div className="absolute inset-0 flex flex-col items-center justify-center gap-3">
          <div className="w-10 h-10 rounded-full bg-white/10 flex items-center justify-center">
            <Loader2 size={20} className="text-white/60 animate-spin" />
          </div>
          <p className="text-white/40 text-xs font-medium">Rendering preview…</p>
        </div>
      ) : thumb ? (
        <img src={`data:image/png;base64,${thumb}`} alt={label}
          className="absolute inset-0 w-full h-full rounded-2xl"
          style={{ objectFit: 'fill' }} draggable={false} />
      ) : (
        <div className="absolute inset-0 flex flex-col items-center justify-center gap-3">
          <div className="w-12 h-12 rounded-2xl bg-white/10 flex items-center justify-center">
            <ImageIcon size={22} className="text-white/30" />
          </div>
          <p className="text-white/30 text-xs font-medium">No preview available</p>
        </div>
      )}

      {label && !loading && thumb && (
        <div className="absolute top-3 left-3 bg-black/40 text-white/80 text-[10px] font-semibold
            px-2.5 py-1 rounded-full backdrop-blur-md pointer-events-none border border-white/10">
          {label}
        </div>
      )}
    </div>
  )
}

// ── File card ──────────────────────────────────────────────────────────────
function FileCard({ entry, index, total, onRemove, onMoveUp, onMoveDown, isActive, onClick }) {
  const sizeKB     = Math.round(entry.file.size / 1024)
  const firstThumb = entry.thumbnails?.[0] ?? null

  return (
    <div
      onClick={onClick}
      className={`group flex items-center gap-2.5 rounded-xl px-2.5 py-2 cursor-pointer
          transition-all duration-150 border
          ${entry.needsReattach
            ? 'bg-amber-50 border-amber-200 shadow-sm'
            : isActive
              ? 'bg-brand-50 border-brand-200 shadow-sm'
              : 'bg-white border-slate-100 hover:border-brand-200 hover:shadow-sm'}`}
    >
      <GripVertical size={13} className="text-slate-200 group-hover:text-slate-300 shrink-0 transition" />

      <div className="w-11 shrink-0 rounded-md overflow-hidden bg-slate-100 ring-1 ring-black/5"
        style={{ aspectRatio: '16/9' }}>
        {entry.loading ? (
          <div className="w-full h-full flex items-center justify-center">
            <Loader2 size={9} className="text-slate-300 animate-spin" />
          </div>
        ) : firstThumb ? (
          <img src={`data:image/png;base64,${firstThumb}`} className="w-full h-full"
            style={{ objectFit: 'fill' }} alt="" draggable={false} />
        ) : (
          <div className="w-full h-full flex items-center justify-center">
            <FileSliders size={9} className="text-slate-300" />
          </div>
        )}
      </div>

      <div className="flex-1 min-w-0">
        <p className="text-xs font-semibold text-slate-700 truncate leading-tight">{entry.file.name}</p>
        <p className="text-[10px] mt-0.5 leading-none">
          {entry.needsReattach ? (
            <span className="text-amber-500 font-semibold flex items-center gap-1">
              <RotateCcw size={8} /> Needs re-attach
            </span>
          ) : entry.loading ? (
            <span className="text-slate-400">Loading…</span>
          ) : (
            <span className="text-slate-400">
              {sizeKB > 1024 ? `${(sizeKB / 1024).toFixed(1)} MB` : `${sizeKB} KB`}
              {entry.slideCount != null && (
                <> · <span className="text-brand-500 font-semibold">{entry.slideCount} slides</span></>
              )}
            </span>
          )}
        </p>
      </div>

      <span className="text-[10px] font-bold text-slate-300 tabular-nums w-4 text-center shrink-0">{index + 1}</span>

      <div className="flex items-center gap-0.5 shrink-0" onClick={e => e.stopPropagation()}>
        <div className="flex flex-col gap-px">
          <button onClick={onMoveUp} disabled={index === 0}
            className="p-0.5 rounded text-slate-200 hover:text-brand-500 hover:bg-brand-50 disabled:opacity-20 transition">
            <ArrowUp size={10} />
          </button>
          <button onClick={onMoveDown} disabled={index === total - 1}
            className="p-0.5 rounded text-slate-200 hover:text-brand-500 hover:bg-brand-50 disabled:opacity-20 transition">
            <ArrowDown size={10} />
          </button>
        </div>
        <button onClick={onRemove}
          className="p-1 rounded text-slate-200 hover:text-red-400 hover:bg-red-50 transition">
          <Trash2 size={11} />
        </button>
      </div>
    </div>
  )
}

// ── Horizontal scroll strip ────────────────────────────────────────────────
function ScrollStrip({ children, innerRef, className = '' }) {
  const localRef             = useRef(null)
  const ref                  = innerRef ?? localRef
  const [canLeft, setLeft]   = useState(false)
  const [canRight, setRight] = useState(false)

  const sync = () => {
    const el = ref.current; if (!el) return
    setLeft(el.scrollLeft > 4)
    setRight(el.scrollLeft < el.scrollWidth - el.clientWidth - 4)
  }

  useEffect(() => {
    const el = ref.current; if (!el) return
    sync()
    el.addEventListener('scroll', sync, { passive: true })
    const ro = new ResizeObserver(sync); ro.observe(el)
    return () => { el.removeEventListener('scroll', sync); ro.disconnect() }
  }, [children])

  const scroll = dir => ref.current?.scrollBy({ left: dir * Math.max(180, (ref.current?.clientWidth ?? 300) * 0.6), behavior: 'smooth' })

  return (
    <div className={`relative ${className}`}>
      <div className={`absolute left-0 top-0 bottom-0 z-10 flex items-center pointer-events-none transition-opacity duration-150 ${canLeft ? 'opacity-100' : 'opacity-0'}`}>
        <div className="absolute inset-y-0 left-0 w-8 bg-gradient-to-r from-white to-transparent" />
        <button onClick={() => scroll(-1)} className="pointer-events-auto relative z-10 w-6 h-6 rounded-full bg-white border border-slate-200 shadow-sm flex items-center justify-center text-slate-400 hover:text-brand-500 hover:border-brand-300 transition -ml-1">
          <ChevronLeft size={12} />
        </button>
      </div>
      <div className={`absolute right-0 top-0 bottom-0 z-10 flex items-center pointer-events-none transition-opacity duration-150 ${canRight ? 'opacity-100' : 'opacity-0'}`}>
        <div className="absolute inset-y-0 right-0 w-8 bg-gradient-to-l from-white to-transparent" />
        <button onClick={() => scroll(1)} className="pointer-events-auto relative z-10 w-6 h-6 rounded-full bg-white border border-slate-200 shadow-sm flex items-center justify-center text-slate-400 hover:text-brand-500 hover:border-brand-300 transition -mr-1">
          <ChevronRight size={12} />
        </button>
      </div>
      <div ref={ref} className="flex gap-2 overflow-x-auto py-0.5 px-0.5 no-scrollbar">
        {children}
      </div>
    </div>
  )
}

// ── Drop zone ──────────────────────────────────────────────────────────────
function DropZone({ onFiles, disabled, compact = false }) {
  const inputRef        = useRef(null)
  const [drag, setDrag] = useState(false)

  const handleDrop = useCallback(e => {
    e.preventDefault(); setDrag(false)
    if (disabled) return
    const files = Array.from(e.dataTransfer?.files ?? []).filter(f => f.name.toLowerCase().endsWith('.pptx'))
    if (files.length) onFiles(files)
  }, [onFiles, disabled])

  return (
    <div
      onDrop={handleDrop}
      onDragOver={e => { e.preventDefault(); setDrag(true) }}
      onDragLeave={() => setDrag(false)}
      onClick={() => !disabled && inputRef.current?.click()}
      className={`relative flex flex-col items-center justify-center gap-2 rounded-xl border-2 border-dashed
          cursor-pointer transition-all duration-200 text-center select-none
          ${compact ? 'py-4 px-3' : 'py-8 px-5'}
          ${drag ? 'border-brand-400 bg-brand-50/80 scale-[1.01]' : 'border-slate-200 bg-slate-50/60 hover:border-brand-300 hover:bg-brand-50/40'}
          ${disabled ? 'opacity-50 pointer-events-none' : ''}`}
    >
      <input ref={inputRef} type="file" accept=".pptx" multiple className="hidden"
        onChange={e => {
          const files = Array.from(e.target.files ?? []).filter(f => f.name.toLowerCase().endsWith('.pptx'))
          if (files.length) onFiles(files)
          e.target.value = ''
        }} />
      <div className={`rounded-xl bg-brand-100 flex items-center justify-center transition-transform duration-200
          ${drag ? 'scale-110' : ''} ${compact ? 'w-8 h-8' : 'w-12 h-12'}`}>
        <Upload size={compact ? 15 : 20} className="text-brand-500" />
      </div>
      <div>
        <p className={`font-semibold text-slate-700 ${compact ? 'text-xs' : 'text-sm'}`}>
          {drag ? 'Drop your PPTX files here' : 'Click or drag & drop PPTX files'}
        </p>
        {!compact && <p className="text-xs text-slate-400 mt-1">Slide thumbnails are generated automatically</p>}
      </div>
    </div>
  )
}

// ── Section title ──────────────────────────────────────────────────────────
function SectionTitle({ icon: Icon, children, meta }) {
  return (
    <div className="flex items-center gap-2 mb-3">
      <div className="w-6 h-6 rounded-lg bg-brand-50 flex items-center justify-center shrink-0">
        <Icon size={12} className="text-brand-500" />
      </div>
      <span className="text-sm font-semibold text-slate-700 flex-1">{children}</span>
      {meta && <span className="text-[10px] font-semibold text-brand-600 bg-brand-50 border border-brand-100 px-2 py-0.5 rounded-full">{meta}</span>}
    </div>
  )
}

// ── Main page ──────────────────────────────────────────────────────────────
export default function MergeSlidesPage() {
  const [files, setFiles]                     = useState([])
  const [activeFileIdx, setActiveFileIdx]     = useState(0)
  const [activeSlideIdx, setActiveSlideIdx]   = useState(0)
  const [merging, setMerging]                 = useState(false)
  const [outputName, setOutputName]           = useState('merged_slides')
  const [toast, setToast]                     = useState(null)
  const [restoredSession, setRestoredSession] = useState(false)
  const stripRef = useRef(null)

  // ── Session restore ────────────────────────────────────────────
  useEffect(() => {
    const session = loadSession()
    if (!session?.files?.length) return
    setFiles(session.files.map(s => ({
      id: s.id, file: { name: s.name, size: s.size },
      slideCount: s.slideCount, thumbnails: s.thumbnails, loading: false, needsReattach: true,
    })))
    setOutputName(session.outputName ?? 'merged_slides')
    setActiveFileIdx(session.activeFileIdx ?? 0)
    setRestoredSession(true)
  }, [])

  useEffect(() => {
    const handler = e => { if (!files.length) return; e.preventDefault(); e.returnValue = '' }
    window.addEventListener('beforeunload', handler)
    return () => window.removeEventListener('beforeunload', handler)
  }, [files.length])

  useEffect(() => {
    if (!files.length) { clearSession(); return }
    saveSession(
      files.map(e => ({ ...e, file: e.file instanceof File ? e.file : { name: e.file.name, size: e.file.size } })),
      outputName, activeFileIdx,
    )
  }, [files, outputName, activeFileIdx])

  const totalSlides = files.reduce((s, f) => s + (f.slideCount ?? 0), 0)

  const allSlides = files.flatMap((entry, fi) =>
    (entry.thumbnails ?? Array.from({ length: entry.slideCount ?? 1 })).map((thumb, si) => ({
      fileIdx: fi, slideIdx: si, thumb: thumb ?? null, loading: entry.loading,
      label: `${entry.file.name.replace(/\.pptx$/i, '')} · ${si + 1}`,
    }))
  )

  const globalIdx    = allSlides.findIndex(s => s.fileIdx === activeFileIdx && s.slideIdx === activeSlideIdx)
  const currentSlide = allSlides[Math.max(0, globalIdx)] ?? null

  const selectFile = fi => {
    setActiveFileIdx(fi); setActiveSlideIdx(0)
    setTimeout(() => stripRef.current?.scrollTo({ left: 0, behavior: 'smooth' }), 50)
  }

  const goTo = flatIdx => {
    const slide = allSlides[flatIdx]; if (!slide) return
    setActiveFileIdx(slide.fileIdx); setActiveSlideIdx(slide.slideIdx)
    setTimeout(() => {
      const btns = stripRef.current?.querySelectorAll('button')
      const vi   = allSlides.slice(0, flatIdx + 1).filter(s => s.fileIdx === slide.fileIdx).length - 1
      btns?.[vi]?.scrollIntoView({ behavior: 'smooth', inline: 'nearest', block: 'nearest' })
    }, 30)
  }

  const addFiles = useCallback(newFiles => {
    const reattaching = [], brandNew = []
    newFiles.forEach(f => {
      const si = files.findIndex(e => e.needsReattach && e.file.name === f.name && e.file.size === f.size)
      si >= 0 ? reattaching.push({ file: f, stubIdx: si }) : brandNew.push(f)
    })
    if (reattaching.length) {
      setFiles(prev => {
        const arr = [...prev]
        reattaching.forEach(({ file, stubIdx }) => { arr[stubIdx] = { ...arr[stubIdx], file, needsReattach: false } })
        return arr
      })
      if (!brandNew.length) {
        setRestoredSession(false)
        setToast({ message: `${reattaching.length} file${reattaching.length > 1 ? 's' : ''} re-attached!`, type: 'success' })
      }
    }
    if (!brandNew.length) return

    const entries = brandNew.map(f => ({
      id: crypto.randomUUID(), file: f, slideCount: null, thumbnails: null, loading: true, needsReattach: false,
    }))
    setFiles(prev => {
      const seen = new Set(prev.map(e => `${e.file.name}-${e.file.size}`))
      return [...prev, ...entries.filter(e => !seen.has(`${e.file.name}-${e.file.size}`))]
    })

    entries.forEach(async entry => {
      try {
        const data = await previewSlides(entry.file)
        setFiles(prev => prev.map(e => e.id === entry.id
          ? { ...e, loading: false, slideCount: data.slide_count, thumbnails: data.thumbnails } : e))
      } catch {
        try {
          const fd = new FormData(); fd.append('file', entry.file)
          const res = await fetch('/api/merge/count-slides', { method: 'POST', body: fd })
          if (res.ok) {
            const { slide_count } = await res.json()
            setFiles(prev => prev.map(e => e.id === entry.id
              ? { ...e, loading: false, slideCount: slide_count, thumbnails: Array(slide_count).fill(null) } : e))
          } else setFiles(prev => prev.map(e => e.id === entry.id ? { ...e, loading: false } : e))
        } catch {
          setFiles(prev => prev.map(e => e.id === entry.id ? { ...e, loading: false } : e))
        }
      }
    })
  }, [files])

  const removeFile = idx => {
    setFiles(prev => { const n = prev.filter((_, i) => i !== idx); if (!n.length) clearSession(); return n })
    setActiveFileIdx(i => Math.max(0, Math.min(i, files.length - 2)))
    setActiveSlideIdx(0)
  }

  const clearAll = () => { setFiles([]); setActiveFileIdx(0); setActiveSlideIdx(0); setRestoredSession(false); clearSession() }

  const moveFile = (idx, dir) => {
    setFiles(prev => {
      const arr = [...prev], t = idx + dir
      if (t < 0 || t >= arr.length) return arr
      ;[arr[idx], arr[t]] = [arr[t], arr[idx]]
      return arr
    })
  }

  const needsReattach = files.some(e => e.needsReattach)

  const handleMerge = async () => {
    if (files.length < 2) { setToast({ message: 'Add at least 2 PPTX files to merge.', type: 'error' }); return }
    if (needsReattach)    { setToast({ message: 'Re-attach highlighted files before merging.', type: 'error' }); return }
    setMerging(true)
    try {
      const blob = await mergePptxFiles(files.map(e => e.file), outputName.trim() || 'merged_slides')
      const url  = URL.createObjectURL(blob)
      const a    = document.createElement('a'); a.href = url; a.download = `${outputName.trim() || 'merged_slides'}.pptx`; a.click()
      URL.revokeObjectURL(url)
      setToast({ message: `✓ Merged ${totalSlides} slides from ${files.length} files!`, type: 'success' })
    } catch (err) {
      setToast({ message: `Merge failed — ${err.message}`, type: 'error' })
    } finally { setMerging(false) }
  }

  return (
    <div className="flex-1 flex flex-col min-h-0 overflow-hidden bg-slate-50">

      {/* ── Session restore banner ─────────────────────────────────────── */}
      {restoredSession && (
        <div className="shrink-0 flex items-center gap-3 px-5 py-2.5 bg-amber-50 border-b border-amber-100 text-amber-800 text-xs">
          <RotateCcw size={13} className="shrink-0 text-amber-400" />
          <span className="flex-1 leading-relaxed">
            <strong>Session restored.</strong> Previews are cached — re-attach the same files <em>(highlighted)</em> to enable merging.
          </span>
          <label className="flex items-center gap-1.5 font-semibold cursor-pointer
              bg-white hover:bg-amber-100 border border-amber-200 rounded-lg px-3 py-1.5 transition shrink-0 shadow-sm">
            <RefreshCw size={11} className="text-amber-500" />
            Re-attach files
            <input type="file" accept=".pptx" multiple className="hidden"
              onChange={e => { addFiles(Array.from(e.target.files ?? []).filter(f => f.name.toLowerCase().endsWith('.pptx'))); e.target.value = '' }} />
          </label>
          <button onClick={() => setRestoredSession(false)} className="text-amber-300 hover:text-amber-600 transition shrink-0">
            <X size={14} />
          </button>
        </div>
      )}

      {/* ── Body ──────────────────────────────────────────────────────── */}
      <div className="flex-1 min-h-0 overflow-hidden">
        <div className="h-full max-w-[1600px] mx-auto px-4 sm:px-6 py-5
            flex flex-col xl:flex-row gap-5">

          {/* ══ LEFT SIDEBAR ════════════════════════════════════════════ */}
          <aside className="xl:w-72 2xl:w-80 shrink-0 flex flex-col gap-4
              xl:overflow-y-auto xl:overflow-x-hidden xl:min-h-0 pb-2"
            style={{ scrollbarWidth: 'thin' }}>

            {/* Upload card */}
            <div className="card p-4">
              <SectionTitle icon={Upload}>Upload PPTX Files</SectionTitle>
              <DropZone onFiles={addFiles} disabled={merging} />
            </div>

            {/* File list card */}
            {files.length > 0 && (
              <div className="card p-4">
                <SectionTitle
                  icon={Layers}
                  meta={`${files.length} file${files.length !== 1 ? 's' : ''}${totalSlides > 0 ? ` · ${totalSlides} slides` : ''}`}>
                  Merge Order
                </SectionTitle>
                <div className="space-y-1.5">
                  {files.map((entry, idx) => (
                    <FileCard key={entry.id} entry={entry} index={idx} total={files.length}
                      isActive={idx === activeFileIdx}
                      onClick={() => selectFile(idx)}
                      onRemove={() => removeFile(idx)}
                      onMoveUp={() => moveFile(idx, -1)}
                      onMoveDown={() => moveFile(idx, 1)} />
                  ))}
                </div>
                <div className="flex items-center justify-between pt-3 mt-2 border-t border-slate-50">
                  <label className="flex items-center gap-1.5 text-xs text-brand-600 font-semibold cursor-pointer hover:text-brand-700 transition group">
                    <input type="file" accept=".pptx" multiple className="hidden"
                      onChange={e => { addFiles(Array.from(e.target.files ?? []).filter(f => f.name.toLowerCase().endsWith('.pptx'))); e.target.value = '' }} />
                    <Plus size={12} className="group-hover:scale-110 transition-transform" /> Add more
                  </label>
                  <button onClick={clearAll} className="flex items-center gap-1.5 text-xs text-slate-400 hover:text-red-500 transition">
                    <Trash2 size={11} /> Clear all
                  </button>
                </div>
              </div>
            )}

            {/* Export card */}
            <div className="card p-4">
              <SectionTitle icon={Sparkles}>Export</SectionTitle>
              <div className="space-y-3">
                <div>
                  <label className="label">Output File Name</label>
                  <div className="flex items-center gap-2">
                    <input type="text" className="input flex-1" placeholder="merged_slides"
                      value={outputName} onChange={e => setOutputName(e.target.value)} />
                    <span className="text-xs text-slate-400 font-mono shrink-0">.pptx</span>
                  </div>
                </div>
                <button onClick={handleMerge}
                  disabled={files.length < 2 || merging || needsReattach}
                  className="btn-primary w-full">
                  {merging
                    ? <><Loader2 size={14} className="animate-spin" />Merging…</>
                    : <><Download size={14} />Merge &amp; Download</>}
                </button>
                {files.length < 2 && (
                  <p className="text-[11px] text-slate-400 text-center">Add at least 2 PPTX files to begin.</p>
                )}
                {files.length >= 2 && needsReattach && (
                  <p className="text-[11px] text-amber-500 text-center flex items-center justify-center gap-1">
                    <RotateCcw size={10} /> Re-attach highlighted files first.
                  </p>
                )}
                {files.length >= 2 && !needsReattach && totalSlides > 0 && (
                  <p className="text-[11px] text-slate-400 text-center">
                    {totalSlides} slides · {files.length} files will be merged
                  </p>
                )}
              </div>
            </div>
          </aside>

          {/* ══ MAIN CONTENT ════════════════════════════════════════════ */}
          <main className="flex-1 min-w-0 min-h-0 flex flex-col gap-4 overflow-hidden">

            {files.length === 0 ? (
              /* ── Empty state ── */
              <div className="flex-1 card flex items-center justify-center p-8">
                <div className="max-w-sm w-full flex flex-col items-center gap-6 text-center">
                  <div className="relative">
                    <div className="w-20 h-20 rounded-3xl bg-gradient-to-br from-brand-50 to-slate-100 flex items-center justify-center shadow-inner">
                      <Layers size={32} className="text-brand-300" />
                    </div>
                    <div className="absolute -top-1 -right-2 w-8 h-8 rounded-xl bg-brand-500 flex items-center justify-center shadow-lg">
                      <Merge size={14} className="text-white" />
                    </div>
                  </div>
                  <div>
                    <h2 className="text-lg font-bold text-slate-700 mb-2">Merge PPTX Files</h2>
                    <p className="text-sm text-slate-400 leading-relaxed">
                      Upload two or more PowerPoint files. Slide thumbnails are generated automatically so you can preview and reorder before exporting.
                    </p>
                  </div>
                  <div className="w-full">
                    <DropZone onFiles={addFiles} disabled={false} />
                  </div>
                  <div className="flex flex-wrap justify-center gap-2">
                    {['Live slide preview', 'Drag to reorder', 'Style preserved', 'One-click export'].map(f => (
                      <span key={f} className="text-[11px] font-medium bg-white border border-slate-100 text-slate-500 rounded-full px-3 py-1 shadow-sm">
                        {f}
                      </span>
                    ))}
                  </div>
                </div>
              </div>
            ) : (
              <>
                {/* ── Slide preview card ── */}
                <div className="card p-4 flex flex-col gap-3 flex-1 min-h-0 overflow-hidden">
                  {/* Header row */}
                  <div className="flex items-center gap-3 shrink-0 flex-wrap">
                    <div className="flex items-center gap-2">
                      <div className="w-6 h-6 rounded-lg bg-brand-50 flex items-center justify-center">
                        <Eye size={12} className="text-brand-500" />
                      </div>
                      <span className="text-sm font-semibold text-slate-700">Preview</span>
                    </div>
                    {/* File tabs */}
                    <div className="flex items-center gap-1.5 flex-wrap">
                      {files.map((entry, fi) => (
                        <button key={entry.id} onClick={() => selectFile(fi)}
                          className={`flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-semibold
                              border transition-all duration-150
                              ${fi === activeFileIdx
                                ? 'bg-brand-500 text-white border-brand-500 shadow-sm'
                                : 'bg-white text-slate-500 border-slate-200 hover:border-brand-300 hover:text-brand-600'}`}>
                          <span className="opacity-60 font-mono">#{fi + 1}</span>
                          <span className="max-w-[100px] truncate">{entry.file.name.replace(/\.pptx$/i, '')}</span>
                          {entry.slideCount != null && (
                            <span className={`text-[9px] rounded px-1 font-bold
                                ${fi === activeFileIdx ? 'bg-white/25 text-white' : 'bg-slate-100 text-slate-400'}`}>
                              {entry.slideCount}
                            </span>
                          )}
                        </button>
                      ))}
                    </div>
                    {allSlides.length > 0 && (
                      <span className="ml-auto text-xs font-semibold tabular-nums px-2.5 py-1
                          rounded-lg bg-slate-50 text-slate-500 border border-slate-100 shrink-0">
                        {Math.max(1, globalIdx + 1)}&thinsp;/&thinsp;{allSlides.length}
                      </span>
                    )}
                  </div>

                  {/* Viewer */}
                  <div className="flex-1 min-h-0 flex items-center justify-center overflow-hidden rounded-xl bg-slate-900/5 p-2">
                    <div className="relative w-full h-full flex items-center justify-center">
                      <div className="w-full" style={{ aspectRatio: '16/9', maxHeight: '100%', maxWidth: '100%' }}>
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
                        className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold
                            bg-white border border-slate-200 text-slate-600
                            hover:border-brand-300 hover:text-brand-600 disabled:opacity-30 transition shadow-sm">
                        <ChevronLeft size={13} /> Prev
                      </button>
                      <span className="text-xs text-slate-400 tabular-nums min-w-[70px] text-center">
                        Slide {Math.max(1, globalIdx + 1)} of {allSlides.length}
                      </span>
                      <button onClick={() => goTo(Math.min(allSlides.length - 1, globalIdx + 1))} disabled={globalIdx >= allSlides.length - 1}
                        className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold
                            bg-white border border-slate-200 text-slate-600
                            hover:border-brand-300 hover:text-brand-600 disabled:opacity-30 transition shadow-sm">
                        Next <ChevronRight size={13} />
                      </button>
                    </div>
                  )}
                </div>

                {/* ── Slide strips card ── */}
                <div className="card p-4 shrink-0 space-y-3">
                  {/* Per-file strip */}
                  <div>
                    <p className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-2">
                      {files[activeFileIdx]?.file.name.replace(/\.pptx$/i, '')}
                      {files[activeFileIdx]?.slideCount != null && (
                        <span className="ml-1.5 text-brand-500 normal-case font-semibold">
                          {files[activeFileIdx].slideCount} slides
                        </span>
                      )}
                    </p>
                    <ScrollStrip innerRef={stripRef}>
                      {allSlides.map((slide, flatIdx) => {
                        if (slide.fileIdx !== activeFileIdx) return null
                        return (
                          <SlideThumbnail
                            key={`${slide.fileIdx}-${slide.slideIdx}`}
                            thumb={slide.thumb}
                            index={slide.slideIdx}
                            loading={slide.loading}
                            active={slide.fileIdx === activeFileIdx && slide.slideIdx === activeSlideIdx}
                            onClick={() => goTo(flatIdx)}
                          />
                        )
                      })}
                    </ScrollStrip>
                  </div>

                  {/* Merge-order strip */}
                  {files.length >= 2 && (
                    <div className="border-t border-slate-100 pt-3">
                      <p className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-2">
                        Merge order
                        <span className="ml-1.5 text-brand-500 normal-case font-semibold">{totalSlides} slides total</span>
                      </p>
                      <ScrollStrip>
                        {allSlides.map((slide, flatIdx) => {
                          const isActive = slide.fileIdx === activeFileIdx && slide.slideIdx === activeSlideIdx
                          const accents  = ['ring-blue-300','ring-violet-300','ring-emerald-300','ring-amber-300','ring-pink-300','ring-teal-300']
                          const accent   = accents[slide.fileIdx % accents.length]
                          return (
                            <button key={`all-${slide.fileIdx}-${slide.slideIdx}`}
                              onClick={() => goTo(flatIdx)} title={slide.label}
                              className={`relative flex-shrink-0 rounded-md overflow-hidden transition-all duration-150 ring-2
                                ${isActive ? 'ring-brand-500 shadow-md scale-105' : `${accent} hover:ring-brand-400`}`}
                              style={{ width: 56, aspectRatio: '16/9' }}>
                              {slide.loading ? (
                                <div className="w-full h-full bg-slate-100 animate-pulse" />
                              ) : slide.thumb ? (
                                <img src={`data:image/png;base64,${slide.thumb}`} className="w-full h-full"
                                  style={{ objectFit: 'fill' }} alt="" draggable={false} />
                              ) : (
                                <div className="w-full h-full bg-slate-100" />
                              )}
                              <div className="absolute bottom-0 right-0 bg-black/50 text-white text-[7px] font-bold px-0.5 leading-snug">
                                {flatIdx + 1}
                              </div>
                            </button>
                          )
                        })}
                        {/* Output marker */}
                        <div className="flex-shrink-0 flex flex-col items-center justify-center gap-0.5
                            bg-gradient-to-br from-green-50 to-emerald-50 ring-2 ring-emerald-300
                            rounded-md px-2 text-emerald-700"
                          style={{ minWidth: 60, aspectRatio: '16/9' }}>
                          <CheckCircle2 size={10} className="text-emerald-500" />
                          <span className="text-[7px] font-bold text-center leading-tight truncate max-w-[52px]">
                            {(outputName || 'merged').slice(0, 12)}
                          </span>
                          <span className="text-[7px] opacity-60">{totalSlides}s</span>
                        </div>
                      </ScrollStrip>
                    </div>
                  )}
                </div>
              </>
            )}
          </main>
        </div>
      </div>

      {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
    </div>
  )
}
