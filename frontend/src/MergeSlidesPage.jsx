import { useState, useCallback, useRef } from 'react'
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
} from 'lucide-react'
import { mergePptxFiles } from './api.js'

// ── Toast ────────────────────────────────────────────────────────
function Toast({ message, type = 'success', onClose }) {
  return (
    <div
      className={`fixed bottom-5 right-5 z-50 flex items-center gap-3 px-4 py-3 rounded-xl shadow-panel border text-sm font-medium animate-fade-in
        ${type === 'success' ? 'bg-white border-green-200 text-green-700' : 'bg-white border-red-200 text-red-700'}`}
    >
      {type === 'success' ? <CheckCircle2 size={15} /> : <AlertCircle size={15} />}
      {message}
      <button onClick={onClose} className="ml-1 text-slate-400 hover:text-slate-600 transition">
        <X size={13} />
      </button>
    </div>
  )
}

// ── File card ────────────────────────────────────────────────────
function FileCard({ file, index, total, onRemove, onMoveUp, onMoveDown, slideCount }) {
  const sizeKB = (file.size / 1024).toFixed(1)

  return (
    <div className="flex items-center gap-3 bg-white border border-slate-200 rounded-xl px-3 py-2.5 shadow-sm group hover:border-brand-300 transition-all">
      {/* drag handle (visual only) */}
      <GripVertical size={16} className="text-slate-300 shrink-0 cursor-grab" />

      {/* icon */}
      <div className="w-8 h-8 rounded-lg bg-brand-50 border border-brand-100 flex items-center justify-center shrink-0">
        <FileSliders size={14} className="text-brand-500" />
      </div>

      {/* name + meta */}
      <div className="flex-1 min-w-0">
        <p className="text-sm font-medium text-slate-800 truncate">{file.name}</p>
        <p className="text-[11px] text-slate-400 mt-0.5">
          {sizeKB} KB
          {slideCount != null && (
            <span className="ml-2 text-brand-500 font-semibold">{slideCount} slide{slideCount !== 1 ? 's' : ''}</span>
          )}
        </p>
      </div>

      {/* order badge */}
      <span className="text-[11px] font-bold text-slate-400 tabular-nums w-5 text-center">
        {index + 1}
      </span>

      {/* move up / down */}
      <div className="flex flex-col gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity">
        <button
          onClick={onMoveUp}
          disabled={index === 0}
          className="p-0.5 rounded text-slate-400 hover:text-brand-600 disabled:opacity-20 transition"
        >
          <ArrowUp size={12} />
        </button>
        <button
          onClick={onMoveDown}
          disabled={index === total - 1}
          className="p-0.5 rounded text-slate-400 hover:text-brand-600 disabled:opacity-20 transition"
        >
          <ArrowDown size={12} />
        </button>
      </div>

      {/* remove */}
      <button
        onClick={onRemove}
        className="p-1 rounded-lg text-slate-300 hover:text-red-500 hover:bg-red-50 transition opacity-0 group-hover:opacity-100"
      >
        <Trash2 size={13} />
      </button>
    </div>
  )
}

// ── Drop Zone ────────────────────────────────────────────────────
function DropZone({ onFiles, disabled }) {
  const inputRef = useRef(null)
  const [drag, setDrag] = useState(false)

  const handleDrop = useCallback(
    e => {
      e.preventDefault()
      setDrag(false)
      if (disabled) return
      const files = Array.from(e.dataTransfer?.files ?? []).filter(f =>
        f.name.toLowerCase().endsWith('.pptx'),
      )
      if (files.length) onFiles(files)
    },
    [onFiles, disabled],
  )

  return (
    <div
      onDrop={handleDrop}
      onDragOver={e => { e.preventDefault(); setDrag(true) }}
      onDragLeave={() => setDrag(false)}
      onClick={() => !disabled && inputRef.current?.click()}
      className={`relative flex flex-col items-center justify-center gap-3 rounded-2xl border-2 border-dashed cursor-pointer transition-all py-8 px-4
        ${drag ? 'border-brand-400 bg-brand-50 scale-[1.01]' : 'border-slate-200 bg-slate-50 hover:border-brand-300 hover:bg-brand-50/40'}
        ${disabled ? 'opacity-60 cursor-not-allowed' : ''}`}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".pptx"
        multiple
        className="hidden"
        onChange={e => {
          const files = Array.from(e.target.files ?? []).filter(f =>
            f.name.toLowerCase().endsWith('.pptx'),
          )
          if (files.length) onFiles(files)
          e.target.value = ''
        }}
      />
      <div className="w-12 h-12 rounded-2xl bg-brand-100 flex items-center justify-center">
        <Upload size={22} className="text-brand-500" />
      </div>
      <div className="text-center">
        <p className="text-sm font-semibold text-slate-700">
          {drag ? 'Drop your PPTX files here' : 'Click or drag & drop PPTX files'}
        </p>
        <p className="text-xs text-slate-400 mt-1">Add as many slides as you like — reorder them below</p>
      </div>
    </div>
  )
}

// ── Slide thumbnail preview ───────────────────────────────────────
function PreviewPlaceholder({ index, total }) {
  return (
    <div
      className="w-full rounded-lg bg-slate-800 flex items-center justify-center select-none"
      style={{ aspectRatio: '16/9' }}
    >
      <div className="text-center">
        <Layers size={28} className="text-slate-500 mx-auto mb-1" />
        <p className="text-slate-400 text-xs font-mono">Slide {index + 1} / {total}</p>
      </div>
    </div>
  )
}

// ── Main Page ─────────────────────────────────────────────────────
export default function MergeSlidesPage() {
  const [files, setFiles] = useState([])       // { file: File, slideCount: number|null }[]
  const [previewIndex, setPreviewIndex] = useState(0)
  const [merging, setMerging] = useState(false)
  const [outputName, setOutputName] = useState('merged_slides')
  const [toast, setToast] = useState(null)

  const totalSlides = files.reduce((s, f) => s + (f.slideCount ?? 0), 0)

  // ── Add files ────────────────────────────────────────────────
  const addFiles = useCallback(newFiles => {
    const entries = newFiles.map(f => ({ file: f, slideCount: null, id: crypto.randomUUID() }))
    setFiles(prev => {
      // dedupe by name+size
      const existing = new Set(prev.map(e => `${e.file.name}-${e.file.size}`))
      return [...prev, ...entries.filter(e => !existing.has(`${e.file.name}-${e.file.size}`))]
    })
    // Count slides via backend (lightweight)
    entries.forEach(async entry => {
      try {
        const fd = new FormData()
        fd.append('file', entry.file)
        const res = await fetch('/api/merge/count-slides', { method: 'POST', body: fd })
        if (res.ok) {
          const { slide_count } = await res.json()
          setFiles(prev =>
            prev.map(e => (e.id === entry.id ? { ...e, slideCount: slide_count } : e)),
          )
        }
      } catch { /* ignore count errors */ }
    })
  }, [])

  // ── Remove / move ────────────────────────────────────────────
  const removeFile = idx => setFiles(prev => prev.filter((_, i) => i !== idx))

  const moveFile = (idx, dir) => {
    setFiles(prev => {
      const arr = [...prev]
      const target = idx + dir
      if (target < 0 || target >= arr.length) return arr
      ;[arr[idx], arr[target]] = [arr[target], arr[idx]]
      return arr
    })
    setPreviewIndex(idx => Math.max(0, Math.min(idx, files.length - 2)))
  }

  // ── Merge & download ─────────────────────────────────────────
  const handleMerge = async () => {
    if (files.length < 2) {
      setToast({ message: 'Add at least 2 PPTX files to merge.', type: 'error' })
      return
    }
    setMerging(true)
    try {
      const blob = await mergePptxFiles(
        files.map(e => e.file),
        outputName.trim() || 'merged_slides',
      )
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `${outputName.trim() || 'merged_slides'}.pptx`
      a.click()
      URL.revokeObjectURL(url)
      setToast({ message: `Merged ${files.length} files into one PPTX!`, type: 'success' })
    } catch (err) {
      setToast({ message: `Merge failed — ${err.message}`, type: 'error' })
    } finally {
      setMerging(false)
    }
  }

  const canMerge = files.length >= 2 && !merging

  return (
    <div className="flex-1 max-w-[1400px] mx-auto w-full px-4 py-5
                    grid grid-cols-1
                    lg:grid-cols-[420px_1fr]
                    gap-5 items-start">

      {/* ── LEFT PANEL ── */}
      <div className="space-y-4 lg:sticky lg:top-20">

        {/* Drop zone */}
        <div className="card p-4 space-y-3">
          <div className="flex items-center gap-2">
            <Merge size={14} className="text-brand-500" />
            <h2 className="font-semibold text-sm text-slate-800">Attach Slides</h2>
          </div>
          <DropZone onFiles={addFiles} disabled={merging} />
          {files.length > 0 && (
            <button
              onClick={() => { setFiles([]); setPreviewIndex(0) }}
              className="text-xs text-slate-400 hover:text-red-500 flex items-center gap-1 transition"
            >
              <Trash2 size={11} /> Clear all
            </button>
          )}
        </div>

        {/* File list */}
        {files.length > 0 && (
          <div className="card p-4 space-y-2">
            <div className="flex items-center justify-between mb-1">
              <h3 className="text-[10px] font-bold uppercase tracking-widest text-slate-400">
                Files — in order
              </h3>
              <span className="text-[11px] font-semibold text-brand-600 bg-brand-50 border border-brand-100 px-2 py-0.5 rounded-full">
                {files.length} file{files.length !== 1 ? 's' : ''}
                {totalSlides > 0 && ` · ${totalSlides} slides`}
              </span>
            </div>

            <div className="space-y-2 max-h-[45vh] overflow-y-auto pr-0.5">
              {files.map((entry, idx) => (
                <FileCard
                  key={entry.id}
                  file={entry.file}
                  index={idx}
                  total={files.length}
                  slideCount={entry.slideCount}
                  onRemove={() => removeFile(idx)}
                  onMoveUp={() => moveFile(idx, -1)}
                  onMoveDown={() => moveFile(idx, 1)}
                />
              ))}
            </div>

            {/* add more */}
            <label className="flex items-center gap-2 text-xs text-brand-600 font-semibold cursor-pointer hover:text-brand-700 pt-1 transition">
              <input
                type="file"
                accept=".pptx"
                multiple
                className="hidden"
                onChange={e => {
                  addFiles(Array.from(e.target.files ?? []).filter(f => f.name.toLowerCase().endsWith('.pptx')))
                  e.target.value = ''
                }}
              />
              <Plus size={13} /> Add more files
            </label>
          </div>
        )}

        {/* Output settings */}
        <div className="card p-4 space-y-3">
          <h3 className="text-[10px] font-bold uppercase tracking-widest text-slate-400">Output</h3>
          <div>
            <label className="label">File Name</label>
            <div className="flex gap-2">
              <input
                type="text"
                className="input flex-1"
                placeholder="merged_slides"
                value={outputName}
                onChange={e => setOutputName(e.target.value)}
              />
              <span className="flex items-center text-xs text-slate-400 font-mono shrink-0">.pptx</span>
            </div>
          </div>
          <button
            onClick={handleMerge}
            disabled={!canMerge}
            className="btn-primary w-full"
          >
            {merging ? (
              <><Loader2 size={14} className="animate-spin" />Merging…</>
            ) : (
              <><Download size={14} />Merge &amp; Download PPTX</>
            )}
          </button>
          {files.length < 2 && (
            <p className="text-[11px] text-slate-400 text-center">
              Add at least 2 PPTX files to enable merging.
            </p>
          )}
        </div>
      </div>

      {/* ── RIGHT PANEL — Preview ── */}
      <div className="card p-6">
        <div className="flex items-center gap-2 mb-4">
          <Eye size={15} className="text-brand-500" />
          <span className="font-semibold text-sm text-slate-700">File Preview</span>
          {files.length > 0 && (
            <span className="ml-auto text-xs font-semibold tabular-nums px-2.5 py-1 rounded-lg bg-brand-50 text-brand-600 border border-brand-100">
              {previewIndex + 1} / {files.length}
            </span>
          )}
        </div>

        {files.length === 0 ? (
          /* Empty state */
          <div
            className="w-full rounded-2xl border-2 border-dashed border-slate-200 flex flex-col items-center justify-center gap-4 bg-slate-50 text-slate-400"
            style={{ minHeight: '320px' }}
          >
            <div className="w-16 h-16 rounded-2xl bg-slate-100 flex items-center justify-center">
              <Layers size={28} className="text-slate-300" />
            </div>
            <div className="text-center">
              <p className="text-sm font-semibold text-slate-500">No files attached yet</p>
              <p className="text-xs text-slate-400 mt-1">Upload PPTX files on the left to get started</p>
            </div>
          </div>
        ) : (
          <div className="space-y-4">
            {/* File info card */}
            <div className="bg-slate-50 rounded-xl border border-slate-200 px-4 py-3">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-lg bg-brand-100 flex items-center justify-center shrink-0">
                  <FileSliders size={16} className="text-brand-600" />
                </div>
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-semibold text-slate-800 truncate">
                    {files[previewIndex].file.name}
                  </p>
                  <p className="text-xs text-slate-400 mt-0.5">
                    {(files[previewIndex].file.size / 1024).toFixed(1)} KB
                    {files[previewIndex].slideCount != null && (
                      <span className="ml-2 text-brand-500 font-semibold">
                        {files[previewIndex].slideCount} slide{files[previewIndex].slideCount !== 1 ? 's' : ''}
                      </span>
                    )}
                    <span className="ml-2 text-slate-300">·</span>
                    <span className="ml-2">Position {previewIndex + 1} of {files.length}</span>
                  </p>
                </div>
                <span className="text-[11px] font-bold text-slate-400 bg-white border border-slate-200 rounded-lg px-2 py-1">
                  #{previewIndex + 1}
                </span>
              </div>
            </div>

            {/* Visual placeholder for the file's slides */}
            <div className="relative">
              <PreviewPlaceholder index={previewIndex} total={files.length} />
              <div className="absolute bottom-3 left-1/2 -translate-x-1/2 bg-black/60 text-white/80 text-xs font-semibold px-3 py-1 rounded-full backdrop-blur-sm pointer-events-none">
                {files[previewIndex].file.name}
              </div>
            </div>

            {/* Merge order visualiser */}
            <div>
              <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-2">
                Merge Order Preview
              </p>
              <div className="flex flex-wrap gap-2">
                {files.map((entry, idx) => (
                  <button
                    key={entry.id}
                    onClick={() => setPreviewIndex(idx)}
                    className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-semibold border transition-all
                      ${idx === previewIndex
                        ? 'bg-brand-500 text-white border-brand-500 shadow'
                        : 'bg-white text-slate-600 border-slate-200 hover:border-brand-300 hover:text-brand-600'}`}
                  >
                    <span className="opacity-60">#{idx + 1}</span>
                    <span className="max-w-[120px] truncate">{entry.file.name.replace(/\.pptx$/i, '')}</span>
                    {entry.slideCount != null && (
                      <span className={`rounded px-1 py-0.5 text-[9px] ${idx === previewIndex ? 'bg-white/20' : 'bg-slate-100 text-slate-400'}`}>
                        {entry.slideCount}s
                      </span>
                    )}
                  </button>
                ))}
                {files.length >= 2 && (
                  <div className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-green-50 text-green-700 border border-green-200">
                    <CheckCircle2 size={11} />
                    {outputName || 'merged_slides'}.pptx
                    {totalSlides > 0 && <span className="opacity-70">· {totalSlides} slides</span>}
                  </div>
                )}
              </div>
            </div>

            {/* Navigation */}
            {files.length > 1 && (
              <div className="flex items-center justify-center gap-3">
                <button
                  onClick={() => setPreviewIndex(i => Math.max(0, i - 1))}
                  disabled={previewIndex === 0}
                  className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white border border-slate-200 text-slate-600 hover:border-brand-300 hover:text-brand-600 disabled:opacity-30 transition"
                >
                  <ChevronLeft size={13} /> Previous
                </button>
                <span className="text-xs text-slate-400 tabular-nums">{previewIndex + 1} / {files.length}</span>
                <button
                  onClick={() => setPreviewIndex(i => Math.min(files.length - 1, i + 1))}
                  disabled={previewIndex === files.length - 1}
                  className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white border border-slate-200 text-slate-600 hover:border-brand-300 hover:text-brand-600 disabled:opacity-30 transition"
                >
                  Next <ChevronRight size={13} />
                </button>
              </div>
            )}
          </div>
        )}
      </div>

      {toast && (
        <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />
      )}
    </div>
  )
}
