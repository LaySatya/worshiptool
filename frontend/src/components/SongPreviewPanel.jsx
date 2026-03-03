import { ChevronLeft, ChevronRight, Monitor } from 'lucide-react'

function SongSlideCanvas({ slide, settings }) {
  const alignMap  = { left: 'flex-start', center: 'center', right: 'flex-end' }
  const textAlign = settings.align ?? 'center'
  const scaleFactor = 0.52

  return (
    <div
      className="relative w-full rounded-xl overflow-hidden select-none shadow-lg"
      style={{ aspectRatio: '16/9', background: settings.bg_color }}
    >
      {/* Body lyrics — vertically centred */}
      <div
        className="absolute inset-0 flex items-center px-[5%] py-[6%]"
        style={{ justifyContent: alignMap[textAlign] ?? 'center' }}
      >
        <p
          style={{
            color:      settings.text_color,
            fontFamily: settings.font_family,
            fontSize:   `clamp(10px, calc(${settings.font_size}px * ${scaleFactor}), 9vw)`,
            fontWeight: settings.bold ? 700 : 400,
            lineHeight: settings.line_spacing,
            whiteSpace: 'pre-wrap',
            textAlign,
            width: '100%',
          }}
        >
          {(slide.lines ?? []).join('\n')}
        </p>
      </div>

      {/* Song title — bottom-right */}
      {slide.title && (
        <div
          className="absolute bottom-[4%] right-[4%] max-w-[70%] text-right"
          style={{
            color:      settings.ref_color,
            fontFamily: settings.font_family,
            fontSize:   `clamp(7px, calc(${settings.ref_font_size}px * ${scaleFactor}), 3vw)`,
          }}
        >
          {slide.title}
        </div>
      )}

      {/* 16:9 badge */}
      <div className="absolute top-2 right-2 bg-black/20 text-black/30 text-[8px] px-1.5 py-0.5 rounded font-mono pointer-events-none">
        16:9
      </div>
    </div>
  )
}

export default function SongPreviewPanel({ slides, settings, current, onNavigate, isLoading }) {
  const total = slides.length
  const slide = slides[current]

  return (
    <div className="flex flex-col gap-4">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-2">
          <Monitor size={15} className="text-brand-500" />
          <span className="font-semibold text-sm text-slate-700">Live Preview</span>
        </div>
        {total > 0 && (
          <span className="text-xs font-semibold tabular-nums px-2.5 py-1 rounded-lg bg-brand-50 text-brand-600 border border-brand-100">
            {current + 1} / {total}
          </span>
        )}
      </div>

      {/* Canvas */}
      {isLoading ? (
        <div className="w-full rounded-xl bg-slate-100 animate-pulse" style={{ aspectRatio: '16/9' }} />
      ) : slide ? (
        <SongSlideCanvas slide={slide} settings={settings} />
      ) : (
        <div
          className="w-full rounded-xl border-2 border-dashed border-slate-200 flex flex-col items-center justify-center gap-3 text-slate-400 bg-slate-50"
          style={{ aspectRatio: '16/9' }}
        >
          <Monitor size={40} className="opacity-20" />
          <p className="text-sm font-medium">Upload a file or paste lyrics to preview</p>
          <p className="text-xs opacity-60">Supports PDF, JPG, PNG music sheets</p>
        </div>
      )}

      {/* Navigation */}
      {total > 1 && (
        <div className="flex items-center gap-3">
          <button
            onClick={() => onNavigate(Math.max(0, current - 1))}
            disabled={current === 0}
            className="btn-secondary !p-2"
          >
            <ChevronLeft size={16} />
          </button>

          <div className="flex-1 flex items-center justify-center gap-1 overflow-x-auto py-1 no-scrollbar">
            {Array.from({ length: Math.min(total, 20) }, (_, i) => {
              const idx    = total > 20 ? Math.round((i / 19) * (total - 1)) : i
              const active = current === idx
              return (
                <button
                  key={i}
                  onClick={() => onNavigate(idx)}
                  className={`rounded-full shrink-0 transition-all duration-150 ${
                    active ? 'w-5 h-2.5 bg-brand-500' : 'w-2 h-2 bg-slate-200 hover:bg-slate-400'
                  }`}
                />
              )
            })}
          </div>

          <button
            onClick={() => onNavigate(Math.min(total - 1, current + 1))}
            disabled={current === total - 1}
            className="btn-secondary !p-2"
          >
            <ChevronRight size={16} />
          </button>
        </div>
      )}

      {/* Current slide summary */}
      {slide && (
        <div className="rounded-xl bg-slate-50 border border-slate-100 px-4 py-3 text-xs text-slate-500 space-y-0.5">
          <p className="font-semibold text-slate-700">Slide {current + 1}</p>
          {(slide.lines ?? []).map((ln, i) => (
            <p key={i} className="truncate">{ln}</p>
          ))}
          {slide.title && (
            <p className="italic truncate" style={{ color: settings.ref_color }}>
              🎵 {slide.title}
            </p>
          )}
        </div>
      )}
    </div>
  )
}
