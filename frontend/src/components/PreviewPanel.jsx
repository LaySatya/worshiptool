import { useLayoutEffect, useRef, useState } from "react";
import { ChevronLeft, ChevronRight, Monitor } from "lucide-react";

// PowerPoint slide is 13.33in x 7.5in (16:9 widescreen)
// 13.33in = 1280px at 96dpi; 1pt = 96/72 px.
// N pt on the slide corresponds to N * (96/72) / 1280 * 100 cqw of the slide container width.
const PT_TO_CQW = ((96 / 72) / 1280) * 100;

/**
 * Automatically shrinks text font-size via binary search until it fits within
 * the available slide text-box, matching PowerPoint's "shrink text on overflow".
 */
function useFitText(deps) {
  const boxRef = useRef(null);
  const textRef = useRef(null);
  const [scale, setScale] = useState(1);

  useLayoutEffect(() => {
    const box = boxRef.current;
    const text = textRef.current;
    if (!box || !text) return;

    const fit = () => {
      text.style.setProperty("--fit", "1");
      if (
        text.scrollHeight <= box.clientHeight &&
        text.scrollWidth <= box.clientWidth
      ) {
        setScale(1);
        return;
      }
      let lo = 0.15;
      let hi = 1;
      let best = 0.15;
      for (let i = 0; i < 8; i++) {
        const mid = (lo + hi) / 2;
        text.style.setProperty("--fit", String(mid));
        if (
          text.scrollHeight <= box.clientHeight &&
          text.scrollWidth <= box.clientWidth
        ) {
          best = mid;
          lo = mid;
        } else {
          hi = mid;
        }
      }
      text.style.setProperty("--fit", String(best));
      setScale(best);
    };

    fit();
    const ro = new ResizeObserver(fit);
    ro.observe(box);
    return () => ro.disconnect();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  return { boxRef, textRef, scale };
}

function SlideCanvas({ slide, settings }) {
  const alignMap = { left: "flex-start", center: "center", right: "flex-end" };
  const textAlign = settings.align ?? "center";
  const { boxRef, textRef, scale } = useFitText([slide, settings]);

  // Padding as % of slide width (13.33in standard)
  const padPct = Math.max(1, ((settings.padding ?? 0.55) / 13.33) * 100);

  const lines = slide.lines ?? [];
  const firstLine = lines[0] ?? "";
  const restLines = lines.slice(1);

  return (
    <div
      className="relative w-full rounded-xl overflow-hidden select-none shadow-lg"
      style={{
        aspectRatio: "16/9",
        background: settings.bg_color,
        containerType: "inline-size",
      }}
    >
      {/* Body text container: fixed padding, shrinks font to fit on overflow */}
      <div
        ref={boxRef}
        className="absolute flex items-center overflow-hidden"
        style={{
          left: `${padPct}%`,
          right: `${padPct}%`,
          top: `${(padPct * 16) / 9}%`,
          bottom: `calc(${((padPct * 16) / 9)}% + 9%)`,
          justifyContent: alignMap[textAlign] ?? "center",
        }}
      >
        <p
          ref={textRef}
          style={{
            "--fit": scale,
            color: settings.text_color,
            fontFamily: settings.font_family,
            fontSize: `calc(${settings.font_size * PT_TO_CQW}cqw * var(--fit, 1))`,
            fontWeight: settings.bold ? 700 : 400,
            lineHeight: settings.line_spacing,
            whiteSpace: "pre-wrap",
            textAlign,
            width: "100%",
            margin: 0,
            wordBreak: "break-word",
            overflowWrap: "anywhere",
          }}
        >
          {slide.verse_num ? (
            <>
              <span
                style={{ color: settings.verse_num_color, fontWeight: 700 }}
              >
                {slide.verse_num}
              </span>
              {"  " + firstLine}
              {restLines.length > 0 ? "\n" + restLines.join("\n") : ""}
            </>
          ) : (
            lines.join("\n")
          )}
        </p>
      </div>

      {/* Reference — bottom-right */}
      {slide.ref && (
        <div
          className="absolute bottom-[4%] whitespace-nowrap overflow-hidden text-ellipsis text-right pointer-events-none"
          style={{
            right: `${padPct}%`,
            left: `${padPct}%`,
            color: settings.ref_color,
            fontFamily: settings.font_family,
            fontSize: `${(settings.ref_font_size ?? 20) * PT_TO_CQW}cqw`,
          }}
        >
          {slide.ref}
        </div>
      )}

      {/* Badges: shrink status + aspect */}
      <div className="absolute top-2 right-2 flex items-center gap-1 pointer-events-none">
        {scale < 0.99 && (
          <span className="bg-amber-400/90 text-amber-950 text-[9px] px-1.5 py-0.5 rounded font-semibold shadow-sm">
            Fitted {Math.round(scale * 100)}%
          </span>
        )}
        <span className="bg-black/25 text-white/70 text-[8px] px-1.5 py-0.5 rounded font-mono">
          16:9
        </span>
      </div>
    </div>
  );
}

export default function PreviewPanel({
  slides,
  settings,
  current,
  onNavigate,
  isLoading,
}) {
  const total = slides.length;
  const slide = slides[current];

  return (
    <div className="flex flex-col gap-3">
      {/* Header row */}
      <div className="flex items-center justify-between flex-shrink-0">
        <div className="flex items-center gap-2">
          <Monitor size={15} className="text-brand-500" />
          <span className="font-semibold text-sm text-slate-700">
            Live Preview
          </span>
        </div>
        {total > 0 && (
          <span className="text-xs font-semibold tabular-nums px-2.5 py-0.5 rounded-lg bg-brand-50 text-brand-600 border border-brand-100">
            {current + 1} / {total}
          </span>
        )}
      </div>

      {/* Canvas - Fixed 16:9 container */}
      <div className="w-full flex-shrink-0">
        {isLoading ? (
          <div
            className="w-full rounded-xl bg-slate-100 animate-pulse"
            style={{ aspectRatio: "16/9" }}
          />
        ) : slide ? (
          <SlideCanvas slide={slide} settings={settings} />
        ) : (
          <div
            className="w-full rounded-xl border-2 border-dashed border-slate-200 flex flex-col items-center justify-center gap-2 text-slate-400 bg-slate-50"
            style={{ aspectRatio: "16/9" }}
          >
            <Monitor size={32} className="opacity-25" />
            <p className="text-xs font-medium">Select verses from the reader to preview</p>
          </div>
        )}
      </div>

      {/* Navigation */}
      {total > 1 && (
        <div className="flex items-center gap-2 flex-shrink-0">
          <button
            onClick={() => onNavigate(Math.max(0, current - 1))}
            disabled={current === 0}
            className="btn-secondary !p-1.5"
            title="Previous slide"
          >
            <ChevronLeft size={15} />
          </button>

          {/* Scrollable dot strip */}
          <div className="flex-1 flex items-center justify-center gap-1 overflow-x-auto py-1 no-scrollbar">
            {Array.from({ length: Math.min(total, 24) }, (_, i) => {
              const idx = total > 24 ? Math.round((i / 23) * (total - 1)) : i;
              const active = current === idx;
              return (
                <button
                  key={i}
                  onClick={() => onNavigate(idx)}
                  className={`rounded-full shrink-0 transition-all duration-150 ${
                    active
                      ? "w-4 h-2 bg-brand-500"
                      : "w-1.5 h-1.5 bg-slate-200 hover:bg-slate-400"
                  }`}
                />
              );
            })}
          </div>

          <button
            onClick={() => onNavigate(Math.min(total - 1, current + 1))}
            disabled={current === total - 1}
            className="btn-secondary !p-1.5"
            title="Next slide"
          >
            <ChevronRight size={15} />
          </button>
        </div>
      )}

      {/* Current slide text summary */}
      {slide && (
        <div className="rounded-xl bg-slate-50 border border-slate-100 px-3 py-2 text-xs text-slate-500 max-h-24 overflow-y-auto">
          {slide.verse_num && (
            <span
              className="font-bold mr-1.5"
              style={{ color: settings.verse_num_color }}
            >
              v.{slide.verse_num}
            </span>
          )}
          <span className="leading-relaxed break-words">
            {slide.lines.join(" ")}
          </span>
          {slide.ref && (
            <span className="block pt-1 font-medium text-slate-400 text-[11px]">
              {slide.ref}
            </span>
          )}
        </div>
      )}
    </div>
  );
}
