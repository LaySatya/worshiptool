import { useEffect, useRef, useState } from "react";
import {
  ArrowLeft,
  ChevronLeft,
  ChevronRight,
  ExternalLink,
  Maximize2,
  Minimize2,
  RefreshCw,
} from "lucide-react";
import { bookSrc } from "../../books.js";

const LAST_PAGE_KEY = "worshiptool:bookpage:";

/** Remember where the teacher stopped reading, per book. */
function readLastPage(bookId) {
  try {
    const v = parseInt(localStorage.getItem(LAST_PAGE_KEY + bookId) || "1", 10);
    return Number.isFinite(v) && v > 0 ? v : 1;
  } catch {
    return 1;
  }
}

function saveLastPage(bookId, page) {
  try {
    localStorage.setItem(LAST_PAGE_KEY + bookId, String(page));
  } catch {
    /* storage unavailable — paging still works for this session */
  }
}

/** Pick the speaker icon for the audio the book page is currently playing. */
function detectAudioIcon(doc) {
  if (!doc) return "⏹";
  try {
    const icon = doc.querySelector(
      "#flipbook_audio_icon, .flipbook_audio_icon, #audio_icon, .audio_icon",
    );
    if (!icon) return "";
    const active = /stop|pause|playing|active|on/i.test(icon.className || "");
    if (active) return "⏸";
    return /play/i.test(icon.className || "")
      ? "▶"
      : icon.textContent?.trim() || "";
  } catch {
    return "";
  }
}

/**
 * FlipHTML5 reader for one textbook.
 * Loaded through the same-origin /api/flipbook proxy so the book can be
 * framed, paged, and its reading position remembered.
 */
export default function BookReader({
  book,
  grade,
  subject,
  onBack,
  options = [],
  onSelectBook,
}) {
  const [page, setPage] = useState(() => readLastPage(book.id));
  const [zoom, setZoom] = useState(false);
  const [loading, setLoading] = useState(true);
  const [reloadKey, setReloadKey] = useState(0);
  const [audioIcon, setAudioIcon] = useState("");
  const frameRef = useRef(null);
  const src = bookSrc(book, page);

  // Opening a different book starts at that book's saved page.
  useEffect(() => {
    setLoading(true);
    setPage(readLastPage(book.id));
  }, [book.id]);

  useEffect(() => {
    saveLastPage(book.id, page);
  }, [book.id, page]);

  const syncFromFrame = () => {
    const f = frameRef.current;
    if (!f) return;
    try {
      const doc = f.contentDocument;
      const hash = f.contentWindow?.location?.hash || "";
      const m = hash.match(/[#&?]p=(\d+)/);
      if (m) {
        const p = parseInt(m[1], 10);
        if (p > 0 && p !== page) setPage(p);
      }
      setAudioIcon(detectAudioIcon(doc));
    } catch {
      /* cross-origin during a redirect — ignore */
    }
  };

  useEffect(() => {
    const onHash = (e) => {
      if (e.target === frameRef.current?.contentWindow) syncFromFrame();
    };
    window.addEventListener("hashchange", onHash, true);
    return () => window.removeEventListener("hashchange", onHash, true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [page]);

  const go = (delta) => {
    const next = Math.max(1, page + delta);
    if (next === page) return;
    setLoading(true);
    setPage(next);
  };

  const reload = () => {
    setLoading(true);
    setReloadKey((k) => k + 1);
  };

  const btn =
    "inline-flex items-center justify-center gap-1.5 h-8 px-2.5 rounded-lg border border-slate-200 bg-white text-xs font-semibold text-slate-600 hover:bg-slate-50 hover:text-slate-900 transition disabled:opacity-40 disabled:cursor-not-allowed";

  return (
    <div className="flex flex-col h-full min-h-0 rounded-2xl border border-slate-200 bg-white shadow-xs overflow-hidden">
      {/* Toolbar */}
      <div className="flex flex-wrap items-center gap-2 px-3 py-2 border-b border-slate-100 bg-slate-50/80 flex-shrink-0">
        <button
          type="button"
          onClick={onBack}
          className={btn}
          title="Back to the book library"
        >
          <ArrowLeft size={14} />
          <span className="hidden sm:inline">Library</span>
        </button>

        <div className="flex items-center gap-2 min-w-0">
          <span className="w-2 h-2 rounded-full bg-brand-500 flex-shrink-0" />
          <div className="min-w-0">
            <p className="text-xs font-bold text-slate-800 truncate leading-tight">
              {grade.label} · {subject.label}
              {options.length > 1 && book.label ? ` · ${book.label}` : ""}
            </p>
            <p className="text-[11px] text-slate-400 leading-tight truncate">
              {grade.labelKm} · {subject.labelKm}
              {options.length > 1 && book.labelKm ? ` · ${book.labelKm}` : ""}
            </p>
          </div>
        </div>

        {options.length > 1 && onSelectBook && (
          <div className="flex items-center gap-1 ml-1">
            {options.map((b) => (
              <button
                key={b.id}
                type="button"
                onClick={() => onSelectBook(b)}
                className={
                  b.id === book.id
                    ? "h-8 px-2.5 rounded-lg bg-slate-900 text-white text-[11px] font-semibold"
                    : "h-8 px-2.5 rounded-lg border border-slate-200 bg-white text-[11px] font-semibold text-slate-600 hover:bg-slate-50 transition"
                }
                title={b.labelKm || b.label}
              >
                {b.label || b.id}
              </button>
            ))}
          </div>
        )}

        <div className="ml-auto flex items-center gap-2">
          <div className="flex items-center rounded-lg border border-slate-200 bg-white overflow-hidden h-8">
            <button
              type="button"
              onClick={() => go(-1)}
              disabled={page <= 1}
              className="h-full px-1.5 text-slate-500 hover:bg-slate-50 disabled:opacity-30 transition"
              title="Previous page"
            >
              <ChevronLeft size={15} />
            </button>
            <span className="px-2 text-[11px] font-semibold text-slate-600 tabular-nums whitespace-nowrap">
              Page {page}
            </span>
            <button
              type="button"
              onClick={() => go(1)}
              className="h-full px-1.5 text-slate-500 hover:bg-slate-50 transition"
              title="Next page"
            >
              <ChevronRight size={15} />
            </button>
          </div>

          {audioIcon && (
            <span
              className="h-8 inline-flex items-center px-2 rounded-lg bg-emerald-50 border border-emerald-200 text-xs text-emerald-700"
              title="Book audio is playing"
            >
              {audioIcon}
            </span>
          )}

          <button
            type="button"
            onClick={reload}
            className={btn}
            title="Reload book"
          >
            <RefreshCw size={13} className={loading ? "animate-spin" : ""} />
          </button>
          <button
            type="button"
            onClick={() => setZoom((z) => !z)}
            className={btn}
            title={zoom ? "Exit full screen" : "Full screen reading"}
          >
            {zoom ? <Minimize2 size={13} /> : <Maximize2 size={13} />}
            <span className="hidden sm:inline">
              {zoom ? "Exit" : "Full screen"}
            </span>
          </button>
          <a
            href={book.url}
            target="_blank"
            rel="noreferrer"
            className={btn}
            title="Open the original book in a new tab"
          >
            <ExternalLink size={13} />
          </a>
        </div>
      </div>

      {/* Book frame */}
      <div
        className={
          zoom
            ? "fixed inset-0 z-[60] bg-slate-900 p-0 flex flex-col"
            : "relative flex-1 min-h-0 bg-slate-100"
        }
      >
        {zoom && (
          <div className="h-10 flex items-center justify-between px-3 bg-slate-900 text-slate-200 flex-shrink-0">
            <span className="text-xs font-semibold truncate">
              {grade.label} · {subject.label} book
            </span>
            <button
              type="button"
              onClick={() => setZoom(false)}
              className="inline-flex items-center gap-1.5 px-2.5 h-7 rounded-md bg-white/10 hover:bg-white/20 text-xs font-semibold transition"
            >
              <Minimize2 size={13} /> Exit full screen
            </button>
          </div>
        )}

        {loading && (
          <div className="absolute inset-0 z-10 bg-white/80 flex flex-col items-center justify-center gap-2">
            <RefreshCw size={22} className="animate-spin text-brand-500" />
            <span className="text-xs font-medium text-slate-500">
              Opening the book…
            </span>
          </div>
        )}

        <iframe
          key={`${book.id}-${page}-${reloadKey}`}
          ref={frameRef}
          src={src}
          title={`${grade.label} ${subject.label} book`}
          className="w-full h-full border-0 bg-white"
          allow="autoplay; fullscreen; encrypted-media"
          allowFullScreen
          onLoad={() => {
            setLoading(false);
            syncFromFrame();
          }}
        />
      </div>
    </div>
  );
}
