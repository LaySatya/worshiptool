import { useState, useEffect, useCallback, useRef } from "react";
import {
  ScanText,
  Upload,
  FileText,
  FileUp,
  Image as ImageIcon,
  FileType2,
  Loader2,
  CheckCircle2,
  AlertCircle,
  X,
  Copy,
  Sliders,
  Download,
  Monitor,
  Sparkles,
  ChevronLeft,
  ChevronRight,
  Layers,
  Maximize2,
  ImageIcon as ImageIcon2,
  FileImage,
} from "lucide-react";
import SongSettingsPanel from "./components/SongSettingsPanel.jsx";
import {
  extractSongLyrics,
  fetchSongPreview,
  generateSongPptx,
} from "./api.js";
import { DEFAULT_SONG_SETTINGS } from "./constants.js";

function useDebounce(value, delay) {
  const [d, setD] = useState(value);
  useEffect(() => {
    const t = setTimeout(() => setD(value), delay);
    return () => clearTimeout(t);
  }, [value, delay]);
  return d;
}

function Toast({ message, type = "success", onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 4500);
    return () => clearTimeout(t);
  }, [onClose]);
  return (
    <div
      className={`fixed bottom-5 right-5 z-50 flex items-center gap-3 px-4 py-3 rounded-xl shadow-panel border text-sm font-medium animate-fade-in
      ${type === "success" ? "bg-white border-green-200 text-green-700" : "bg-white border-red-200 text-red-700"}`}
    >
      {type === "success" ? (
        <CheckCircle2 size={15} />
      ) : (
        <AlertCircle size={15} />
      )}
      {message}
      <button
        onClick={onClose}
        className="ml-1 text-slate-400 hover:text-slate-600"
      >
        <X size={13} />
      </button>
    </div>
  );
}

/** Full-size 16:9 preview canvas for a single extracted slide. */
function ExtractSlideCanvas({ slide, settings }) {
  const alignMap = { left: "flex-start", center: "center", right: "flex-end" };
  const textAlign = settings.align ?? "center";
  const scaleFactor = 0.52;

  if (slide.is_title) {
    return (
      <div
        className="relative w-full rounded-xl overflow-hidden select-none shadow-lg flex flex-col items-center justify-center p-8 text-center"
        style={{ aspectRatio: "16/9", background: settings.bg_color }}
      >
        {slide.category && (
          <span className="inline-block px-3 py-1 rounded-full text-[11px] font-extrabold uppercase tracking-widest bg-brand-50 text-brand-600 border border-brand-200 mb-3 shadow-xs">
            {slide.category}
          </span>
        )}
        <h2
          style={{
            color: settings.text_color,
            fontFamily: settings.font_family,
            fontSize: `clamp(16px, calc(${settings.font_size + 8}px * ${scaleFactor}), 7vw)`,
            fontWeight: 800,
            lineHeight: 1.25,
            marginBottom: "0.4rem",
            maxWidth: "90%",
          }}
        >
          {slide.title}
        </h2>
        {slide.subtitle && (
          <p
            style={{
              color: settings.ref_color,
              fontFamily: settings.font_family,
              fontSize: `clamp(11px, calc(${Math.max(settings.font_size - 14, 20)}px * ${scaleFactor}), 4.5vw)`,
              fontWeight: 500,
              lineHeight: 1.3,
              maxWidth: "85%",
            }}
          >
            {slide.subtitle}
          </p>
        )}
      </div>
    );
  }

  const lines = slide.lines ?? [];
  const firstLine = lines[0] ?? "";
  const restLines = lines.slice(1);
  const isChorus = slide.prefix === "R." || slide.prefix_color === "red";

  return (
    <div
      className="relative w-full rounded-xl overflow-hidden select-none shadow-lg"
      style={{ aspectRatio: "16/9", background: settings.bg_color }}
    >
      <div
        className="absolute inset-0 flex items-center px-[5%] py-[6%] overflow-hidden"
        style={{ justifyContent: alignMap[textAlign] ?? "center" }}
      >
        <p
          style={{
            color: settings.text_color,
            fontFamily: settings.font_family,
            fontSize: `clamp(10px, calc(${settings.font_size}px * ${scaleFactor}), 9vw)`,
            fontWeight: settings.bold ? 700 : 400,
            lineHeight: settings.line_spacing,
            whiteSpace: "pre-wrap",
            textAlign,
            width: "100%",
            overflowWrap: "break-word",
          }}
        >
          {slide.prefix ? (
            <>
              <span
                style={{
                  color: isChorus
                    ? settings.chorus_r_color || "#ef4444"
                    : settings.verse_num_color || "#ffcc44",
                  fontWeight: 800,
                }}
              >
                {slide.prefix}
              </span>
              {"  " + firstLine}
              {restLines.length > 0 ? "\n" + restLines.join("\n") : ""}
            </>
          ) : (
            lines.join("\n")
          )}
        </p>
      </div>

      {slide.title && (
        <div
          className="absolute bottom-[4%] right-[4%] max-w-[70%] text-right truncate pointer-events-none"
          style={{
            color: settings.ref_color,
            fontFamily: settings.font_family,
            fontSize: `clamp(7px, calc(${settings.ref_font_size}px * ${scaleFactor}), 3vw)`,
          }}
        >
          {slide.title}
        </div>
      )}
    </div>
  );
}

const ACCEPTED = ".pdf,.png,.jpg,.jpeg,.webp,.txt";

export default function SongExtractPage() {
  // ── Upload / extraction state ───────────────────────────────────
  const [file, setFile] = useState(null);
  const [extracting, setExtracting] = useState(false);
  const [pages, setPages] = useState([]);           // base64 page previews
  const [titleImage, setTitleImage] = useState(""); // song-title block crop
  const [pageOpen, setPageOpen] = useState(0);      // which page is shown
  const [previewOpen, setPreviewOpen] = useState(false); // 75% slide modal
  const [extractedText, setExtractedText] = useState("");
  const [extractError, setExtractError] = useState("");
  const [dragActive, setDragActive] = useState(false);
  const [copied, setCopied] = useState(false);

  // ── Slide builder state ─────────────────────────────────────────
  const [songTitle, setSongTitle] = useState("");
  const [settings, setSettings] = useState(DEFAULT_SONG_SETTINGS);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [slides, setSlides] = useState([]);
  const [currentSlide, setCurrentSlide] = useState(0);
  const [previewLoad, setPreviewLoad] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [toast, setToast] = useState(null);

  const fileInputRef = useRef(null);

  const dText = useDebounce(extractedText, 350);
  const dSettings = useDebounce(settings, 250);
  const dTitle = useDebounce(songTitle, 250);

  const hasText = extractedText.trim().length > 0;

  // ── Live preview of the extracted lyrics ────────────────────────
  useEffect(() => {
    if (!dText.trim()) {
      setSlides([]);
      setCurrentSlide(0);
      return;
    }
    let cancelled = false;
    setPreviewLoad(true);
    fetchSongPreview({
      lyrics_text: dText,
      settings: dSettings,
      song_title: dTitle,
    })
      .then((data) => {
        if (cancelled) return;
        setSlides(data.slides ?? []);
        setCurrentSlide(0);
      })
      .catch(() => {
        if (!cancelled)
          setToast({
            message: "Preview error — check backend.",
            type: "error",
          });
      })
      .finally(() => {
        if (!cancelled) setPreviewLoad(false);
      });
    return () => {
      cancelled = true;
    };
  }, [dText, dSettings, dTitle]);

  // ── Extraction ──────────────────────────────
  const runExtraction = useCallback(async (f) => {
    if (!f) return;
    setFile(f);
    setExtractedText("");
    setExtractError("");
    setSongTitle(
      f.name
        .replace(/\.[^.]+$/, "")
        .replace(/[-_]+/g, " ")
        .trim(),
    );
    setExtracting(true);
    setPages([]);
    setTitleImage("");
    try {
      const data = await extractSongLyrics(f);
      // Page previews of the uploaded PDF / image (shown beside the text)
      setPages(data.pages ?? []);
      // The song-title block is a PICTURE for the first slide — it is never
      // OCR'd into the lyric text box.
      setTitleImage(data.title_image || "");
      const text = (data.text || "").trim();
      if (!text || /^\(no .*text detected\)$/i.test(text)) {
        setExtractError(
          "No readable Khmer or English text was detected in this file.",
        );
        setToast({
          message: "No text found — try a clearer scan.",
          type: "error",
        });
      } else {
        setExtractedText(text);
        setToast({
          message: "Lyrics extracted successfully!",
          type: "success",
        });
      }
    } catch (err) {
      setExtractError(err?.message || "Extraction failed — check backend.");
      setToast({
        message: "Extraction failed — check backend.",
        type: "error",
      });
    } finally {
      setExtracting(false);
    }
  }, []);

  // ── Drag & drop ─────────────────────────────────────────────────
  const onDrop = (e) => {
    e.preventDefault();
    setDragActive(false);
    const f = e.dataTransfer?.files?.[0];
    if (f) runExtraction(f);
  };

  const handleClear = () => {
    setFile(null);
    setExtractedText("");
    setExtractError("");
    setSongTitle("");
    setSlides([]);
    setCurrentSlide(0);
    setPages([]);
    setTitleImage("");
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  // ── Copy extracted text ─────────────────────────────────────────
  const handleCopy = async () => {
    if (!extractedText) return;
    try {
      await navigator.clipboard.writeText(extractedText);
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
      setToast({ message: "Lyrics copied to clipboard!", type: "success" });
    } catch {
      setToast({
        message: "Copy failed — select the text manually.",
        type: "error",
      });
    }
  };

  // ── Generate PPTX from the extracted text ───────────────────────
  const handleGenerate = async () => {
    if (!extractedText.trim()) {
      setToast({ message: "Extract lyrics first.", type: "error" });
      return;
    }
    setGenerating(true);
    try {
      const blob = await generateSongPptx({
        lyrics_text: extractedText,
        settings,
        song_title: songTitle,
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${songTitle || "song"}_slides.pptx`;
      a.click();
      URL.revokeObjectURL(url);
      setToast({
        message: `${slides.length} slides downloaded!`,
        type: "success",
      });
    } catch {
      setToast({
        message: "PPTX generation failed — check backend.",
        type: "error",
      });
    } finally {
      setGenerating(false);
    }
  };

  const fileIcon = file ? (
    file.name.toLowerCase().endsWith(".pdf") ? (
      <FileType2 size={16} className="text-rose-500" />
    ) : file.name.toLowerCase().endsWith(".txt") ? (
      <FileText size={16} className="text-slate-500" />
    ) : (
      <ImageIcon size={16} className="text-brand-500" />
    )
  ) : null;


  return (
    <div className="flex-1 overflow-hidden flex flex-col h-full bg-slate-50">
      <input
        ref={fileInputRef}
        type="file"
        accept={ACCEPTED}
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) runExtraction(f);
        }}
      />

      {/* ── Top toolbar ─────────────────────────────────────────── */}
      <div className="bg-white border-b border-slate-200 px-6 py-3 flex flex-wrap items-center justify-between gap-3 flex-shrink-0 z-10 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-brand-50 text-brand-600 flex items-center justify-center">
            <ScanText size={16} />
          </div>
          <div>
            <h2 className="text-sm font-bold text-slate-800">
              Upload Song File → Extract Lyrics
            </h2>
            <p className="text-xs text-slate-400">
              Lyrics under the melody are extracted — the title block becomes a
              picture on the first slide
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            disabled={extracting}
            className="btn-secondary !py-2 !px-3 text-xs flex items-center gap-1.5"
          >
            {extracting ? (
              <Loader2 size={13} className="animate-spin" />
            ) : (
              <FileUp size={13} className="text-brand-500" />
            )}
            <span>{extracting ? "Extracting…" : "Upload File"}</span>
          </button>

          <button
            type="button"
            onClick={() => setSettingsOpen((prev) => !prev)}
            className={`btn-secondary !py-2 !px-3 text-xs flex items-center gap-1.5 ${
              settingsOpen ? "bg-brand-50 text-brand-600 border-brand-200" : ""
            }`}
          >
            <Sliders size={14} />
            <span>{settingsOpen ? "Hide Settings" : "Slide Settings"}</span>
          </button>

          <button
            type="button"
            onClick={handleGenerate}
            disabled={generating || !hasText}
            className="btn-primary !py-2 !px-4 text-xs flex items-center gap-1.5"
          >
            {generating ? (
              <>
                <Loader2 size={14} className="animate-spin" />
                Generating…
              </>
            ) : (
              <>
                <Download size={14} />
                Create Slides (PPTX)
              </>
            )}
          </button>
        </div>
      </div>

      {/* ── Workspace ───────────────────────────────────────────── */}
      <div className="flex-1 overflow-hidden flex">
        {/* scrollable content */}
        <div className="flex-1 overflow-y-auto p-5 min-w-0">
          <div className="max-w-[1500px] mx-auto grid grid-cols-1 lg:grid-cols-2 gap-5">

            {/* ── LEFT: upload + extracted text ─────────────────── */}
            <div className="space-y-5 min-w-0">
              {/* upload card */}
              <div className="card p-5">
                <div className="flex items-center justify-between pb-2 border-b border-slate-100">
                  <div className="flex items-center gap-2">
                    <Upload size={16} className="text-brand-500" />
                    <h3 className="font-semibold text-slate-800 text-sm">
                      Song File
                    </h3>
                  </div>
                  {file && (
                    <button
                      onClick={handleClear}
                      className="text-xs text-slate-400 hover:text-red-500 transition font-medium"
                    >
                      Clear
                    </button>
                  )}
                </div>

                <div
                  onDragOver={(e) => {
                    e.preventDefault();
                    setDragActive(true);
                  }}
                  onDragLeave={() => setDragActive(false)}
                  onDrop={onDrop}
                  onClick={() => !extracting && fileInputRef.current?.click()}
                  className={`mt-4 rounded-xl border-2 border-dashed px-5 py-8 text-center cursor-pointer transition ${
                    dragActive
                      ? "border-brand-400 bg-brand-50/60"
                      : "border-slate-200 bg-slate-50/60 hover:border-brand-300 hover:bg-brand-50/30"
                  }`}
                >
                  {extracting ? (
                    <div className="flex flex-col items-center gap-2 text-slate-500">
                      <Loader2 size={26} className="animate-spin text-brand-500" />
                      <p className="text-sm font-medium">Extracting lyrics…</p>
                      <p className="text-[11px] text-slate-400">
                        Reading the file, this can take a few seconds per page
                      </p>
                    </div>
                  ) : file ? (
                    <div className="flex flex-col items-center gap-1.5">
                      <div className="w-10 h-10 rounded-lg bg-white border border-slate-200 flex items-center justify-center">
                        {fileIcon}
                      </div>
                      <p className="text-sm font-semibold text-slate-800 break-all">
                        {file.name}
                      </p>
                      <p className="text-[11px] text-slate-400">
                        {(file.size / 1024).toFixed(0)} KB · click to choose
                        another file
                      </p>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center gap-1.5">
                      <div className="w-10 h-10 rounded-lg bg-white border border-slate-200 flex items-center justify-center">
                        <Upload size={16} className="text-brand-500" />
                      </div>
                      <p className="text-sm font-semibold text-slate-800">
                        Drop your song file here, or click to browse
                      </p>
                      <p className="text-[11px] text-slate-400">
                        PDF · PNG · JPG · WEBP · TXT (max 20 MB)
                      </p>
                    </div>
                  )}
                </div>

                <div className="mt-3 p-3 bg-slate-50 rounded-xl border border-slate-100 text-[11px] text-slate-500 space-y-1">
                  <p className="font-semibold text-slate-700">
                    💡 What gets extracted:
                  </p>
                  <ul className="list-disc list-inside space-y-0.5 text-slate-600">
                    <li>Only the Khmer lyrics printed under the melody.</li>
                    <li>
                      Staff lines, chords and solfège note rows are removed.
                    </li>
                    <li>
                      The song title block at the top becomes a picture on the
                      first slide — it is never OCR'd into the lyrics.
                    </li>
                  </ul>
                </div>
              </div>

              {/* extracted text card */}
              <div className="card p-5">
                <div className="flex items-center justify-between pb-2 border-b border-slate-100">
                  <div className="flex items-center gap-2">
                    <FileText size={16} className="text-brand-500" />
                    <h3 className="font-semibold text-slate-800 text-sm">
                      Extracted Lyrics
                    </h3>
                    {hasText && (
                      <span className="text-[11px] font-semibold px-2 py-0.5 rounded-lg bg-emerald-50 text-emerald-700 border border-emerald-100">
                        Ready
                      </span>
                    )}
                  </div>
                  <button
                    type="button"
                    onClick={handleCopy}
                    disabled={!hasText}
                    className={`text-xs flex items-center gap-1 font-medium transition ${
                      hasText
                        ? "text-brand-600 hover:underline"
                        : "text-slate-300 cursor-not-allowed"
                    }`}
                  >
                    {copied ? <CheckCircle2 size={12} /> : <Copy size={12} />}
                    {copied ? "Copied!" : "Copy"}
                  </button>
                </div>

                <div className="mt-3">
                  <input
                    type="text"
                    value={songTitle}
                    onChange={(e) => setSongTitle(e.target.value)}
                    placeholder="Song title (used on the title slide)"
                    className="w-full mb-2 px-3 py-2 rounded-lg border border-slate-200 text-xs font-semibold focus:outline-none focus:border-brand-400"
                  />
                  <textarea
                    className="input resize-y text-[13px] w-full font-sans min-h-[340px]"
                    style={{ lineHeight: "1.15", whiteSpace: "pre-wrap" }}
                    placeholder="Extracted lyrics will appear here after processing your file…"
                    value={extractedText}
                    onChange={(e) => setExtractedText(e.target.value)}
                  />
                </div>

                {extractError && (
                  <p className="mt-2 text-[11px] text-rose-600 flex items-start gap-1">
                    <AlertCircle size={12} className="mt-0.5 flex-shrink-0" />
                    {extractError}
                  </p>
                )}

                {hasText && (
                  <p className="mt-2 text-[11px] text-slate-400 flex items-center gap-1">
                    <Sparkles size={11} className="text-amber-500" />
                    You can edit the text above before creating slides.
                  </p>
                )}
              </div>
            </div>

            {/* ── RIGHT: uploaded page + slide preview ──────────── */}
            <div className="flex flex-col gap-5 h-fit lg:sticky lg:top-0 min-w-0">

              {/* uploaded file pages */}
              {pages.length > 0 && (
                <div className="card p-5">
                  <div className="flex items-center justify-between pb-2 border-b border-slate-100">
                    <div className="flex items-center gap-2">
                      <FileImage size={16} className="text-rose-500" />
                      <h3 className="font-semibold text-slate-800 text-sm">
                        Uploaded File
                      </h3>
                      <span className="text-[11px] font-semibold px-2 py-0.5 rounded-lg bg-slate-100 text-slate-600">
                        {pages.length} page{pages.length > 1 ? "s" : ""}
                      </span>
                    </div>
                    {pages.length > 1 && (
                      <div className="flex items-center gap-1.5">
                        <button
                          onClick={() => setPageOpen((p) => Math.max(0, p - 1))}
                          disabled={pageOpen === 0}
                          className="btn-secondary !p-1.5"
                        >
                          <ChevronLeft size={14} />
                        </button>
                        <span className="text-xs font-semibold tabular-nums text-slate-600">
                          {pageOpen + 1} / {pages.length}
                        </span>
                        <button
                          onClick={() =>
                            setPageOpen((p) => Math.min(pages.length - 1, p + 1))
                          }
                          disabled={pageOpen === pages.length - 1}
                          className="btn-secondary !p-1.5"
                        >
                          <ChevronRight size={14} />
                        </button>
                      </div>
                    )}
                  </div>

                  <div className="mt-3 rounded-lg overflow-hidden border border-slate-200 bg-slate-100">
                    <img
                      src={pages[pageOpen]}
                      alt={`Uploaded page ${pageOpen + 1}`}
                      className="w-full h-auto block"
                    />
                  </div>

                  {titleImage && (
                    <div className="mt-3 p-3 rounded-lg bg-slate-50 border border-slate-200">
                      <p className="text-[11px] font-semibold text-slate-600 mb-2 flex items-center gap-1">
                        <ImageIcon2 size={11} className="text-brand-500" />
                        Song title detected — used as the first slide (not
                        extracted as lyrics)
                      </p>
                      <img
                        src={titleImage}
                        alt="Detected song title"
                        className="w-full h-auto rounded border border-slate-200 bg-white"
                      />
                    </div>
                  )}
                </div>
              )}

              {/* slide preview */}
              <div className="card p-5">
                <div className="flex items-center justify-between pb-2 border-b border-slate-100">
                  <div className="flex items-center gap-2">
                    <Monitor size={16} className="text-brand-500" />
                    <h3 className="font-semibold text-slate-800 text-sm">
                      Slide Preview
                    </h3>
                  </div>
                  {slides.length > 0 && (
                    <span className="text-xs font-semibold tabular-nums px-2.5 py-1 rounded-lg bg-brand-50 text-brand-600 border border-brand-100 flex items-center gap-1">
                      <Layers size={11} />
                      {currentSlide + 1} / {slides.length}
                    </span>
                  )}
                </div>

                <div className="mt-4">
                  {previewLoad ? (
                    <div
                      className="w-full rounded-xl bg-slate-100 animate-pulse"
                      style={{ aspectRatio: "16/9" }}
                    />
                  ) : slides[currentSlide] ? (
                    <button
                      type="button"
                      onClick={() => setPreviewOpen(true)}
                      className="relative w-full text-left group"
                      title="Click to open the large preview"
                    >
                      <ExtractSlideCanvas
                        slide={slides[currentSlide]}
                        settings={settings}
                      />
                      <span className="absolute inset-0 flex items-center justify-center rounded-xl bg-black/0 group-hover:bg-black/30 transition">
                        <span className="opacity-0 group-hover:opacity-100 transition px-3 py-1.5 rounded-lg bg-white text-slate-800 text-xs font-semibold flex items-center gap-1.5 shadow-lg">
                          <Maximize2 size={12} /> Large Preview
                        </span>
                      </span>
                    </button>
                  ) : (
                    <div
                      className="w-full rounded-xl border-2 border-dashed border-slate-200 flex flex-col items-center justify-center gap-2 text-slate-400 bg-slate-50"
                      style={{ aspectRatio: "16/9" }}
                    >
                      <Monitor size={36} className="opacity-20" />
                      <p className="text-sm font-medium">
                        Upload a file to build slides
                      </p>
                      <p className="text-xs opacity-60">
                        Extracted lyrics preview here instantly
                      </p>
                    </div>
                  )}
                </div>

                {slides.length > 1 && (
                  <div className="flex items-center gap-3 mt-4">
                    <button
                      onClick={() => setCurrentSlide((c) => Math.max(0, c - 1))}
                      disabled={currentSlide === 0}
                      className="btn-secondary !p-2"
                    >
                      <ChevronLeft size={16} />
                    </button>
                    <div className="flex-1 flex items-center justify-center gap-1 overflow-x-auto py-1 no-scrollbar">
                      {Array.from(
                        { length: Math.min(slides.length, 20) },
                        (_, i) => {
                          const idx =
                            slides.length > 20
                              ? Math.round((i / 19) * (slides.length - 1))
                              : i;
                          const active = currentSlide === idx;
                          return (
                            <button
                              key={i}
                              onClick={() => setCurrentSlide(idx)}
                              className={`rounded-full shrink-0 transition-all duration-150 ${
                                active
                                  ? "w-5 h-2.5 bg-brand-500"
                                  : "w-2 h-2 bg-slate-200 hover:bg-slate-400"
                              }`}
                            />
                          );
                        },
                      )}
                    </div>
                    <button
                      onClick={() =>
                        setCurrentSlide((c) => Math.min(slides.length - 1, c + 1))
                      }
                      disabled={currentSlide === slides.length - 1}
                      className="btn-secondary !p-2"
                    >
                      <ChevronRight size={16} />
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* ── settings drawer ─────────────────────────────────── */}
        <div
          className={`bg-white border-l border-slate-200 transition-all duration-300 overflow-hidden flex-shrink-0 z-20 ${
            settingsOpen ? "w-80" : "w-0"
          }`}
        >
          {settingsOpen && (
            <div className="w-80 h-full overflow-y-auto">
              <div className="sticky top-0 bg-white border-b border-slate-100 px-5 py-4 z-10 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Sliders size={16} className="text-brand-500" />
                  <span className="font-semibold text-sm text-slate-800">
                    Song Slide Settings
                  </span>
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

      {/* ── Large slide preview modal (75% of the screen) ───────── */}
      {previewOpen && slides.length > 0 && (
        <div
          className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4"
          onClick={() => setPreviewOpen(false)}
        >
          <div
            className="bg-white rounded-2xl shadow-2xl flex flex-col overflow-hidden"
            style={{ width: "75vw", height: "75vh" }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between px-5 py-3 border-b border-slate-200 flex-shrink-0">
              <div className="flex items-center gap-2">
                <Monitor size={16} className="text-brand-500" />
                <span className="font-semibold text-sm text-slate-800">
                  Slide Preview
                </span>
                <span className="text-xs font-semibold tabular-nums px-2 py-0.5 rounded-lg bg-brand-50 text-brand-600 border border-brand-100">
                  {currentSlide + 1} / {slides.length}
                </span>
              </div>
              <button
                onClick={() => setPreviewOpen(false)}
                className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100 transition"
                title="Close"
              >
                <X size={18} />
              </button>
            </div>

            <div className="flex-1 min-h-0 bg-slate-100 flex items-center justify-center p-6 overflow-hidden">
              <div className="max-w-full max-h-full w-full">
                <ExtractSlideCanvas
                  slide={slides[currentSlide]}
                  settings={settings}
                />
              </div>
            </div>

            <div className="flex items-center justify-center gap-4 px-5 py-3 border-t border-slate-200 flex-shrink-0 bg-white">
              <button
                onClick={() => setCurrentSlide((c) => Math.max(0, c - 1))}
                disabled={currentSlide === 0}
                className="btn-secondary !p-2"
              >
                <ChevronLeft size={16} />
              </button>
              <div className="flex items-center gap-1 overflow-x-auto py-1 no-scrollbar max-w-md">
                {Array.from({ length: Math.min(slides.length, 24) }, (_, i) => {
                  const idx =
                    slides.length > 24
                      ? Math.round((i / 23) * (slides.length - 1))
                      : i;
                  const active = currentSlide === idx;
                  return (
                    <button
                      key={i}
                      onClick={() => setCurrentSlide(idx)}
                      className={`rounded-full shrink-0 transition-all duration-150 ${
                        active
                          ? "w-5 h-2.5 bg-brand-500"
                          : "w-2 h-2 bg-slate-200 hover:bg-slate-400"
                      }`}
                    />
                  );
                })}
              </div>
              <button
                onClick={() =>
                  setCurrentSlide((c) => Math.min(slides.length - 1, c + 1))
                }
                disabled={currentSlide === slides.length - 1}
                className="btn-secondary !p-2"
              >
                <ChevronRight size={16} />
              </button>
            </div>
          </div>
        </div>
      )}

      {toast && (
        <Toast
          message={toast.message}
          type={toast.type}
          onClose={() => setToast(null)}
        />
      )}
    </div>
  );
}
