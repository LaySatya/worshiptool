import { useState, useEffect, useCallback, useRef } from "react";
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
  ChevronRight,
  Settings,
  ArrowDown,
  ArrowUp,
  Menu,
} from "lucide-react";
import SettingsPanel from "./components/SettingsPanel.jsx";
import PreviewPanel from "./components/PreviewPanel.jsx";
import BibleReader from "./components/BibleReader.jsx";
import SongSlidePage from "./SongSlidePage.jsx";
import MergeSlidesPage from "./MergeSlidesPage.jsx";
import { fetchPreview, generatePptx } from "./api.js";
import { DEFAULT_SETTINGS } from "./constants.js";

function useDebounce(value, delay) {
  const [debounced, setDebounced] = useState(value);
  useEffect(() => {
    const t = setTimeout(() => setDebounced(value), delay);
    return () => clearTimeout(t);
  }, [value, delay]);
  return debounced;
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
        className="ml-1 text-slate-400 hover:text-slate-600 transition"
      >
        <X size={13} />
      </button>
    </div>
  );
}

const PLACEHOLDER = `Example — paste Khmer numbered verses:

12 ដំក់ដិណាដាតិទាំទូរយេ្រីង គំជាពួកស្លោដែលមានក្រាប់គាម
13 នោះក័មានល្ហាចមានព្រីករេ្យីង ជាថ្ងៃទីរំ

Or English references:

John 3:16
For God so loved the world that he gave his one and only Son.

Romans 8:28
And we know that in all things God works for the good of those who love him.`;

export default function App() {
  const [page, setPage] = useState("bible"); // 'bible' | 'song' | 'merge'
  const [rawText, setRawText] = useState("");
  const [reference, setReference] = useState("");
  const [perSlide, setPerSlide] = useState(1);
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [slides, setSlides] = useState([]);
  const [currentSlide, setCurrentSlide] = useState(0);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [toast, setToast] = useState(null);
  const [detected, setDetected] = useState(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(false);

  const bottomSectionRef = useRef(null);
  const readerSectionRef = useRef(null);
  const lastSelectedTextRef = useRef(null);

  const dRawText = useDebounce(rawText, 350);
  const dReference = useDebounce(reference, 300);
  const dPerSlide = useDebounce(perSlide, 200);
  const dSettings = useDebounce(settings, 200);

  useEffect(() => {
    if (!dRawText.trim()) {
      setSlides([]);
      setCurrentSlide(0);
      setDetected(null);
      return;
    }
    let cancelled = false;
    setPreviewLoading(true);
    fetchPreview({
      raw_text: dRawText,
      per_slide: dPerSlide,
      settings: dSettings,
      reference: dReference,
    })
      .then((data) => {
        if (cancelled) return;
        setSlides(data.slides);
        setCurrentSlide(0);
        if (data.slides.some((s) => s.verse_num)) setDetected("khmer");
        else if (data.slides.some((s) => s.ref)) setDetected("english");
        else setDetected(null);
      })
      .catch(() => {
        if (!cancelled)
          setToast({
            message: "Preview error — is the backend running?",
            type: "error",
          });
      })
      .finally(() => {
        if (!cancelled) setPreviewLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [dRawText, dReference, dPerSlide, dSettings]);

  const handleBibleSelection = useCallback((text) => {
    if (!text) return;

    // Detect if exact same verse was selected twice without an intervening deselection
    if (lastSelectedTextRef.current === text) {
      return;
    }
    lastSelectedTextRef.current = text;

    setRawText((prev) => (prev ? prev + "\n" + text : text));

    // Do NOT scroll down automatically — keep user focused in the reader
    setToast({
      message: "Verse added to slides!",
      type: "success",
    });
  }, []);

  const handleBibleDeselection = useCallback(() => {
    // When the user unselects, reset the tracker so selecting the same verse again will work
    lastSelectedTextRef.current = null;
  }, []);

  const scrollToBottom = () => {
    bottomSectionRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const scrollToTop = () => {
    readerSectionRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const handleGenerate = async () => {
    if (!rawText.trim()) {
      setToast({
        message: "Please paste some Bible verses first.",
        type: "error",
      });
      return;
    }
    setGenerating(true);
    try {
      const blob = await generatePptx({
        raw_text: rawText,
        per_slide: perSlide,
        settings,
        reference,
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "bible_slides.pptx";
      a.click();
      URL.revokeObjectURL(url);
      setToast({
        message: `${slides.length} slides downloaded!`,
        type: "success",
      });
    } catch {
      setToast({
        message: "Generation failed — check backend.",
        type: "error",
      });
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="bg-slate-50 flex h-screen overflow-hidden relative">
      {/* Responsive Backdrop Overlay */}
      {sidebarOpen && (
        <div
          className="fixed inset-0 bg-black/40 backdrop-blur-xs z-40 transition-opacity"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      {/* Left Sidebar Drawer (Default Closed, slides out when opened) */}
      <aside
        className={`fixed top-0 bottom-0 left-0 w-72 max-w-[85vw] bg-white border-r border-slate-200 flex flex-col z-50 shadow-2xl transition-transform duration-300 ease-in-out
          ${sidebarOpen ? "translate-x-0" : "-translate-x-full"}`}
      >
        {/* Sidebar Header */}
        <div className="h-16 flex items-center justify-between px-5 border-b border-slate-100">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-lg bg-brand-500 flex items-center justify-center shadow-xs">
              <BookOpen size={16} className="text-white" />
            </div>
            <span className="font-bold text-slate-800 tracking-tight text-lg">
              ChurchTool
            </span>
          </div>
          <button
            onClick={() => setSidebarOpen(false)}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition"
            title="Close Sidebar"
          >
            <X size={18} />
          </button>
        </div>

        {/* Navigation */}
        <nav className="flex-1 p-4 space-y-1.5 overflow-y-auto">
          <button
            onClick={() => {
              setPage("bible");
              setSidebarOpen(false);
            }}
            className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-semibold transition-all
              ${
                page === "bible"
                  ? "bg-brand-50 text-brand-600 border border-brand-200"
                  : "text-slate-600 hover:bg-slate-50"
              }`}
          >
            <BookOpen size={18} />
            <span>Bible Slides</span>
            {page === "bible" && <ChevronRight size={16} className="ml-auto" />}
          </button>

          <button
            onClick={() => {
              setPage("song");
              setSidebarOpen(false);
            }}
            className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-semibold transition-all
              ${
                page === "song"
                  ? "bg-brand-50 text-brand-600 border border-brand-200"
                  : "text-slate-600 hover:bg-slate-50"
              }`}
          >
            <Music size={18} />
            <span>Song Slides</span>
            {page === "song" && <ChevronRight size={16} className="ml-auto" />}
          </button>

          <button
            onClick={() => {
              setPage("merge");
              setSidebarOpen(false);
            }}
            className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-semibold transition-all
              ${
                page === "merge"
                  ? "bg-brand-50 text-brand-600 border border-brand-200"
                  : "text-slate-600 hover:bg-slate-50"
              }`}
          >
            <Merge size={18} />
            <span>Merge Slides</span>
            {page === "merge" && <ChevronRight size={16} className="ml-auto" />}
          </button>
        </nav>

        {/* Footer info */}
        <div className="p-4 border-t border-slate-100 bg-slate-50/50">
          <p className="text-[11px] font-medium text-slate-400">
            ChurchTool — Worship Presentation Suite
          </p>
        </div>
      </aside>

      {/* Main Content */}
      <div className="flex-1 flex flex-col overflow-hidden w-full">
        {/* Top Header */}
        <header className="h-16 bg-white border-b border-slate-100 flex items-center justify-between px-4 sm:px-6 flex-shrink-0 z-30">
          <div className="flex items-center gap-3">
            <button
              onClick={() => setSidebarOpen((prev) => !prev)}
              className="p-2 -ml-1 rounded-xl text-slate-600 hover:bg-slate-100 hover:text-slate-900 transition flex items-center justify-center"
              title="Open Navigation Menu"
            >
              <Menu size={20} />
            </button>
            <div className="flex items-center gap-2">
              <div className="w-8 h-8 rounded-lg bg-brand-500 flex items-center justify-center shadow-xs">
                {page === "bible" ? (
                  <BookOpen size={16} className="text-white" />
                ) : page === "song" ? (
                  <Music size={16} className="text-white" />
                ) : (
                  <Merge size={16} className="text-white" />
                )}
              </div>
              <h1 className="text-base sm:text-lg font-bold text-slate-800 tracking-tight">
                {page === "bible"
                  ? "Bible Slides"
                  : page === "song"
                    ? "Song Slides"
                    : "Merge Slides"}
              </h1>
            </div>
            {page === "bible" && detected && (
              <span className="hidden md:inline-flex text-xs text-slate-500 ml-2">
                {detected === "khmer"
                  ? "🇰🇭 Khmer verses detected"
                  : "📖 English references detected"}
              </span>
            )}
          </div>

          <div className="flex items-center gap-3">
            {page === "bible" && (
              <>
                <button
                  onClick={() => setSettingsOpen(!settingsOpen)}
                  className={`btn-secondary ${settingsOpen ? "bg-brand-50 text-brand-600 border-brand-200" : ""}`}
                >
                  <Settings size={16} />
                  {settingsOpen ? "Hide Settings" : "Slide Settings"}
                </button>
                <button
                  onClick={handleGenerate}
                  disabled={generating || !rawText.trim()}
                  className="btn-primary"
                >
                  {generating ? (
                    <>
                      <Loader2 size={14} className="animate-spin" />
                      Generating…
                    </>
                  ) : (
                    <>
                      <Download size={14} />
                      Download PPTX
                    </>
                  )}
                </button>
              </>
            )}
          </div>
        </header>

        {/* Page Content */}
        {page === "song" ? (
          <SongSlidePage />
        ) : page === "merge" ? (
          <MergeSlidesPage />
        ) : (
          /* ── Bible Slides Content ── */
          <div className="flex-1 overflow-hidden flex">
            {/* Bible Slides Workspace (scrolls vertically: full-screen reader on top, text & preview below) */}
            <div className="flex-1 overflow-y-auto scroll-smooth flex flex-col">
              {/* SECTION 1: Full screen height Bible Reader */}
              <section
                ref={readerSectionRef}
                className="h-[calc(100vh-4rem)] min-h-[640px] flex-shrink-0 p-4 pb-2 flex flex-col"
              >
                <div className="flex-1 min-h-0 flex flex-col">
                  <BibleReader
                    onSelection={handleBibleSelection}
                    onDeselection={handleBibleDeselection}
                  />
                </div>

                {/* Jump bar at the bottom of the iframe */}
                <div className="pt-2 flex items-center justify-between px-2 text-xs text-slate-500 flex-shrink-0">
                  <span className="flex items-center gap-1.5 text-slate-500">
                    <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                    Highlight any verse in the reader above to automatically add slides below
                  </span>
                  <button
                    type="button"
                    onClick={scrollToBottom}
                    className="flex items-center gap-1.5 px-3 py-1 rounded-lg bg-brand-50 hover:bg-brand-100 text-brand-700 font-medium transition shadow-xs"
                  >
                    <span>Go to Text &amp; Preview</span>
                    <ArrowDown size={14} />
                  </button>
                </div>
              </section>

              {/* SECTION 2: Bible Text & Live Preview in a row */}
              <section
                ref={bottomSectionRef}
                className="min-h-[calc(100vh-4rem)] p-4 pt-2 flex flex-col gap-4 flex-shrink-0 bg-slate-50 border-t border-slate-200/80"
              >
                {/* Header bar with Back to Reader button */}
                <div className="flex items-center justify-between bg-white px-5 py-3 rounded-xl border border-slate-200/80 shadow-xs flex-shrink-0">
                  <div className="flex items-center gap-3">
                    <div className="w-8 h-8 rounded-lg bg-brand-50 text-brand-600 flex items-center justify-center font-bold">
                      <FileText size={16} />
                    </div>
                    <div>
                      <h2 className="font-bold text-slate-800 text-sm">Bible Slides Editor &amp; Live Preview</h2>
                      <p className="text-xs text-slate-400">Edit captured text, adjust slide options, or live preview slides</p>
                    </div>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={scrollToTop}
                      className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-100 text-slate-600 text-xs font-semibold transition"
                    >
                      <ArrowUp size={13} />
                      <span>Back to Bible Reader</span>
                    </button>
                  </div>
                </div>

                {/* Side-by-side row: Bible Text (left) and Live Preview (right) */}
                <div className="flex-1 grid grid-cols-1 lg:grid-cols-2 gap-5 min-h-[520px]">
                  {/* Left card: Bible Text Input */}
                  <div className="card p-5 flex flex-col h-full">
                    <div className="flex items-center justify-between flex-shrink-0">
                      <h2 className="font-semibold text-slate-800 text-sm flex items-center gap-2">
                        <FileText size={16} className="text-brand-500" />
                        Bible Text
                      </h2>
                      <div className="flex items-center gap-3">
                        <div className="flex items-center gap-1.5 text-xs text-slate-500">
                          <span>Per slide:</span>
                          <select
                            className="px-2 py-0.5 rounded border border-slate-200 bg-white text-xs font-medium"
                            value={perSlide}
                            onChange={(e) =>
                              setPerSlide(parseInt(e.target.value, 10))
                            }
                          >
                            {[1, 2, 3, 4, 5, 6].map((n) => (
                              <option key={n} value={n}>
                                {n}
                              </option>
                            ))}
                          </select>
                        </div>
                        {rawText && (
                          <button
                            onClick={() => setRawText("")}
                            className="text-xs text-slate-400 hover:text-red-500 transition font-medium"
                          >
                            Clear
                          </button>
                        )}
                      </div>
                    </div>
                    <textarea
                      className="input resize-none text-[13px] leading-relaxed flex-1 mt-3 min-h-[280px]"
                      placeholder={PLACEHOLDER}
                      value={rawText}
                      onChange={(e) => setRawText(e.target.value)}
                    />
                  </div>

                  {/* Right card: Live Preview */}
                  <div className="card p-5 flex flex-col h-full overflow-hidden">
                    <PreviewPanel
                      slides={slides}
                      settings={settings}
                      current={currentSlide}
                      onNavigate={setCurrentSlide}
                      isLoading={previewLoading}
                    />
                  </div>
                </div>
              </section>
            </div>

            {/* RIGHT: Collapsible Settings Panel */}
            <div
              className={`bg-white border-l border-slate-200 transition-all duration-300 overflow-hidden flex-shrink-0
              ${settingsOpen ? "w-80" : "w-0"}`}
            >
              {settingsOpen && (
                <div className="w-80 h-full overflow-y-auto">
                  <div className="sticky top-0 bg-white border-b border-slate-100 px-5 py-4 z-10">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <Sliders size={16} className="text-brand-500" />
                        <span className="font-semibold text-sm text-slate-800">
                          Slide Settings
                        </span>
                      </div>
                      <button
                        onClick={() => setSettingsOpen(false)}
                        className="text-slate-400 hover:text-slate-600 transition"
                      >
                        <X size={18} />
                      </button>
                    </div>
                  </div>
                  <div className="px-5 py-5">
                    <SettingsPanel settings={settings} onChange={setSettings} />
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </div>

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
