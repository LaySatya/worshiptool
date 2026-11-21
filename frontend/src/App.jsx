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
  Menu,
  FileMusic,
  ExternalLink,
  Calculator,
  ScanText,
  SlidersHorizontal,
  Church,
  School,
  Library,
  Clock,
  BookMarked,
} from "lucide-react";
import SettingsPanel from "./components/SettingsPanel.jsx";
import PreviewPanel from "./components/PreviewPanel.jsx";
import BibleReader from "./components/BibleReader.jsx";
import SongSlidePage from "./SongSlidePage.jsx";
import MergeSlidesPage from "./MergeSlidesPage.jsx";
import MelodySlidePage from "./MelodySlidePage.jsx";
import MathematicPage from "./MathematicPage.jsx";
import SongExtractPage from "./SongExtractPage.jsx";
import SlideEditorPage from "./SlideEditorPage.jsx";
import BooksLibrary from "./components/books/BooksLibrary.jsx";
import BookCoverPage from "./components/cover/BookCoverPage.jsx";
import { fetchPreview, generatePptx } from "./api.js";
import { DEFAULT_SETTINGS } from "./constants.js";
import {
  CHURCH_PAGES,
  MAINTAINING_PAGES,
  buildHash,
  parseHash,
} from "./navigation.js";

// Sidebar entries. Everything worship-related lives under "Church";
// teaching tools live under "School"; the textbooks have their own "Books" page.
const CHURCH_NAV = [
  { id: "bible", label: "Bible Slides", icon: BookOpen },
  { id: "song", label: "Song Slides", icon: Music },
  { id: "cover", label: "Book Cover", icon: BookMarked },
  { id: "melody", label: "Melody Slide", icon: FileMusic },
  { id: "extract", label: "Song File Extract", icon: ScanText },
  { id: "editor", label: "Slide Editor", icon: SlidersHorizontal },
  { id: "merge", label: "Merge Slides", icon: Merge },
];

const SCHOOL_NAV = [
  { id: "math", label: "Mathematic", icon: Calculator },
];

const PAGE_META = {
  bible: { title: "Bible Slides", icon: BookOpen },
  melody: { title: "Melody Slide", icon: FileMusic },
  song: { title: "Song Slides", icon: Music },
  extract: { title: "Song File Extract", icon: ScanText },
  editor: { title: "Slide Editor", icon: SlidersHorizontal },
  merge: { title: "Merge Slides", icon: Merge },
  cover: { title: "Book Cover", icon: BookMarked },
  math: { title: "Mathematic", icon: Calculator },
  books: { title: "Books", icon: Library },
};

const pageOf = (id) => PAGE_META[id] || PAGE_META.bible;

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
  // The whole navigation lives in the URL hash, so a refresh (or a bookmark)
  // restores exactly the page, tab and book the user was on.
  const [nav, setNav] = useState(() => parseHash(window.location.hash));
  const { page } = nav; // 'bible' | 'song' | 'cover' | 'melody' | 'extract' | 'editor' | 'merge' | 'math' | 'books'
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
  const [bibleModalOpen, setBibleModalOpen] = useState(false);

  const lastSelectedTextRef = useRef(null);

  /** Change navigation: update state, restore it on refresh via the hash. */
  const navigate = useCallback((patch) => {
    setNav((prev) => {
      const next = { ...prev, ...patch };
      if (patch.books) next.books = { ...prev.books, ...patch.books };
      const hash = buildHash(next);
      if (window.location.hash !== hash) {
        window.history.replaceState(null, "", hash);
      }
      return next;
    });
  }, []);

  const goPage = useCallback(
    (id) => {
      navigate({ page: id });
      setSidebarOpen(false);
    },
    [navigate],
  );

  // Keep the app in sync with manual hash edits and the browser back button.
  useEffect(() => {
    const onHashChange = () => setNav(parseHash(window.location.hash));
    window.addEventListener("hashchange", onHashChange);
    return () => window.removeEventListener("hashchange", onHashChange);
  }, []);

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
        className={`fixed top-0 bottom-0 left-0 w-80 max-w-[85vw] bg-white border-r border-slate-200 flex flex-col z-50 shadow-2xl transition-transform duration-300 ease-in-out
          ${sidebarOpen ? "translate-x-0" : "-translate-x-full"}`}
      >
        {/* Sidebar Header */}
        <div className="h-16 flex items-center justify-between px-5 border-b border-slate-100">
          <div className="flex items-center gap-3 min-w-0">
            <div className="w-9 h-9 rounded-lg bg-brand-500 flex items-center justify-center shadow-xs flex-shrink-0">
              <BookOpen size={16} className="text-white" />
            </div>
            <div className="min-w-0">
              <span className="block font-extrabold text-slate-800 tracking-tight text-lg leading-tight">
                SYASKYTOOL
              </span>
              <span className="block text-[10px] font-medium text-slate-400 leading-tight truncate">
                Church &amp; School suite
              </span>
            </div>
          </div>
          <button
            onClick={() => setSidebarOpen(false)}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition"
            title="Close Sidebar"
          >
            <X size={18} />
          </button>
        </div>

        {/* Navigation — grouped: Church · School · Books */}
        <nav className="flex-1 p-3 overflow-y-auto space-y-4">
          {/* ── Church group ── */}
          <div>
            <div className="flex items-center gap-2 px-2 pb-1.5">
              <Church size={13} className="text-slate-400" />
              <span className="text-[11px] font-bold uppercase tracking-wider text-slate-400">
                Church
              </span>
            </div>
            <div className="space-y-1">
              {CHURCH_NAV.map(({ id, label, icon: Icon }) => {
                const active = page === id;
                const pending = MAINTAINING_PAGES.includes(id);
                return (
                  <button
                    key={id}
                    type="button"
                    onClick={() => goPage(id)}
                    className={`w-full flex items-center gap-2.5 pl-2.5 pr-2 py-2.5 rounded-xl text-[13px] font-semibold transition-all
                      ${
                        active
                          ? "bg-brand-50 text-brand-600 border border-brand-200"
                          : "text-slate-600 border border-transparent hover:bg-slate-50"
                      }`}
                  >
                    <Icon size={17} className={`flex-shrink-0 ${active ? "text-brand-500" : "text-slate-400"}`} />
                    <span className="truncate text-left">{label}</span>
                    {pending && (
                      <span
                        className="ml-auto inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded-md bg-amber-50 text-amber-700 border border-amber-200 text-[9px] font-bold uppercase tracking-tight whitespace-nowrap flex-shrink-0"
                        title="This tool is still being maintained"
                      >
                        <Clock size={9} />
                        Maintaining
                      </span>
                    )}
                    {active && !pending && (
                      <ChevronRight size={15} className="ml-auto text-brand-500 flex-shrink-0" />
                    )}
                  </button>
                );
              })}
            </div>
          </div>

          {/* ── School group (Mathematic) ── */}
          <div>
            <div className="flex items-center gap-2 px-2 pb-1.5">
              <School size={13} className="text-slate-400" />
              <span className="text-[11px] font-bold uppercase tracking-wider text-slate-400">
                School
              </span>
            </div>
            <div className="space-y-1">
              {SCHOOL_NAV.map(({ id, label, icon: Icon }) => {
                const active = page === id;
                return (
                  <button
                    key={id}
                    type="button"
                    onClick={() => goPage(id)}
                    className={`w-full flex items-center gap-3 pl-3 pr-2.5 py-2.5 rounded-xl text-sm font-semibold transition-all
                      ${
                        active
                          ? "bg-brand-50 text-brand-600 border border-brand-200"
                          : "text-slate-600 border border-transparent hover:bg-slate-50"
                      }`}
                  >
                    <Icon size={17} className={active ? "text-brand-500" : "text-slate-400"} />
                    <span className="truncate">{label}</span>
                    {active && <ChevronRight size={15} className="ml-auto text-brand-500" />}
                  </button>
                );
              })}
            </div>
          </div>

          {/* ── Books: grade + subject textbooks ── */}
          <div>
            <div className="flex items-center gap-2 px-2 pb-1.5">
              <Library size={13} className="text-slate-400" />
              <span className="text-[11px] font-bold uppercase tracking-wider text-slate-400">
                Books
              </span>
            </div>
            <button
              type="button"
              onClick={() => {
                navigate({ page: "books" });
                setSidebarOpen(false);
              }}
              className={`w-full flex items-center gap-3 pl-3 pr-2.5 py-2.5 rounded-xl text-sm font-semibold transition-all
                ${
                  page === "books"
                    ? "bg-brand-50 text-brand-600 border border-brand-200"
                    : "text-slate-600 border border-transparent hover:bg-slate-50"
                }`}
            >
              <Library size={17} className={page === "books" ? "text-brand-500" : "text-slate-400"} />
              <div className="text-left min-w-0">
                <span className="block truncate">Textbook Library</span>
                <span className="block text-[10px] font-medium text-slate-400 truncate">
                  Grade 4–6 · Khmer, Math, English
                </span>
              </div>
              {page === "books" && (
                <ChevronRight size={15} className="ml-auto text-brand-500 flex-shrink-0" />
              )}
            </button>
          </div>
        </nav>

        {/* Footer info */}
        <div className="p-4 border-t border-slate-100 bg-slate-50/50">
          <p className="text-[11px] font-semibold text-slate-500">SYASKYTOOL</p>
          <p className="text-[10px] font-medium text-slate-400">
            Worship presentation &amp; teaching suite
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
            <div className="flex items-center gap-2 min-w-0">
              <div className="w-8 h-8 rounded-lg bg-brand-500 flex items-center justify-center shadow-xs flex-shrink-0">
                {(() => {
                  const Icon = pageOf(page).icon;
                  return <Icon size={16} className="text-white" />;
                })()}
              </div>
              <div className="min-w-0">
                <p className="text-[10px] font-bold uppercase tracking-wider text-slate-400 leading-none hidden sm:block">
                  {CHURCH_PAGES.includes(page)
                    ? "SYASKYTOOL · Church"
                    : page === "books"
                      ? "SYASKYTOOL · Books"
                      : "SYASKYTOOL · School"}
                </p>
                <h1 className="text-base sm:text-lg font-bold text-slate-800 tracking-tight truncate leading-tight">
                  {pageOf(page).title}
                </h1>
              </div>
              {MAINTAINING_PAGES.includes(page) && (
                <span className="hidden md:inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-amber-50 text-amber-700 border border-amber-200 text-[10px] font-bold uppercase tracking-wide">
                  <Clock size={10} />
                  Maintenance
                </span>
              )}
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

        {/* Page Content — keyed by page so it re-animates on every navigation click */}
        <div key={page} className="page-enter flex-1 flex flex-col min-h-0 overflow-y-auto">
        {page === "books" ? (
          <BooksLibrary
            selection={nav.books}
            onSelect={(books) => navigate({ page: "books", books })}
          />
        ) : page === "math" ? (
          <MathematicPage
            tab={nav.mathTab}
            onTabChange={(mathTab) => navigate({ mathTab })}
          />
        ) : page === "melody" ? (
          <MelodySlidePage />
        ) : page === "song" ? (
          <SongSlidePage />
        ) : page === "cover" ? (
          <BookCoverPage />
        ) : page === "extract" ? (
          <SongExtractPage />
        ) : page === "editor" ? (
          <SlideEditorPage />
        ) : page === "merge" ? (
          <MergeSlidesPage />
        ) : (
          /* ── Bible Slides Content ── */
          <div className="flex-1 overflow-hidden flex">
            {/* Bible Slides Workspace (Side-by-side row: Bible Text & Live Preview) */}
            <div className="flex-1 overflow-y-auto p-4 flex flex-col gap-4">
              {/* Header bar */}
              <div className="flex items-center justify-between bg-white px-5 py-3 rounded-xl border border-slate-200/80 shadow-xs flex-shrink-0">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-lg bg-brand-50 text-brand-600 flex items-center justify-center font-bold">
                    <FileText size={16} />
                  </div>
                  <div>
                    <h2 className="font-bold text-slate-800 text-sm">Bible Slides Editor &amp; Live Preview</h2>
                    <p className="text-xs text-slate-400">Edit text, adjust slide options, or copy verses directly from online Bible</p>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setBibleModalOpen(true)}
                    className="btn-secondary !py-1.5 !px-3 text-xs flex items-center gap-1.5 text-brand-700 bg-brand-50 hover:bg-brand-100 border-brand-200 font-semibold transition"
                    title="Open Bible Reader popup to select and copy verses"
                  >
                    <BookOpen size={14} className="text-brand-600" />
                    <span>Copy Verse from Website</span>
                  </button>
                </div>
              </div>

              {/* Side-by-side row: Bible Text (left) and Live Preview (right) */}
              <div className="flex-1 grid grid-cols-1 lg:grid-cols-2 gap-5 min-h-[520px]">
                {/* Left card: Bible Text Input */}
                <div className="card p-5 flex flex-col h-full">
                  <div className="flex items-center justify-between flex-shrink-0">
                    <div className="flex items-center gap-2">
                      <FileText size={16} className="text-brand-500" />
                      <h2 className="font-semibold text-slate-800 text-sm">
                        Bible Text
                      </h2>
                    </div>
                    <div className="flex items-center gap-3">
                      {/* Copy Verse from Website Button */}
                      <button
                        type="button"
                        onClick={() => setBibleModalOpen(true)}
                        className="px-2.5 py-1 rounded-lg bg-brand-50 hover:bg-brand-100 text-brand-700 border border-brand-200 text-xs font-semibold flex items-center gap-1.5 transition shadow-2xs hover:shadow-xs"
                        title="Open Bible Reader popup to select and copy verses directly"
                      >
                        <BookOpen size={13} className="text-brand-600" />
                        <span>Copy Verse from Website</span>
                      </button>

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
      </div>

      {/* Popup Iframe Modal for copying Bible Verse */}
      {bibleModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-6 bg-black/60 backdrop-blur-xs animate-fade-in">
          <div
            className="w-full max-w-5xl h-[88vh] bg-white rounded-2xl shadow-2xl flex flex-col overflow-hidden border border-slate-200"
            onClick={(e) => e.stopPropagation()}
          >
            {/* Modal Header */}
            <div className="h-12 bg-white border-b border-slate-200 px-4 flex items-center justify-between flex-shrink-0">
              <div className="flex items-center gap-2">
                <div className="w-7 h-7 rounded-lg bg-brand-50 text-brand-600 flex items-center justify-center">
                  <BookOpen size={15} />
                </div>
                <div>
                  <h3 className="text-xs sm:text-sm font-bold text-slate-800 leading-tight">
                    Bible Reader — Select &amp; Copy Verse
                  </h3>
                  <p className="text-[11px] text-slate-400 hidden sm:block leading-none">
                    Highlight any verse: text is automatically captured into your slides
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setBibleModalOpen(false)}
                  className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100 transition"
                  title="Close Modal"
                >
                  <X size={18} />
                </button>
              </div>
            </div>

            {/* Modal Body: Embedded Bible Reader iframe with proxy */}
            <div className="flex-1 min-h-0 p-3 bg-slate-50 flex flex-col">
              <BibleReader
                onSelection={handleBibleSelection}
                onDeselection={handleBibleDeselection}
                onClose={() => setBibleModalOpen(false)}
              />
            </div>

            {/* Modal Footer */}
            <div className="p-3 border-t border-slate-200 bg-white flex items-center justify-between text-xs text-slate-600 flex-shrink-0">
              <span className="flex items-center gap-1.5 text-emerald-700 font-medium text-[11px] sm:text-xs">
                <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                Select any text to copy directly to your slides
              </span>
              <button
                type="button"
                onClick={() => setBibleModalOpen(false)}
                className="btn-primary !py-1.5 !px-4 text-xs"
              >
                Done
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
