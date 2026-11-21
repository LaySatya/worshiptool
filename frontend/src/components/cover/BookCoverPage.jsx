import { useEffect, useMemo, useState } from "react";
import { BookMarked, Printer, RotateCcw } from "lucide-react";

/**
 * Book Cover tool — a printable cover for teachers.
 *
 * Editable:   name, book title, grade (ថ្នាក់), teacher name, years for the logo.
 * Fixed:      school name, the word "សៀវភៅ", and the border frame.
 * Print:      A4 / A3 / A5 via @page size rules injected at print time.
 */

const STORAGE_KEY = "worshiptool:cover";

const DEFAULT_COVER = {
  name: "",
  bookTitle: "",
  grade: "ថ្នាក់ទី​៤",
  teacherName: "",
  years: "",
};

const GRADE_OPTIONS = [
  "ថ្នាក់ទី​១",
  "ថ្នាក់ទី​២",
  "ថ្នាក់ទី​៣",
  "ថ្នាក់ទី​៤",
  "ថ្នាក់ទី​៥",
  "ថ្នាក់ទី​៦",
  "ថ្នាក់ទី​៧",
  "ថ្នាក់ទី​៨",
  "ថ្នាក់ទី​៩",
  "ថ្នាក់ទី​១០",
  "ថ្នាក់ទី​១១",
  "ថ្នាក់ទី​១២",
];

const PRINT_SIZES = [
  { id: "A4", label: "A4", mm: "210 × 297 mm" },
  { id: "A3", label: "A3", mm: "297 × 420 mm" },
  { id: "A5", label: "A5", mm: "148 × 210 mm" },
];

// Page dimensions in mm (portrait) — used to scale the on-screen preview.
const PAGE_MM = { A4: [210, 297], A3: [297, 420], A5: [148, 210] };

function loadCover() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return { ...DEFAULT_COVER, ...JSON.parse(raw) };
  } catch {
    /* ignore */
  }
  return DEFAULT_COVER;
}

/** The cover artwork. Fixed pieces (border, សៀវភៅ, school name) are locked. */
function CoverArt({ cover, fontSize }) {
  const s = (mm) => `calc(${mm} * ${fontSize})`; // mm → preview px scale
  return (
    <div
      className="cover-art relative text-slate-900 overflow-hidden"
      style={{ width: "100%", aspectRatio: "210 / 297", background: "#fdf6e3" }}
    >
      {/* Thin blue rounded outer border — never changes */}
      <div
        className="absolute pointer-events-none rounded-lg"
        style={{
          inset: `calc(2.5 * ${fontSize})`,
          border: `calc(0.8 * ${fontSize}) solid #1e3a8a`,
        }}
      />
      {/* Red inner rounded border — never changes */}
      <div
        className="absolute pointer-events-none rounded-md"
        style={{
          inset: `calc(5.5 * ${fontSize})`,
          border: `calc(0.6 * ${fontSize}) solid #dc2626`,
        }}
      />

      {/* Header band — round logo left, green school name center, red years right */}
      <div
        className="relative flex items-center justify-between"
        style={{
          margin: `calc(9 * ${fontSize})`,
          padding: `calc(2.5 * ${fontSize}) calc(3.5 * ${fontSize})`,
          borderBottom: `calc(0.5 * ${fontSize}) solid #dc2626`,
        }}
      >
        <div
          className="cover-anim-pop flex items-center justify-center rounded-full bg-white border border-red-200 shadow-sm"
          style={{ width: s(14), height: s(14) }}
        >
          <BookMarked
            className="cover-anim-logo"
            style={{ width: s(8), height: s(8), color: "#dc2626" }}
          />
        </div>
        <div
          className="font-muol text-center"
          style={{ fontSize: s(6), color: "#15803d" }}
        >
          សាលាបឋមសិក្សា ស៊ាងស្កី
        </div>
        <div
          className="font-khmer text-right"
          style={{ fontSize: s(4.5), color: "#dc2626", lineHeight: 1.4 }}
        >
          {cover.years || "\u00A0"}
        </div>
      </div>

      <div
        className="relative h-full flex flex-col items-center"
        style={{ padding: `${s(4)} ${s(12)} ${s(10)}` }}
      >
        {/* សៀវភៅ — green bold heading, fixed, never changes */}
        <h1
          className="cover-anim font-muol text-center"
          style={{ fontSize: s(15), color: "#15803d", marginTop: s(4) }}
        >
          សៀវភៅ
        </h1>

        {/* Book title — editable, red accent */}
        <div
          className="cover-anim-d1 font-khmer font-bold text-center"
          style={{ fontSize: s(9), color: "#dc2626", marginTop: s(5) }}
        >
          {cover.bookTitle || "\u00A0"}
        </div>

        {/* Red divider accent */}
        <div
          className="cover-anim-d2"
          style={{
            width: s(40), height: s(0.6),
            background: "#dc2626", margin: `${s(3)} 0`,
          }}
        />

        <div style={{ flex: 1 }} />

        {/* Grade — editable, blue heading with red bullet */}
        <div
          className="cover-anim-d3 flex items-center gap-2"
          style={{ fontSize: s(7) }}
        >
          <span
            className="rounded-full"
            style={{ width: s(1.6), height: s(1.6), background: "#dc2626" }}
          />
          <span className="font-khmer font-bold" style={{ color: "#1e3a8a" }}>
            {cover.grade}
          </span>
        </div>

        {/* Teacher name — editable */}
        {cover.teacherName && (
          <div
            className="cover-anim-d4 flex items-center gap-2"
            style={{ fontSize: s(6), marginTop: s(3) }}
          >
            <span
              className="rounded-full"
              style={{ width: s(1.6), height: s(1.6), background: "#dc2626" }}
            />
            <span className="font-khmer" style={{ color: "#0f766e" }}>
              គ្រូបង្រៀន៖ {cover.teacherName}
            </span>
          </div>
        )}

        {/* Name — editable */}
        {cover.name && (
          <div
            className="cover-anim-d5 flex items-center gap-2"
            style={{ fontSize: s(6), marginTop: s(2) }}
          >
            <span
              className="rounded-full"
              style={{ width: s(1.6), height: s(1.6), background: "#dc2626" }}
            />
            <span className="font-khmer" style={{ color: "#0f766e" }}>
              ឈ្មោះ៖ {cover.name}
            </span>
          </div>
        )}

        <div style={{ flex: 1 }} />

        {/* Red bottom band accent */}
        <div
          className="cover-anim-d2 rounded-full"
          style={{
            width: s(24), height: s(1.2), background: "#dc2626",
            marginBottom: s(2),
          }}
        />
      </div>
    </div>
  );
}

export default function BookCoverPage() {
  const [cover, setCover] = useState(loadCover);
  const [printSize, setPrintSize] = useState("A4");

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(cover));
    } catch {
      /* ignore */
    }
  }, [cover]);

  const set = (key) => (e) =>
    setCover((c) => ({ ...c, [key]: e.target.value }));

  // Preview scale: 1 mm ≈ 3.2px on screen, so the A4 preview fits ~700px.
  const fontSize = useMemo(() => "3.2px", []);

  const print = () => {
    const [w, h] = PAGE_MM[printSize];
    const style = document.createElement("style");
    style.id = "cover-print-style";
    style.textContent = `
      @page { size: ${w}mm ${h}mm; margin: 0; }
      @media print {
        body * { visibility: hidden !important; }
        .cover-print-root, .cover-print-root * { visibility: visible !important; }
        .cover-print-root {
          position: absolute; left: 0; top: 0;
          width: ${w}mm; height: ${h}mm;
          margin: 0; box-shadow: none; border-radius: 0;
        }
      }
    `;
    document.getElementById("cover-print-style")?.remove();
    document.head.appendChild(style);
    const cleanup = () => {
      document.getElementById("cover-print-style")?.remove();
      window.removeEventListener("afterprint", cleanup);
    };
    window.addEventListener("afterprint", cleanup);
    window.print();
  };

  return (
    <div className="max-w-6xl mx-auto animate-fade-in">
      <div className="grid lg:grid-cols-[380px_1fr] gap-6">
        {/* Editor panel */}
        <div className="card p-5 space-y-4 h-fit">
          <div>
            <h2 className="text-sm font-semibold text-slate-800 flex items-center gap-2">
              <BookMarked size={16} className="text-brand-500" />
              Book Cover Editor
            </h2>
            <p className="text-xs text-slate-500 mt-1">
              Fill in the details. The school name, «សៀវភៅ» and the border are
              fixed.
            </p>
          </div>

          <div>
            <label className="label">ឈ្មោះ (Name)</label>
            <input
              className="input"
              value={cover.name}
              onChange={set("name")}
              placeholder="ឈ្មោះសិស្ស"
            />
          </div>
          <div>
            <label className="label">ចំណងជើងសៀវភៅ (Book title)</label>
            <input
              className="input"
              value={cover.bookTitle}
              onChange={set("bookTitle")}
              placeholder="ឈ្មោះសៀវភៅ"
            />
          </div>
          <div>
            <label className="label">ថ្នាក់ (Grade)</label>
            <select
              className="select"
              value={cover.grade}
              onChange={set("grade")}
            >
              {GRADE_OPTIONS.map((g) => (
                <option key={g} value={g}>
                  {g}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className="label">ឈ្មោះគ្រូបង្រៀន (Teacher)</label>
            <input
              className="input"
              value={cover.teacherName}
              onChange={set("teacherName")}
              placeholder="ឈ្មោះគ្រូ"
            />
          </div>
          <div>
            <label className="label">ឆ្នាំ (Years for logo)</label>
            <input
              className="input"
              value={cover.years}
              onChange={set("years")}
              placeholder="២០២៥ - ២០២៦"
            />
          </div>

          <div>
            <label className="label">Print size</label>
            <div className="flex gap-2">
              {PRINT_SIZES.map((p) => (
                <button
                  key={p.id}
                  type="button"
                  onClick={() => setPrintSize(p.id)}
                  className={`flex-1 px-3 py-2 rounded-xl text-sm font-semibold transition-all border ${
                    printSize === p.id
                      ? "bg-brand-50 text-brand-600 border-brand-200"
                      : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"
                  }`}
                >
                  <div>{p.label}</div>
                  <div className="text-[10px] font-normal text-slate-400">
                    {p.mm}
                  </div>
                </button>
              ))}
            </div>
          </div>

          <div className="flex gap-2 pt-1">
            <button className="btn-primary flex-1" onClick={print}>
              <Printer size={16} />
              Print {printSize}
            </button>
            <button
              className="btn-secondary"
              title="Reset fields"
              onClick={() => setCover(DEFAULT_COVER)}
            >
              <RotateCcw size={16} />
            </button>
          </div>
        </div>

        {/* Live preview */}
        <div className="card p-6">
          <div className="flex items-center justify-between mb-4">
            <h2 className="text-sm font-semibold text-slate-800">
              Live Preview
            </h2>
            <span className="text-xs text-slate-400">
              {PAGE_MM[printSize][0]} × {PAGE_MM[printSize][1]} mm
            </span>
          </div>
          <div className="cover-print-root shadow-panel rounded-sm overflow-hidden">
            <CoverArt cover={cover} fontSize={fontSize} />
          </div>
        </div>
      </div>
    </div>
  );
}
