import { useState, useEffect, useRef, useCallback } from "react";
import {
  Layers,
  Upload,
  FileUp,
  Loader2,
  CheckCircle2,
  AlertCircle,
  X,
  Download,
  Palette,
  Type,
  Sliders,
  Eye,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  Bold,
  Italic,
  Underline,
  Image as ImageIcon,
  ArrowUpToLine,
  ArrowDownToLine,
  AlignVerticalJustifyCenter,
  Trash2,
  FileText,
} from "lucide-react";
import Toggle from "./components/Toggle.jsx";
import ColorPicker from "./components/ColorPicker.jsx";
import SliderField from "./components/SliderField.jsx";
import { FONT_OPTIONS, ALIGN_OPTIONS } from "./constants.js";
import { inspectSlideDeck, applySlideRestyle } from "./api.js";

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

function Section({ title, children }) {
  return (
    <div className="space-y-4">
      <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 pb-0.5 border-b border-slate-100">
        {title}
      </p>
      {children}
    </div>
  );
}

const EMPTY = {
  bg_color: "",
  font_family: "",
  font_size: "",
  text_color: "",
  bold: null,
  italic: null,
  underline: null,
  align: "",
  line_spacing: "",
  outline_color: "",
  outline_width: "",
  shadow: null,
  vertical_anchor: "",
  scale_font_by_pct: "",
};

export default function SlideEditorPage() {
  const [file, setFile] = useState(null);
  const [deck, setDeck] = useState(null); // { slide_count, thumbnails, ... }
  const [loading, setLoading] = useState(false);
  const [applying, setApplying] = useState(false);
  const [previewing, setPreviewing] = useState(false);
  const [f, setF] = useState(EMPTY);
  const [selected, setSelected] = useState([]); // 1-based indices; [] = all
  const [bgImage, setBgImage] = useState(""); // data URL
  const [previewThumbs, setPreviewThumbs] = useState(null);
  const [settingsOpen, setSettingsOpen] = useState(true);
  const [toast, setToast] = useState(null);
  const [dragActive, setDragActive] = useState(false);

  const fileInputRef = useRef(null);
  const bgInputRef = useRef(null);

  const set = (k, v) => setF((prev) => ({ ...prev, [k]: v }));

  // ── Load a PPTX ─────────────────────────────────────────────────
  const loadFile = useCallback(async (pf) => {
    if (!pf) return;
    if (!pf.name.toLowerCase().endsWith(".pptx")) {
      setToast({ message: "Only .pptx files are supported.", type: "error" });
      return;
    }
    setFile(pf);
    setDeck(null);
    setPreviewThumbs(null);
    setSelected([]);
    setLoading(true);
    try {
      const data = await inspectSlideDeck(pf);
      setDeck(data);
      setToast({
        message: `${data.slide_count} slides loaded.`,
        type: "success",
      });
    } catch (err) {
      setToast({
        message: err.message || "Could not read the PPTX.",
        type: "error",
      });
      setFile(null);
    } finally {
      setLoading(false);
    }
  }, []);

  const onDrop = (e) => {
    e.preventDefault();
    setDragActive(false);
    const pf = e.dataTransfer?.files?.[0];
    if (pf) loadFile(pf);
  };

  // ── Build the settings payload (strip empty values) ─────────────
  const buildSettings = () => {
    const out = {};
    for (const [k, v] of Object.entries(f)) {
      if (v === "" || v === null) continue;
      if (k === "font_size" || k === "outline_width") {
        out[k] = Number(v);
      } else if (k === "line_spacing" || k === "scale_font_by_pct") {
        out[k] = Number(v);
      } else {
        out[k] = v;
      }
    }
    if (bgImage) out.bg_image = bgImage;
    return out;
  };

  const hasAnyChange = () =>
    Object.values(f).some((v) => v !== "" && v !== null) || Boolean(bgImage);

  // ── Preview: apply to the server and render new thumbnails ──────
  const handlePreview = async () => {
    if (!file) return;
    if (!hasAnyChange()) {
      setToast({ message: "Choose at least one change first.", type: "error" });
      return;
    }
    setPreviewing(true);
    try {
      const res = await applySlideRestyle(
        file,
        buildSettings(),
        selected,
        true,
      );
      setPreviewThumbs(res.thumbnails ?? []);
      setToast({ message: "Preview updated.", type: "success" });
    } catch (err) {
      setToast({ message: err.message || "Preview failed.", type: "error" });
    } finally {
      setPreviewing(false);
    }
  };

  // ── Apply + download ────────────────────────────────────────────
  const handleApply = async () => {
    if (!file) return;
    if (!hasAnyChange()) {
      setToast({ message: "Choose at least one change first.", type: "error" });
      return;
    }
    setApplying(true);
    try {
      const blob = await applySlideRestyle(
        file,
        buildSettings(),
        selected,
        false,
      );
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download =
        (file.name || "slides").replace(/\.pptx$/i, "") + "_edited.pptx";
      a.click();
      URL.revokeObjectURL(url);
      setToast({
        message: `Applied to ${selected.length ? selected.length + " slides" : "all slides"} — downloaded!`,
        type: "success",
      });
    } catch (err) {
      setToast({ message: err.message || "Apply failed.", type: "error" });
    } finally {
      setApplying(false);
    }
  };

  // ── Background image upload → data URL ──────────────────────────
  const handleBgImage = (pf) => {
    if (!pf) return;
    const reader = new FileReader();
    reader.onload = () => {
      // Setting the image triggers the debounced preview effect above, so the
      // thumbnails refresh automatically — no extra click needed.
      setBgImage(String(reader.result));
      setToast({
        message: "Background image applied — preview updating…",
        type: "success",
      });
    };
    reader.readAsDataURL(pf);
  };

  // ── Selection helpers ───────────────────────────────────────────
  const toggleSlide = (idx) => {
    setSelected((prev) =>
      prev.includes(idx) ? prev.filter((x) => x !== idx) : [...prev, idx],
    );
  };

  // ── Real-time preview ───────────────────────────────────────────
  // Every formatting change (including picking/changing a background image)
  // re-renders the thumbnails automatically.  The request is debounced so
  // dragging a colour slider or typing a number does not fire dozens of
  // renders, and each new run cancels the previous one.
  useEffect(() => {
    if (!file || !deck) return;
    if (!hasAnyChange()) {
      setPreviewThumbs(null);
      return;
    }

    let cancelled = false;
    const timer = setTimeout(async () => {
      setPreviewing(true);
      try {
        const res = await applySlideRestyle(
          file,
          buildSettings(),
          selected,
          true,
        );
        if (!cancelled) setPreviewThumbs(res.thumbnails ?? []);
      } catch {
        // keep the previous thumbnails; the manual Preview button still works
      } finally {
        if (!cancelled) setPreviewing(false);
      }
    }, 550);

    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [file, deck, f, bgImage, selected]);

  const thumbs = previewThumbs ?? deck?.thumbnails ?? [];
  const slideCount = deck?.slide_count ?? 0;
  const scope = selected.length ? `${selected.length} selected` : "all slides";

  return (
    <div className="flex-1 overflow-hidden flex flex-col h-full bg-slate-50">
      <input
        ref={fileInputRef}
        type="file"
        accept=".pptx"
        className="hidden"
        onChange={(e) => {
          const pf = e.target.files?.[0];
          if (pf) loadFile(pf);
        }}
      />
      <input
        ref={bgInputRef}
        type="file"
        accept="image/*"
        className="hidden"
        onChange={(e) => {
          const pf = e.target.files?.[0];
          if (pf) handleBgImage(pf);
          e.target.value = "";
        }}
      />

      {/* ── Toolbar ─────────────────────────────────────────────── */}
      <div className="bg-white border-b border-slate-200 px-6 py-3 flex flex-wrap items-center justify-between gap-3 flex-shrink-0 z-10 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-brand-50 text-brand-600 flex items-center justify-center">
            <Sliders size={16} />
          </div>
          <div>
            <h2 className="text-sm font-bold text-slate-800">
              Slide Editor — change every slide at once
            </h2>
            <p className="text-xs text-slate-400">
              Upload a .pptx, set background, font, colour and outline, then
              apply to all pages
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            disabled={loading}
            className="btn-secondary !py-2 !px-3 text-xs flex items-center gap-1.5"
          >
            {loading ? (
              <Loader2 size={13} className="animate-spin" />
            ) : (
              <FileUp size={13} className="text-brand-500" />
            )}
            <span>{loading ? "Reading…" : "Upload .pptx"}</span>
          </button>

          <button
            type="button"
            onClick={() => setSettingsOpen((p) => !p)}
            className={`btn-secondary !py-2 !px-3 text-xs flex items-center gap-1.5 ${
              settingsOpen ? "bg-brand-50 text-brand-600 border-brand-200" : ""
            }`}
          >
            <Palette size={14} />
            <span>{settingsOpen ? "Hide Formatting" : "Formatting"}</span>
          </button>

          <button
            type="button"
            onClick={handlePreview}
            disabled={previewing || !deck || !hasAnyChange()}
            className="btn-secondary !py-2 !px-3 text-xs flex items-center gap-1.5"
            title="Preview updates automatically — click to refresh now"
          >
            {previewing ? (
              <Loader2 size={13} className="animate-spin" />
            ) : (
              <Eye size={13} />
            )}
            <span>{previewing ? "Updating…" : "Preview"}</span>
          </button>

          <button
            type="button"
            onClick={handleApply}
            disabled={applying || !deck || !hasAnyChange()}
            className="btn-primary !py-2 !px-4 text-xs flex items-center gap-1.5"
          >
            {applying ? (
              <>
                <Loader2 size={14} className="animate-spin" />
                Applying…
              </>
            ) : (
              <>
                <Download size={14} />
                Apply &amp; Download
              </>
            )}
          </button>
        </div>
      </div>

      {/* ── Workspace ───────────────────────────────────────────── */}
      <div className="flex-1 overflow-hidden flex">
        <div className="flex-1 overflow-y-auto p-5 min-w-0">
          {!deck && !loading ? (
            /* empty state drop zone */
            <div
              onDragOver={(e) => {
                e.preventDefault();
                setDragActive(true);
              }}
              onDragLeave={() => setDragActive(false)}
              onDrop={onDrop}
              onClick={() => fileInputRef.current?.click()}
              className={`max-w-2xl mx-auto mt-10 card p-12 text-center cursor-pointer transition border-2 border-dashed ${
                dragActive
                  ? "border-brand-400 bg-brand-50/60"
                  : "border-slate-200 hover:border-brand-300"
              }`}
            >
              <div className="w-14 h-14 rounded-xl bg-brand-50 text-brand-600 flex items-center justify-center mx-auto mb-3">
                <Upload size={24} />
              </div>
              <h3 className="text-sm font-bold text-slate-800">
                Drop a PowerPoint file here
              </h3>
              <p className="text-xs text-slate-400 mt-1">
                .pptx only · every page can be restyled at once
              </p>
            </div>
          ) : loading ? (
            <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-4 gap-4">
              {Array.from({ length: 8 }, (_, i) => (
                <div
                  key={i}
                  className="rounded-xl bg-slate-100 animate-pulse"
                  style={{ aspectRatio: "16/9" }}
                />
              ))}
            </div>
          ) : (
            <>
              {/* deck summary + selection controls */}
              <div className="card p-4 mb-4 flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-lg bg-slate-100 text-slate-700 flex items-center justify-center">
                    <FileText size={15} />
                  </div>
                  <div>
                    <p className="text-sm font-semibold text-slate-800">
                      {file?.name}
                    </p>
                    <p className="text-xs text-slate-500">
                      {slideCount} slides · applying to{" "}
                      <span className="font-semibold text-brand-600">
                        {scope}
                      </span>
                      {previewThumbs && " · showing preview"}
                    </p>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setSelected([])}
                    className={`text-xs font-medium px-2.5 py-1.5 rounded-lg border transition ${
                      !selected.length
                        ? "bg-brand-50 text-brand-600 border-brand-200"
                        : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"
                    }`}
                  >
                    All slides
                  </button>
                  <button
                    onClick={() =>
                      setSelected(
                        Array.from({ length: slideCount }, (_, i) => i + 1),
                      )
                    }
                    className="text-xs font-medium px-2.5 py-1.5 rounded-lg border bg-white text-slate-600 border-slate-200 hover:bg-slate-50 transition"
                  >
                    Select all
                  </button>
                  {previewThumbs && (
                    <button
                      onClick={() => setPreviewThumbs(null)}
                      className="text-xs font-medium text-slate-400 hover:text-rose-500 px-2 transition"
                    >
                      Reset preview
                    </button>
                  )}
                </div>
              </div>

              {/* slide thumbnails */}
              <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
                {thumbs.map((src, i) => {
                  const idx = i + 1;
                  const isSel = selected.includes(idx);
                  return (
                    <button
                      key={i}
                      onClick={() => toggleSlide(idx)}
                      className={`relative rounded-xl overflow-hidden border-2 text-left transition ${
                        isSel
                          ? "border-brand-500 ring-2 ring-brand-200"
                          : "border-slate-200 hover:border-slate-300"
                      }`}
                    >
                      <img
                        src={
                          src.startsWith("data:")
                            ? src
                            : `data:image/png;base64,${src}`
                        }
                        alt={`Slide ${idx}`}
                        className="w-full h-auto block bg-white"
                      />
                      <span className="absolute top-2 left-2 px-2 py-0.5 rounded-md bg-black/60 text-white text-[11px] font-semibold">
                        {idx}
                      </span>
                      {isSel && (
                        <span className="absolute top-2 right-2 w-5 h-5 rounded-full bg-brand-500 text-white flex items-center justify-center">
                          <CheckCircle2 size={13} />
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>
            </>
          )}
        </div>

        {/* ── Formatting drawer ─────────────────────────────────── */}
        <div
          className={`bg-white border-l border-slate-200 transition-all duration-300 overflow-hidden flex-shrink-0 z-20 ${
            settingsOpen ? "w-80" : "w-0"
          }`}
        >
          {settingsOpen && (
            <div className="w-80 h-full overflow-y-auto">
              <div className="sticky top-0 bg-white border-b border-slate-100 px-5 py-4 z-10 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Palette size={16} className="text-brand-500" />
                  <span className="font-semibold text-sm text-slate-800">
                    Format all slides
                  </span>
                </div>
                <button
                  onClick={() => setSettingsOpen(false)}
                  className="p-1 rounded-lg text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition"
                >
                  <X size={16} />
                </button>
              </div>

              <div className="p-5 space-y-6 text-sm">
                {/* Background */}
                <Section title="Background">
                  <ColorPicker
                    label="Background Color"
                    value={f.bg_color || "#ffffff"}
                    onChange={(v) => set("bg_color", v)}
                  />
                  {f.bg_color && (
                    <button
                      onClick={() => set("bg_color", "")}
                      className="text-[11px] text-slate-400 hover:text-rose-500 transition"
                    >
                      Clear background colour
                    </button>
                  )}

                  <div>
                    <label className="label">Background Image</label>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => bgInputRef.current?.click()}
                        className="btn-secondary !py-1.5 !px-3 text-xs flex items-center gap-1.5"
                      >
                        <ImageIcon size={12} />
                        Choose image…
                      </button>
                      {bgImage && (
                        <button
                          onClick={() => setBgImage("")}
                          className="text-[11px] text-rose-500 hover:underline flex items-center gap-1"
                        >
                          <Trash2 size={11} /> Remove
                        </button>
                      )}
                    </div>
                    {bgImage && (
                      <img
                        src={bgImage}
                        alt="background"
                        className="mt-2 w-full h-16 object-cover rounded-lg border border-slate-200"
                      />
                    )}
                    <p className="mt-1 text-[10px] text-slate-400">
                      An image set here is placed behind the content on every
                      selected slide.
                    </p>
                  </div>
                </Section>

                {/* Text */}
                <Section title="Text">
                  <div>
                    <label className="label">Font Family</label>
                    <select
                      className="select"
                      value={f.font_family}
                      onChange={(e) => set("font_family", e.target.value)}
                    >
                      <option value="">— keep original —</option>
                      {FONT_OPTIONS.map((font) => (
                        <option key={font} value={font}>
                          {font}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="label">Font Size (pt)</label>
                    <div className="flex items-center gap-2">
                      <input
                        type="number"
                        min={6}
                        max={200}
                        value={f.font_size}
                        placeholder="keep"
                        onChange={(e) => set("font_size", e.target.value)}
                        className="input !py-1.5 !text-xs"
                      />
                      <span className="text-[11px] text-slate-400">pt</span>
                    </div>
                  </div>

                  <ColorPicker
                    label="Text Color"
                    value={f.text_color || "#000000"}
                    onChange={(v) => set("text_color", v)}
                  />

                  <div>
                    <label className="label">Style</label>
                    <div className="flex gap-1.5">
                      {[
                        {
                          key: "bold",
                          icon: <Bold size={13} />,
                          label: "Bold",
                        },
                        {
                          key: "italic",
                          icon: <Italic size={13} />,
                          label: "Italic",
                        },
                        {
                          key: "underline",
                          icon: <Underline size={13} />,
                          label: "Underline",
                        },
                      ].map(({ key, icon, label }) => (
                        <button
                          key={key}
                          onClick={() =>
                            set(key, f[key] === true ? null : true)
                          }
                          title={label}
                          className={`flex-1 py-2 rounded-lg border text-xs flex items-center justify-center gap-1 transition ${
                            f[key] === true
                              ? "bg-brand-500 text-white border-brand-500"
                              : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"
                          }`}
                        >
                          {icon}
                          <span>{label}</span>
                        </button>
                      ))}
                    </div>
                  </div>

                  <div>
                    <label className="label">Alignment</label>
                    <div className="flex gap-1.5">
                      {ALIGN_OPTIONS.map(({ value, label }) => (
                        <button
                          key={value}
                          onClick={() =>
                            set("align", f.align === value ? "" : value)
                          }
                          title={label}
                          className={`flex-1 py-2 rounded-lg border text-xs flex items-center justify-center transition ${
                            f.align === value
                              ? "bg-brand-500 text-white border-brand-500"
                              : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"
                          }`}
                        >
                          {value === "left" && <AlignLeft size={13} />}
                          {value === "center" && <AlignCenter size={13} />}
                          {value === "right" && <AlignRight size={13} />}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div>
                    <label className="label">Vertical Position</label>
                    <div className="flex gap-1.5">
                      {[
                        {
                          v: "top",
                          icon: <ArrowUpToLine size={13} />,
                          l: "Top",
                        },
                        {
                          v: "middle",
                          icon: <AlignVerticalJustifyCenter size={13} />,
                          l: "Middle",
                        },
                        {
                          v: "bottom",
                          icon: <ArrowDownToLine size={13} />,
                          l: "Bottom",
                        },
                      ].map(({ v, icon, l }) => (
                        <button
                          key={v}
                          onClick={() =>
                            set(
                              "vertical_anchor",
                              f.vertical_anchor === v ? "" : v,
                            )
                          }
                          className={`flex-1 py-2 rounded-lg border text-xs flex items-center justify-center gap-1 transition ${
                            f.vertical_anchor === v
                              ? "bg-brand-500 text-white border-brand-500"
                              : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"
                          }`}
                        >
                          {icon}
                          <span>{l}</span>
                        </button>
                      ))}
                    </div>
                  </div>

                  <div>
                    <label className="label">Line Spacing</label>
                    <div className="flex items-center gap-2">
                      <input
                        type="number"
                        min={0.5}
                        max={4}
                        step={0.1}
                        value={f.line_spacing}
                        placeholder="keep"
                        onChange={(e) => set("line_spacing", e.target.value)}
                        className="input !py-1.5 !text-xs"
                      />
                    </div>
                  </div>
                </Section>

                {/* Outline & shadow */}
                <Section title="Text Outline & Shadow">
                  <ColorPicker
                    label="Outline Color"
                    value={f.outline_color || "#000000"}
                    onChange={(v) => set("outline_color", v)}
                  />
                  <div>
                    <label className="label">Outline Width (pt)</label>
                    <input
                      type="number"
                      min={0}
                      max={20}
                      step={0.25}
                      value={f.outline_width}
                      placeholder="0 = no outline"
                      onChange={(e) => set("outline_width", e.target.value)}
                      className="input !py-1.5 !text-xs"
                    />
                  </div>

                  <Toggle
                    label="Text Shadow"
                    checked={f.shadow === true}
                    onChange={(v) => set("shadow", v ? true : null)}
                  />

                  <div>
                    <label className="label">Scale Font</label>
                    <div className="flex items-center gap-2">
                      <input
                        type="number"
                        min={25}
                        max={400}
                        step={5}
                        value={f.scale_font_by_pct}
                        placeholder="keep"
                        onChange={(e) =>
                          set("scale_font_by_pct", e.target.value)
                        }
                        className="input !py-1.5 !text-xs"
                      />
                      <span className="text-[11px] text-slate-400">%</span>
                    </div>
                    <p className="mt-1 text-[10px] text-slate-400">
                      Shrinks or grows every existing font size by this
                      percentage.
                    </p>
                  </div>
                </Section>

                <div className="pt-2 border-t border-slate-100">
                  <button
                    onClick={() => {
                      setF(EMPTY);
                      setBgImage("");
                    }}
                    className="w-full py-2 rounded-lg border border-slate-200 text-xs text-slate-600 hover:bg-slate-50 transition"
                  >
                    Reset all formatting
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>
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
