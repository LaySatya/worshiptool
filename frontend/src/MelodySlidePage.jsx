import { useState, useRef, useCallback, useEffect } from "react";
import {
  Music,
  Upload,
  Download,
  Scissors,
  Crop,
  Plus,
  Trash2,
  ArrowUp,
  ArrowDown,
  Loader2,
  CheckCircle2,
  AlertCircle,
  X,
  FileMusic,
  Eye,
  RefreshCw,
  Sliders,
  ChevronLeft,
  ChevronRight,
  Maximize2,
} from "lucide-react";
import {
  processMelodyFile,
  cropMelodyPiece,
  generateMelodyPptx,
} from "./api.js";

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
      {type === "success" ? <CheckCircle2 size={15} /> : <AlertCircle size={15} />}
      {message}
      <button onClick={onClose} className="ml-1 text-slate-400 hover:text-slate-600">
        <X size={13} />
      </button>
    </div>
  );
}

export default function MelodySlidePage() {
  const [file, setFile] = useState(null);
  const [fileName, setFileName] = useState("");
  const [title, setTitle] = useState("");
  const [loading, setLoading] = useState(false);
  const [cropping, setCropping] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [toast, setToast] = useState(null);

  // Settings
  const [linesPerSlide, setLinesPerSlide] = useState(1);
  const [bgColor, setBgColor] = useState("#ffffff");

  // Server response
  const [pages, setPages] = useState([]);
  const [pieces, setPieces] = useState([]);
  const [selectedPageIndex, setSelectedPageIndex] = useState(0);
  const [currentSlideIndex, setCurrentSlideIndex] = useState(0);

  // Crop / Selection tool on page
  const [cropMode, setCropMode] = useState(false);
  const [cropDrag, setCropDrag] = useState(null); // { startX, startY, currentX, currentY }
  const pageContainerRef = useRef(null);
  const fileInputRef = useRef(null);

  // Handle file upload
  const handleUpload = useCallback(
    async (uploadedFile, lps = linesPerSlide) => {
      if (!uploadedFile) return;
      setFile(uploadedFile);
      setFileName(uploadedFile.name);
      const guessedTitle = uploadedFile.name
        .replace(/\.[^.]+$/, "")
        .replace(/[-_]/g, " ");
      setTitle(guessedTitle);
      setLoading(true);
      try {
        const data = await processMelodyFile(uploadedFile, lps, 200);
        setPages(data.pages || []);
        setPieces(data.pieces || []);
        setSelectedPageIndex(0);
        setCurrentSlideIndex(0);
        setToast({
          message: `Detected ${data.pieces?.length || 0} melody piece${(data.pieces?.length || 0) === 1 ? "" : "s"} across ${data.page_count} page(s)!`,
          type: "success",
        });
      } catch (err) {
        setToast({
          message: err.message || "Could not process melody sheet.",
          type: "error",
        });
      } finally {
        setLoading(false);
      }
    },
    [linesPerSlide]
  );

  const onDrop = (e) => {
    e.preventDefault();
    const dropped = e.dataTransfer?.files?.[0];
    if (dropped) handleUpload(dropped);
  };

  // Re-detect with new lines_per_slide
  const handleReDetect = (newLines) => {
    setLinesPerSlide(newLines);
    if (file) {
      handleUpload(file, newLines);
    }
  };

  // Reorder pieces
  const movePiece = (index, direction) => {
    const nextIdx = index + direction;
    if (nextIdx < 0 || nextIdx >= pieces.length) return;
    const updated = [...pieces];
    const temp = updated[index];
    updated[index] = updated[nextIdx];
    updated[nextIdx] = temp;
    setPieces(updated);
  };

  // Delete piece
  const removePiece = (index) => {
    setPieces((prev) => prev.filter((_, i) => i !== index));
    if (currentSlideIndex >= pieces.length - 1) {
      setCurrentSlideIndex(Math.max(0, pieces.length - 2));
    }
  };

  // ── Interactive rectangle crop selection on the page ───────────
  const currentPage = pages[selectedPageIndex];

  const handleMouseDown = (e) => {
    if (!cropMode || !pageContainerRef.current) return;
    const rect = pageContainerRef.current.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    setCropDrag({ startX: x, startY: y, currentX: x, currentY: y });
  };

  const handleMouseMove = (e) => {
    if (!cropDrag || !pageContainerRef.current) return;
    const rect = pageContainerRef.current.getBoundingClientRect();
    const x = Math.max(0, Math.min(rect.width, e.clientX - rect.left));
    const y = Math.max(0, Math.min(rect.height, e.clientY - rect.top));
    setCropDrag((prev) => ({ ...prev, currentX: x, currentY: y }));
  };

  const handleMouseUp = async () => {
    if (!cropDrag || !pageContainerRef.current || !currentPage || !file) {
      setCropDrag(null);
      return;
    }

    const rect = pageContainerRef.current.getBoundingClientRect();
    const leftPx = Math.min(cropDrag.startX, cropDrag.currentX);
    const topPx = Math.min(cropDrag.startY, cropDrag.currentY);
    const widthPx = Math.abs(cropDrag.currentX - cropDrag.startX);
    const heightPx = Math.abs(cropDrag.currentY - cropDrag.startY);

    setCropDrag(null);

    // If dragged box is too small, ignore
    if (widthPx < 20 || heightPx < 20) return;

    // Convert DOM coords to original page image pixel coordinates
    const scaleX = currentPage.width / rect.width;
    const scaleY = currentPage.height / rect.height;

    const box = [
      Math.round(leftPx * scaleX),
      Math.round(topPx * scaleY),
      Math.round((leftPx + widthPx) * scaleX),
      Math.round((topPx + heightPx) * scaleY),
    ];

    setCropping(true);
    try {
      const cropped = await cropMelodyPiece(file, selectedPageIndex, box, 200);
      const newPiece = {
        id: cropped.id,
        order: pieces.length + 1,
        page_index: selectedPageIndex,
        box: cropped.box,
        image_data_url: cropped.image_data_url,
        label: `Piece ${pieces.length + 1} (Custom)`,
      };
      setPieces((prev) => [...prev, newPiece]);
      setCurrentSlideIndex(pieces.length);
      setToast({
        message: "Added custom cropped piece!",
        type: "success",
      });
    } catch (err) {
      setToast({ message: "Could not crop piece: " + err.message, type: "error" });
    } finally {
      setCropping(false);
    }
  };

  // ── Download PPTX ───────────────────────────────────────────────
  const handleGenerate = async () => {
    if (!pieces.length) {
      setToast({ message: "No melody pieces to export.", type: "error" });
      return;
    }
    setGenerating(true);
    try {
      const blob = await generateMelodyPptx({
        title: title || "melody_slides",
        bg_color: bgColor,
        pieces,
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${title || "melody"}_slides.pptx`;
      a.click();
      URL.revokeObjectURL(url);
      setToast({
        message: `${pieces.length} melody slides downloaded!`,
        type: "success",
      });
    } catch (err) {
      setToast({ message: "Failed to generate PPTX: " + err.message, type: "error" });
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="flex-1 overflow-hidden flex flex-col h-full bg-slate-50">
      <input
        ref={fileInputRef}
        type="file"
        accept=".pdf,.png,.jpg,.jpeg,.webp"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) handleUpload(f);
          e.target.value = "";
        }}
      />

      {/* Top Header */}
      <div className="bg-white border-b border-slate-200 px-6 py-3 flex items-center justify-between flex-shrink-0 z-10 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-brand-50 text-brand-600 flex items-center justify-center font-bold">
            <FileMusic size={16} />
          </div>
          <div>
            <h2 className="text-sm font-bold text-slate-800">
              Melody &amp; Notation Slide Generator
            </h2>
            <p className="text-xs text-slate-400">
              Split melody + lyric sheet into large, clean presentation slides for church projection
            </p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          {pieces.length > 0 && (
            <div className="flex items-center gap-1.5 text-xs text-slate-600 bg-slate-100 px-2.5 py-1.5 rounded-lg">
              <span className="font-semibold text-brand-600">{pieces.length}</span>
              <span>slides ready</span>
            </div>
          )}

          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            disabled={loading}
            className="btn-secondary !py-2 !px-3 text-xs flex items-center gap-1.5"
          >
            {loading ? (
              <Loader2 size={13} className="animate-spin" />
            ) : (
              <Upload size={13} className="text-brand-500" />
            )}
            <span>{file ? "Change File" : "Upload Melody File"}</span>
          </button>

          <button
            type="button"
            onClick={handleGenerate}
            disabled={generating || !pieces.length}
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
                Download Melody PPTX
              </>
            )}
          </button>
        </div>
      </div>

      {/* Main Workspace */}
      {!file ? (
        /* Empty State Dropzone */
        <div className="flex-1 p-8 flex items-center justify-center">
          <div
            className="max-w-xl w-full border-2 border-dashed border-slate-300 rounded-3xl p-12 text-center bg-white hover:border-brand-400 transition cursor-pointer shadow-sm flex flex-col items-center"
            onClick={() => fileInputRef.current?.click()}
            onDragOver={(e) => e.preventDefault()}
            onDrop={onDrop}
          >
            <div className="w-16 h-16 rounded-2xl bg-brand-50 text-brand-500 flex items-center justify-center mb-4 shadow-xs">
              <FileMusic size={32} />
            </div>
            <h3 className="text-base font-bold text-slate-800 mb-1">
              Upload Melody Music Sheet (PDF or Image)
            </h3>
            <p className="text-xs text-slate-500 max-w-md mb-6 leading-relaxed">
              Upload your scanned hymn or worship song with notes and lyrics. The system will automatically cut it into individual slide-sized pieces, or let you crop custom pieces so notation looks large and readable when projected.
            </p>
            <button type="button" className="btn-primary text-xs !py-2.5 !px-5 pointer-events-none">
              <Upload size={14} />
              Select PDF or Image
            </button>
            <p className="text-[11px] text-slate-400 mt-4">
              Supported formats: .pdf, .png, .jpg, .jpeg, .webp
            </p>
          </div>
        </div>
      ) : (
        /* 3-Column / Split Workspace */
        <div className="flex-1 overflow-hidden grid grid-cols-1 lg:grid-cols-12 gap-4 p-4 min-h-0">
          {/* COLUMN 1: Original Music Sheet & Visual Cropper (5 cols) */}
          <div className="lg:col-span-5 bg-white rounded-2xl border border-slate-200 flex flex-col overflow-hidden shadow-card">
            {/* Toolbar */}
            <div className="p-3 border-b border-slate-100 bg-slate-50/80 flex items-center justify-between flex-shrink-0">
              <div className="flex items-center gap-2">
                <span className="font-semibold text-xs text-slate-700">Original Sheet</span>
                {pages.length > 1 && (
                  <div className="flex items-center gap-1 text-xs">
                    {pages.map((p, idx) => (
                      <button
                        key={idx}
                        onClick={() => setSelectedPageIndex(idx)}
                        className={`px-2 py-0.5 rounded text-[11px] font-semibold transition ${
                          selectedPageIndex === idx
                            ? "bg-brand-500 text-white"
                            : "bg-slate-200 text-slate-600 hover:bg-slate-300"
                        }`}
                      >
                        P.{idx + 1}
                      </button>
                    ))}
                  </div>
                )}
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setCropMode(!cropMode)}
                  className={`text-xs px-2.5 py-1 rounded-lg font-medium border flex items-center gap-1 transition ${
                    cropMode
                      ? "bg-brand-50 border-brand-300 text-brand-700 shadow-xs"
                      : "bg-white border-slate-200 text-slate-600 hover:bg-slate-100"
                  }`}
                  title="Enable click-and-drag box cropper on image"
                >
                  <Crop size={13} className={cropMode ? "text-brand-600" : "text-slate-500"} />
                  <span>{cropMode ? "Cropper Active" : "Select & Crop"}</span>
                </button>
              </div>
            </div>

            {/* Instruction strip */}
            <div className="px-3 py-1.5 bg-brand-50/50 border-b border-slate-100 text-[11px] text-slate-600 flex items-center justify-between">
              <span>
                {cropMode
                  ? "🎯 Click & drag on the sheet below to cut a custom piece"
                  : "💡 Turn on 'Select & Crop' above to highlight any melody part"}
              </span>
              {cropping && (
                <span className="flex items-center gap-1 text-brand-600 font-medium">
                  <Loader2 size={11} className="animate-spin" /> Cropping…
                </span>
              )}
            </div>

            {/* Scrollable Sheet Image */}
            <div
              className={`flex-1 overflow-auto p-3 flex justify-center bg-slate-100/60 select-none relative ${
                cropMode ? "cursor-crosshair" : "cursor-default"
              }`}
              onMouseDown={handleMouseDown}
              onMouseMove={handleMouseMove}
              onMouseUp={handleMouseUp}
            >
              {currentPage ? (
                <div ref={pageContainerRef} className="relative inline-block shadow-sm">
                  <img
                    src={currentPage.preview_data_url}
                    alt={`Melody page ${selectedPageIndex + 1}`}
                    className="max-w-full block rounded border border-slate-200 pointer-events-none"
                    draggable={false}
                  />

                  {/* Active Drag Selection Overlay */}
                  {cropDrag && (
                    <div
                      className="absolute border-2 border-brand-500 bg-brand-500/20 pointer-events-none z-20"
                      style={{
                        left: Math.min(cropDrag.startX, cropDrag.currentX),
                        top: Math.min(cropDrag.startY, cropDrag.currentY),
                        width: Math.abs(cropDrag.currentX - cropDrag.startX),
                        height: Math.abs(cropDrag.currentY - cropDrag.startY),
                      }}
                    />
                  )}

                  {/* Existing detected piece markers */}
                  {pieces
                    .filter((p) => p.page_index === selectedPageIndex && p.box)
                    .map((p, pIdx) => {
                      if (!pageContainerRef.current) return null;
                      const rect = pageContainerRef.current.getBoundingClientRect();
                      const scaleY = rect.height / currentPage.height;
                      const top = p.box[1] * scaleY;
                      const height = (p.box[3] - p.box[1]) * scaleY;
                      return (
                        <div
                          key={p.id || pIdx}
                          className="absolute left-0 right-0 border border-brand-400/40 bg-brand-500/10 pointer-events-none"
                          style={{ top, height }}
                        >
                          <span className="absolute top-1 left-2 text-[10px] font-bold bg-white/90 text-brand-700 px-1.5 py-0.5 rounded shadow-xs">
                            Slide {pIdx + 1}
                          </span>
                        </div>
                      );
                    })}
                </div>
              ) : (
                <div className="flex items-center justify-center text-slate-400 text-xs">
                  Loading page preview…
                </div>
              )}
            </div>
          </div>

          {/* COLUMN 2: Cut Pieces List & Reorder (3 cols) */}
          <div className="lg:col-span-3 bg-white rounded-2xl border border-slate-200 flex flex-col overflow-hidden shadow-card">
            {/* Header */}
            <div className="p-3 border-b border-slate-100 bg-slate-50/80 flex items-center justify-between flex-shrink-0">
              <div className="flex items-center gap-1.5">
                <Scissors size={14} className="text-brand-500" />
                <span className="font-semibold text-xs text-slate-700">Slide Pieces</span>
                <span className="text-[11px] text-slate-400 font-mono">({pieces.length})</span>
              </div>

              {/* Lines per slide select */}
              <div className="flex items-center gap-1 text-[11px] text-slate-500">
                <span>Cut:</span>
                <select
                  value={linesPerSlide}
                  onChange={(e) => handleReDetect(parseInt(e.target.value, 10))}
                  className="px-1.5 py-0.5 text-xs rounded border border-slate-200 bg-white"
                  title="Number of musical lines per slide"
                >
                  <option value={1}>1 line/slide</option>
                  <option value={2}>2 lines/slide</option>
                  <option value={3}>3 lines/slide</option>
                </select>
              </div>
            </div>

            {/* Pieces list */}
            <div className="flex-1 overflow-y-auto p-3 space-y-2">
              {pieces.length === 0 ? (
                <div className="p-6 text-center text-xs text-slate-400">
                  No pieces detected. Use "Select &amp; Crop" on the left to highlight a melody line.
                </div>
              ) : (
                pieces.map((piece, idx) => (
                  <div
                    key={piece.id || idx}
                    onClick={() => setCurrentSlideIndex(idx)}
                    className={`group p-2 rounded-xl border transition-all cursor-pointer flex items-center gap-2.5 ${
                      currentSlideIndex === idx
                        ? "border-brand-500 bg-brand-50/40 shadow-xs"
                        : "border-slate-200 hover:border-slate-300 bg-white"
                    }`}
                  >
                    {/* Index badge */}
                    <div
                      className={`w-6 h-6 rounded-lg text-[11px] font-bold flex items-center justify-center flex-shrink-0 ${
                        currentSlideIndex === idx
                          ? "bg-brand-500 text-white"
                          : "bg-slate-100 text-slate-600"
                      }`}
                    >
                      {idx + 1}
                    </div>

                    {/* Thumbnail */}
                    <div className="w-20 h-10 bg-white border border-slate-200 rounded-md overflow-hidden flex items-center justify-center flex-shrink-0">
                      <img
                        src={piece.image_data_url}
                        alt={`Slide piece ${idx + 1}`}
                        className="max-h-full max-w-full object-contain"
                      />
                    </div>

                    {/* Info */}
                    <div className="flex-1 min-w-0">
                      <div className="text-xs font-semibold text-slate-800 truncate">
                        Slide {idx + 1}
                      </div>
                      <div className="text-[10px] text-slate-400">
                        Page {piece.page_index + 1}
                      </div>
                    </div>

                    {/* Actions: Move up/down, delete */}
                    <div className="flex items-center gap-0.5 opacity-80 group-hover:opacity-100">
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          movePiece(idx, -1);
                        }}
                        disabled={idx === 0}
                        className="p-1 rounded text-slate-400 hover:text-slate-700 disabled:opacity-20"
                        title="Move Up"
                      >
                        <ArrowUp size={12} />
                      </button>
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          movePiece(idx, 1);
                        }}
                        disabled={idx === pieces.length - 1}
                        className="p-1 rounded text-slate-400 hover:text-slate-700 disabled:opacity-20"
                        title="Move Down"
                      >
                        <ArrowDown size={12} />
                      </button>
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          removePiece(idx);
                        }}
                        className="p-1 rounded text-slate-400 hover:text-red-500"
                        title="Remove Slide"
                      >
                        <Trash2 size={12} />
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>

          {/* COLUMN 3: Live 16:9 Presentation Slide Preview (4 cols) */}
          <div className="lg:col-span-4 bg-white rounded-2xl border border-slate-200 flex flex-col overflow-hidden shadow-card">
            {/* Header */}
            <div className="p-3 border-b border-slate-100 bg-slate-50/80 flex items-center justify-between flex-shrink-0">
              <div className="flex items-center gap-1.5">
                <Eye size={14} className="text-brand-500" />
                <span className="font-semibold text-xs text-slate-700">Projector Slide Preview</span>
              </div>
              <div className="text-xs text-slate-500 font-mono">
                {pieces.length > 0 ? `${currentSlideIndex + 1} / ${pieces.length}` : "–"}
              </div>
            </div>

            {/* Slide Stage (16:9 aspect ratio) */}
            <div className="flex-1 p-4 flex flex-col justify-center items-center bg-slate-900/5">
              <div
                className="w-full aspect-[16/9] rounded-xl shadow-lg border border-slate-300/80 flex items-center justify-center p-4 overflow-hidden transition-all relative"
                style={{ backgroundColor: bgColor }}
              >
                {pieces[currentSlideIndex] ? (
                  <img
                    src={pieces[currentSlideIndex].image_data_url}
                    alt={`Slide ${currentSlideIndex + 1}`}
                    className="max-h-full max-w-full object-contain filter drop-shadow-xs"
                  />
                ) : (
                  <div className="text-slate-400 text-xs text-center p-4">
                    <FileMusic size={32} className="mx-auto mb-2 opacity-40" />
                    No slide piece selected
                  </div>
                )}

                {/* Slide Number Badge */}
                {pieces.length > 0 && (
                  <span className="absolute bottom-2 right-3 text-[10px] font-mono text-slate-400 bg-white/70 px-1.5 py-0.5 rounded shadow-xs">
                    {currentSlideIndex + 1}
                  </span>
                )}
              </div>

              {/* Navigation Carousel Arrows */}
              {pieces.length > 1 && (
                <div className="flex items-center justify-center gap-3 mt-4">
                  <button
                    type="button"
                    onClick={() => setCurrentSlideIndex((prev) => Math.max(0, prev - 1))}
                    disabled={currentSlideIndex === 0}
                    className="p-1.5 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-30 shadow-xs transition"
                  >
                    <ChevronLeft size={16} />
                  </button>

                  <span className="text-xs text-slate-600 font-medium font-mono">
                    Slide {currentSlideIndex + 1} of {pieces.length}
                  </span>

                  <button
                    type="button"
                    onClick={() =>
                      setCurrentSlideIndex((prev) => Math.min(pieces.length - 1, prev + 1))
                    }
                    disabled={currentSlideIndex === pieces.length - 1}
                    className="p-1.5 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-30 shadow-xs transition"
                  >
                    <ChevronRight size={16} />
                  </button>
                </div>
              )}
            </div>

            {/* Slide Settings Footer */}
            <div className="p-3 border-t border-slate-100 bg-slate-50 text-xs text-slate-600 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span>Slide BG:</span>
                <input
                  type="color"
                  value={bgColor}
                  onChange={(e) => setBgColor(e.target.value)}
                  className="w-6 h-6 rounded cursor-pointer border border-slate-300"
                  title="Slide Background Color"
                />
                <span className="font-mono text-[11px] text-slate-500">{bgColor}</span>
              </div>

              <button
                type="button"
                onClick={handleGenerate}
                disabled={generating || !pieces.length}
                className="btn-primary !py-1.5 !px-3 text-xs"
              >
                <Download size={13} />
                <span>Export PPTX</span>
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
