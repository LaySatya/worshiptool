import { useState, useRef, useMemo } from "react";
import {
  FileText,
  Download,
  Printer,
  Image,
  Calendar,
  Sparkles,
  CheckCircle,
  XCircle,
  FileDown,
  Loader2,
} from "lucide-react";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import { calculateGrade, KHMER_MONTHS } from "../../services/mathStorage.js";

export default function MathReport({
  students = [],
  scores = [],
  rules = [],
  selectedMonth,
  selectedYear,
  onMonthChange,
  onYearChange,
}) {
  const reportRef = useRef(null);
  const [downloadingPng, setDownloadingPng] = useState(false);
  const [downloadingPdf, setDownloadingPdf] = useState(false);

  // Custom metadata for report header
  const [reportTitle, setReportTitle] = useState("លទ្ធផលប្រឡងគណិតវិទ្យា");
  const [className, setClassName] = useState("ថ្នាក់ទី១០");
  const [teacherName, setTeacherName] = useState("");
  const [sortBy, setSortBy] = useState("score_desc"); // 'score_desc' | 'name'

  // Current month label
  const currentMonthLabel = useMemo(() => {
    const found = KHMER_MONTHS.find((m) => m.value === selectedMonth);
    return found ? found.label.split(" ")[0] : selectedMonth;
  }, [selectedMonth]);

  // Combine students and scores
  const reportData = useMemo(() => {
    const scoreMap = {};
    scores.forEach((sc) => {
      scoreMap[sc.studentId] = sc;
    });

    const hasAnyScoreRecorded = scores.length > 0;

    const rows = students.map((st) => {
      const entry = scoreMap[st.id];

      // If no score recorded for this student in this month, show empty fields!
      if (!entry) {
        return {
          id: st.id,
          name: st.name,
          gender: st.gender || "ប",
          examScore: "",
          totalScore: "",
          displayScore: "",
          grade: "",
          isPass: false,
          isAbsent: false,
          hasRecord: false,
        };
      }

      const isAbsent = Boolean(entry.isAbsent);
      const scoreVal = isAbsent
        ? 0
        : entry.totalScore === ""
          ? ""
          : Number(entry.totalScore);
      const examVal = isAbsent
        ? 0
        : entry.examScore === ""
          ? ""
          : Number(entry.examScore);

      const gradeInfo =
        scoreVal !== "" || isAbsent
          ? calculateGrade(scoreVal === "" ? 0 : scoreVal, rules, isAbsent)
          : { grade: "", isPass: false, color: "#64748b" };

      return {
        id: st.id,
        name: st.name,
        gender: st.gender || "ប",
        examScore: examVal,
        totalScore: scoreVal,
        displayScore: isAbsent ? 0 : scoreVal,
        grade: isAbsent ? "Absent" : gradeInfo.grade,
        isPass: gradeInfo.isPass,
        isAbsent,
        gradeColor: gradeInfo.color,
        hasRecord: true,
      };
    });

    if (sortBy === "score_desc") {
      rows.sort((a, b) => {
        if (!a.hasRecord && !b.hasRecord) return 0;
        if (!a.hasRecord) return 1;
        if (!b.hasRecord) return -1;
        if (a.isAbsent) return 1;
        if (b.isAbsent) return -1;
        return Number(b.totalScore) - Number(a.totalScore);
      });
    } else {
      rows.sort((a, b) => a.name.localeCompare(b.name, "km"));
    }

    // Summary counts for recorded students
    const passedCount = rows.filter((r) => r.hasRecord && !r.isAbsent && r.isPass).length;
    const failedCount = rows.filter((r) => r.hasRecord && !r.isAbsent && !r.isPass).length;
    const absentCount = rows.filter((r) => r.hasRecord && r.isAbsent).length;

    return {
      rows,
      passedCount,
      failedCount,
      absentCount,
      totalCount: rows.length,
      hasRecords: hasAnyScoreRecorded,
    };
  }, [students, scores, rules, sortBy]);

  // Download high-resolution PNG
  const handleDownloadPng = async () => {
    if (!reportRef.current) return;
    setDownloadingPng(true);
    try {
      if (document.fonts && document.fonts.ready) {
        await document.fonts.ready;
      }
      const canvas = await html2canvas(reportRef.current, {
        scale: 2.5, // Crisp retina quality
        useCORS: true,
        backgroundColor: "#ffffff",
        logging: false,
      });

      const url = canvas.toDataURL("image/png");
      const a = document.createElement("a");
      a.href = url;
      a.download = `Math_Report_${currentMonthLabel}_${selectedYear}.png`;
      a.click();
    } catch (err) {
      console.error("Failed to generate PNG:", err);
      alert("Error generating PNG: " + err.message);
    } finally {
      setDownloadingPng(false);
    }
  };

  // Download PDF
  const handleDownloadPdf = async () => {
    if (!reportRef.current) return;
    setDownloadingPdf(true);
    try {
      if (document.fonts && document.fonts.ready) {
        await document.fonts.ready;
      }
      const canvas = await html2canvas(reportRef.current, {
        scale: 2.2,
        useCORS: true,
        backgroundColor: "#ffffff",
        logging: false,
      });

      const imgData = canvas.toDataURL("image/jpeg", 0.95);
      const pdf = new jsPDF("p", "mm", "a4");
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (canvas.height * pdfWidth) / canvas.width;

      pdf.addImage(imgData, "JPEG", 0, 8, pdfWidth, pdfHeight);
      pdf.save(`Math_Report_${currentMonthLabel}_${selectedYear}.pdf`);
    } catch (err) {
      console.error("Failed to generate PDF:", err);
      alert("Error generating PDF: " + err.message);
    } finally {
      setDownloadingPdf(false);
    }
  };

  // Browser Print
  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="space-y-5 font-sans">
      {/* Top Controls Toolbar */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-white p-4 rounded-xl border border-slate-200/80 shadow-xs print:hidden">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-slate-100 text-slate-700 flex items-center justify-center font-bold">
            <FileText size={16} />
          </div>
          <div>
            <h3 className="font-semibold text-slate-900 text-sm">
              Monthly Examination Report
            </h3>
            <p className="text-xs text-slate-500">
              Generate, download (PNG / PDF), or print the official mathematics score sheet
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {/* Month & Year */}
          <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1 text-xs">
            <Calendar size={13} className="text-slate-400" />
            <select
              value={selectedMonth}
              onChange={(e) => onMonthChange(e.target.value)}
              className="bg-transparent font-medium text-slate-800 outline-none cursor-pointer"
            >
              {KHMER_MONTHS.map((m) => (
                <option key={m.value} value={m.value}>
                  {m.label}
                </option>
              ))}
            </select>
            <select
              value={selectedYear}
              onChange={(e) => onYearChange(parseInt(e.target.value, 10))}
              className="bg-transparent font-medium text-slate-800 outline-none cursor-pointer ml-1"
            >
              {[2024, 2025, 2026, 2027].map((y) => (
                <option key={y} value={y}>
                  {y}
                </option>
              ))}
            </select>
          </div>

          {/* Sort */}
          <select
            value={sortBy}
            onChange={(e) => setSortBy(e.target.value)}
            className="text-xs px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-700"
          >
            <option value="score_desc">Sort by Score (Highest First)</option>
            <option value="name">Sort by Name (Alphabetical)</option>
          </select>

          {/* Download PNG */}
          <button
            onClick={handleDownloadPng}
            disabled={downloadingPng}
            className="px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-700 text-xs font-medium flex items-center gap-1.5 transition"
            title="Download high-resolution PNG image"
          >
            {downloadingPng ? (
              <Loader2 size={13} className="animate-spin text-slate-600" />
            ) : (
              <Image size={13} className="text-slate-600" />
            )}
            <span>Export PNG</span>
          </button>

          {/* Download PDF */}
          <button
            onClick={handleDownloadPdf}
            disabled={downloadingPdf}
            className="px-3.5 py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white font-medium text-xs flex items-center gap-1.5 transition shadow-xs"
            title="Download printable PDF"
          >
            {downloadingPdf ? (
              <Loader2 size={13} className="animate-spin" />
            ) : (
              <Download size={13} />
            )}
            <span>Export PDF</span>
          </button>

          {/* Print */}
          <button
            onClick={handlePrint}
            className="p-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50 transition"
            title="Print Report"
          >
            <Printer size={14} />
          </button>
        </div>
      </div>

      {/* Customizable Metadata options bar */}
      <div className="bg-white p-3.5 rounded-xl border border-slate-200/80 flex flex-wrap items-center gap-4 text-xs text-slate-600 print:hidden shadow-xs">
        <span className="font-medium text-slate-800">Report Settings:</span>
        <div className="flex items-center gap-1.5">
          <span>Title:</span>
          <input
            type="text"
            value={reportTitle}
            onChange={(e) => setReportTitle(e.target.value)}
            className="px-2.5 py-1 rounded-md border border-slate-200 bg-slate-50 text-xs font-khmer font-semibold w-52 focus:outline-none focus:border-slate-400"
          />
        </div>
        <div className="flex items-center gap-1.5">
          <span>Class:</span>
          <input
            type="text"
            value={className}
            onChange={(e) => setClassName(e.target.value)}
            className="px-2 py-1 rounded-md border border-slate-200 bg-slate-50 text-xs w-28 focus:outline-none focus:border-slate-400"
          />
        </div>
        <div className="flex items-center gap-1.5">
          <span>Teacher:</span>
          <input
            type="text"
            placeholder="Teacher name…"
            value={teacherName}
            onChange={(e) => setTeacherName(e.target.value)}
            className="px-2 py-1 rounded-md border border-slate-200 bg-slate-50 text-xs w-36 focus:outline-none focus:border-slate-400"
          />
        </div>
      </div>

      {/* ─────────────────────────────────────────────────────────────
          REPORT CANVAS / PREVIEW CONTAINER
          Professional Khmer Academic Report Layout
         ───────────────────────────────────────────────────────────── */}
      <div className="flex justify-center overflow-x-auto p-2">
        <div
          ref={reportRef}
          className="bg-white p-12 sm:p-16 shadow-lg border border-slate-300 w-full max-w-[960px] text-slate-900"
          style={{
            fontFamily:
              "'Kantumruy Pro', 'Battambang', 'Khmer OS Battambang', Arial, sans-serif",
          }}
        >
          {/* Main Title Header */}
          <div className="text-center mb-9">
            <h1
              className="text-2xl sm:text-3xl font-bold tracking-tight text-slate-900 uppercase"
              style={{
                fontFamily: "'Moul', 'Khmer OS Muol', 'Kantumruy Pro', serif",
                letterSpacing: "0.5px",
              }}
            >
              {reportTitle}
            </h1>
            <div className="flex items-center justify-center gap-6 mt-3 text-xs text-slate-600 font-medium">
              {className && (
                <span>
                  កម្រិត/ថ្នាក់: <strong className="text-slate-900">{className}</strong>
                </span>
              )}
              <span>
                កាលបរិច្ឆេទ: <strong className="text-slate-900">ខែ {currentMonthLabel} ឆ្នាំ {selectedYear}</strong>
              </span>
              {teacherName && (
                <span>
                  គ្រូបង្រៀន: <strong className="text-slate-900">{teacherName}</strong>
                </span>
              )}
            </div>
            <div className="w-24 h-0.5 bg-slate-900 mx-auto mt-3" />
          </div>

          {/* Table */}
          <div className="w-full">
            <table
              className="w-full border-collapse text-center text-xs sm:text-sm"
              style={{
                border: "1px solid #334155",
              }}
            >
              <thead>
                <tr
                  className="bg-slate-50"
                  style={{
                    borderBottom: "1px solid #334155",
                    height: "40px",
                  }}
                >
                  <th
                    className="font-bold text-center px-4 text-slate-900"
                    style={{
                      borderRight: "1px solid #334155",
                      width: "32%",
                    }}
                  >
                    ឈ្មោះ
                  </th>
                  <th
                    className="font-bold text-center px-2 text-slate-900"
                    style={{
                      borderRight: "1px solid #334155",
                      width: "10%",
                    }}
                  >
                    ភេទ
                  </th>
                  <th
                    className="font-bold text-center px-2 text-slate-900"
                    style={{
                      borderRight: "1px solid #334155",
                      width: "14%",
                    }}
                  >
                    ប្រឡង
                  </th>
                  <th
                    className="font-bold text-center px-2 text-slate-900"
                    style={{
                      borderRight: "1px solid #334155",
                      width: "14%",
                    }}
                  >
                    សរុប
                  </th>
                  <th
                    className="font-bold text-center px-2 text-slate-900"
                    style={{
                      borderRight: "1px solid #334155",
                      width: "14%",
                    }}
                  >
                    ពិន្ទុ
                  </th>
                  <th
                    className="font-bold text-center px-2 text-slate-900"
                    style={{
                      width: "16%",
                    }}
                  >
                    និទ្ទេស
                  </th>
                </tr>
              </thead>
              <tbody>
                {reportData.rows.map((row) => {
                  const isFail = row.grade === "F" || row.isAbsent;
                  return (
                    <tr
                      key={row.id}
                      style={{
                        borderBottom: "1px solid #334155",
                        height: "36px",
                      }}
                    >
                      {/* Name */}
                      <td
                        className="px-4 text-center font-medium font-khmer text-slate-900"
                        style={{
                          borderRight: "1px solid #334155",
                        }}
                      >
                        {row.name}
                      </td>

                      {/* Gender (ប / ស) */}
                      <td
                        className="px-2 text-center text-slate-700 font-medium font-khmer"
                        style={{
                          borderRight: "1px solid #334155",
                        }}
                      >
                        {row.gender}
                      </td>

                      {/* Exam */}
                      <td
                        className="px-2 text-center font-mono font-medium text-slate-800"
                        style={{
                          borderRight: "1px solid #334155",
                        }}
                      >
                        {row.hasRecord ? (row.isAbsent ? 0 : row.examScore) : ""}
                      </td>

                      {/* Total */}
                      <td
                        className="px-2 text-center font-mono font-medium text-slate-800"
                        style={{
                          borderRight: "1px solid #334155",
                        }}
                      >
                        {row.hasRecord ? (row.isAbsent ? 0 : row.totalScore) : ""}
                      </td>

                      {/* Score */}
                      <td
                        className="px-2 text-center font-mono font-semibold text-slate-900"
                        style={{
                          borderRight: "1px solid #334155",
                        }}
                      >
                        {row.hasRecord ? row.displayScore : ""}
                      </td>

                      {/* Grade (Red background for Fail/Absent like in sample image, crisp bold text) */}
                      <td
                        className="px-2 text-center font-bold font-mono"
                        style={{
                          backgroundColor: row.hasRecord && isFail ? "#dc2626" : "transparent",
                          color: row.hasRecord && isFail ? "#ffffff" : "#0f172a",
                        }}
                      >
                        {row.hasRecord ? row.grade : ""}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {/* Bottom Summary (Clean, elegant layout matching exact sample image) */}
          <div className="mt-8 flex justify-between items-end pt-4 border-t border-slate-200">
            <div className="text-xs text-slate-500 font-mono">
              Total Students: {reportData.totalCount}
            </div>

            {reportData.hasRecords ? (
              <div
                className="text-right text-sm sm:text-base font-bold space-y-1 text-slate-900"
                style={{
                  fontFamily: "'Kantumruy Pro', 'Battambang', sans-serif",
                }}
              >
                <div>
                  ជាប់ <span className="font-mono text-emerald-700">{reportData.passedCount}</span> នាក់
                </div>
                <div>
                  ធ្លាក់ <span className="font-mono text-rose-600">{reportData.failedCount}</span> នាក់
                </div>
                {reportData.absentCount > 0 && (
                  <div className="text-slate-600 font-medium text-xs sm:text-sm">
                    អវត្តមាន <span className="font-mono text-slate-700">{reportData.absentCount}</span> នាក់
                  </div>
                )}
              </div>
            ) : (
              <div className="text-xs text-slate-400 italic">
                (No examination records for this month yet)
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
