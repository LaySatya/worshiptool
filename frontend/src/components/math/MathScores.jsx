import { useState, useEffect, useMemo } from "react";
import {
  FileText,
  Save,
  CheckCircle2,
  Calendar,
  AlertTriangle,
  RotateCcw,
  Sparkles,
} from "lucide-react";
import { calculateGrade, KHMER_MONTHS } from "../../services/mathStorage.js";

export default function MathScores({
  students = [],
  scores = [],
  rules = [],
  selectedMonth,
  selectedYear,
  onMonthChange,
  onYearChange,
  onSaveScores,
}) {
  // Local editable buffer of scores for the month
  const [localScores, setLocalScores] = useState({});
  const [hasChanges, setHasChanges] = useState(false);
  const [savedSuccess, setSavedSuccess] = useState(false);

  // Initialize localScores from students & passed-in scores
  useEffect(() => {
    const map = {};
    const scoreMap = {};
    scores.forEach((sc) => {
      scoreMap[sc.studentId] = sc;
    });

    students.forEach((st) => {
      const existing = scoreMap[st.id];
      map[st.id] = {
        studentId: st.id,
        examScore: existing ? existing.examScore : 0,
        totalScore: existing ? existing.totalScore : 0,
        isAbsent: existing ? Boolean(existing.isAbsent) : false,
        notes: existing?.notes || "",
      };
    });

    setLocalScores(map);
    setHasChanges(false);
  }, [students, scores, selectedMonth, selectedYear]);

  // Handle cell edits
  const handleScoreChange = (studentId, field, value) => {
    setLocalScores((prev) => {
      const current = prev[studentId] || {
        studentId,
        examScore: 0,
        totalScore: 0,
        isAbsent: false,
      };
      let updated = { ...current };

      if (field === "examScore") {
        const val =
          value === "" ? "" : Math.max(0, Math.min(100, Number(value)));
        updated.examScore = val;
        updated.totalScore = val; // default sync exam with total
        if (Number(val) > 0) updated.isAbsent = false;
      } else if (field === "totalScore") {
        const val =
          value === "" ? "" : Math.max(0, Math.min(100, Number(value)));
        updated.totalScore = val;
      } else if (field === "isAbsent") {
        updated.isAbsent = Boolean(value);
        if (value) {
          updated.examScore = 0;
          updated.totalScore = 0;
        }
      }

      return {
        ...prev,
        [studentId]: updated,
      };
    });
    setHasChanges(true);
    setSavedSuccess(false);
  };

  // Save changes
  const handleSave = () => {
    const list = Object.values(localScores).map((item) => ({
      ...item,
      examScore: item.examScore === "" ? 0 : Number(item.examScore),
      totalScore: item.totalScore === "" ? 0 : Number(item.totalScore),
    }));
    onSaveScores(list);
    setHasChanges(false);
    setSavedSuccess(true);
    setTimeout(() => setSavedSuccess(false), 3000);
  };

  // Quick live summary
  const summary = useMemo(() => {
    const items = Object.values(localScores);
    const total = items.length;
    let pass = 0;
    let fail = 0;
    let absent = 0;
    let sum = 0;
    let countActive = 0;

    items.forEach((item) => {
      if (item.isAbsent) {
        absent += 1;
      } else {
        const scoreVal = Number(item.totalScore) || 0;
        const res = calculateGrade(scoreVal, rules, false);
        if (res.isPass) pass += 1;
        else fail += 1;
        sum += scoreVal;
        countActive += 1;
      }
    });

    const avg = countActive > 0 ? (sum / countActive).toFixed(1) : 0;
    return { total, pass, fail, absent, avg };
  }, [localScores, rules]);

  return (
    <div className="space-y-4 font-sans">
      {/* Top Filter & Action Bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white p-4 rounded-xl border border-slate-200/80 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-slate-100 text-slate-700 flex items-center justify-center font-bold">
            <FileText size={16} />
          </div>
          <div>
            <h3 className="font-semibold text-slate-900 text-sm">
              Monthly Scores Entry
            </h3>
            <p className="text-xs text-slate-500">
              Select month and record math examination scores
            </p>
          </div>
        </div>

        {/* Month & Year Selectors + Save Button */}
        <div className="flex items-center gap-2.5">
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

          <button
            type="button"
            onClick={handleSave}
            disabled={!hasChanges && !savedSuccess}
            className={`px-4 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 transition ${
              hasChanges
                ? "bg-slate-900 hover:bg-slate-800 text-white shadow-xs"
                : "bg-slate-100 text-slate-400 border border-slate-200"
            }`}
          >
            {savedSuccess ? (
              <>
                <CheckCircle2 size={13} className="text-emerald-500" />
                <span className="text-slate-700">Saved</span>
              </>
            ) : (
              <>
                <Save size={13} />
                <span>{hasChanges ? "Save Scores" : "Saved"}</span>
              </>
            )}
          </button>
        </div>
      </div>

      {/* Summary KPI Strip */}
      <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 bg-white p-3 rounded-xl border border-slate-200/80 text-xs shadow-xs">
        <div className="text-center p-2 rounded-lg bg-slate-50">
          <span className="text-slate-500 block text-[11px]">Total Students</span>
          <span className="font-bold text-slate-900 text-sm font-mono">{summary.total}</span>
        </div>
        <div className="text-center p-2 rounded-lg bg-emerald-50/70 text-emerald-800 border border-emerald-100">
          <span className="block text-[11px] font-medium">Passed</span>
          <span className="font-bold text-sm font-mono">{summary.pass}</span>
        </div>
        <div className="text-center p-2 rounded-lg bg-rose-50/70 text-rose-800 border border-rose-100">
          <span className="block text-[11px] font-medium">Failed</span>
          <span className="font-bold text-sm font-mono">{summary.fail}</span>
        </div>
        <div className="text-center p-2 rounded-lg bg-slate-50 text-slate-600">
          <span className="block text-[11px] font-medium">Absent</span>
          <span className="font-bold text-sm font-mono">{summary.absent}</span>
        </div>
        <div className="text-center p-2 rounded-lg bg-slate-50 text-slate-800 col-span-2 sm:col-span-1">
          <span className="block text-[11px] font-medium">Average</span>
          <span className="font-bold text-sm font-mono">{summary.avg}</span>
        </div>
      </div>

      {/* Scores Input Table */}
      <div className="bg-white rounded-xl border border-slate-200/80 shadow-xs overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-semibold text-[11px]">
              <tr>
                <th className="py-2.5 px-4 w-12 text-center">#</th>
                <th className="py-2.5 px-4 min-w-[160px]">Student Name</th>
                <th className="py-2.5 px-3 w-16 text-center">Gender</th>
                <th className="py-2.5 px-4 w-32 text-center">Exam Score</th>
                <th className="py-2.5 px-4 w-32 text-center">Total Score</th>
                <th className="py-2.5 px-4 w-28 text-center">Absent</th>
                <th className="py-2.5 px-4 w-24 text-center">Grade</th>
                <th className="py-2.5 px-4 w-28 text-center">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {students.map((st, idx) => {
                const entry = localScores[st.id] || {
                  examScore: 0,
                  totalScore: 0,
                  isAbsent: false,
                };
                const isAbsent = Boolean(entry.isAbsent);
                const scoreNum = isAbsent
                  ? 0
                  : entry.totalScore === ""
                    ? 0
                    : Number(entry.totalScore);
                const gradeInfo = calculateGrade(scoreNum, rules, isAbsent);
                const isMale = st.gender === "ប" || st.gender === "M";

                return (
                  <tr
                    key={st.id}
                    className={`hover:bg-slate-50/70 transition ${
                      isAbsent ? "bg-slate-50/40 text-slate-400" : ""
                    }`}
                  >
                    <td className="py-2.5 px-4 text-center font-mono text-slate-400 text-[11px]">
                      {idx + 1}
                    </td>

                    <td className="py-2.5 px-4">
                      <span className="font-medium text-slate-900 text-xs font-khmer">
                        {st.name}
                      </span>
                    </td>

                    <td className="py-2.5 px-3 text-center">
                      <span className="text-[11px] font-medium text-slate-500">
                        {isMale ? "ប" : "ស"}
                      </span>
                    </td>

                    {/* Exam Score input */}
                    <td className="py-2 px-4 text-center">
                      <input
                        type="number"
                        min="0"
                        max="100"
                        disabled={isAbsent}
                        value={isAbsent ? "" : entry.examScore}
                        onChange={(e) =>
                          handleScoreChange(st.id, "examScore", e.target.value)
                        }
                        placeholder={isAbsent ? "—" : "0"}
                        className="w-20 text-center py-1 px-2 rounded-md border border-slate-200 focus:border-slate-400 text-xs font-medium font-mono bg-white disabled:bg-slate-100 disabled:text-slate-400"
                      />
                    </td>

                    {/* Total Score input */}
                    <td className="py-2 px-4 text-center">
                      <input
                        type="number"
                        min="0"
                        max="100"
                        disabled={isAbsent}
                        value={isAbsent ? "" : entry.totalScore}
                        onChange={(e) =>
                          handleScoreChange(st.id, "totalScore", e.target.value)
                        }
                        placeholder={isAbsent ? "—" : "0"}
                        className="w-20 text-center py-1 px-2 rounded-md border border-slate-200 focus:border-slate-400 text-xs font-medium font-mono bg-white disabled:bg-slate-100 disabled:text-slate-400"
                      />
                    </td>

                    {/* Absent checkbox */}
                    <td className="py-2.5 px-4 text-center">
                      <label className="inline-flex items-center gap-1.5 cursor-pointer select-none">
                        <input
                          type="checkbox"
                          checked={isAbsent}
                          onChange={(e) =>
                            handleScoreChange(
                              st.id,
                              "isAbsent",
                              e.target.checked,
                            )
                          }
                          className="rounded border-slate-300 text-slate-800 focus:ring-slate-500 w-3.5 h-3.5 cursor-pointer"
                        />
                        <span className="text-[11px] text-slate-500">
                          Absent
                        </span>
                      </label>
                    </td>

                    {/* Grade Badge */}
                    <td className="py-2.5 px-4 text-center">
                      <span
                        className={`inline-block min-w-[32px] py-0.5 px-2 rounded text-[11px] font-bold font-mono ${
                          gradeInfo.isPass
                            ? "bg-slate-900 text-white"
                            : "bg-rose-600 text-white"
                        }`}
                      >
                        {gradeInfo.grade}
                      </span>
                    </td>

                    {/* Result (Passed / Failed) */}
                    <td className="py-2.5 px-4 text-center">
                      {isAbsent ? (
                        <span className="inline-block px-2 py-0.5 rounded text-[11px] font-medium bg-slate-100 text-slate-500">
                          Absent
                        </span>
                      ) : gradeInfo.isPass ? (
                        <span className="inline-block px-2 py-0.5 rounded text-[11px] font-medium bg-emerald-50 text-emerald-700 border border-emerald-200">
                          Passed
                        </span>
                      ) : (
                        <span className="inline-block px-2 py-0.5 rounded text-[11px] font-medium bg-rose-50 text-rose-700 border border-rose-200">
                          Failed
                        </span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
