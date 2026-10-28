import { useState, useEffect, useCallback } from "react";
import {
  Calculator,
  LayoutDashboard,
  Users,
  FileText,
  FileSpreadsheet,
  Sliders,
  CheckCircle2,
  AlertCircle,
  X,
  Database,
} from "lucide-react";
import MathDashboard from "./components/math/MathDashboard.jsx";
import MathStudents from "./components/math/MathStudents.jsx";
import MathScores from "./components/math/MathScores.jsx";
import MathReport from "./components/math/MathReport.jsx";
import MathRules from "./components/math/MathRules.jsx";
import {
  getStudents,
  saveStudent,
  deleteStudent,
  getScoresForMonth,
  saveScoresForMonth,
  getGradeRules,
  saveGradeRules,
  DEFAULT_GRADE_RULES,
} from "./services/mathStorage.js";
import { isInitialized, firebaseConfig } from "./firebase.js";

function Toast({ message, type = "success", onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 3500);
    return () => clearTimeout(t);
  }, [onClose]);
  return (
    <div
      className={`fixed bottom-5 right-5 z-50 flex items-center gap-2.5 px-4 py-3 rounded-xl shadow-lg border text-xs font-medium animate-fade-in
      ${type === "success" ? "bg-white border-emerald-200 text-emerald-800" : "bg-white border-rose-200 text-rose-800"}`}
    >
      {type === "success" ? (
        <CheckCircle2 size={15} className="text-emerald-600" />
      ) : (
        <AlertCircle size={15} className="text-rose-600" />
      )}
      <span>{message}</span>
      <button
        onClick={onClose}
        className="ml-2 text-slate-400 hover:text-slate-600 p-0.5"
      >
        <X size={12} />
      </button>
    </div>
  );
}

export default function MathematicPage() {
  const [activeTab, setActiveTab] = useState("dashboard"); // 'dashboard' | 'students' | 'scores' | 'report' | 'rules'
  const [students, setStudents] = useState([]);
  const [scores, setScores] = useState([]);
  const [rules, setRules] = useState(DEFAULT_GRADE_RULES);
  const [selectedMonth, setSelectedMonth] = useState("10"); // October
  const [selectedYear, setSelectedYear] = useState(2026);
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState(null);

  // Load initial data
  const loadData = useCallback(async () => {
    setLoading(true);
    try {
      const [stList, rList, scList] = await Promise.all([
        getStudents(),
        getGradeRules(),
        getScoresForMonth(selectedYear, selectedMonth),
      ]);
      setStudents(stList);
      setRules(rList);
      setScores(scList);
    } catch (err) {
      console.error("Error loading math data:", err);
      setToast({
        message: "Failed to load data: " + err.message,
        type: "error",
      });
    } finally {
      setLoading(false);
    }
  }, [selectedYear, selectedMonth]);

  useEffect(() => {
    loadData();
  }, [loadData]);

  const handleMonthChange = (month) => {
    setSelectedMonth(month);
  };

  const handleYearChange = (year) => {
    setSelectedYear(year);
  };

  // Student CRUD
  const handleAddStudent = async (studentData) => {
    try {
      const created = await saveStudent(studentData);
      setStudents((prev) => [...prev, created]);
      setToast({
        message: `Student "${created.name}" added successfully`,
        type: "success",
      });
    } catch (err) {
      setToast({
        message: "Error adding student: " + err.message,
        type: "error",
      });
    }
  };

  const handleUpdateStudent = async (studentData) => {
    try {
      const updated = await saveStudent(studentData);
      setStudents((prev) =>
        prev.map((s) => (s.id === updated.id ? updated : s)),
      );
      setToast({
        message: `Student "${updated.name}" updated`,
        type: "success",
      });
    } catch (err) {
      setToast({
        message: "Error updating student: " + err.message,
        type: "error",
      });
    }
  };

  const handleDeleteStudent = async (studentId) => {
    try {
      await deleteStudent(studentId);
      setStudents((prev) => prev.filter((s) => s.id !== studentId));
      setToast({ message: "Student deleted", type: "success" });
    } catch (err) {
      setToast({
        message: "Error deleting student: " + err.message,
        type: "error",
      });
    }
  };

  // Save Scores
  const handleSaveScores = async (updatedScoresList) => {
    try {
      const saved = await saveScoresForMonth(
        selectedYear,
        selectedMonth,
        updatedScoresList,
      );
      setScores(saved);
      setToast({
        message: `Scores for ${selectedMonth}/${selectedYear} saved`,
        type: "success",
      });
    } catch (err) {
      setToast({
        message: "Failed to save scores: " + err.message,
        type: "error",
      });
    }
  };

  // Save Rules
  const handleSaveRules = async (updatedRules) => {
    try {
      const saved = await saveGradeRules(updatedRules);
      setRules(saved);
      setToast({ message: "Grade rules updated", type: "success" });
    } catch (err) {
      setToast({
        message: "Failed to save rules: " + err.message,
        type: "error",
      });
    }
  };

  return (
    <div className="flex-1 overflow-hidden flex flex-col h-full bg-slate-50/60 font-sans">
      {/* Top Header & Sub-navigation bar */}
      <div className="bg-white border-b border-slate-200/80 px-6 pt-3.5 pb-0 flex flex-col flex-shrink-0 z-10 shadow-xs">
        <div className="flex items-center justify-between pb-3">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-lg bg-slate-900 text-white flex items-center justify-center font-medium shadow-xs">
              <Calculator size={16} />
            </div>
            <div>
              <h2 className="text-sm font-semibold text-slate-900 tracking-tight">
                Mathematics Management
              </h2>
              <p className="text-xs text-slate-500">
                Manage students, record monthly scores, and generate reports
              </p>
            </div>
          </div>

          {/* Firebase Connection status */}
          <div className="hidden sm:flex items-center gap-1.5 px-2.5 py-1 rounded-md text-[11px] font-medium bg-slate-100/80 text-slate-600 border border-slate-200">
            <span
              className={`w-1.5 h-1.5 rounded-full ${
                isInitialized ? "bg-emerald-500" : "bg-slate-400"
              }`}
            />
            <Database size={11} className="text-slate-400" />
            <span>Firebase: immanuel-songs</span>
          </div>
        </div>

        {/* Sub-Navigation Tabs */}
        <div className="flex items-center gap-1 overflow-x-auto border-t border-slate-100 pt-1 -mb-[1px]">
          <button
            type="button"
            onClick={() => setActiveTab("dashboard")}
            className={`flex items-center gap-2 px-3.5 py-2.5 text-xs font-medium border-b-2 transition-all whitespace-nowrap ${
              activeTab === "dashboard"
                ? "border-slate-900 text-slate-900 font-semibold"
                : "border-transparent text-slate-500 hover:text-slate-800 hover:border-slate-200"
            }`}
          >
            <LayoutDashboard size={14} />
            <span>Dashboard</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab("students")}
            className={`flex items-center gap-2 px-3.5 py-2.5 text-xs font-medium border-b-2 transition-all whitespace-nowrap ${
              activeTab === "students"
                ? "border-slate-900 text-slate-900 font-semibold"
                : "border-transparent text-slate-500 hover:text-slate-800 hover:border-slate-200"
            }`}
          >
            <Users size={14} />
            <span>Students</span>
            <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-slate-100 text-slate-600 font-mono">
              {students.length}
            </span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab("scores")}
            className={`flex items-center gap-2 px-3.5 py-2.5 text-xs font-medium border-b-2 transition-all whitespace-nowrap ${
              activeTab === "scores"
                ? "border-slate-900 text-slate-900 font-semibold"
                : "border-transparent text-slate-500 hover:text-slate-800 hover:border-slate-200"
            }`}
          >
            <FileText size={14} />
            <span>Scores</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab("report")}
            className={`flex items-center gap-2 px-3.5 py-2.5 text-xs font-medium border-b-2 transition-all whitespace-nowrap ${
              activeTab === "report"
                ? "border-slate-900 text-slate-900 font-semibold"
                : "border-transparent text-slate-500 hover:text-slate-800 hover:border-slate-200"
            }`}
          >
            <FileSpreadsheet size={14} />
            <span>Report</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab("rules")}
            className={`flex items-center gap-2 px-3.5 py-2.5 text-xs font-medium border-b-2 transition-all whitespace-nowrap ${
              activeTab === "rules"
                ? "border-slate-900 text-slate-900 font-semibold"
                : "border-transparent text-slate-500 hover:text-slate-800 hover:border-slate-200"
            }`}
          >
            <Sliders size={14} />
            <span>Grade Rules</span>
          </button>
        </div>
      </div>

      {/* Main Tab Content */}
      <div className="flex-1 overflow-y-auto p-5">
        <div className="max-w-[1400px] mx-auto">
          {activeTab === "dashboard" && (
            <MathDashboard
              students={students}
              scores={scores}
              rules={rules}
              currentMonth={selectedMonth}
              currentYear={selectedYear}
              onNavigateTab={setActiveTab}
            />
          )}

          {activeTab === "students" && (
            <MathStudents
              students={students}
              scores={scores}
              rules={rules}
              onAddStudent={handleAddStudent}
              onUpdateStudent={handleUpdateStudent}
              onDeleteStudent={handleDeleteStudent}
            />
          )}

          {activeTab === "scores" && (
            <MathScores
              students={students}
              scores={scores}
              rules={rules}
              selectedMonth={selectedMonth}
              selectedYear={selectedYear}
              onMonthChange={handleMonthChange}
              onYearChange={handleYearChange}
              onSaveScores={handleSaveScores}
            />
          )}

          {activeTab === "report" && (
            <MathReport
              students={students}
              scores={scores}
              rules={rules}
              selectedMonth={selectedMonth}
              selectedYear={selectedYear}
              onMonthChange={handleMonthChange}
              onYearChange={handleYearChange}
            />
          )}

          {activeTab === "rules" && (
            <MathRules rules={rules} onSaveRules={handleSaveRules} />
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
