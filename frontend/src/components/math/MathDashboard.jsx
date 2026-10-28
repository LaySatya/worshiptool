import { useMemo } from "react";
import {
  Users,
  Award,
  TrendingUp,
  AlertTriangle,
  CheckCircle,
  XCircle,
  Clock,
  ArrowUpRight,
  Sparkles,
  BarChart3,
  Calendar,
} from "lucide-react";
import { calculateGrade } from "../../services/mathStorage.js";

export default function MathDashboard({
  students = [],
  scores = [],
  rules = [],
  currentMonth,
  currentYear,
  onNavigateTab,
}) {
  // Aggregate stats for current month
  const stats = useMemo(() => {
    const totalStudents = students.length;
    const maleCount = students.filter(
      (s) => s.gender === "ប" || s.gender === "M",
    ).length;
    const femaleCount = students.filter(
      (s) => s.gender === "ស" || s.gender === "F",
    ).length;

    // Map student with score
    const studentScoreMap = {};
    scores.forEach((sc) => {
      studentScoreMap[sc.studentId] = sc;
    });

    const evaluated = students.map((s) => {
      const entry = studentScoreMap[s.id] || {
        examScore: 0,
        totalScore: 0,
        isAbsent: false,
      };
      const scoreVal =
        typeof entry.totalScore === "number"
          ? entry.totalScore
          : parseFloat(entry.totalScore) || 0;
      const gradeInfo = calculateGrade(scoreVal, rules, entry.isAbsent);
      return {
        ...s,
        score: scoreVal,
        isAbsent: Boolean(entry.isAbsent),
        grade: gradeInfo.grade,
        isPass: gradeInfo.isPass,
        color: gradeInfo.color,
      };
    });

    const activeScorers = evaluated.filter((e) => !e.isAbsent);
    const absentCount = evaluated.filter((e) => e.isAbsent).length;
    const passedCount = evaluated.filter((e) => !e.isAbsent && e.isPass).length;
    const failedCount = evaluated.filter(
      (e) => !e.isAbsent && !e.isPass,
    ).length;
    const passRate =
      activeScorers.length > 0
        ? Math.round((passedCount / activeScorers.length) * 100)
        : 0;

    const scoresList = activeScorers.map((e) => e.score);
    const highestScore = scoresList.length > 0 ? Math.max(...scoresList) : 0;
    const lowestScore = scoresList.length > 0 ? Math.min(...scoresList) : 0;
    const avgScore =
      scoresList.length > 0
        ? (scoresList.reduce((a, b) => a + b, 0) / scoresList.length).toFixed(1)
        : 0;

    // Top students
    const sortedStudents = [...evaluated].sort((a, b) => {
      if (a.isAbsent) return 1;
      if (b.isAbsent) return -1;
      return b.score - a.score;
    });

    // Grade frequency distribution
    const gradeCounts = {};
    rules.forEach((r) => {
      gradeCounts[r.grade] = 0;
    });
    gradeCounts["Absent"] = 0;

    evaluated.forEach((e) => {
      if (e.isAbsent) {
        gradeCounts["Absent"] = (gradeCounts["Absent"] || 0) + 1;
      } else {
        gradeCounts[e.grade] = (gradeCounts[e.grade] || 0) + 1;
      }
    });

    return {
      totalStudents,
      maleCount,
      femaleCount,
      passedCount,
      failedCount,
      absentCount,
      passRate,
      highestScore,
      lowestScore,
      avgScore,
      topStudents: sortedStudents.slice(0, 5),
      needAttention: sortedStudents.filter((s) => !s.isPass),
      gradeCounts,
    };
  }, [students, scores, rules]);

  return (
    <div className="space-y-5 font-sans">
      {/* KPI Cards Grid */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        {/* Card 1: Total Students */}
        <div className="bg-white rounded-xl p-5 border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-500 mb-3">
            <span className="text-xs font-semibold text-slate-600">Total Students</span>
            <div className="w-8 h-8 rounded-lg bg-slate-100 text-slate-700 flex items-center justify-center">
              <Users size={15} />
            </div>
          </div>
          <div>
            <div className="text-3xl font-bold text-slate-900 font-mono tracking-tight">
              {stats.totalStudents}
            </div>
            <div className="flex items-center gap-2 mt-3 pt-2.5 border-t border-slate-100 text-[11px] text-slate-500">
              <span>Male: <strong className="text-slate-800 font-mono">{stats.maleCount}</strong></span>
              <span>•</span>
              <span>Female: <strong className="text-slate-800 font-mono">{stats.femaleCount}</strong></span>
            </div>
          </div>
        </div>

        {/* Card 2: Highest Score */}
        <div className="bg-white rounded-xl p-5 border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-500 mb-3">
            <span className="text-xs font-semibold text-slate-600">Highest Score</span>
            <div className="w-8 h-8 rounded-lg bg-slate-100 text-slate-700 flex items-center justify-center">
              <Award size={15} />
            </div>
          </div>
          <div>
            <div className="text-3xl font-bold text-slate-900 font-mono tracking-tight">
              {stats.highestScore}{" "}
              <span className="text-xs font-normal text-slate-400">/ 100</span>
            </div>
            <div className="mt-3 pt-2.5 border-t border-slate-100 text-[11px] text-slate-500 truncate">
              {stats.topStudents[0] ? (
                <span>
                  Top: <strong className="text-slate-800 font-khmer">{stats.topStudents[0].name}</strong>
                </span>
              ) : (
                "No scores yet"
              )}
            </div>
          </div>
        </div>

        {/* Card 3: Average Score */}
        <div className="bg-white rounded-xl p-5 border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-500 mb-3">
            <span className="text-xs font-semibold text-slate-600">Average Score</span>
            <div className="w-8 h-8 rounded-lg bg-slate-100 text-slate-700 flex items-center justify-center">
              <TrendingUp size={15} />
            </div>
          </div>
          <div>
            <div className="text-3xl font-bold text-slate-900 font-mono tracking-tight">
              {stats.avgScore}
            </div>
            <div className="flex items-center justify-between mt-3 pt-2.5 border-t border-slate-100 text-[11px] text-slate-500">
              <span>Lowest: <strong className="font-mono text-slate-700">{stats.lowestScore}</strong></span>
              <span className="text-slate-400">Out of 100</span>
            </div>
          </div>
        </div>

        {/* Card 4: Pass Rate */}
        <div className="bg-white rounded-xl p-5 border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-500 mb-3">
            <span className="text-xs font-semibold text-slate-600">Pass Rate</span>
            <div className="w-8 h-8 rounded-lg bg-emerald-50 text-emerald-700 flex items-center justify-center">
              <CheckCircle size={15} />
            </div>
          </div>
          <div>
            <div className="text-3xl font-bold text-slate-900 font-mono tracking-tight flex items-baseline gap-2">
              <span>{stats.passRate}%</span>
              <span className="text-xs font-medium text-emerald-700">
                ({stats.passedCount} passed)
              </span>
            </div>
            <div className="flex items-center gap-2 mt-3 pt-2.5 border-t border-slate-100 text-[11px] text-slate-500">
              <span className="text-rose-600 font-medium">Failed: {stats.failedCount}</span>
              <span>•</span>
              <span>Absent: {stats.absentCount}</span>
            </div>
          </div>
        </div>
      </div>

      {/* Main Breakdown Row */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
        {/* Left Column: Grade Distribution (7 cols) */}
        <div className="lg:col-span-7 bg-white rounded-xl p-5 border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div className="flex items-center justify-between pb-3 border-b border-slate-100">
            <div className="flex items-center gap-2">
              <BarChart3 size={15} className="text-slate-600" />
              <h3 className="font-semibold text-slate-900 text-sm">
                Grade Distribution
              </h3>
            </div>
            <span className="text-xs text-slate-400 font-mono">
              {stats.totalStudents} total students
            </span>
          </div>

          {/* Visual clean bars */}
          <div className="space-y-2.5 my-4">
            {rules.map((rule) => {
              const count = stats.gradeCounts[rule.grade] || 0;
              const pct =
                stats.totalStudents > 0
                  ? Math.round((count / stats.totalStudents) * 100)
                  : 0;
              return (
                <div key={rule.grade} className="flex items-center gap-3 text-xs">
                  {/* Grade label */}
                  <span className="w-9 font-semibold text-slate-700 text-left font-mono">
                    {rule.grade}
                  </span>

                  {/* Score Range */}
                  <span className="w-16 text-slate-400 font-mono text-[11px]">
                    {rule.min}–{rule.max}
                  </span>

                  {/* Clean Gray/Slate Bar with green for pass and red for fail */}
                  <div className="flex-1 h-2 bg-slate-100 rounded-full overflow-hidden">
                    <div
                      className="h-full rounded-full transition-all duration-300"
                      style={{
                        width: `${pct}%`,
                        backgroundColor: rule.isPass ? "#16a34a" : "#dc2626",
                      }}
                    />
                  </div>

                  {/* Count */}
                  <span className="w-16 text-right font-medium text-slate-700 font-mono">
                    {count} <span className="text-[10px] text-slate-400">({pct}%)</span>
                  </span>
                </div>
              );
            })}

            {/* Absent row */}
            <div className="flex items-center gap-3 text-xs pt-1.5 border-t border-slate-100">
              <span className="w-9 font-semibold text-slate-500 text-left text-[11px]">
                Abs
              </span>
              <span className="w-16 text-slate-400 text-[11px]">Absent</span>
              <div className="flex-1 h-2 bg-slate-100 rounded-full overflow-hidden">
                <div
                  className="h-full bg-slate-400 rounded-full transition-all duration-300"
                  style={{
                    width: `${stats.totalStudents > 0 ? (stats.absentCount / stats.totalStudents) * 100 : 0}%`,
                  }}
                />
              </div>
              <span className="w-16 text-right font-medium text-slate-700 font-mono">
                {stats.absentCount}
              </span>
            </div>
          </div>

          {/* Footer Summary */}
          <div className="flex items-center justify-between pt-3 border-t border-slate-100 text-xs text-slate-500">
            <div className="flex items-center gap-3">
              <span className="flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-emerald-600" />
                <span>Passed: <strong className="text-emerald-700">{stats.passedCount}</strong></span>
              </span>
              <span className="flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-rose-600" />
                <span>Failed: <strong className="text-rose-700">{stats.failedCount}</strong></span>
              </span>
            </div>
            <button
              onClick={() => onNavigateTab("rules")}
              className="text-slate-700 hover:text-slate-900 font-medium hover:underline text-xs"
            >
              Configure Grade Rules &rarr;
            </button>
          </div>
        </div>

        {/* Right Column: Top Performing Students (5 cols) */}
        <div className="lg:col-span-5 bg-white rounded-xl p-5 border border-slate-200/80 shadow-xs flex flex-col justify-between">
          <div className="flex items-center justify-between pb-3 border-b border-slate-100">
            <div className="flex items-center gap-2">
              <Award size={15} className="text-slate-600" />
              <h3 className="font-semibold text-slate-900 text-sm">
                Top Performing Students
              </h3>
            </div>
            <span className="text-xs text-slate-400 font-medium">Top 5</span>
          </div>

          <div className="space-y-2 my-3">
            {stats.topStudents.map((st, idx) => (
              <div
                key={st.id}
                className="flex items-center justify-between p-2.5 rounded-lg bg-slate-50 border border-slate-100 hover:bg-slate-100/60 transition"
              >
                <div className="flex items-center gap-3">
                  <span className="w-5 text-center font-mono text-xs font-semibold text-slate-400">
                    #{idx + 1}
                  </span>
                  <div>
                    <div className="text-xs font-semibold text-slate-900 font-khmer">
                      {st.name}
                    </div>
                    <div className="text-[11px] text-slate-500">
                      {st.gender === "ប" ? "Male" : "Female"} • {st.gradeLevel || "Grade 10"}
                    </div>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  <span className="text-xs font-bold text-slate-900 font-mono">
                    {st.score}
                  </span>
                  <span
                    className={`px-2 py-0.5 rounded text-[10px] font-bold font-mono ${
                      st.isPass
                        ? "bg-emerald-600 text-white"
                        : "bg-rose-600 text-white"
                    }`}
                  >
                    {st.grade}
                  </span>
                </div>
              </div>
            ))}
          </div>

          {/* Attention alert */}
          <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg flex items-center justify-between text-xs text-slate-700">
            <div className="flex items-center gap-2">
              <AlertTriangle size={14} className="text-rose-600 flex-shrink-0" />
              <span>
                Students needing support: <strong className="text-rose-600">{stats.failedCount}</strong>
              </span>
            </div>
            <button
              onClick={() => onNavigateTab("report")}
              className="text-slate-900 font-medium hover:underline text-xs"
            >
              View Report &rarr;
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
