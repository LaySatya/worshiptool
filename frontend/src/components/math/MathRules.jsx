import { useState } from "react";
import {
  Sliders,
  Plus,
  Trash2,
  RotateCcw,
  Save,
  CheckCircle2,
  Sparkles,
  Calculator,
  HelpCircle,
} from "lucide-react";
import {
  DEFAULT_GRADE_RULES,
  calculateGrade,
} from "../../services/mathStorage.js";

export default function MathRules({ rules = [], onSaveRules }) {
  const [localRules, setLocalRules] = useState(rules);
  const [testScore, setTestScore] = useState(82);
  const [savedSuccess, setSavedSuccess] = useState(false);

  const handleRuleChange = (index, field, value) => {
    setLocalRules((prev) => {
      const copy = [...prev];
      copy[index] = {
        ...copy[index],
        [field]:
          field === "min" || field === "max"
            ? Math.max(0, Math.min(100, Number(value) || 0))
            : value,
      };
      return copy;
    });
    setSavedSuccess(false);
  };

  const handleAddRule = () => {
    const newRule = {
      id: "r_" + Date.now(),
      grade: "New",
      min: 0,
      max: 10,
      isPass: true,
      color: "#3b82f6",
      label: "Custom Grade",
    };
    setLocalRules((prev) => [...prev, newRule]);
    setSavedSuccess(false);
  };

  const handleDeleteRule = (index) => {
    if (localRules.length <= 1) {
      alert("អ្នកត្រូវតែរក្សាទុកកម្រិតពិន្ទុយ៉ាងហោចណាស់មួយ។");
      return;
    }
    setLocalRules((prev) => prev.filter((_, i) => i !== index));
    setSavedSuccess(false);
  };

  const handleResetDefaults = () => {
    if (confirm("Reset grade rules back to default ranges?")) {
      setLocalRules(DEFAULT_GRADE_RULES);
      onSaveRules(DEFAULT_GRADE_RULES);
      setSavedSuccess(true);
      setTimeout(() => setSavedSuccess(false), 2500);
    }
  };

  const handleSave = () => {
    onSaveRules(localRules);
    setSavedSuccess(true);
    setTimeout(() => setSavedSuccess(false), 3000);
  };

  const testResult = calculateGrade(testScore, localRules, false);

  return (
    <div className="space-y-5 font-sans">
      {/* Header Bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white p-4 rounded-xl border border-slate-200/80 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-slate-100 text-slate-700 flex items-center justify-center font-bold">
            <Sliders size={16} />
          </div>
          <div>
            <h3 className="font-semibold text-slate-900 text-sm">
              Grade Rules &amp; Score Ranges
            </h3>
            <p className="text-xs text-slate-500">
              Customize dynamic score thresholds for each grade level (A+, A, B+, B, C+, C, E, F)
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={handleResetDefaults}
            className="px-3 py-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50 text-xs flex items-center gap-1.5 transition"
            title="Reset to default Ministry standards"
          >
            <RotateCcw size={12} className="text-slate-400" />
            <span>Reset Defaults</span>
          </button>

          <button
            type="button"
            onClick={handleAddRule}
            className="px-3 py-1.5 rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-50 text-xs flex items-center gap-1.5 transition"
          >
            <Plus size={12} />
            <span>Add Rule</span>
          </button>

          <button
            type="button"
            onClick={handleSave}
            className="px-4 py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white font-medium text-xs flex items-center gap-1.5 transition shadow-xs"
          >
            {savedSuccess ? (
              <>
                <CheckCircle2 size={13} className="text-emerald-400" />
                <span>Saved</span>
              </>
            ) : (
              <>
                <Save size={13} />
                <span>Save Rules</span>
              </>
            )}
          </button>
        </div>
      </div>

      {/* Grid: Rules Table on Left, Live Score Tester on Right */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
        {/* Rules Table (8 cols) */}
        <div className="lg:col-span-8 bg-white rounded-xl border border-slate-200/80 shadow-xs overflow-hidden">
          <div className="p-3.5 border-b border-slate-100 flex items-center justify-between">
            <span className="font-semibold text-xs text-slate-800">
              Active Score Bands
            </span>
            <span className="text-xs text-slate-400 font-mono">
              {localRules.length} grade levels
            </span>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs">
              <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-semibold text-[11px]">
                <tr>
                  <th className="py-2.5 px-4 w-20">Grade</th>
                  <th className="py-2.5 px-4 w-24 text-center">Min Score</th>
                  <th className="py-2.5 px-4 w-24 text-center">Max Score</th>
                  <th className="py-2.5 px-4 w-24 text-center">Status</th>
                  <th className="py-2.5 px-4">Description</th>
                  <th className="py-2.5 px-3 w-10 text-center"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {localRules.map((rule, idx) => (
                  <tr key={rule.id || idx} className="hover:bg-slate-50/70 transition">
                    {/* Grade Name */}
                    <td className="py-2 px-4">
                      <input
                        type="text"
                        value={rule.grade}
                        onChange={(e) => handleRuleChange(idx, "grade", e.target.value)}
                        className="w-14 py-1 px-2 rounded-md border border-slate-200 font-bold font-mono text-xs text-center focus:outline-none focus:border-slate-400"
                      />
                    </td>

                    {/* Min Score */}
                    <td className="py-2 px-4 text-center">
                      <input
                        type="number"
                        min="0"
                        max="100"
                        value={rule.min}
                        onChange={(e) => handleRuleChange(idx, "min", e.target.value)}
                        className="w-16 py-1 px-2 rounded-md border border-slate-200 font-mono text-center text-xs focus:outline-none focus:border-slate-400"
                      />
                    </td>

                    {/* Max Score */}
                    <td className="py-2 px-4 text-center">
                      <input
                        type="number"
                        min="0"
                        max="100"
                        value={rule.max}
                        onChange={(e) => handleRuleChange(idx, "max", e.target.value)}
                        className="w-16 py-1 px-2 rounded-md border border-slate-200 font-mono text-center text-xs focus:outline-none focus:border-slate-400"
                      />
                    </td>

                    {/* Pass / Fail Toggle */}
                    <td className="py-2 px-4 text-center">
                      <button
                        type="button"
                        onClick={() => handleRuleChange(idx, "isPass", !rule.isPass)}
                        className={`px-2.5 py-1 rounded text-[11px] font-medium transition ${
                          rule.isPass
                            ? "bg-emerald-50 text-emerald-700 border border-emerald-200"
                            : "bg-rose-50 text-rose-700 border border-rose-200"
                        }`}
                      >
                        {rule.isPass ? "Passed" : "Failed"}
                      </button>
                    </td>

                    {/* Label / Meaning */}
                    <td className="py-2 px-4">
                      <input
                        type="text"
                        value={rule.label || ""}
                        onChange={(e) => handleRuleChange(idx, "label", e.target.value)}
                        placeholder="Description…"
                        className="w-full py-1 px-2 rounded-md border border-slate-200 text-xs focus:outline-none focus:border-slate-400"
                      />
                    </td>

                    {/* Delete */}
                    <td className="py-2 px-3 text-center">
                      <button
                        type="button"
                        onClick={() => handleDeleteRule(idx)}
                        className="p-1 rounded text-slate-400 hover:text-rose-600 transition"
                        title="Delete Rule"
                      >
                        <Trash2 size={13} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Right: Live Interactive Rule Tester (4 cols) */}
        <div className="lg:col-span-4 space-y-4">
          <div className="bg-white rounded-xl border border-slate-200/80 shadow-xs p-5">
            <div className="flex items-center gap-2 pb-3 border-b border-slate-100 mb-4">
              <Calculator size={15} className="text-slate-700" />
              <h4 className="font-semibold text-slate-900 text-sm">
                Live Score Tester
              </h4>
            </div>

            <p className="text-xs text-slate-500 mb-3">
              Type any score to test the calculated grade in real time:
            </p>

            <div className="space-y-4">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Test Score (0 – 100):
                </label>
                <input
                  type="number"
                  min="0"
                  max="100"
                  value={testScore}
                  onChange={(e) => setTestScore(Number(e.target.value) || 0)}
                  className="w-full py-2 px-3 rounded-lg border border-slate-200 text-lg font-bold font-mono text-center focus:outline-none focus:border-slate-400 bg-slate-50"
                />
              </div>

              {/* Calculated Result Card */}
              <div
                className={`p-5 rounded-xl border text-center transition-all ${
                  testResult.isPass
                    ? "bg-emerald-50/50 border-emerald-200 text-slate-900"
                    : "bg-rose-50/60 border-rose-200 text-rose-900"
                }`}
              >
                <span className="text-[11px] font-medium text-slate-500 block uppercase tracking-wider mb-1">
                  Calculated Grade
                </span>
                <div
                  className={`text-4xl font-extrabold font-mono my-1 ${
                    testResult.isPass ? "text-emerald-700" : "text-rose-600"
                  }`}
                >
                  {testResult.grade}
                </div>
                <div className="text-xs font-medium text-slate-600 mt-1">
                  {testResult.label}
                </div>
                <div className="mt-3">
                  <span
                    className={`inline-block px-3 py-1 rounded-full text-xs font-semibold ${
                      testResult.isPass
                        ? "bg-emerald-600 text-white"
                        : "bg-rose-600 text-white"
                    }`}
                  >
                    {testResult.isPass ? "✓ Passed" : "✕ Failed"}
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* Quick Info Box */}
          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 text-xs text-slate-700 space-y-1.5">
            <div className="font-semibold flex items-center gap-1.5 text-slate-900">
              <HelpCircle size={13} className="text-slate-500" />
              <span>Standard Grading Guidance</span>
            </div>
            <p className="text-[11px] text-slate-600 leading-relaxed">
              • <strong>A+ / A</strong>: 85 – 100 (Excellent)<br />
              • <strong>B+ / B</strong>: 65 – 84 (Good)<br />
              • <strong>C+ / C</strong>: 55 – 64 (Average)<br />
              • <strong>E</strong>: 50 – 54 (Passing threshold)<br />
              • <strong>F</strong>: Below 50 (Failing grade)
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}
