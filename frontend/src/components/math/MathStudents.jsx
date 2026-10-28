import { useState, useMemo } from "react";
import {
  Users,
  UserPlus,
  Search,
  Edit2,
  Trash2,
  X,
  History,
  AlertTriangle,
} from "lucide-react";

export default function MathStudents({
  students = [],
  scores = [],
  rules = [],
  onAddStudent,
  onUpdateStudent,
  onDeleteStudent,
}) {
  const [searchTerm, setSearchTerm] = useState("");
  const [genderFilter, setGenderFilter] = useState("all");
  const [modalOpen, setModalOpen] = useState(false);
  const [editingStudent, setEditingStudent] = useState(null);
  const [historyStudent, setHistoryStudent] = useState(null);
  const [deletingStudent, setDeletingStudent] = useState(null);

  // Form state
  const [formName, setFormName] = useState("");
  const [formGender, setFormGender] = useState("ប");
  const [formGradeLevel, setFormGradeLevel] = useState("Grade 10");
  const [formPhone, setFormPhone] = useState("");
  const [formNotes, setFormNotes] = useState("");

  const filtered = useMemo(() => {
    return students.filter((s) => {
      const matchName = s.name.toLowerCase().includes(searchTerm.toLowerCase());
      const matchGender =
        genderFilter === "all" ||
        (genderFilter === "male" && (s.gender === "ប" || s.gender === "M")) ||
        (genderFilter === "female" && (s.gender === "ស" || s.gender === "F"));
      return matchName && matchGender;
    });
  }, [students, searchTerm, genderFilter]);

  const openAddModal = () => {
    setEditingStudent(null);
    setFormName("");
    setFormGender("ប");
    setFormGradeLevel("Grade 10");
    setFormPhone("");
    setFormNotes("");
    setModalOpen(true);
  };

  const openEditModal = (st) => {
    setEditingStudent(st);
    setFormName(st.name);
    setFormGender(st.gender || "ប");
    setFormGradeLevel(st.gradeLevel || "Grade 10");
    setFormPhone(st.phone || "");
    setFormNotes(st.notes || "");
    setModalOpen(true);
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!formName.trim()) return;

    const payload = {
      name: formName.trim(),
      gender: formGender,
      gradeLevel: formGradeLevel,
      phone: formPhone.trim(),
      notes: formNotes.trim(),
    };

    if (editingStudent) {
      onUpdateStudent({ ...editingStudent, ...payload });
    } else {
      onAddStudent(payload);
    }
    setModalOpen(false);
  };

  const handleConfirmDelete = () => {
    if (deletingStudent) {
      onDeleteStudent(deletingStudent.id);
      setDeletingStudent(null);
    }
  };

  const getStudentScoresHistory = (studentId) => {
    return scores.filter((sc) => sc.studentId === studentId);
  };

  return (
    <div className="space-y-4 font-sans">
      {/* Top action toolbar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white p-4 rounded-xl border border-slate-200/80 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-slate-100 text-slate-700 flex items-center justify-center font-bold">
            <Users size={16} />
          </div>
          <div>
            <h3 className="font-semibold text-slate-900 text-sm">
              Students List
            </h3>
            <p className="text-xs text-slate-500">
              {students.length} students enrolled ({students.filter((s) => s.gender === "ប" || s.gender === "M").length} ប, {students.filter((s) => s.gender === "ស" || s.gender === "F").length} ស)
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          {/* Search */}
          <div className="relative">
            <Search
              size={13}
              className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
            />
            <input
              type="text"
              placeholder="Search student name…"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="pl-8 pr-3 py-1.5 text-xs rounded-lg border border-slate-200 bg-slate-50 focus:bg-white focus:border-slate-400 w-44 sm:w-56"
            />
          </div>

          {/* Gender Filter */}
          <select
            value={genderFilter}
            onChange={(e) => setGenderFilter(e.target.value)}
            className="text-xs px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-700"
          >
            <option value="all">All</option>
            <option value="male">ប</option>
            <option value="female">ស</option>
          </select>

          {/* Add Student Button */}
          <button
            onClick={openAddModal}
            className="px-3.5 py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white font-medium text-xs flex items-center gap-1.5 shadow-xs transition"
          >
            <UserPlus size={13} />
            <span>Add Student</span>
          </button>
        </div>
      </div>

      {/* Students Table */}
      <div className="bg-white rounded-xl border border-slate-200/80 shadow-xs overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-semibold text-[11px]">
              <tr>
                <th className="py-2.5 px-4 w-12 text-center">#</th>
                <th className="py-2.5 px-4">Student Name</th>
                <th className="py-2.5 px-4 w-20 text-center">ភេទ</th>
                <th className="py-2.5 px-4 w-32">Class</th>
                <th className="py-2.5 px-4 w-36">Contact</th>
                <th className="py-2.5 px-4">Notes</th>
                <th className="py-2.5 px-4 w-28 text-center">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={7} className="py-12 text-center text-slate-400">
                    <Users size={28} className="mx-auto mb-2 opacity-30" />
                    No students found matching your search
                  </td>
                </tr>
              ) : (
                filtered.map((st, idx) => {
                  const isMale = st.gender === "ប" || st.gender === "M";
                  return (
                    <tr key={st.id} className="hover:bg-slate-50/70 transition">
                      <td className="py-3 px-4 text-center font-mono text-slate-400 text-[11px]">
                        {idx + 1}
                      </td>
                      <td className="py-3 px-4">
                        <span className="font-semibold text-slate-900 text-xs font-khmer">
                          {st.name}
                        </span>
                      </td>
                      <td className="py-3 px-4 text-center">
                        <span className="inline-block px-2.5 py-0.5 rounded text-xs font-semibold bg-slate-100 text-slate-800 font-khmer">
                          {isMale ? "ប" : "ស"}
                        </span>
                      </td>
                      <td className="py-3 px-4 text-slate-600">
                        {st.gradeLevel || "Grade 10"}
                      </td>
                      <td className="py-3 px-4 text-slate-500 font-mono text-[11px]">
                        {st.phone || "—"}
                      </td>
                      <td className="py-3 px-4 text-slate-400 truncate max-w-xs text-[11px]">
                        {st.notes || "—"}
                      </td>
                      <td className="py-3 px-4 text-center">
                        <div className="flex items-center justify-center gap-1">
                          <button
                            type="button"
                            onClick={() => setHistoryStudent(st)}
                            className="p-1.5 rounded text-slate-400 hover:text-slate-800 hover:bg-slate-100 transition"
                            title="Score History"
                          >
                            <History size={13} />
                          </button>
                          <button
                            type="button"
                            onClick={() => openEditModal(st)}
                            className="p-1.5 rounded text-slate-400 hover:text-slate-800 hover:bg-slate-100 transition"
                            title="Edit Student"
                          >
                            <Edit2 size={13} />
                          </button>
                          <button
                            type="button"
                            onClick={() => setDeletingStudent(st)}
                            className="p-1.5 rounded text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition"
                            title="Delete Student"
                          >
                            <Trash2 size={13} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Add / Edit Student Modal */}
      {modalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40 backdrop-blur-xs animate-fade-in">
          <div className="bg-white rounded-xl max-w-md w-full p-5 shadow-xl border border-slate-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <h3 className="font-semibold text-slate-900 text-sm">
                {editingStudent ? "Edit Student Details" : "Add New Student"}
              </h3>
              <button
                onClick={() => setModalOpen(false)}
                className="text-slate-400 hover:text-slate-600 p-1"
              >
                <X size={15} />
              </button>
            </div>

            <form onSubmit={handleSubmit} className="space-y-3.5 mt-4 text-xs">
              <div>
                <label className="block font-medium text-slate-700 mb-1">
                  Full Name (Khmer) *
                </label>
                <input
                  type="text"
                  required
                  placeholder="e.g. អោ ចាន់អន"
                  value={formName}
                  onChange={(e) => setFormName(e.target.value)}
                  className="w-full px-3 py-2 rounded-lg border border-slate-200 text-xs font-khmer focus:outline-none focus:border-slate-400"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-medium text-slate-700 mb-1">
                    ភេទ (Gender) *
                  </label>
                  <select
                    value={formGender}
                    onChange={(e) => setFormGender(e.target.value)}
                    className="w-full px-3 py-2 rounded-lg border border-slate-200 text-xs focus:outline-none focus:border-slate-400 bg-white font-khmer font-medium"
                  >
                    <option value="ប">ប</option>
                    <option value="ស">ស</option>
                  </select>
                </div>

                <div>
                  <label className="block font-medium text-slate-700 mb-1">
                    Grade Level (7 - 12) *
                  </label>
                  <select
                    value={formGradeLevel}
                    onChange={(e) => setFormGradeLevel(e.target.value)}
                    className="w-full px-3 py-2 rounded-lg border border-slate-200 text-xs focus:outline-none focus:border-slate-400 bg-white font-medium"
                  >
                    <option value="Grade 7">Grade 7 (ថ្នាក់ទី ៧)</option>
                    <option value="Grade 8">Grade 8 (ថ្នាក់ទី ៨)</option>
                    <option value="Grade 9">Grade 9 (ថ្នាក់ទី ៩)</option>
                    <option value="Grade 10">Grade 10 (ថ្នាក់ទី ១០)</option>
                    <option value="Grade 11">Grade 11 (ថ្នាក់ទី ១១)</option>
                    <option value="Grade 12">Grade 12 (ថ្នាក់ទី ១២)</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">
                  Contact Phone
                </label>
                <input
                  type="text"
                  placeholder="012 345 678"
                  value={formPhone}
                  onChange={(e) => setFormPhone(e.target.value)}
                  className="w-full px-3 py-2 rounded-lg border border-slate-200 text-xs font-mono focus:outline-none focus:border-slate-400"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">
                  Notes
                </label>
                <textarea
                  rows={2}
                  placeholder="Optional remarks…"
                  value={formNotes}
                  onChange={(e) => setFormNotes(e.target.value)}
                  className="w-full px-3 py-2 rounded-lg border border-slate-200 text-xs focus:outline-none focus:border-slate-400 resize-none"
                />
              </div>

              <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setModalOpen(false)}
                  className="px-3.5 py-2 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50 transition"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-white font-medium transition"
                >
                  {editingStudent ? "Save Changes" : "Create Student"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Delete Confirmation Modal */}
      {deletingStudent && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40 backdrop-blur-xs animate-fade-in">
          <div className="bg-white rounded-xl max-w-sm w-full p-5 shadow-xl border border-slate-200">
            <div className="flex items-center gap-3 mb-3">
              <div className="w-9 h-9 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center flex-shrink-0">
                <AlertTriangle size={17} />
              </div>
              <div>
                <h3 className="font-semibold text-slate-900 text-sm">
                  Confirm Delete
                </h3>
                <p className="text-xs text-slate-500">
                  Are you sure you want to delete this student?
                </p>
              </div>
            </div>

            <div className="p-3 bg-slate-50 rounded-lg border border-slate-200 text-xs mb-4">
              <div className="font-bold text-slate-900 font-khmer">
                {deletingStudent.name}
              </div>
              <div className="text-slate-500 text-[11px] mt-0.5">
                {deletingStudent.gradeLevel} • ភេទ: {deletingStudent.gender}
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100 text-xs">
              <button
                type="button"
                onClick={() => setDeletingStudent(null)}
                className="px-3 py-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleConfirmDelete}
                className="px-3.5 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-700 text-white font-medium"
              >
                Delete Student
              </button>
            </div>
          </div>
        </div>
      )}

      {/* History Modal */}
      {historyStudent && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40 backdrop-blur-xs animate-fade-in">
          <div className="bg-white rounded-xl max-w-md w-full p-5 shadow-xl border border-slate-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div>
                <h3 className="font-semibold text-slate-900 text-sm font-khmer">
                  {historyStudent.name}
                </h3>
                <p className="text-xs text-slate-500">
                  ភេទ: {historyStudent.gender} • {historyStudent.gradeLevel || "Grade 10"}
                </p>
              </div>
              <button
                onClick={() => setHistoryStudent(null)}
                className="text-slate-400 hover:text-slate-600 p-1"
              >
                <X size={15} />
              </button>
            </div>

            <div className="py-4">
              {(() => {
                const hist = getStudentScoresHistory(historyStudent.id);
                if (!hist.length) {
                  return (
                    <div className="p-8 text-center text-xs text-slate-400">
                      No score history recorded for this student yet
                    </div>
                  );
                }
                return (
                  <div className="space-y-2">
                    {hist.map((sc, i) => (
                      <div
                        key={i}
                        className="flex items-center justify-between p-2.5 rounded-lg bg-slate-50 border border-slate-100 text-xs"
                      >
                        <span className="font-medium text-slate-700">
                          {sc.monthKey || "Current Month"}
                        </span>
                        <div className="flex items-center gap-3">
                          <span className="font-bold font-mono text-slate-900">
                            {sc.isAbsent ? "Absent" : `${sc.totalScore} pts`}
                          </span>
                        </div>
                      </div>
                    ))}
                  </div>
                );
              })()}
            </div>

            <div className="flex justify-end pt-3 border-t border-slate-100">
              <button
                onClick={() => setHistoryStudent(null)}
                className="px-3.5 py-1.5 rounded-lg border border-slate-200 text-slate-700 text-xs hover:bg-slate-50"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
