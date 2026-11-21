import { useCallback } from "react";
import {
  ArrowLeft,
  BookOpen,
  ChevronRight,
  Clock,
  GraduationCap,
  Library,
  ShieldCheck,
} from "lucide-react";
import {
  BOOKS,
  GRADES,
  SUBJECTS,
  SUBJECT_STYLES,
  booksFor,
  subjectById,
} from "../../books.js";
import BookReader from "./BookReader.jsx";

const LAST_PAGE_KEY = "worshiptool:bookpage:";

/** Total number of books across every grade and subject. */
function totalBooks() {
  return GRADES.reduce(
    (sum, g) =>
      sum +
      SUBJECTS.reduce((n, s) => n + booksFor(g.id, s.id).length, 0),
    0,
  );
}

function lastPageOf(book) {
  try {
    const v = parseInt(
      localStorage.getItem(LAST_PAGE_KEY + book.id) || "1",
      10,
    );
    return Number.isFinite(v) && v > 0 ? v : 1;
  } catch {
    return 1;
  }
}

function SubjectIcon({ subject, size = 16 }) {
  const style = SUBJECT_STYLES[subject.color];
  return (
    <div
      className={`w-8 h-8 rounded-lg flex items-center justify-center flex-shrink-0 ${style.icon}`}
    >
      <BookOpen size={size} />
    </div>
  );
}

function BookCard({ grade, subject, book, onOpen }) {
  const style = SUBJECT_STYLES[subject.color];
  const last = lastPageOf(book);
  const levelKm = book.levelKm || null;
  const levelEn = book.levelEn || null;
  return (
    <button
      type="button"
      onClick={onOpen}
      className="group text-left bg-white border border-slate-200 rounded-xl p-4 hover:border-brand-300 hover:shadow-md hover:-translate-y-0.5 transition-all flex flex-col gap-3"
    >
      <div className="flex items-start justify-between gap-2">
        <div className="flex items-center gap-2.5 min-w-0">
          <SubjectIcon subject={subject} />
          <div className="min-w-0">
            <p className="text-sm font-bold text-slate-800 leading-tight truncate">
              {levelEn ? `${subject.label} — ${levelEn}` : subject.label}
            </p>
            <p className="text-[11px] text-slate-400 truncate">
              {levelKm ? `${subject.labelKm} · ${levelKm}` : subject.labelKm}
            </p>
          </div>
        </div>
        <span
          className={`text-[10px] font-semibold px-1.5 py-0.5 rounded border ${style.chip} whitespace-nowrap`}
        >
          {grade.label}
        </span>
      </div>

      <div className="flex items-center justify-between border-t border-slate-100 pt-2.5 mt-auto">
        <span className="inline-flex items-center gap-1 text-[11px] text-slate-500 font-medium">
          <Clock size={11} />
          {last > 1 ? `Resume p. ${last}` : "Start reading"}
        </span>
        <span className="inline-flex items-center gap-0.5 text-[11px] font-semibold text-brand-600 group-hover:gap-1.5 transition-all">
          Open <ChevronRight size={13} />
        </span>
      </div>
    </button>
  );
}

function SubjectCards({ grade, onOpen, highlightSubject }) {
  return SUBJECTS.map((subject) => {
    const list = booksFor(grade.id, subject.id);
    if (!list.length) return null;
    return (
      <div
        key={subject.id}
        className={
          highlightSubject === subject.id
            ? "ring-2 ring-brand-300 rounded-xl"
            : "contents"
        }
      >
        {list.map((book) => (
          <BookCard
            key={book.id}
            grade={grade}
            subject={subject}
            book={book}
            onOpen={() => onOpen(grade.id, subject.id, book.id)}
          />
        ))}
      </div>
    );
  });
}

function GradeSection({ grade, onOpen, highlightSubject }) {
  return (
    <section>
      <div className="flex items-center gap-2.5 mb-3">
        <div className="w-9 h-9 rounded-xl bg-slate-900 text-white flex items-center justify-center shadow-xs">
          <GraduationCap size={17} />
        </div>
        <div>
          <h3 className="text-sm font-bold text-slate-800 leading-tight">
            {grade.label}
          </h3>
          <p className="text-[11px] text-slate-400 leading-tight">
            {grade.labelKm}
          </p>
        </div>
        <div className="flex-1 h-px bg-slate-200 ml-2" />
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-3">
        <SubjectCards
          grade={grade}
          onOpen={onOpen}
          highlightSubject={highlightSubject}
        />
      </div>
    </section>
  );
}

/**
 * "Books" navigation — every textbook categorised by grade and subject.
 *
 * Props:
 *   selection = { grade, subject, book }  (subject null -> grade view)
 *   onSelect({ grade, subject, book })    persists the selection to the URL hash
 */
export default function BooksLibrary({ selection, onSelect }) {
  const gradeId = selection?.grade || null;
  const subjectId = selection?.subject || null;
  const bookId = selection?.book || null;

  const grade = GRADES.find((g) => g.id === gradeId) || null;
  const subject = subjectById(subjectId);
  const options = grade && subject ? booksFor(grade.id, subject.id) : [];
  const book = bookId
    ? options.find((b) => b.id === bookId) || options[0] || null
    : options.length === 1
      ? options[0]
      : null;

  const open = useCallback(
    (g, s, b = null) => onSelect({ grade: g, subject: s, book: b }),
    [onSelect],
  );

  // ── Reader view ────────────────────────────────────────────────
  if (grade && subject && book) {
    return (
      <div className="flex-1 min-h-0 flex flex-col p-3 sm:p-4 bg-slate-50/60">
        <div className="flex-1 min-h-0">
          <BookReader
            book={book}
            grade={grade}
            subject={subject}
            onBack={() =>
              open(grade.id, options.length > 1 ? subject.id : null)
            }
            onSelectBook={
              options.length > 1
                ? (b) => open(grade.id, subject.id, b.id)
                : undefined
            }
            options={options}
          />
        </div>
      </div>
    );
  }

  // ── Subject with several books but none chosen yet ─────────────
  if (grade && subject && options.length > 1) {
    return (
      <div className="flex-1 overflow-y-auto p-4 sm:p-6 bg-slate-50/60">
        <div className="max-w-4xl mx-auto space-y-5">
          <button
            type="button"
            onClick={() => open(grade.id, null)}
            className="inline-flex items-center gap-1.5 h-8 px-3 rounded-lg border border-slate-200 bg-white text-xs font-semibold text-slate-600 hover:bg-slate-50 transition"
          >
            <ArrowLeft size={14} /> {grade.labelKm}
          </button>
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
            <h2 className="text-base font-bold text-slate-800">
              {subject.label} — {grade.label}
            </h2>
            <p className="text-xs text-slate-400 mb-4">
              {subject.labelKm} — choose a level
            </p>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {options.map((b) => (
                <BookCard
                  key={b.id}
                  grade={grade}
                  subject={subject}
                  book={b}
                  onOpen={() => open(grade.id, subject.id, b.id)}
                />
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ── Single grade view (its three subjects) ─────────────────────
  if (grade) {
    const others = GRADES.filter((g) => g.id !== grade.id);
    return (
      <div className="flex-1 overflow-y-auto p-4 sm:p-6 bg-slate-50/60">
        <div className="max-w-5xl mx-auto space-y-5">
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => open(null, null)}
              className="inline-flex items-center gap-1.5 h-8 px-3 rounded-lg border border-slate-200 bg-white text-xs font-semibold text-slate-600 hover:bg-slate-50 transition"
            >
              <ArrowLeft size={14} /> All grades
            </button>
            {others.map((g) => (
              <button
                key={g.id}
                type="button"
                onClick={() => open(g.id, null)}
                className="inline-flex items-center gap-1 h-8 px-3 rounded-lg border border-slate-200 bg-white text-xs font-semibold text-slate-500 hover:bg-slate-50 hover:text-slate-800 transition"
              >
                {g.labelKm}
              </button>
            ))}
          </div>

          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
            <div className="flex items-center gap-3 mb-4">
              <div className="w-11 h-11 rounded-xl bg-slate-900 text-white flex items-center justify-center font-bold">
                {grade.label.replace(/\D/g, "")}
              </div>
              <div>
                <h2 className="text-base font-bold text-slate-800">
                  {grade.label} books
                </h2>
                <p className="text-xs text-slate-400">
                  {grade.labelKm} — choose a subject to start teaching
                </p>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
              <SubjectCards grade={grade} onOpen={open} />
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ── Library view (all grades x subjects) ──────────────────────
  return (
    <div className="flex-1 overflow-y-auto p-4 sm:p-6 bg-slate-50/60">
      <div className="max-w-6xl mx-auto space-y-6">
        <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="w-11 h-11 rounded-xl bg-brand-500 text-white flex items-center justify-center shadow-xs">
              <Library size={20} />
            </div>
            <div>
              <h2 className="text-base font-bold text-slate-800">
                Textbook Library
              </h2>
              <p className="text-xs text-slate-400">
                Khmer · Mathematics · English — Grades 4 to 6 and 12, ready to
                teach
              </p>
            </div>
          </div>
          <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-emerald-50 border border-emerald-200 text-[11px] font-semibold text-emerald-700">
            <ShieldCheck size={13} /> {totalBooks()} books available
          </span>
        </div>

        {GRADES.map((g) => (
          <GradeSection key={g.id} grade={g} onOpen={open} />
        ))}

        <p className="text-[11px] text-slate-400 text-center pt-2">
          Books are served through the {""}
          <span className="font-semibold text-slate-500">SYASKYTOOL</span>{" "}
          reader proxy. Your last page is remembered for each book.
        </p>
      </div>
    </div>
  );
}
