/**
 * Textbook library — FlipHTML5 readers grouped by grade and subject.
 *
 * Each book opens in a same-origin proxy (/api/flipbook) so the pages render
 * inside the app instead of being blocked by the publisher's frame policy.
 */

export const SUBJECTS = [
  { id: "khmer", label: "Khmer", labelKm: "ភាសាខ្មែរ", color: "rose" },
  { id: "math", label: "Mathematics", labelKm: "គណិតវិទ្យា", color: "sky" },
  { id: "english", label: "English", labelKm: "អង់គ្លេស", color: "amber" },
];

export const GRADES = [
  { id: "grade-4", label: "Grade 4", labelKm: "ថ្នាក់ទី ៤" },
  { id: "grade-5", label: "Grade 5", labelKm: "ថ្នាក់ទី ៥" },
  { id: "grade-6", label: "Grade 6", labelKm: "ថ្នាក់ទី ៦" },
  { id: "grade-12", label: "Grade 12", labelKm: "ថ្នាក់ទី ១២" },
];

// gradeId -> subjectId -> book
export const BOOKS = {
  "grade-4": {
    khmer: {
      id: "riaxr-gura",
      url: "https://online.fliphtml5.com/riaxr/gura/",
    },
    math: { id: "riaxr-nhni", url: "https://online.fliphtml5.com/riaxr/nhni/" },
    english: {
      id: "qtjoy-pgpo",
      url: "https://online.fliphtml5.com/qtjoy/pgpo/",
    },
  },
  "grade-5": {
    khmer: {
      id: "riaxr-uxtz",
      url: "https://online.fliphtml5.com/riaxr/uxtz/",
    },
    math: { id: "vtzsn-jtwd", url: "https://online.fliphtml5.com/vtzsn/jtwd/" },
    english: {
      id: "qtjoy-eser",
      url: "https://online.fliphtml5.com/qtjoy/eser/",
    },
  },
  "grade-6": {
    khmer: {
      id: "vtzsn-fwmi",
      url: "https://online.fliphtml5.com/vtzsn/fwmi/",
    },
    math: { id: "vtzsn-hqme", url: "https://online.fliphtml5.com/vtzsn/hqme/" },
    // English Grade 5 and Grade 6 share the same book.
    english: {
      id: "qtjoy-eser",
      url: "https://online.fliphtml5.com/qtjoy/eser/",
    },
  },
  // Grade 12 Mathematics — Khmer textbooks, two levels of the same subject.
  "grade-12": {
    math: [
      {
        id: "mylzw-ouck",
        url: "https://online.fliphtml5.com/mylzw/ouck/",
        label: "Basic",
        labelKm: "កម្រិតមូលដ្ឋាន",
        levelKm: "កម្រិតមូលដ្ឋាន",
        levelEn: "Basic level",
      },
      {
        id: "mylzw-ibyu",
        url: "https://online.fliphtml5.com/mylzw/ibyu/",
        label: "Advanced",
        labelKm: "កម្រិតខ្ពស់",
        levelKm: "កម្រិតខ្ពស់",
        levelEn: "Advanced level",
      },
    ],
  },
};

export const SUBJECT_STYLES = {
  rose: {
    chip: "bg-rose-50 text-rose-700 border-rose-200",
    dot: "bg-rose-500",
    icon: "bg-rose-100 text-rose-600",
  },
  sky: {
    chip: "bg-sky-50 text-sky-700 border-sky-200",
    dot: "bg-sky-500",
    icon: "bg-sky-100 text-sky-600",
  },
  amber: {
    chip: "bg-amber-50 text-amber-700 border-amber-200",
    dot: "bg-amber-500",
    icon: "bg-amber-100 text-amber-600",
  },
};

export const gradeById = (id) => GRADES.find((g) => g.id === id) || null;
export const subjectById = (id) => SUBJECTS.find((s) => s.id === id) || null;
/** All books of one subject (a subject may have several volumes/levels). */
export const booksFor = (gradeId, subjectId) => {
  const entry = BOOKS[gradeId]?.[subjectId];
  if (!entry) return [];
  return Array.isArray(entry) ? entry : [entry];
};

/** First book of a subject — the default one. */
export const bookFor = (gradeId, subjectId) => booksFor(gradeId, subjectId)[0] || null;

/** More than one book exists for this subject. */
export const hasMultipleBooks = (gradeId, subjectId) =>
  booksFor(gradeId, subjectId).length > 1;

/** "https://online.fliphtml5.com/riaxr/gura/" -> "riaxr/gura" */
export const bookKey = (book) =>
  book.url
    .replace(/^https?:\/\/online\.fliphtml5\.com\//, "")
    .replace(/\/$/, "");

/** Same-origin proxy URL for a book, with an optional page number. */
export const bookSrc = (book, page = 1) =>
  `/api/flipbook/${bookKey(book)}/#p=${Math.max(1, page || 1)}`;
