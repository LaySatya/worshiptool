/**
 * Navigation state that survives a page refresh.
 *
 * The app is a single-page app served from "/" (the backend catch-all serves
 * index.html for every path), so the active page is kept in the URL hash:
 *
 *   #/bible
 *   #/church/melody
 *   #/school/math/scores
 *   #/books/grade-4/khmer
 *
 * On reload the hash is read back and the same navigation is restored.
 */

// Pages that live under the "Church" group in the sidebar.
export const CHURCH_PAGES = [
  "bible",
  "song",
  "cover",
  "melody",
  "extract",
  "editor",
  "merge",
];

// Pages that live under the "School" group in the sidebar.
export const SCHOOL_PAGES = ["math", "books"];

// Pages that are still being maintained (shown with a badge).
export const MAINTAINING_PAGES = ["melody", "extract", "editor", "merge"];

export const DEFAULT_NAV = {
  page: "bible",
  mathTab: "dashboard",
  books: { grade: null, subject: null, book: null },
};

/**
 * Parse a location.hash into navigation state.
 * Anything unrecognised falls back to the default so a hand-edited hash
 * can never leave the app in a broken state.
 */
export function parseHash(hash) {
  const raw = String(hash || "").replace(/^#\/?/, "");
  const [section, second, third] = raw.split("/").filter(Boolean);
  const nav = { ...DEFAULT_NAV, books: { ...DEFAULT_NAV.books } };

  if (!section) return nav;

  if (section === "church") {
    if (CHURCH_PAGES.includes(second)) nav.page = second;
  } else if (section === "school") {
    if (SCHOOL_PAGES.includes(second)) nav.page = second;
    if (second === "math" && third) nav.mathTab = third;
  } else if (section === "books") {
    nav.page = "books";
    const [grade, subject, book] = raw.split("/").filter(Boolean).slice(1);
    if (grade) nav.books.grade = grade;
    if (subject) nav.books.subject = subject;
    if (book) nav.books.book = book;
  } else if (CHURCH_PAGES.includes(section) || SCHOOL_PAGES.includes(section)) {
    // Legacy flat hashes such as "#/bible" or "#/math".
    nav.page = section;
    if (section === "math" && second) nav.mathTab = second;
  }

  return nav;
}

/** Build the hash string for a navigation state. */
export function buildHash({ page, mathTab, books }) {
  if (page === "bible") return "#/bible";
  if (page === "math") return `#/school/math/${mathTab || "dashboard"}`;
  if (page === "books") {
    const b = books || {};
    if (b.grade && b.subject && b.book)
      return `#/books/${b.grade}/${b.subject}/${b.book}`;
    if (b.grade && b.subject) return `#/books/${b.grade}/${b.subject}`;
    if (b.grade) return `#/books/${b.grade}`;
    return "#/books";
  }
  if (CHURCH_PAGES.includes(page)) return `#/church/${page}`;
  return "#/bible";
}
