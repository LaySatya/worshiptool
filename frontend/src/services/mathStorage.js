/**
 * Math Storage Service
 * Handles persistence for Students, Scores, and Dynamic Grade Rules.
 * Synchronizes with Firebase Firestore (Project: immanuel-songs)
 * with robust local cache & offline fallback.
 */

import {
  db,
  isInitialized,
  collection,
  doc,
  getDocs,
  setDoc,
  deleteDoc,
} from "../firebase.js";

export const DEFAULT_GRADE_RULES = [
  { id: "r1", grade: "A+", min: 90, max: 100, isPass: true, color: "#166534", label: "Excellent+" },
  { id: "r2", grade: "A",  min: 85, max: 89,  isPass: true, color: "#15803d", label: "Excellent" },
  { id: "r3", grade: "B+", min: 75, max: 84,  isPass: true, color: "#0369a1", label: "Very Good" },
  { id: "r4", grade: "B",  min: 65, max: 74,  isPass: true, color: "#0284c7", label: "Good" },
  { id: "r5", grade: "C+", min: 60, max: 64,  isPass: true, color: "#b45309", label: "Fairly Good" },
  { id: "r6", grade: "C",  min: 55, max: 59,  isPass: true, color: "#d97706", label: "Average" },
  { id: "r7", grade: "E",  min: 50, max: 54,  isPass: true, color: "#64748b", label: "Passing" },
  { id: "r8", grade: "F",  min: 0,  max: 49,  isPass: false, color: "#dc2626", label: "Fail" },
];

export const INITIAL_STUDENTS = [
  {
    id: "s1",
    name: "អោ ចាន់អន",
    gender: "ប",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s2",
    name: "ហូ ស្រីពី",
    gender: "ស",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s3",
    name: "ស៊ិត រក្សា",
    gender: "ប",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s4",
    name: "ស្រីមាស",
    gender: "ស",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s5",
    name: "យូ ពៅរ័ត្ន",
    gender: "ស",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s6",
    name: "ទូច វ៉ានី",
    gender: "ស",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s7",
    name: "ឌិន សុយា",
    gender: "ស",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s8",
    name: "ភក្ត្រ័វឌ្ឍី សែនឌី",
    gender: "ស",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s9",
    name: "ភីត ស្រីពៅ",
    gender: "ស",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s10",
    name: "ម៉ប់ និមល",
    gender: "ប",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s11",
    name: "ខែម ហាន់ណា",
    gender: "ស",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s12",
    name: "ជីហួត",
    gender: "ប",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s13",
    name: "យឿន ស្រីចាន់",
    gender: "ស",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
  {
    id: "s14",
    name: "សុផាត ជីងជីង",
    gender: "ស",
    gradeLevel: "ថ្នាក់ទី១០",
    phone: "",
    notes: "",
  },
];

export const INITIAL_SCORES_SAMPLE = [
  { studentId: "s1", examScore: 82, totalScore: 82, isAbsent: false },
  { studentId: "s2", examScore: 71, totalScore: 71, isAbsent: false },
  { studentId: "s3", examScore: 54, totalScore: 54, isAbsent: false },
  { studentId: "s4", examScore: 52, totalScore: 52, isAbsent: false },
  { studentId: "s5", examScore: 46, totalScore: 46, isAbsent: false },
  { studentId: "s6", examScore: 46, totalScore: 46, isAbsent: false },
  { studentId: "s7", examScore: 43, totalScore: 43, isAbsent: false },
  { studentId: "s8", examScore: 39, totalScore: 39, isAbsent: false },
  { studentId: "s9", examScore: 24, totalScore: 24, isAbsent: false },
  { studentId: "s10", examScore: 20, totalScore: 20, isAbsent: false },
  { studentId: "s11", examScore: 13, totalScore: 13, isAbsent: false },
  { studentId: "s12", examScore: 8, totalScore: 8, isAbsent: false },
  { studentId: "s13", examScore: 5, totalScore: 5, isAbsent: false },
  { studentId: "s14", examScore: 0, totalScore: 0, isAbsent: true },
];

export const KHMER_MONTHS = [
  { value: "01", label: "January (មករា)" },
  { value: "02", label: "February (កុម្ភៈ)" },
  { value: "03", label: "March (មីនា)" },
  { value: "04", label: "April (មេសា)" },
  { value: "05", label: "May (ឧសភា)" },
  { value: "06", label: "June (មិថុនា)" },
  { value: "07", label: "July (កក្កដា)" },
  { value: "08", label: "August (សីហា)" },
  { value: "09", label: "September (កញ្ញា)" },
  { value: "10", label: "October (តុលា)" },
  { value: "11", label: "November (វិច្ឆិកា)" },
  { value: "12", label: "December (ធ្នូ)" },
];

const STORAGE_KEYS = {
  STUDENTS: "churchtool_math_students",
  SCORES: "churchtool_math_scores",
  RULES: "churchtool_math_rules",
  SETTINGS: "churchtool_math_settings",
};

/**
 * Calculate grade dynamically based on score & rules
 */
export function calculateGrade(
  score,
  rules = DEFAULT_GRADE_RULES,
  isAbsent = false,
) {
  if (isAbsent) {
    return {
      grade: "Absent",
      isPass: false,
      color: "#dc2626",
      label: "Absent",
    };
  }

  const num = typeof score === "number" ? score : parseFloat(score) || 0;
  // Sort rules descending by min score
  const sorted = [...rules].sort((a, b) => b.min - a.min);

  for (const rule of sorted) {
    if (num >= rule.min && num <= rule.max) {
      return {
        grade: rule.grade,
        isPass: rule.isPass,
        color: rule.color,
        label: rule.label || rule.grade,
      };
    }
  }

  // Fallback for scores below lowest minimum
  return {
    grade: "F",
    isPass: false,
    color: "#dc2626",
    label: "Fail",
  };
}

// ── Students Storage ──────────────────────────────────────────────

export async function getStudents() {
  // Try Firestore if available
  if (isInitialized && db) {
    try {
      const snap = await getDocs(collection(db, "math_students"));
      if (!snap.empty) {
        const list = [];
        snap.forEach((d) => list.push({ id: d.id, ...d.data() }));
        localStorage.setItem(STORAGE_KEYS.STUDENTS, JSON.stringify(list));
        return list;
      }
    } catch (err) {
      console.warn("[Firestore] getStudents fallback to local:", err.message);
    }
  }

  // Local storage fallback
  const cached = localStorage.getItem(STORAGE_KEYS.STUDENTS);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  // Seed default students
  localStorage.setItem(STORAGE_KEYS.STUDENTS, JSON.stringify(INITIAL_STUDENTS));
  return INITIAL_STUDENTS;
}

export async function saveStudent(student) {
  const students = await getStudents();
  const id = student.id || "s_" + Date.now();
  const record = { ...student, id, updatedAt: new Date().toISOString() };

  const existingIdx = students.findIndex((s) => s.id === id);
  if (existingIdx >= 0) {
    students[existingIdx] = record;
  } else {
    students.push(record);
  }

  localStorage.setItem(STORAGE_KEYS.STUDENTS, JSON.stringify(students));

  // Sync to Firestore
  if (isInitialized && db) {
    try {
      await setDoc(doc(db, "math_students", id), record);
    } catch (err) {
      console.warn("[Firestore] saveStudent error:", err.message);
    }
  }

  return record;
}

export async function deleteStudent(studentId) {
  const students = await getStudents();
  const filtered = students.filter((s) => s.id !== studentId);
  localStorage.setItem(STORAGE_KEYS.STUDENTS, JSON.stringify(filtered));

  if (isInitialized && db) {
    try {
      await deleteDoc(doc(db, "math_students", studentId));
    } catch (err) {
      console.warn("[Firestore] deleteStudent error:", err.message);
    }
  }

  return true;
}

// ── Scores Storage ────────────────────────────────────────────────

export function getMonthKey(year, month) {
  const m = String(month).padStart(2, "0");
  return `${year}-${m}`;
}

export async function getScoresForMonth(year, month) {
  const monthKey = getMonthKey(year, month);

  // Try Firestore
  if (isInitialized && db) {
    try {
      const snap = await getDocs(collection(db, `math_scores_${monthKey}`));
      if (!snap.empty) {
        const list = [];
        snap.forEach((d) => list.push({ id: d.id, ...d.data() }));
        return list;
      }
    } catch (err) {
      console.warn("[Firestore] getScores fallback:", err.message);
    }
  }

  // Local storage
  const allScores = JSON.parse(
    localStorage.getItem(STORAGE_KEYS.SCORES) || "{}",
  );
  if (allScores[monthKey]) {
    return allScores[monthKey];
  }

  // If viewing current year/month sample, seed the sample scores from image ONLY for 2026-10
  if (monthKey === "2026-10") {
    const defaultScores = INITIAL_SCORES_SAMPLE.map((s) => ({
      id: `${monthKey}_${s.studentId}`,
      studentId: s.studentId,
      monthKey,
      year,
      month,
      examScore: s.examScore,
      totalScore: s.totalScore,
      isAbsent: s.isAbsent,
      notes: "",
    }));

    allScores[monthKey] = defaultScores;
    localStorage.setItem(STORAGE_KEYS.SCORES, JSON.stringify(allScores));
    return defaultScores;
  }

  // Any other month without saved records starts empty
  return [];
}

export async function saveScoresForMonth(year, month, scoresList) {
  const monthKey = getMonthKey(year, month);
  const allScores = JSON.parse(
    localStorage.getItem(STORAGE_KEYS.SCORES) || "{}",
  );
  allScores[monthKey] = scoresList;
  localStorage.setItem(STORAGE_KEYS.SCORES, JSON.stringify(allScores));

  // Sync to Firestore
  if (isInitialized && db) {
    try {
      for (const item of scoresList) {
        const docId = item.id || `${monthKey}_${item.studentId}`;
        await setDoc(doc(db, `math_scores_${monthKey}`, docId), {
          ...item,
          monthKey,
          year,
          month,
          updatedAt: new Date().toISOString(),
        });
      }
    } catch (err) {
      console.warn("[Firestore] saveScores sync error:", err.message);
    }
  }

  return scoresList;
}

// ── Grade Rules Storage ───────────────────────────────────────────

export async function getGradeRules() {
  if (isInitialized && db) {
    try {
      const snap = await getDocs(collection(db, "math_rules"));
      if (!snap.empty) {
        const list = [];
        snap.forEach((d) => list.push({ id: d.id, ...d.data() }));
        list.sort((a, b) => b.min - a.min);
        localStorage.setItem(STORAGE_KEYS.RULES, JSON.stringify(list));
        return list;
      }
    } catch (err) {
      console.warn("[Firestore] getGradeRules fallback:", err.message);
    }
  }

  const cached = localStorage.getItem(STORAGE_KEYS.RULES);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  localStorage.setItem(STORAGE_KEYS.RULES, JSON.stringify(DEFAULT_GRADE_RULES));
  return DEFAULT_GRADE_RULES;
}

export async function saveGradeRules(rules) {
  const sorted = [...rules].sort((a, b) => b.min - a.min);
  localStorage.setItem(STORAGE_KEYS.RULES, JSON.stringify(sorted));

  if (isInitialized && db) {
    try {
      for (const rule of sorted) {
        await setDoc(doc(db, "math_rules", rule.id || rule.grade), rule);
      }
    } catch (err) {
      console.warn("[Firestore] saveGradeRules error:", err.message);
    }
  }

  return sorted;
}

// ── General Settings ──────────────────────────────────────────────

export function getMathSettings() {
  const cached = localStorage.getItem(STORAGE_KEYS.SETTINGS);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }
  return {
    reportTitle: "លទ្ធផលប្រឡងគណិតវិទ្យា",
    subjectName: "គណិតវិទ្យា (Mathematics)",
    className: "ថ្នាក់ទី១០",
    teacherName: "",
    passingScore: 50,
  };
}

export function saveMathSettings(settings) {
  localStorage.setItem(STORAGE_KEYS.SETTINGS, JSON.stringify(settings));
  return settings;
}
