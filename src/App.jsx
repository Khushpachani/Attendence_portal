import React, { useState, useMemo, useRef } from "react";
import { createPortal } from "react-dom";
import ExcelJS from "exceljs";
import {
  LayoutGrid,
  Users,
  BookOpen,
  CalendarCheck,
  FileBarChart,
  Table2,
  Download,
  Plus,
  X,
  Check,
  Search,
  Stamp,
  Clock,
  Pencil,
  ArrowRightLeft,
  ShieldCheck,
  History,
  Menu,
  Upload,
} from "lucide-react";

// ---------- design tokens ----------
const COLORS = {
  ink: "#1B2A4A",
  inkSoft: "#3A4B6E",
  parchment: "#F7F4EA",
  parchmentDark: "#EFE9D8",
  brass: "#B08D57",
  brassSoft: "#D9C49A",
  present: "#3F6B4F",
  presentSoft: "#E4EDE6",
  warn: "#C98A2C",
  warnSoft: "#F6EBD8",
  absent: "#A34A3D",
  absentSoft: "#F3E4E1",
  proxy: "#7B5EA7",
  proxySoft: "#EDE6F5",
  line: "#D8D0BA",
  slate: "#6B7280",
};

// Only two ways a lecture can be marked: Present or Absent. An unmarked
// student simply has no record (shown as blank/"—" everywhere).
const STATUS_META = {
  present: { label: "Present", abbr: "P", color: "#3F6B4F", argb: "FFC6EFCE", font: "FF006100" },
  absent: { label: "Absent", abbr: "A", color: "#A34A3D", argb: "FFFFC7CE", font: "FF9C0006" },
};
const STATUS_ORDER = ["present", "absent"];

// attendance color tiers: >=75 green, 60-74 orange, <60 red
function pctColor(pct) {
  if (pct >= 75) return COLORS.present;
  if (pct >= 60) return COLORS.warn;
  return COLORS.absent;
}

const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');`;

// Mobile/responsive support. Most of this component uses inline styles, so
// rather than rewrite every style object, targeted utility classes here
// carry the responsive behavior (with !important where it needs to beat an
// inline style) for the handful of layout patterns that repeat throughout:
// multi-column grids that should stack, tables that need horizontal
// scrolling instead of squeezing, and the sidebar becoming a slide-in
// drawer on narrow screens.
const RESPONSIVE_CSS = `
  .rp-topbar { display: none; }
  .rp-scroll-x { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  @media (max-width: 860px) {
    .rp-sidebar {
      position: fixed !important;
      left: 0; top: 0; bottom: 0;
      transform: translateX(-100%);
      transition: transform 0.2s ease;
      z-index: 200;
      box-shadow: 4px 0 24px rgba(0,0,0,0.25);
    }
    .rp-sidebar.rp-open { transform: translateX(0); }
    .rp-topbar {
      display: flex; align-items: center; justify-content: space-between;
      padding: 12px 16px; background: #1B2A4A; color: #F7F4EA;
      position: sticky; top: 0; z-index: 100;
    }
    .rp-main { padding: 16px !important; }
    .rp-grid-2, .rp-grid-3, .rp-grid-4, .rp-form-grid { grid-template-columns: 1fr !important; }
    .rp-hide-mobile { display: none !important; }
    .rp-stack-mobile { flex-direction: column !important; align-items: stretch !important; }
    .rp-login-card { width: 92vw !important; max-width: 360px; padding: 24px !important; }
  }
  @media (min-width: 861px) {
    .rp-backdrop { display: none !important; }
  }
`;

// A small, cohesive set of motion — tab transitions, modal entrances,
// button/table feedback, hover lifts — applied globally via plain element
// selectors (buttons, table rows) plus a handful of utility classes for the
// spots that need something more specific (modals, cards, the sync dot).
const ANIMATION_CSS = `
  @keyframes rpFadeInUp { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes rpFadeIn { from { opacity: 0; } to { opacity: 1; } }
  @keyframes rpScaleIn { from { opacity: 0; transform: scale(0.96) translateY(4px); } to { opacity: 1; transform: scale(1) translateY(0); } }
  @keyframes rpPulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.35; } }
  @keyframes rpSpin { to { transform: rotate(360deg); } }
  @keyframes rpSlideDown { from { opacity: 0; transform: translateY(-8px); } to { opacity: 1; transform: translateY(0); } }

  button, select, input {
    transition: background-color 0.15s ease, color 0.15s ease, border-color 0.15s ease,
      box-shadow 0.15s ease, opacity 0.15s ease, transform 0.1s ease;
  }
  button:active { transform: scale(0.96); }
  select:focus, input:focus { outline: none; box-shadow: 0 0 0 3px rgba(176,141,87,0.25); }

  .rp-tab-content { animation: rpFadeInUp 0.22s ease both; }
  .rp-modal-backdrop { animation: rpFadeIn 0.15s ease both; }
  .rp-modal-card { animation: rpScaleIn 0.2s cubic-bezier(0.16, 1, 0.3, 1) both; }
  .rp-banner { animation: rpSlideDown 0.25s ease both; }

  .rp-card-hover { transition: transform 0.18s ease, box-shadow 0.18s ease; }
  .rp-card-hover:hover { transform: translateY(-3px); box-shadow: 0 10px 24px rgba(27,42,74,0.14); }

  .rp-sync-dot-saving { animation: rpPulse 1s ease-in-out infinite; }
  .rp-spin { animation: rpSpin 0.8s linear infinite; }

  .rp-nav-btn { transition: background-color 0.15s ease, color 0.15s ease, transform 0.1s ease; }
  .rp-nav-btn:hover:not([data-active="true"]) { background: rgba(247,244,234,0.08); }

  .rp-bar-fill { transition: width 0.5s cubic-bezier(0.16, 1, 0.3, 1); }

  tbody tr { transition: background-color 0.12s ease; }
  .rp-status-btn { transition: background-color 0.15s ease, color 0.15s ease, transform 0.1s ease, border-color 0.15s ease; }
  .rp-status-btn:hover { transform: translateY(-1px); }
`;

// The class periods as they actually appear on the RRU timetable, including
// the combined double-periods (e.g. lab sessions spanning two slots).
// Must match api/users.js — the single hardcoded admin account.
const ADMIN_USERNAME = "User_Admin";
const ADMIN_PASSWORD = "Pass_Admin";

const TIME_SLOTS = [
  "09:15 - 10:00",
  "10:00 - 10:45",
  "10:45 - 11:30",
  "11:30 - 12:15",
  "12:15 - 01:00",
  "02:00 - 02:45",
  "02:45 - 03:30",
  "03:30 - 04:15",
  "04:15 - 05:00",
  "09:15 - 10:45",
  "10:45 - 12:15",
  "02:00 - 03:30",
  "02:45 - 04:15",
];

// ---------- seed data ----------
// Real roster: RRU MSc CS&DF + PGD, Semester 3, 2026-27 — taken from the
// uploaded timetable & attendance workbooks. The timetable shows two
// elective slots shared across subjects: "OS&MF/SCS" (pick 1 of 2) and
// "CS&F/SMF/ACS" (pick 1 of 3). The remaining 4 subjects (RM, CL, AI&ML,
// SARC) are compulsory for everyone.
const ELECTIVE_GROUPS = [
  { id: "G1", label: "Elective Group 1", pick: "Choose 1 of 2" },
  { id: "G2", label: "Elective Group 2", pick: "Choose 1 of 3" },
];

const seedSubjects = [
  { id: "S1", code: "RM", name: "Research Methodology", type: "core", program: "Both", faculty: "Mr. Ashish Revar (AGR)", defaultTime: "09:15 - 10:00" },
  { id: "S2", code: "CL", name: "Cyber Law", type: "core", program: "Both", faculty: "Mr. Vivek Joshi (VJ)", defaultTime: "03:30 - 04:15" },
  { id: "S3", code: "AI&ML", name: "Artificial Intelligence and Machine Learning", type: "core", program: "Both", faculty: "Ms. Ankita Shah (AHS)", defaultTime: "10:45 - 11:30" },
  { id: "S7", code: "SARC", name: "Security Auditing, Risk and Compliance", type: "core", program: "Both", faculty: "Ms. Richa Sharma (RS)", defaultTime: "09:15 - 10:00" },
  { id: "S8", code: "OS&MF", name: "OS and Multimedia Forensics", type: "elective", program: "Both", group: "G1", faculty: "Dr. Nitin Padariya (NP)", defaultTime: "12:15 - 01:00" },
  { id: "S9", code: "SCS", name: "Scripting for Cyber Security", type: "elective", program: "Both", group: "G1", faculty: "Mr. Naveen Kandwal (NK)", defaultTime: "12:15 - 01:00" },
  { id: "S4", code: "CS&F", name: "Cloud Security and Forensics", type: "elective", program: "Both", group: "G2", faculty: "Ms. Dixa Koradia (DK)", defaultTime: "10:00 - 10:45" },
  { id: "S5", code: "ACS", name: "Applied Cyber Security", type: "elective", program: "Both", group: "G2", faculty: "Ms. Simran Sharma (SDS)", defaultTime: "10:00 - 10:45" },
  { id: "S6", code: "SMF", name: "Social Media Forensics", type: "elective", program: "Both", group: "G2", faculty: "Ms. Prakruti Parmar (PRP)", defaultTime: "10:00 - 10:45" },
];

// Batch A = MSc CS&DF (53 students, emails from the Student Detail sheet).
// Batch B = PGD (9 students) — the source workbook didn't list official
// emails for this batch, so a placeholder "<enrollment>@student.rru.ac.in"
// address is used; swap these for real ones whenever you have them.
// Electives were not recorded per-student in the source data, so each
// student was assigned one subject from each group on a rotating basis —
// edit any student's picks in the Students tab to correct them.
const seedStudents = [
  { id: "25003111025081001", name: "Aachal Ashok Godse", email: "25mcsdf001@student.rru.ac.in", program: "MSc", electives: ["S9", "S6"] },
  { id: "25003111025081002", name: "Adit Sonone", email: "25mcsdf002@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081004", name: "Anuj Deepak Rathod", email: "25mcsdf004@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081005", name: "BHARGAVI RAMESH BHAI RATHVA", email: "25mcsdf005@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081006", name: "Dhruv Mukeshbhai Kachhadiya", email: "25mcsdf006@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081007", name: "DRASHTI VISHALKUMAR PATEL", email: "25mcsdf007@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081008", name: "Jaykumar Hareshbhai Patel", email: "25mcsdf008@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081009", name: "Kashish Choudhary", email: "25mcsdf009@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081010", name: "Kathiya Mayurdhvajsinh Kirtisinh", email: "25mcsdf010@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081011", name: "KHUSHI BHARAT PATEL", email: "25mcsdf011@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081013", name: "Manish Dipak Dhaygude", email: "25mcsdf013@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081014", name: "Naman Nilesh Gaur", email: "25mcsdf014@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081015", name: "Nishka Bijalkumar Bhatt", email: "25mcsdf015@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081016", name: "Parth Sanjaykumar Darji", email: "25mcsdf016@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081017", name: "PATEL AYUSH BAKULBHAI", email: "25mcsdf017@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081018", name: "Raj Karamata Lakhamanbhai", email: "25mcsdf018@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081019", name: "SMITKUMAR HARESHBHAI CHUADHARY", email: "25mcsdf019@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081020", name: "Souharda Patra", email: "25mcsdf020@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111025081021", name: "SUMIT Kumar Khuman Bhai Solanki", email: "25mcsdf021@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081023", name: "Taksh shah", email: "25mcsdf023@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081024", name: "USHVI NIRMALBHAI SHAH", email: "25mcsdf024@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081025", name: "VASOYA KESHVI RAJESHKUMAR", email: "25mcsdf025@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081026", name: "Adesh Dipeshbhai Teraiya", email: "25mcsdf026@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081027", name: "Ajay Gohali", email: "25mcsdf027@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081028", name: "Aksha Dharmesh Bhai Chudasama", email: "25mcsdf028@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081029", name: "AKSHITA BHAVSAR", email: "25mcsdf029@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081030", name: "Ashwani Kaushik", email: "25mcsdf030@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081031", name: "Garima", email: "25mcsdf031@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081032", name: "HARSH SANJAY SAWANT", email: "25mcsdf032@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081034", name: "Jay Jigneshbhai Mangroliya", email: "25mcsdf034@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081035", name: "Keval Jagdishbhai Vaghani", email: "25mcsdf035@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081036", name: "Keval Kinnarkumar Patel", email: "25mcsdf036@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081038", name: "Mitaliraj Ghosh", email: "25mcsdf038@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081039", name: "Mukta Bhausaheb Kodagapatil", email: "25mcsdf039@student.rru.ac.in", program: "MSc", electives: ["S9", "S6"] },
  { id: "25003111025081040", name: "Nitin Dinesh Bhai Prajapati", email: "25mcsdf040@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081041", name: "Parth Ghanshyambhai Jethva", email: "25mcsdf041@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081043", name: "Renuka Devi Palli", email: "25mcsdf043@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081044", name: "Riya Pradip Dalvi", email: "25mcsdf044@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081046", name: "Sourav Das", email: "25mcsdf046@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111025081047", name: "Swarnava Pal", email: "25mcsdf047@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081048", name: "Tiyasha Ray", email: "25mcsdf048@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111025081049", name: "Vrushti Paragkumar MEHTA", email: "25mcsdf049@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081050", name: "Vishvaraj Upendrasinh chauhan", email: "25mcsdf050@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081051", name: "Aayush Arun Nishad", email: "25mcsdf051@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081052", name: "Pachani Khush Bhavesbhai", email: "25mcsdf052@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081053", name: "Pranav Prakash Nair", email: "25mcsdf053@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081055", name: "Unnat Patel", email: "25mcsdf055@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081056", name: "Chaudhary suresh premaram", email: "25mcsdf056@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003311025081001", name: "KATSANDE HUDSON WALTER", email: "25mcsdf001-f@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003311025081002", name: "MUTUKU GRASIANO TERERAI", email: "25mcsdf002-f@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003311025081003", name: "Mohamed Suma", email: "25mcsdf003-f@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003311025081004", name: "Maxwell Chukwuekezie Onyebueke", email: "25mcsdf004-f@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003311025081005", name: "Opeyemi Olaolu Ayegbo", email: "25mcsdf005-f@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111018071001", name: "Harshkumar Jagdishbhai Patel", email: "25003111018071001@student.rru.ac.in", program: "PGD", electives: ["S8", "S4"] },
  { id: "25003111018071003", name: "Radhika Utpal Yagnik", email: "25003111018071003@student.rru.ac.in", program: "PGD", electives: ["S8", "S4"] }, // no real choice data found — defaulted, please correct
  { id: "25003111018071004", name: "Sanskruti Sandeep Patil", email: "25003111018071004@student.rru.ac.in", program: "PGD", electives: ["S8", "S4"] }, // no real choice data found — defaulted, please correct
  { id: "25003111018071005", name: "SHIVAM DWIVEDI", email: "25003111018071005@student.rru.ac.in", program: "PGD", electives: ["S8", "S5"] },
  { id: "25003111018071006", name: "Abhishek Bharat Parmar", email: "25003111018071006@student.rru.ac.in", program: "PGD", electives: ["S9", "S4"] },
  { id: "25003111018071007", name: "Anshraj Navdeepbhai Dodiya", email: "25003111018071007@student.rru.ac.in", program: "PGD", electives: ["S8", "S4"] },
  { id: "25003111018071008", name: "Jesika khurshed Turel", email: "25003111018071008@student.rru.ac.in", program: "PGD", electives: ["S8", "S6"] },
  { id: "25003111018071009", name: "Krunal Pareshkumar Mewada", email: "25003111018071009@student.rru.ac.in", program: "PGD", electives: ["S8", "S4"] },
  { id: "25003111018071010", name: "Riya jagatbahi Sojitra", email: "25003111018071010@student.rru.ac.in", program: "PGD", electives: ["S8", "S6"] },
];

function pastDates(n) {
  const out = [];
  const d = new Date();
  d.setDate(d.getDate() - 1);
  while (out.length < n) {
    if (d.getDay() !== 0) out.push(d.toISOString().slice(0, 10));
    d.setDate(d.getDate() - 1);
  }
  return out.reverse();
}

// ---------- roster import (Excel / CSV) ----------
// Minimal CSV parser that handles quoted fields (including embedded commas
// and escaped quotes) — good enough for a typical roster export, without
// pulling in another dependency alongside exceljs.
function parseCSVText(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ",") {
      row.push(field);
      field = "";
    } else if (c === "\n" || c === "\r") {
      if (c === "\r" && text[i + 1] === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += c;
    }
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows.filter((r) => r.some((cell) => String(cell).trim() !== ""));
}

// Header aliases so the import isn't picky about exact column naming —
// matches case-insensitively against any of these per field.
const IMPORT_COLUMN_ALIASES = {
  id: ["enrollment no", "enrollment no.", "enrollment number", "enrollment", "id", "student id", "roll no", "roll number"],
  name: ["name", "student name", "full name"],
  email: ["email", "email address", "email id"],
  program: ["program", "course", "programme"],
  electives: ["electives", "elective subjects", "elective"],
};

function findColumnIndex(headerRow, aliases) {
  const normalized = headerRow.map((h) => String(h || "").trim().toLowerCase());
  for (const alias of aliases) {
    const idx = normalized.indexOf(alias);
    if (idx !== -1) return idx;
  }
  return -1;
}

function normalizeProgram(value) {
  const v = String(value || "").trim().toLowerCase();
  if (v.startsWith("pgd") || v.includes("diploma")) return "PGD";
  if (v.startsWith("msc") || v.includes("master")) return "MSc";
  return value ? String(value).trim() : "";
}

// Turns raw sheet/CSV rows (array-of-arrays, first row = header) into
// student objects, matching an "Electives" column's text against known
// subjects by code or name.
function rowsToStudents(rows, subjects) {
  if (!rows.length) return [];
  const header = rows[0];
  const col = {
    id: findColumnIndex(header, IMPORT_COLUMN_ALIASES.id),
    name: findColumnIndex(header, IMPORT_COLUMN_ALIASES.name),
    email: findColumnIndex(header, IMPORT_COLUMN_ALIASES.email),
    program: findColumnIndex(header, IMPORT_COLUMN_ALIASES.program),
    electives: findColumnIndex(header, IMPORT_COLUMN_ALIASES.electives),
  };
  const matchSubject = (text) => {
    const t = text.trim().toLowerCase();
    if (!t) return null;
    const found = subjects.find((s) => s.code.toLowerCase() === t || s.name.toLowerCase() === t);
    return found ? found.id : null;
  };

  return rows.slice(1).map((r) => {
    const electivesRaw = col.electives !== -1 ? String(r[col.electives] || "") : "";
    const electives = electivesRaw
      .split(/[,;]/)
      .map((s) => matchSubject(s))
      .filter(Boolean);
    return {
      id: col.id !== -1 ? String(r[col.id] || "").trim() : "",
      name: col.name !== -1 ? String(r[col.name] || "").trim() : "",
      email: col.email !== -1 ? String(r[col.email] || "").trim() : "",
      program: col.program !== -1 ? normalizeProgram(r[col.program]) : "",
      electives,
    };
  });
}

async function parseImportFile(file, subjects) {
  const ext = file.name.split(".").pop().toLowerCase();
  let rows;
  if (ext === "csv") {
    const text = await file.text();
    rows = parseCSVText(text);
  } else {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];
    rows = [];
    ws.eachRow({ includeEmpty: false }, (row) => {
      // row.values is 1-indexed with a leading empty slot in ExcelJS
      rows.push(row.values.slice(1).map((v) => (v && typeof v === "object" && "text" in v ? v.text : v)));
    });
  }
  return rowsToStudents(rows, subjects);
}

// ---------- attendance register import (multi-sheet Excel) ----------
// Reads a workbook shaped like a real attendance register: one sheet per
// subject (sheet name = subject code, e.g. "RM", "OS&MF", "SMF"), with
// dates as columns and a boolean/True-False (or P/A) grid of students ×
// dates. Auto-detects the header row and date columns rather than assuming
// a fixed layout, and treats two same-day columns as separate sessions.
const norm = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");

function cellToStatus(val) {
  if (val === null || val === undefined || val === "") return null;
  if (typeof val === "boolean") return val ? "present" : "absent";
  const s = String(val).trim().toLowerCase();
  if (["true", "p", "present", "1", "yes", "y"].includes(s)) return "present";
  if (["false", "a", "absent", "0", "no", "n"].includes(s)) return "absent";
  return null;
}

async function parseAttendanceWorkbook(file, subjects) {
  const buffer = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const records = []; // {enroll, subjectId, date, slot, status}
  const sessions = []; // {subjectId, date, slot, time}
  const matchedSheets = [];
  const unmatchedSheets = [];

  wb.worksheets.forEach((ws) => {
    const sheetNorm = norm(ws.name);
    const subject = subjects.find((s) => norm(s.code) === sheetNorm) || subjects.find((s) => norm(s.name) === sheetNorm);
    if (!subject) {
      unmatchedSheets.push(ws.name);
      return;
    }

    // find the header row by scanning for a cell containing "enrollment"
    let headerRow = -1,
      enrollCol = -1;
    for (let r = 1; r <= Math.min(12, ws.rowCount); r++) {
      for (let c = 1; c <= ws.columnCount; c++) {
        const cellText = norm(ws.getRow(r).getCell(c).value);
        if (cellText.includes("enrollment")) {
          headerRow = r;
          enrollCol = c;
          break;
        }
      }
      if (headerRow !== -1) break;
    }
    if (headerRow === -1) {
      unmatchedSheets.push(`${ws.name} (couldn't find an "Enrollment" column)`);
      return;
    }

    // date columns: scan the couple of rows above/at the header for actual
    // Date-typed cells to the right of the enrollment column
    const dateCols = []; // {col, date, time}
    for (let scanRow = Math.max(1, headerRow - 2); scanRow <= headerRow; scanRow++) {
      for (let c = enrollCol + 1; c <= ws.columnCount; c++) {
        const cell = ws.getRow(scanRow).getCell(c);
        if (cell.value instanceof Date) {
          const dateStr = cell.value.toISOString().slice(0, 10);
          const timeVal = ws.getRow(headerRow).getCell(c).value;
          dateCols.push({ col: c, date: dateStr, time: timeVal ? String(timeVal).trim() : "" });
        }
      }
    }
    if (dateCols.length === 0) {
      unmatchedSheets.push(`${ws.name} (couldn't find any date columns)`);
      return;
    }

    // same date appearing more than once in a sheet = separate sessions
    const seenDates = {};
    dateCols.forEach((dc) => {
      seenDates[dc.date] = (seenDates[dc.date] || 0) + 1;
      dc.slot = String(seenDates[dc.date]);
      sessions.push({ subjectId: subject.id, date: dc.date, slot: dc.slot, time: dc.time });
    });

    matchedSheets.push(ws.name);

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const enroll = String(ws.getRow(r).getCell(enrollCol).value || "").trim();
      if (!/^\d{4,}$/.test(enroll)) continue; // skip blanks, "Total" rows, batch-label rows, etc.
      dateCols.forEach((dc) => {
        const status = cellToStatus(ws.getRow(r).getCell(dc.col).value);
        if (status) records.push({ enroll, subjectId: subject.id, date: dc.date, slot: dc.slot, status });
      });
    }
  });

  return { records, sessions, matchedSheets, unmatchedSheets };
}

// Attendance holds ONE status per student per lecture. Time and remark
// belong to the *session* (subject + date) as a whole — e.g. "faculty on
// leave, proxy lecture taken by Mr. X" or "class rescheduled to lab" — not
// to any individual student, so they live in a separate `sessions` map.
//
// This is REAL attendance data, taken directly from the uploaded
// "Sem-3 Msc Attendance" workbook (all 9 subjects, 01–08 Jul 2026) — not
// randomly generated demo data. It's only ever used the very first time
// the database is empty (see applyServerData / resetToSampleData), so it
// effectively acts as a one-time import: once it's saved to the shared
// database, every future login reads the real, growing data back, and
// this hardcoded snapshot is never touched again.
const REAL_SEED_ATTENDANCE = {
  "25003111025081001__S1__2026-07-01__1": "absent",
  "25003111025081001__S1__2026-07-02__1": "absent",
  "25003111025081001__S1__2026-07-06__1": "present",
  "25003111025081001__S1__2026-07-07__1": "absent",
  "25003111025081002__S1__2026-07-01__1": "absent",
  "25003111025081002__S1__2026-07-02__1": "absent",
  "25003111025081002__S1__2026-07-06__1": "absent",
  "25003111025081002__S1__2026-07-07__1": "absent",
  "25003111025081004__S1__2026-07-01__1": "absent",
  "25003111025081004__S1__2026-07-02__1": "absent",
  "25003111025081004__S1__2026-07-06__1": "present",
  "25003111025081004__S1__2026-07-07__1": "present",
  "25003111025081005__S1__2026-07-01__1": "absent",
  "25003111025081005__S1__2026-07-02__1": "absent",
  "25003111025081005__S1__2026-07-06__1": "absent",
  "25003111025081005__S1__2026-07-07__1": "absent",
  "25003111025081006__S1__2026-07-01__1": "present",
  "25003111025081006__S1__2026-07-02__1": "present",
  "25003111025081006__S1__2026-07-06__1": "present",
  "25003111025081006__S1__2026-07-07__1": "present",
  "25003111025081007__S1__2026-07-01__1": "absent",
  "25003111025081007__S1__2026-07-02__1": "absent",
  "25003111025081007__S1__2026-07-06__1": "present",
  "25003111025081007__S1__2026-07-07__1": "present",
  "25003111025081008__S1__2026-07-01__1": "absent",
  "25003111025081008__S1__2026-07-02__1": "absent",
  "25003111025081008__S1__2026-07-06__1": "absent",
  "25003111025081008__S1__2026-07-07__1": "absent",
  "25003111025081009__S1__2026-07-01__1": "absent",
  "25003111025081009__S1__2026-07-02__1": "absent",
  "25003111025081009__S1__2026-07-06__1": "present",
  "25003111025081009__S1__2026-07-07__1": "absent",
  "25003111025081010__S1__2026-07-01__1": "absent",
  "25003111025081010__S1__2026-07-02__1": "absent",
  "25003111025081010__S1__2026-07-06__1": "present",
  "25003111025081010__S1__2026-07-07__1": "absent",
  "25003111025081011__S1__2026-07-01__1": "present",
  "25003111025081011__S1__2026-07-02__1": "present",
  "25003111025081011__S1__2026-07-06__1": "present",
  "25003111025081011__S1__2026-07-07__1": "present",
  "25003111025081013__S1__2026-07-01__1": "absent",
  "25003111025081013__S1__2026-07-02__1": "absent",
  "25003111025081013__S1__2026-07-06__1": "absent",
  "25003111025081013__S1__2026-07-07__1": "absent",
  "25003111025081014__S1__2026-07-01__1": "absent",
  "25003111025081014__S1__2026-07-02__1": "absent",
  "25003111025081014__S1__2026-07-06__1": "present",
  "25003111025081014__S1__2026-07-07__1": "present",
  "25003111025081015__S1__2026-07-01__1": "absent",
  "25003111025081015__S1__2026-07-02__1": "absent",
  "25003111025081015__S1__2026-07-06__1": "present",
  "25003111025081015__S1__2026-07-07__1": "present",
  "25003111025081016__S1__2026-07-01__1": "absent",
  "25003111025081016__S1__2026-07-02__1": "absent",
  "25003111025081016__S1__2026-07-06__1": "absent",
  "25003111025081016__S1__2026-07-07__1": "absent",
  "25003111025081017__S1__2026-07-01__1": "present",
  "25003111025081017__S1__2026-07-02__1": "present",
  "25003111025081017__S1__2026-07-06__1": "present",
  "25003111025081017__S1__2026-07-07__1": "present",
  "25003111025081018__S1__2026-07-01__1": "absent",
  "25003111025081018__S1__2026-07-02__1": "absent",
  "25003111025081018__S1__2026-07-06__1": "absent",
  "25003111025081018__S1__2026-07-07__1": "absent",
  "25003111025081019__S1__2026-07-01__1": "present",
  "25003111025081019__S1__2026-07-02__1": "present",
  "25003111025081019__S1__2026-07-06__1": "absent",
  "25003111025081019__S1__2026-07-07__1": "present",
  "25003111025081020__S1__2026-07-01__1": "absent",
  "25003111025081020__S1__2026-07-02__1": "absent",
  "25003111025081020__S1__2026-07-06__1": "absent",
  "25003111025081020__S1__2026-07-07__1": "absent",
  "25003111025081021__S1__2026-07-01__1": "present",
  "25003111025081021__S1__2026-07-02__1": "absent",
  "25003111025081021__S1__2026-07-06__1": "absent",
  "25003111025081021__S1__2026-07-07__1": "absent",
  "25003111025081023__S1__2026-07-01__1": "absent",
  "25003111025081023__S1__2026-07-02__1": "absent",
  "25003111025081023__S1__2026-07-06__1": "absent",
  "25003111025081023__S1__2026-07-07__1": "absent",
  "25003111025081024__S1__2026-07-01__1": "present",
  "25003111025081024__S1__2026-07-02__1": "present",
  "25003111025081024__S1__2026-07-06__1": "present",
  "25003111025081024__S1__2026-07-07__1": "present",
  "25003111025081025__S1__2026-07-01__1": "present",
  "25003111025081025__S1__2026-07-02__1": "present",
  "25003111025081025__S1__2026-07-06__1": "present",
  "25003111025081025__S1__2026-07-07__1": "present",
  "25003111025081026__S1__2026-07-01__1": "absent",
  "25003111025081026__S1__2026-07-02__1": "absent",
  "25003111025081026__S1__2026-07-06__1": "absent",
  "25003111025081026__S1__2026-07-07__1": "absent",
  "25003111025081027__S1__2026-07-01__1": "absent",
  "25003111025081027__S1__2026-07-02__1": "absent",
  "25003111025081027__S1__2026-07-06__1": "absent",
  "25003111025081027__S1__2026-07-07__1": "absent",
  "25003111025081028__S1__2026-07-01__1": "absent",
  "25003111025081028__S1__2026-07-02__1": "absent",
  "25003111025081028__S1__2026-07-06__1": "absent",
  "25003111025081028__S1__2026-07-07__1": "present",
  "25003111025081029__S1__2026-07-01__1": "absent",
  "25003111025081029__S1__2026-07-02__1": "absent",
  "25003111025081029__S1__2026-07-06__1": "absent",
  "25003111025081029__S1__2026-07-07__1": "absent",
  "25003111025081030__S1__2026-07-01__1": "absent",
  "25003111025081030__S1__2026-07-02__1": "absent",
  "25003111025081030__S1__2026-07-06__1": "present",
  "25003111025081030__S1__2026-07-07__1": "present",
  "25003111025081031__S1__2026-07-01__1": "absent",
  "25003111025081031__S1__2026-07-02__1": "absent",
  "25003111025081031__S1__2026-07-06__1": "absent",
  "25003111025081031__S1__2026-07-07__1": "absent",
  "25003111025081032__S1__2026-07-01__1": "absent",
  "25003111025081032__S1__2026-07-02__1": "absent",
  "25003111025081032__S1__2026-07-06__1": "absent",
  "25003111025081032__S1__2026-07-07__1": "absent",
  "25003111025081034__S1__2026-07-01__1": "absent",
  "25003111025081034__S1__2026-07-02__1": "present",
  "25003111025081034__S1__2026-07-06__1": "present",
  "25003111025081034__S1__2026-07-07__1": "absent",
  "25003111025081035__S1__2026-07-01__1": "present",
  "25003111025081035__S1__2026-07-02__1": "present",
  "25003111025081035__S1__2026-07-06__1": "present",
  "25003111025081035__S1__2026-07-07__1": "present",
  "25003111025081036__S1__2026-07-01__1": "absent",
  "25003111025081036__S1__2026-07-02__1": "absent",
  "25003111025081036__S1__2026-07-06__1": "present",
  "25003111025081036__S1__2026-07-07__1": "present",
  "25003111025081038__S1__2026-07-01__1": "absent",
  "25003111025081038__S1__2026-07-02__1": "absent",
  "25003111025081038__S1__2026-07-06__1": "present",
  "25003111025081038__S1__2026-07-07__1": "present",
  "25003111025081039__S1__2026-07-01__1": "absent",
  "25003111025081039__S1__2026-07-02__1": "absent",
  "25003111025081039__S1__2026-07-06__1": "absent",
  "25003111025081039__S1__2026-07-07__1": "absent",
  "25003111025081040__S1__2026-07-01__1": "present",
  "25003111025081040__S1__2026-07-02__1": "present",
  "25003111025081040__S1__2026-07-06__1": "present",
  "25003111025081040__S1__2026-07-07__1": "present",
  "25003111025081041__S1__2026-07-01__1": "present",
  "25003111025081041__S1__2026-07-02__1": "present",
  "25003111025081041__S1__2026-07-06__1": "present",
  "25003111025081041__S1__2026-07-07__1": "present",
  "25003111025081043__S1__2026-07-01__1": "absent",
  "25003111025081043__S1__2026-07-02__1": "absent",
  "25003111025081043__S1__2026-07-06__1": "present",
  "25003111025081043__S1__2026-07-07__1": "present",
  "25003111025081044__S1__2026-07-01__1": "absent",
  "25003111025081044__S1__2026-07-02__1": "absent",
  "25003111025081044__S1__2026-07-06__1": "absent",
  "25003111025081044__S1__2026-07-07__1": "absent",
  "25003111025081046__S1__2026-07-01__1": "absent",
  "25003111025081046__S1__2026-07-02__1": "absent",
  "25003111025081046__S1__2026-07-06__1": "absent",
  "25003111025081046__S1__2026-07-07__1": "absent",
  "25003111025081047__S1__2026-07-01__1": "absent",
  "25003111025081047__S1__2026-07-02__1": "absent",
  "25003111025081047__S1__2026-07-06__1": "absent",
  "25003111025081047__S1__2026-07-07__1": "absent",
  "25003111025081048__S1__2026-07-01__1": "absent",
  "25003111025081048__S1__2026-07-02__1": "absent",
  "25003111025081048__S1__2026-07-06__1": "absent",
  "25003111025081048__S1__2026-07-07__1": "absent",
  "25003111025081049__S1__2026-07-01__1": "absent",
  "25003111025081049__S1__2026-07-02__1": "absent",
  "25003111025081049__S1__2026-07-06__1": "present",
  "25003111025081049__S1__2026-07-07__1": "present",
  "25003111025081050__S1__2026-07-01__1": "present",
  "25003111025081050__S1__2026-07-02__1": "absent",
  "25003111025081050__S1__2026-07-06__1": "present",
  "25003111025081050__S1__2026-07-07__1": "present",
  "25003111025081051__S1__2026-07-01__1": "absent",
  "25003111025081051__S1__2026-07-02__1": "absent",
  "25003111025081051__S1__2026-07-06__1": "absent",
  "25003111025081051__S1__2026-07-07__1": "absent",
  "25003111025081052__S1__2026-07-01__1": "present",
  "25003111025081052__S1__2026-07-02__1": "present",
  "25003111025081052__S1__2026-07-06__1": "present",
  "25003111025081052__S1__2026-07-07__1": "present",
  "25003111025081053__S1__2026-07-01__1": "absent",
  "25003111025081053__S1__2026-07-02__1": "absent",
  "25003111025081053__S1__2026-07-06__1": "absent",
  "25003111025081053__S1__2026-07-07__1": "absent",
  "25003111025081055__S1__2026-07-01__1": "present",
  "25003111025081055__S1__2026-07-02__1": "present",
  "25003111025081055__S1__2026-07-06__1": "present",
  "25003111025081055__S1__2026-07-07__1": "present",
  "25003111025081056__S1__2026-07-01__1": "present",
  "25003111025081056__S1__2026-07-02__1": "present",
  "25003111025081056__S1__2026-07-06__1": "present",
  "25003111025081056__S1__2026-07-07__1": "present",
  "25003311025081001__S1__2026-07-01__1": "absent",
  "25003311025081001__S1__2026-07-02__1": "absent",
  "25003311025081001__S1__2026-07-06__1": "present",
  "25003311025081001__S1__2026-07-07__1": "present",
  "25003311025081002__S1__2026-07-01__1": "absent",
  "25003311025081002__S1__2026-07-02__1": "absent",
  "25003311025081002__S1__2026-07-06__1": "present",
  "25003311025081002__S1__2026-07-07__1": "present",
  "25003311025081003__S1__2026-07-01__1": "absent",
  "25003311025081003__S1__2026-07-02__1": "absent",
  "25003311025081003__S1__2026-07-06__1": "absent",
  "25003311025081003__S1__2026-07-07__1": "absent",
  "25003311025081004__S1__2026-07-01__1": "absent",
  "25003311025081004__S1__2026-07-02__1": "absent",
  "25003311025081004__S1__2026-07-06__1": "present",
  "25003311025081004__S1__2026-07-07__1": "present",
  "25003311025081005__S1__2026-07-01__1": "absent",
  "25003311025081005__S1__2026-07-02__1": "absent",
  "25003311025081005__S1__2026-07-06__1": "present",
  "25003311025081005__S1__2026-07-07__1": "present",
  "25003111018071001__S1__2026-07-01__1": "absent",
  "25003111018071001__S1__2026-07-02__1": "absent",
  "25003111018071001__S1__2026-07-06__1": "absent",
  "25003111018071001__S1__2026-07-07__1": "present",
  "25003111018071003__S1__2026-07-01__1": "absent",
  "25003111018071003__S1__2026-07-02__1": "absent",
  "25003111018071003__S1__2026-07-06__1": "absent",
  "25003111018071003__S1__2026-07-07__1": "absent",
  "25003111018071004__S1__2026-07-01__1": "absent",
  "25003111018071004__S1__2026-07-02__1": "absent",
  "25003111018071004__S1__2026-07-06__1": "absent",
  "25003111018071004__S1__2026-07-07__1": "absent",
  "25003111018071005__S1__2026-07-01__1": "absent",
  "25003111018071005__S1__2026-07-02__1": "absent",
  "25003111018071005__S1__2026-07-06__1": "absent",
  "25003111018071005__S1__2026-07-07__1": "absent",
  "25003111018071006__S1__2026-07-01__1": "absent",
  "25003111018071006__S1__2026-07-02__1": "absent",
  "25003111018071006__S1__2026-07-06__1": "absent",
  "25003111018071006__S1__2026-07-07__1": "absent",
  "25003111018071007__S1__2026-07-01__1": "absent",
  "25003111018071007__S1__2026-07-02__1": "absent",
  "25003111018071007__S1__2026-07-06__1": "absent",
  "25003111018071007__S1__2026-07-07__1": "present",
  "25003111018071008__S1__2026-07-01__1": "absent",
  "25003111018071008__S1__2026-07-02__1": "absent",
  "25003111018071008__S1__2026-07-06__1": "absent",
  "25003111018071008__S1__2026-07-07__1": "absent",
  "25003111018071009__S1__2026-07-01__1": "absent",
  "25003111018071009__S1__2026-07-02__1": "absent",
  "25003111018071009__S1__2026-07-06__1": "absent",
  "25003111018071009__S1__2026-07-07__1": "present",
  "25003111018071010__S1__2026-07-01__1": "absent",
  "25003111018071010__S1__2026-07-02__1": "absent",
  "25003111018071010__S1__2026-07-06__1": "present",
  "25003111018071010__S1__2026-07-07__1": "absent",
  "25003111025081001__S2__2026-07-01__1": "absent",
  "25003111025081001__S2__2026-07-02__1": "absent",
  "25003111025081001__S2__2026-07-06__1": "absent",
  "25003111025081001__S2__2026-07-07__1": "absent",
  "25003111025081001__S2__2026-07-08__1": "absent",
  "25003111025081002__S2__2026-07-01__1": "absent",
  "25003111025081002__S2__2026-07-02__1": "absent",
  "25003111025081002__S2__2026-07-06__1": "absent",
  "25003111025081002__S2__2026-07-07__1": "absent",
  "25003111025081002__S2__2026-07-08__1": "absent",
  "25003111025081004__S2__2026-07-01__1": "absent",
  "25003111025081004__S2__2026-07-02__1": "absent",
  "25003111025081004__S2__2026-07-06__1": "present",
  "25003111025081004__S2__2026-07-07__1": "present",
  "25003111025081004__S2__2026-07-08__1": "present",
  "25003111025081005__S2__2026-07-01__1": "absent",
  "25003111025081005__S2__2026-07-02__1": "absent",
  "25003111025081005__S2__2026-07-06__1": "absent",
  "25003111025081005__S2__2026-07-07__1": "absent",
  "25003111025081005__S2__2026-07-08__1": "absent",
  "25003111025081006__S2__2026-07-01__1": "present",
  "25003111025081006__S2__2026-07-02__1": "present",
  "25003111025081006__S2__2026-07-06__1": "present",
  "25003111025081006__S2__2026-07-07__1": "present",
  "25003111025081006__S2__2026-07-08__1": "present",
  "25003111025081007__S2__2026-07-01__1": "absent",
  "25003111025081007__S2__2026-07-02__1": "absent",
  "25003111025081007__S2__2026-07-06__1": "absent",
  "25003111025081007__S2__2026-07-07__1": "present",
  "25003111025081007__S2__2026-07-08__1": "present",
  "25003111025081008__S2__2026-07-01__1": "absent",
  "25003111025081008__S2__2026-07-02__1": "absent",
  "25003111025081008__S2__2026-07-06__1": "absent",
  "25003111025081008__S2__2026-07-07__1": "absent",
  "25003111025081008__S2__2026-07-08__1": "absent",
  "25003111025081009__S2__2026-07-01__1": "absent",
  "25003111025081009__S2__2026-07-02__1": "absent",
  "25003111025081009__S2__2026-07-06__1": "absent",
  "25003111025081009__S2__2026-07-07__1": "absent",
  "25003111025081009__S2__2026-07-08__1": "present",
  "25003111025081010__S2__2026-07-01__1": "absent",
  "25003111025081010__S2__2026-07-02__1": "absent",
  "25003111025081010__S2__2026-07-06__1": "absent",
  "25003111025081010__S2__2026-07-07__1": "absent",
  "25003111025081010__S2__2026-07-08__1": "absent",
  "25003111025081011__S2__2026-07-01__1": "present",
  "25003111025081011__S2__2026-07-02__1": "present",
  "25003111025081011__S2__2026-07-06__1": "present",
  "25003111025081011__S2__2026-07-07__1": "present",
  "25003111025081011__S2__2026-07-08__1": "present",
  "25003111025081013__S2__2026-07-01__1": "absent",
  "25003111025081013__S2__2026-07-02__1": "absent",
  "25003111025081013__S2__2026-07-06__1": "absent",
  "25003111025081013__S2__2026-07-07__1": "absent",
  "25003111025081013__S2__2026-07-08__1": "absent",
  "25003111025081014__S2__2026-07-01__1": "absent",
  "25003111025081014__S2__2026-07-02__1": "absent",
  "25003111025081014__S2__2026-07-06__1": "present",
  "25003111025081014__S2__2026-07-07__1": "present",
  "25003111025081014__S2__2026-07-08__1": "present",
  "25003111025081015__S2__2026-07-01__1": "absent",
  "25003111025081015__S2__2026-07-02__1": "absent",
  "25003111025081015__S2__2026-07-06__1": "absent",
  "25003111025081015__S2__2026-07-07__1": "present",
  "25003111025081015__S2__2026-07-08__1": "present",
  "25003111025081016__S2__2026-07-01__1": "absent",
  "25003111025081016__S2__2026-07-02__1": "absent",
  "25003111025081016__S2__2026-07-06__1": "absent",
  "25003111025081016__S2__2026-07-07__1": "absent",
  "25003111025081016__S2__2026-07-08__1": "absent",
  "25003111025081017__S2__2026-07-01__1": "present",
  "25003111025081017__S2__2026-07-02__1": "present",
  "25003111025081017__S2__2026-07-06__1": "present",
  "25003111025081017__S2__2026-07-07__1": "present",
  "25003111025081017__S2__2026-07-08__1": "present",
  "25003111025081018__S2__2026-07-01__1": "absent",
  "25003111025081018__S2__2026-07-02__1": "absent",
  "25003111025081018__S2__2026-07-06__1": "absent",
  "25003111025081018__S2__2026-07-07__1": "absent",
  "25003111025081018__S2__2026-07-08__1": "absent",
  "25003111025081019__S2__2026-07-01__1": "present",
  "25003111025081019__S2__2026-07-02__1": "absent",
  "25003111025081019__S2__2026-07-06__1": "absent",
  "25003111025081019__S2__2026-07-07__1": "present",
  "25003111025081019__S2__2026-07-08__1": "present",
  "25003111025081020__S2__2026-07-01__1": "absent",
  "25003111025081020__S2__2026-07-02__1": "absent",
  "25003111025081020__S2__2026-07-06__1": "absent",
  "25003111025081020__S2__2026-07-07__1": "absent",
  "25003111025081020__S2__2026-07-08__1": "absent",
  "25003111025081021__S2__2026-07-01__1": "present",
  "25003111025081021__S2__2026-07-02__1": "present",
  "25003111025081021__S2__2026-07-06__1": "absent",
  "25003111025081021__S2__2026-07-07__1": "absent",
  "25003111025081021__S2__2026-07-08__1": "present",
  "25003111025081023__S2__2026-07-01__1": "absent",
  "25003111025081023__S2__2026-07-02__1": "absent",
  "25003111025081023__S2__2026-07-06__1": "absent",
  "25003111025081023__S2__2026-07-07__1": "absent",
  "25003111025081023__S2__2026-07-08__1": "absent",
  "25003111025081024__S2__2026-07-01__1": "present",
  "25003111025081024__S2__2026-07-02__1": "present",
  "25003111025081024__S2__2026-07-06__1": "present",
  "25003111025081024__S2__2026-07-07__1": "absent",
  "25003111025081024__S2__2026-07-08__1": "present",
  "25003111025081025__S2__2026-07-01__1": "present",
  "25003111025081025__S2__2026-07-02__1": "present",
  "25003111025081025__S2__2026-07-06__1": "present",
  "25003111025081025__S2__2026-07-07__1": "absent",
  "25003111025081025__S2__2026-07-08__1": "present",
  "25003111025081026__S2__2026-07-01__1": "absent",
  "25003111025081026__S2__2026-07-02__1": "absent",
  "25003111025081026__S2__2026-07-06__1": "absent",
  "25003111025081026__S2__2026-07-07__1": "absent",
  "25003111025081026__S2__2026-07-08__1": "absent",
  "25003111025081027__S2__2026-07-01__1": "absent",
  "25003111025081027__S2__2026-07-02__1": "absent",
  "25003111025081027__S2__2026-07-06__1": "absent",
  "25003111025081027__S2__2026-07-07__1": "absent",
  "25003111025081027__S2__2026-07-08__1": "absent",
  "25003111025081028__S2__2026-07-01__1": "absent",
  "25003111025081028__S2__2026-07-02__1": "absent",
  "25003111025081028__S2__2026-07-06__1": "absent",
  "25003111025081028__S2__2026-07-07__1": "absent",
  "25003111025081028__S2__2026-07-08__1": "present",
  "25003111025081029__S2__2026-07-01__1": "absent",
  "25003111025081029__S2__2026-07-02__1": "absent",
  "25003111025081029__S2__2026-07-06__1": "absent",
  "25003111025081029__S2__2026-07-07__1": "absent",
  "25003111025081029__S2__2026-07-08__1": "absent",
  "25003111025081030__S2__2026-07-01__1": "absent",
  "25003111025081030__S2__2026-07-02__1": "absent",
  "25003111025081030__S2__2026-07-06__1": "absent",
  "25003111025081030__S2__2026-07-07__1": "present",
  "25003111025081030__S2__2026-07-08__1": "absent",
  "25003111025081031__S2__2026-07-01__1": "absent",
  "25003111025081031__S2__2026-07-02__1": "absent",
  "25003111025081031__S2__2026-07-06__1": "absent",
  "25003111025081031__S2__2026-07-07__1": "absent",
  "25003111025081031__S2__2026-07-08__1": "absent",
  "25003111025081032__S2__2026-07-01__1": "absent",
  "25003111025081032__S2__2026-07-02__1": "absent",
  "25003111025081032__S2__2026-07-06__1": "absent",
  "25003111025081032__S2__2026-07-07__1": "absent",
  "25003111025081032__S2__2026-07-08__1": "absent",
  "25003111025081034__S2__2026-07-01__1": "absent",
  "25003111025081034__S2__2026-07-02__1": "present",
  "25003111025081034__S2__2026-07-06__1": "absent",
  "25003111025081034__S2__2026-07-07__1": "absent",
  "25003111025081034__S2__2026-07-08__1": "present",
  "25003111025081035__S2__2026-07-01__1": "present",
  "25003111025081035__S2__2026-07-02__1": "present",
  "25003111025081035__S2__2026-07-06__1": "present",
  "25003111025081035__S2__2026-07-07__1": "present",
  "25003111025081035__S2__2026-07-08__1": "present",
  "25003111025081036__S2__2026-07-01__1": "absent",
  "25003111025081036__S2__2026-07-02__1": "absent",
  "25003111025081036__S2__2026-07-06__1": "present",
  "25003111025081036__S2__2026-07-07__1": "present",
  "25003111025081036__S2__2026-07-08__1": "present",
  "25003111025081038__S2__2026-07-01__1": "absent",
  "25003111025081038__S2__2026-07-02__1": "absent",
  "25003111025081038__S2__2026-07-06__1": "absent",
  "25003111025081038__S2__2026-07-07__1": "present",
  "25003111025081038__S2__2026-07-08__1": "present",
  "25003111025081039__S2__2026-07-01__1": "absent",
  "25003111025081039__S2__2026-07-02__1": "absent",
  "25003111025081039__S2__2026-07-06__1": "absent",
  "25003111025081039__S2__2026-07-07__1": "absent",
  "25003111025081039__S2__2026-07-08__1": "absent",
  "25003111025081040__S2__2026-07-01__1": "present",
  "25003111025081040__S2__2026-07-02__1": "present",
  "25003111025081040__S2__2026-07-06__1": "present",
  "25003111025081040__S2__2026-07-07__1": "present",
  "25003111025081040__S2__2026-07-08__1": "present",
  "25003111025081041__S2__2026-07-01__1": "present",
  "25003111025081041__S2__2026-07-02__1": "present",
  "25003111025081041__S2__2026-07-06__1": "present",
  "25003111025081041__S2__2026-07-07__1": "present",
  "25003111025081041__S2__2026-07-08__1": "present",
  "25003111025081043__S2__2026-07-01__1": "absent",
  "25003111025081043__S2__2026-07-02__1": "absent",
  "25003111025081043__S2__2026-07-06__1": "absent",
  "25003111025081043__S2__2026-07-07__1": "absent",
  "25003111025081043__S2__2026-07-08__1": "absent",
  "25003111025081044__S2__2026-07-01__1": "absent",
  "25003111025081044__S2__2026-07-02__1": "absent",
  "25003111025081044__S2__2026-07-06__1": "absent",
  "25003111025081044__S2__2026-07-07__1": "absent",
  "25003111025081044__S2__2026-07-08__1": "absent",
  "25003111025081046__S2__2026-07-01__1": "absent",
  "25003111025081046__S2__2026-07-02__1": "absent",
  "25003111025081046__S2__2026-07-06__1": "absent",
  "25003111025081046__S2__2026-07-07__1": "absent",
  "25003111025081046__S2__2026-07-08__1": "absent",
  "25003111025081047__S2__2026-07-01__1": "absent",
  "25003111025081047__S2__2026-07-02__1": "absent",
  "25003111025081047__S2__2026-07-06__1": "absent",
  "25003111025081047__S2__2026-07-07__1": "absent",
  "25003111025081047__S2__2026-07-08__1": "absent",
  "25003111025081048__S2__2026-07-01__1": "absent",
  "25003111025081048__S2__2026-07-02__1": "absent",
  "25003111025081048__S2__2026-07-06__1": "absent",
  "25003111025081048__S2__2026-07-07__1": "absent",
  "25003111025081048__S2__2026-07-08__1": "absent",
  "25003111025081049__S2__2026-07-01__1": "absent",
  "25003111025081049__S2__2026-07-02__1": "absent",
  "25003111025081049__S2__2026-07-06__1": "present",
  "25003111025081049__S2__2026-07-07__1": "absent",
  "25003111025081049__S2__2026-07-08__1": "present",
  "25003111025081050__S2__2026-07-01__1": "present",
  "25003111025081050__S2__2026-07-02__1": "absent",
  "25003111025081050__S2__2026-07-06__1": "present",
  "25003111025081050__S2__2026-07-07__1": "present",
  "25003111025081050__S2__2026-07-08__1": "present",
  "25003111025081051__S2__2026-07-01__1": "absent",
  "25003111025081051__S2__2026-07-02__1": "absent",
  "25003111025081051__S2__2026-07-06__1": "absent",
  "25003111025081051__S2__2026-07-07__1": "absent",
  "25003111025081051__S2__2026-07-08__1": "absent",
  "25003111025081052__S2__2026-07-01__1": "present",
  "25003111025081052__S2__2026-07-02__1": "present",
  "25003111025081052__S2__2026-07-06__1": "present",
  "25003111025081052__S2__2026-07-07__1": "present",
  "25003111025081052__S2__2026-07-08__1": "present",
  "25003111025081053__S2__2026-07-01__1": "absent",
  "25003111025081053__S2__2026-07-02__1": "absent",
  "25003111025081053__S2__2026-07-06__1": "absent",
  "25003111025081053__S2__2026-07-07__1": "absent",
  "25003111025081053__S2__2026-07-08__1": "present",
  "25003111025081055__S2__2026-07-01__1": "present",
  "25003111025081055__S2__2026-07-02__1": "present",
  "25003111025081055__S2__2026-07-06__1": "present",
  "25003111025081055__S2__2026-07-07__1": "present",
  "25003111025081055__S2__2026-07-08__1": "present",
  "25003111025081056__S2__2026-07-01__1": "present",
  "25003111025081056__S2__2026-07-02__1": "present",
  "25003111025081056__S2__2026-07-06__1": "present",
  "25003111025081056__S2__2026-07-07__1": "present",
  "25003111025081056__S2__2026-07-08__1": "present",
  "25003311025081001__S2__2026-07-01__1": "absent",
  "25003311025081001__S2__2026-07-02__1": "absent",
  "25003311025081001__S2__2026-07-06__1": "present",
  "25003311025081001__S2__2026-07-07__1": "present",
  "25003311025081001__S2__2026-07-08__1": "present",
  "25003311025081002__S2__2026-07-01__1": "absent",
  "25003311025081002__S2__2026-07-02__1": "absent",
  "25003311025081002__S2__2026-07-06__1": "present",
  "25003311025081002__S2__2026-07-07__1": "present",
  "25003311025081002__S2__2026-07-08__1": "present",
  "25003311025081003__S2__2026-07-01__1": "absent",
  "25003311025081003__S2__2026-07-02__1": "absent",
  "25003311025081003__S2__2026-07-06__1": "absent",
  "25003311025081003__S2__2026-07-07__1": "absent",
  "25003311025081003__S2__2026-07-08__1": "absent",
  "25003311025081004__S2__2026-07-01__1": "absent",
  "25003311025081004__S2__2026-07-02__1": "absent",
  "25003311025081004__S2__2026-07-06__1": "present",
  "25003311025081004__S2__2026-07-07__1": "present",
  "25003311025081004__S2__2026-07-08__1": "absent",
  "25003311025081005__S2__2026-07-01__1": "absent",
  "25003311025081005__S2__2026-07-02__1": "absent",
  "25003311025081005__S2__2026-07-06__1": "present",
  "25003311025081005__S2__2026-07-07__1": "present",
  "25003311025081005__S2__2026-07-08__1": "present",
  "25003111018071001__S2__2026-07-01__1": "absent",
  "25003111018071001__S2__2026-07-02__1": "absent",
  "25003111018071001__S2__2026-07-06__1": "absent",
  "25003111018071001__S2__2026-07-07__1": "present",
  "25003111018071001__S2__2026-07-08__1": "present",
  "25003111018071003__S2__2026-07-01__1": "absent",
  "25003111018071003__S2__2026-07-02__1": "absent",
  "25003111018071003__S2__2026-07-06__1": "absent",
  "25003111018071003__S2__2026-07-07__1": "absent",
  "25003111018071003__S2__2026-07-08__1": "absent",
  "25003111018071004__S2__2026-07-01__1": "absent",
  "25003111018071004__S2__2026-07-02__1": "absent",
  "25003111018071004__S2__2026-07-06__1": "absent",
  "25003111018071004__S2__2026-07-07__1": "absent",
  "25003111018071004__S2__2026-07-08__1": "absent",
  "25003111018071005__S2__2026-07-01__1": "absent",
  "25003111018071005__S2__2026-07-02__1": "absent",
  "25003111018071005__S2__2026-07-06__1": "absent",
  "25003111018071005__S2__2026-07-07__1": "absent",
  "25003111018071005__S2__2026-07-08__1": "absent",
  "25003111018071006__S2__2026-07-01__1": "absent",
  "25003111018071006__S2__2026-07-02__1": "absent",
  "25003111018071006__S2__2026-07-06__1": "absent",
  "25003111018071006__S2__2026-07-07__1": "absent",
  "25003111018071006__S2__2026-07-08__1": "absent",
  "25003111018071007__S2__2026-07-01__1": "absent",
  "25003111018071007__S2__2026-07-02__1": "absent",
  "25003111018071007__S2__2026-07-06__1": "absent",
  "25003111018071007__S2__2026-07-07__1": "present",
  "25003111018071007__S2__2026-07-08__1": "present",
  "25003111018071008__S2__2026-07-01__1": "absent",
  "25003111018071008__S2__2026-07-02__1": "absent",
  "25003111018071008__S2__2026-07-06__1": "absent",
  "25003111018071008__S2__2026-07-07__1": "absent",
  "25003111018071008__S2__2026-07-08__1": "present",
  "25003111018071009__S2__2026-07-01__1": "absent",
  "25003111018071009__S2__2026-07-02__1": "absent",
  "25003111018071009__S2__2026-07-06__1": "absent",
  "25003111018071009__S2__2026-07-07__1": "present",
  "25003111018071009__S2__2026-07-08__1": "present",
  "25003111018071010__S2__2026-07-01__1": "absent",
  "25003111018071010__S2__2026-07-02__1": "absent",
  "25003111018071010__S2__2026-07-06__1": "absent",
  "25003111018071010__S2__2026-07-07__1": "absent",
  "25003111018071010__S2__2026-07-08__1": "present",
  "25003111025081001__S3__2026-07-01__1": "absent",
  "25003111025081001__S3__2026-07-02__1": "absent",
  "25003111025081001__S3__2026-07-06__1": "present",
  "25003111025081001__S3__2026-07-08__1": "absent",
  "25003111025081002__S3__2026-07-01__1": "absent",
  "25003111025081002__S3__2026-07-02__1": "absent",
  "25003111025081002__S3__2026-07-06__1": "absent",
  "25003111025081002__S3__2026-07-08__1": "absent",
  "25003111025081004__S3__2026-07-01__1": "absent",
  "25003111025081004__S3__2026-07-02__1": "absent",
  "25003111025081004__S3__2026-07-06__1": "present",
  "25003111025081004__S3__2026-07-08__1": "present",
  "25003111025081005__S3__2026-07-01__1": "absent",
  "25003111025081005__S3__2026-07-02__1": "absent",
  "25003111025081005__S3__2026-07-06__1": "absent",
  "25003111025081005__S3__2026-07-08__1": "absent",
  "25003111025081006__S3__2026-07-01__1": "present",
  "25003111025081006__S3__2026-07-02__1": "present",
  "25003111025081006__S3__2026-07-06__1": "present",
  "25003111025081006__S3__2026-07-08__1": "present",
  "25003111025081007__S3__2026-07-01__1": "absent",
  "25003111025081007__S3__2026-07-02__1": "absent",
  "25003111025081007__S3__2026-07-06__1": "present",
  "25003111025081007__S3__2026-07-08__1": "present",
  "25003111025081008__S3__2026-07-01__1": "absent",
  "25003111025081008__S3__2026-07-02__1": "absent",
  "25003111025081008__S3__2026-07-06__1": "absent",
  "25003111025081008__S3__2026-07-08__1": "absent",
  "25003111025081009__S3__2026-07-01__1": "absent",
  "25003111025081009__S3__2026-07-02__1": "absent",
  "25003111025081009__S3__2026-07-06__1": "present",
  "25003111025081009__S3__2026-07-08__1": "present",
  "25003111025081010__S3__2026-07-01__1": "absent",
  "25003111025081010__S3__2026-07-02__1": "absent",
  "25003111025081010__S3__2026-07-06__1": "present",
  "25003111025081010__S3__2026-07-08__1": "absent",
  "25003111025081011__S3__2026-07-01__1": "present",
  "25003111025081011__S3__2026-07-02__1": "present",
  "25003111025081011__S3__2026-07-06__1": "present",
  "25003111025081011__S3__2026-07-08__1": "present",
  "25003111025081013__S3__2026-07-01__1": "absent",
  "25003111025081013__S3__2026-07-02__1": "absent",
  "25003111025081013__S3__2026-07-06__1": "absent",
  "25003111025081013__S3__2026-07-08__1": "absent",
  "25003111025081014__S3__2026-07-01__1": "absent",
  "25003111025081014__S3__2026-07-02__1": "absent",
  "25003111025081014__S3__2026-07-06__1": "present",
  "25003111025081014__S3__2026-07-08__1": "present",
  "25003111025081015__S3__2026-07-01__1": "absent",
  "25003111025081015__S3__2026-07-02__1": "absent",
  "25003111025081015__S3__2026-07-06__1": "present",
  "25003111025081015__S3__2026-07-08__1": "present",
  "25003111025081016__S3__2026-07-01__1": "absent",
  "25003111025081016__S3__2026-07-02__1": "absent",
  "25003111025081016__S3__2026-07-06__1": "absent",
  "25003111025081016__S3__2026-07-08__1": "absent",
  "25003111025081017__S3__2026-07-01__1": "present",
  "25003111025081017__S3__2026-07-02__1": "present",
  "25003111025081017__S3__2026-07-06__1": "present",
  "25003111025081017__S3__2026-07-08__1": "present",
  "25003111025081018__S3__2026-07-01__1": "absent",
  "25003111025081018__S3__2026-07-02__1": "absent",
  "25003111025081018__S3__2026-07-06__1": "absent",
  "25003111025081018__S3__2026-07-08__1": "absent",
  "25003111025081019__S3__2026-07-01__1": "present",
  "25003111025081019__S3__2026-07-02__1": "absent",
  "25003111025081019__S3__2026-07-06__1": "absent",
  "25003111025081019__S3__2026-07-08__1": "present",
  "25003111025081020__S3__2026-07-01__1": "absent",
  "25003111025081020__S3__2026-07-02__1": "absent",
  "25003111025081020__S3__2026-07-06__1": "absent",
  "25003111025081020__S3__2026-07-08__1": "absent",
  "25003111025081021__S3__2026-07-01__1": "present",
  "25003111025081021__S3__2026-07-02__1": "present",
  "25003111025081021__S3__2026-07-06__1": "absent",
  "25003111025081021__S3__2026-07-08__1": "present",
  "25003111025081023__S3__2026-07-01__1": "absent",
  "25003111025081023__S3__2026-07-02__1": "absent",
  "25003111025081023__S3__2026-07-06__1": "absent",
  "25003111025081023__S3__2026-07-08__1": "absent",
  "25003111025081024__S3__2026-07-01__1": "present",
  "25003111025081024__S3__2026-07-02__1": "present",
  "25003111025081024__S3__2026-07-06__1": "present",
  "25003111025081024__S3__2026-07-08__1": "present",
  "25003111025081025__S3__2026-07-01__1": "present",
  "25003111025081025__S3__2026-07-02__1": "present",
  "25003111025081025__S3__2026-07-06__1": "present",
  "25003111025081025__S3__2026-07-08__1": "present",
  "25003111025081026__S3__2026-07-01__1": "absent",
  "25003111025081026__S3__2026-07-02__1": "absent",
  "25003111025081026__S3__2026-07-06__1": "absent",
  "25003111025081026__S3__2026-07-08__1": "absent",
  "25003111025081027__S3__2026-07-01__1": "absent",
  "25003111025081027__S3__2026-07-02__1": "absent",
  "25003111025081027__S3__2026-07-06__1": "absent",
  "25003111025081027__S3__2026-07-08__1": "absent",
  "25003111025081028__S3__2026-07-01__1": "absent",
  "25003111025081028__S3__2026-07-02__1": "absent",
  "25003111025081028__S3__2026-07-06__1": "absent",
  "25003111025081028__S3__2026-07-08__1": "present",
  "25003111025081029__S3__2026-07-01__1": "absent",
  "25003111025081029__S3__2026-07-02__1": "absent",
  "25003111025081029__S3__2026-07-06__1": "absent",
  "25003111025081029__S3__2026-07-08__1": "absent",
  "25003111025081030__S3__2026-07-01__1": "absent",
  "25003111025081030__S3__2026-07-02__1": "absent",
  "25003111025081030__S3__2026-07-06__1": "present",
  "25003111025081030__S3__2026-07-08__1": "absent",
  "25003111025081031__S3__2026-07-01__1": "absent",
  "25003111025081031__S3__2026-07-02__1": "absent",
  "25003111025081031__S3__2026-07-06__1": "absent",
  "25003111025081031__S3__2026-07-08__1": "absent",
  "25003111025081032__S3__2026-07-01__1": "absent",
  "25003111025081032__S3__2026-07-02__1": "absent",
  "25003111025081032__S3__2026-07-06__1": "absent",
  "25003111025081032__S3__2026-07-08__1": "absent",
  "25003111025081034__S3__2026-07-01__1": "absent",
  "25003111025081034__S3__2026-07-02__1": "present",
  "25003111025081034__S3__2026-07-06__1": "present",
  "25003111025081034__S3__2026-07-08__1": "present",
  "25003111025081035__S3__2026-07-01__1": "present",
  "25003111025081035__S3__2026-07-02__1": "present",
  "25003111025081035__S3__2026-07-06__1": "present",
  "25003111025081035__S3__2026-07-08__1": "present",
  "25003111025081036__S3__2026-07-01__1": "absent",
  "25003111025081036__S3__2026-07-02__1": "absent",
  "25003111025081036__S3__2026-07-06__1": "present",
  "25003111025081036__S3__2026-07-08__1": "present",
  "25003111025081038__S3__2026-07-01__1": "absent",
  "25003111025081038__S3__2026-07-02__1": "absent",
  "25003111025081038__S3__2026-07-06__1": "present",
  "25003111025081038__S3__2026-07-08__1": "present",
  "25003111025081039__S3__2026-07-01__1": "absent",
  "25003111025081039__S3__2026-07-02__1": "absent",
  "25003111025081039__S3__2026-07-06__1": "absent",
  "25003111025081039__S3__2026-07-08__1": "absent",
  "25003111025081040__S3__2026-07-01__1": "present",
  "25003111025081040__S3__2026-07-02__1": "present",
  "25003111025081040__S3__2026-07-06__1": "present",
  "25003111025081040__S3__2026-07-08__1": "present",
  "25003111025081041__S3__2026-07-01__1": "present",
  "25003111025081041__S3__2026-07-02__1": "present",
  "25003111025081041__S3__2026-07-06__1": "present",
  "25003111025081041__S3__2026-07-08__1": "present",
  "25003111025081043__S3__2026-07-01__1": "absent",
  "25003111025081043__S3__2026-07-02__1": "absent",
  "25003111025081043__S3__2026-07-06__1": "present",
  "25003111025081043__S3__2026-07-08__1": "absent",
  "25003111025081044__S3__2026-07-01__1": "absent",
  "25003111025081044__S3__2026-07-02__1": "absent",
  "25003111025081044__S3__2026-07-06__1": "absent",
  "25003111025081044__S3__2026-07-08__1": "absent",
  "25003111025081046__S3__2026-07-01__1": "absent",
  "25003111025081046__S3__2026-07-02__1": "absent",
  "25003111025081046__S3__2026-07-06__1": "absent",
  "25003111025081046__S3__2026-07-08__1": "absent",
  "25003111025081047__S3__2026-07-01__1": "absent",
  "25003111025081047__S3__2026-07-02__1": "absent",
  "25003111025081047__S3__2026-07-06__1": "absent",
  "25003111025081047__S3__2026-07-08__1": "absent",
  "25003111025081048__S3__2026-07-01__1": "absent",
  "25003111025081048__S3__2026-07-02__1": "absent",
  "25003111025081048__S3__2026-07-06__1": "absent",
  "25003111025081048__S3__2026-07-08__1": "absent",
  "25003111025081049__S3__2026-07-01__1": "absent",
  "25003111025081049__S3__2026-07-02__1": "absent",
  "25003111025081049__S3__2026-07-06__1": "present",
  "25003111025081049__S3__2026-07-08__1": "present",
  "25003111025081050__S3__2026-07-01__1": "present",
  "25003111025081050__S3__2026-07-02__1": "absent",
  "25003111025081050__S3__2026-07-06__1": "present",
  "25003111025081050__S3__2026-07-08__1": "present",
  "25003111025081051__S3__2026-07-01__1": "absent",
  "25003111025081051__S3__2026-07-02__1": "absent",
  "25003111025081051__S3__2026-07-06__1": "absent",
  "25003111025081051__S3__2026-07-08__1": "absent",
  "25003111025081052__S3__2026-07-01__1": "present",
  "25003111025081052__S3__2026-07-02__1": "present",
  "25003111025081052__S3__2026-07-06__1": "present",
  "25003111025081052__S3__2026-07-08__1": "present",
  "25003111025081053__S3__2026-07-01__1": "absent",
  "25003111025081053__S3__2026-07-02__1": "absent",
  "25003111025081053__S3__2026-07-06__1": "absent",
  "25003111025081053__S3__2026-07-08__1": "present",
  "25003111025081055__S3__2026-07-01__1": "present",
  "25003111025081055__S3__2026-07-02__1": "present",
  "25003111025081055__S3__2026-07-06__1": "present",
  "25003111025081055__S3__2026-07-08__1": "present",
  "25003111025081056__S3__2026-07-01__1": "present",
  "25003111025081056__S3__2026-07-02__1": "present",
  "25003111025081056__S3__2026-07-06__1": "present",
  "25003111025081056__S3__2026-07-08__1": "present",
  "25003311025081001__S3__2026-07-01__1": "absent",
  "25003311025081001__S3__2026-07-02__1": "absent",
  "25003311025081001__S3__2026-07-06__1": "present",
  "25003311025081001__S3__2026-07-08__1": "present",
  "25003311025081002__S3__2026-07-01__1": "absent",
  "25003311025081002__S3__2026-07-02__1": "absent",
  "25003311025081002__S3__2026-07-06__1": "present",
  "25003311025081002__S3__2026-07-08__1": "present",
  "25003311025081003__S3__2026-07-01__1": "absent",
  "25003311025081003__S3__2026-07-02__1": "absent",
  "25003311025081003__S3__2026-07-06__1": "absent",
  "25003311025081003__S3__2026-07-08__1": "absent",
  "25003311025081004__S3__2026-07-01__1": "absent",
  "25003311025081004__S3__2026-07-02__1": "absent",
  "25003311025081004__S3__2026-07-06__1": "present",
  "25003311025081004__S3__2026-07-08__1": "absent",
  "25003311025081005__S3__2026-07-01__1": "absent",
  "25003311025081005__S3__2026-07-02__1": "absent",
  "25003311025081005__S3__2026-07-06__1": "present",
  "25003311025081005__S3__2026-07-08__1": "present",
  "25003111018071001__S3__2026-07-01__1": "absent",
  "25003111018071001__S3__2026-07-02__1": "absent",
  "25003111018071001__S3__2026-07-06__1": "absent",
  "25003111018071001__S3__2026-07-08__1": "present",
  "25003111018071003__S3__2026-07-01__1": "absent",
  "25003111018071003__S3__2026-07-02__1": "absent",
  "25003111018071003__S3__2026-07-06__1": "absent",
  "25003111018071003__S3__2026-07-08__1": "absent",
  "25003111018071004__S3__2026-07-01__1": "absent",
  "25003111018071004__S3__2026-07-02__1": "absent",
  "25003111018071004__S3__2026-07-06__1": "absent",
  "25003111018071004__S3__2026-07-08__1": "absent",
  "25003111018071005__S3__2026-07-01__1": "absent",
  "25003111018071005__S3__2026-07-02__1": "absent",
  "25003111018071005__S3__2026-07-06__1": "absent",
  "25003111018071005__S3__2026-07-08__1": "absent",
  "25003111018071006__S3__2026-07-01__1": "absent",
  "25003111018071006__S3__2026-07-02__1": "absent",
  "25003111018071006__S3__2026-07-06__1": "absent",
  "25003111018071006__S3__2026-07-08__1": "absent",
  "25003111018071007__S3__2026-07-01__1": "absent",
  "25003111018071007__S3__2026-07-02__1": "absent",
  "25003111018071007__S3__2026-07-06__1": "absent",
  "25003111018071007__S3__2026-07-08__1": "present",
  "25003111018071008__S3__2026-07-01__1": "absent",
  "25003111018071008__S3__2026-07-02__1": "absent",
  "25003111018071008__S3__2026-07-06__1": "absent",
  "25003111018071008__S3__2026-07-08__1": "present",
  "25003111018071009__S3__2026-07-01__1": "absent",
  "25003111018071009__S3__2026-07-02__1": "absent",
  "25003111018071009__S3__2026-07-06__1": "absent",
  "25003111018071009__S3__2026-07-08__1": "present",
  "25003111018071010__S3__2026-07-01__1": "absent",
  "25003111018071010__S3__2026-07-02__1": "absent",
  "25003111018071010__S3__2026-07-06__1": "present",
  "25003111018071010__S3__2026-07-08__1": "present",
  "25003111025081002__S4__2026-07-01__1": "absent",
  "25003111025081002__S4__2026-07-02__1": "absent",
  "25003111025081002__S4__2026-07-06__1": "absent",
  "25003111025081002__S4__2026-07-07__1": "absent",
  "25003111025081002__S4__2026-07-08__1": "absent",
  "25003111025081004__S4__2026-07-01__1": "absent",
  "25003111025081004__S4__2026-07-02__1": "absent",
  "25003111025081004__S4__2026-07-06__1": "present",
  "25003111025081004__S4__2026-07-07__1": "present",
  "25003111025081004__S4__2026-07-08__1": "present",
  "25003111025081005__S4__2026-07-01__1": "absent",
  "25003111025081005__S4__2026-07-02__1": "absent",
  "25003111025081005__S4__2026-07-06__1": "absent",
  "25003111025081005__S4__2026-07-07__1": "absent",
  "25003111025081005__S4__2026-07-08__1": "absent",
  "25003111025081006__S4__2026-07-01__1": "absent",
  "25003111025081006__S4__2026-07-02__1": "absent",
  "25003111025081006__S4__2026-07-06__1": "absent",
  "25003111025081006__S4__2026-07-07__1": "present",
  "25003111025081006__S4__2026-07-08__1": "absent",
  "25003111025081011__S4__2026-07-01__1": "present",
  "25003111025081011__S4__2026-07-02__1": "present",
  "25003111025081011__S4__2026-07-06__1": "present",
  "25003111025081011__S4__2026-07-07__1": "present",
  "25003111025081011__S4__2026-07-08__1": "present",
  "25003111025081013__S4__2026-07-01__1": "absent",
  "25003111025081013__S4__2026-07-02__1": "absent",
  "25003111025081013__S4__2026-07-06__1": "absent",
  "25003111025081013__S4__2026-07-07__1": "absent",
  "25003111025081013__S4__2026-07-08__1": "absent",
  "25003111025081014__S4__2026-07-01__1": "absent",
  "25003111025081014__S4__2026-07-02__1": "absent",
  "25003111025081014__S4__2026-07-06__1": "present",
  "25003111025081014__S4__2026-07-07__1": "present",
  "25003111025081014__S4__2026-07-08__1": "present",
  "25003111025081016__S4__2026-07-01__1": "absent",
  "25003111025081016__S4__2026-07-02__1": "absent",
  "25003111025081016__S4__2026-07-06__1": "absent",
  "25003111025081016__S4__2026-07-07__1": "absent",
  "25003111025081016__S4__2026-07-08__1": "absent",
  "25003111025081017__S4__2026-07-01__1": "present",
  "25003111025081017__S4__2026-07-02__1": "present",
  "25003111025081017__S4__2026-07-06__1": "present",
  "25003111025081017__S4__2026-07-07__1": "present",
  "25003111025081017__S4__2026-07-08__1": "absent",
  "25003111025081019__S4__2026-07-01__1": "present",
  "25003111025081019__S4__2026-07-02__1": "present",
  "25003111025081019__S4__2026-07-06__1": "absent",
  "25003111025081019__S4__2026-07-07__1": "present",
  "25003111025081019__S4__2026-07-08__1": "present",
  "25003111025081021__S4__2026-07-01__1": "present",
  "25003111025081021__S4__2026-07-02__1": "present",
  "25003111025081021__S4__2026-07-06__1": "absent",
  "25003111025081021__S4__2026-07-07__1": "absent",
  "25003111025081021__S4__2026-07-08__1": "present",
  "25003111025081024__S4__2026-07-01__1": "absent",
  "25003111025081024__S4__2026-07-02__1": "absent",
  "25003111025081024__S4__2026-07-06__1": "absent",
  "25003111025081024__S4__2026-07-07__1": "absent",
  "25003111025081024__S4__2026-07-08__1": "present",
  "25003111025081025__S4__2026-07-01__1": "absent",
  "25003111025081025__S4__2026-07-02__1": "absent",
  "25003111025081025__S4__2026-07-06__1": "absent",
  "25003111025081025__S4__2026-07-07__1": "absent",
  "25003111025081025__S4__2026-07-08__1": "present",
  "25003111025081027__S4__2026-07-01__1": "absent",
  "25003111025081027__S4__2026-07-02__1": "absent",
  "25003111025081027__S4__2026-07-06__1": "absent",
  "25003111025081027__S4__2026-07-07__1": "absent",
  "25003111025081027__S4__2026-07-08__1": "absent",
  "25003111025081029__S4__2026-07-01__1": "absent",
  "25003111025081029__S4__2026-07-02__1": "absent",
  "25003111025081029__S4__2026-07-06__1": "absent",
  "25003111025081029__S4__2026-07-07__1": "absent",
  "25003111025081029__S4__2026-07-08__1": "absent",
  "25003111025081030__S4__2026-07-01__1": "absent",
  "25003111025081030__S4__2026-07-02__1": "absent",
  "25003111025081030__S4__2026-07-06__1": "present",
  "25003111025081030__S4__2026-07-07__1": "present",
  "25003111025081030__S4__2026-07-08__1": "absent",
  "25003111025081031__S4__2026-07-01__1": "absent",
  "25003111025081031__S4__2026-07-02__1": "absent",
  "25003111025081031__S4__2026-07-06__1": "absent",
  "25003111025081031__S4__2026-07-07__1": "absent",
  "25003111025081031__S4__2026-07-08__1": "absent",
  "25003111025081032__S4__2026-07-01__1": "absent",
  "25003111025081032__S4__2026-07-02__1": "absent",
  "25003111025081032__S4__2026-07-06__1": "absent",
  "25003111025081032__S4__2026-07-07__1": "absent",
  "25003111025081032__S4__2026-07-08__1": "absent",
  "25003111025081034__S4__2026-07-01__1": "absent",
  "25003111025081034__S4__2026-07-02__1": "absent",
  "25003111025081034__S4__2026-07-06__1": "present",
  "25003111025081034__S4__2026-07-07__1": "absent",
  "25003111025081034__S4__2026-07-08__1": "present",
  "25003111025081035__S4__2026-07-01__1": "present",
  "25003111025081035__S4__2026-07-02__1": "present",
  "25003111025081035__S4__2026-07-06__1": "present",
  "25003111025081035__S4__2026-07-07__1": "present",
  "25003111025081035__S4__2026-07-08__1": "present",
  "25003111025081036__S4__2026-07-01__1": "absent",
  "25003111025081036__S4__2026-07-02__1": "absent",
  "25003111025081036__S4__2026-07-06__1": "present",
  "25003111025081036__S4__2026-07-07__1": "present",
  "25003111025081036__S4__2026-07-08__1": "present",
  "25003111025081040__S4__2026-07-01__1": "present",
  "25003111025081040__S4__2026-07-02__1": "present",
  "25003111025081040__S4__2026-07-06__1": "present",
  "25003111025081040__S4__2026-07-07__1": "present",
  "25003111025081040__S4__2026-07-08__1": "present",
  "25003111025081041__S4__2026-07-01__1": "present",
  "25003111025081041__S4__2026-07-02__1": "present",
  "25003111025081041__S4__2026-07-06__1": "present",
  "25003111025081041__S4__2026-07-07__1": "present",
  "25003111025081041__S4__2026-07-08__1": "present",
  "25003111025081044__S4__2026-07-01__1": "absent",
  "25003111025081044__S4__2026-07-02__1": "present",
  "25003111025081044__S4__2026-07-06__1": "absent",
  "25003111025081044__S4__2026-07-07__1": "absent",
  "25003111025081044__S4__2026-07-08__1": "absent",
  "25003111025081049__S4__2026-07-01__1": "absent",
  "25003111025081049__S4__2026-07-02__1": "absent",
  "25003111025081049__S4__2026-07-06__1": "absent",
  "25003111025081049__S4__2026-07-07__1": "absent",
  "25003111025081049__S4__2026-07-08__1": "present",
  "25003111025081050__S4__2026-07-01__1": "present",
  "25003111025081050__S4__2026-07-02__1": "absent",
  "25003111025081050__S4__2026-07-06__1": "present",
  "25003111025081050__S4__2026-07-07__1": "present",
  "25003111025081050__S4__2026-07-08__1": "present",
  "25003111025081052__S4__2026-07-01__1": "present",
  "25003111025081052__S4__2026-07-02__1": "present",
  "25003111025081052__S4__2026-07-06__1": "present",
  "25003111025081052__S4__2026-07-07__1": "present",
  "25003111025081052__S4__2026-07-08__1": "present",
  "25003111025081053__S4__2026-07-01__1": "absent",
  "25003111025081053__S4__2026-07-02__1": "absent",
  "25003111025081053__S4__2026-07-06__1": "absent",
  "25003111025081053__S4__2026-07-07__1": "absent",
  "25003111025081053__S4__2026-07-08__1": "present",
  "25003111025081056__S4__2026-07-01__1": "present",
  "25003111025081056__S4__2026-07-02__1": "present",
  "25003111025081056__S4__2026-07-06__1": "present",
  "25003111025081056__S4__2026-07-07__1": "present",
  "25003111025081056__S4__2026-07-08__1": "present",
  "25003111018071001__S4__2026-07-01__1": "absent",
  "25003111018071001__S4__2026-07-02__1": "absent",
  "25003111018071001__S4__2026-07-06__1": "absent",
  "25003111018071001__S4__2026-07-07__1": "present",
  "25003111018071001__S4__2026-07-08__1": "present",
  "25003111018071003__S4__2026-07-01__1": "absent",
  "25003111018071003__S4__2026-07-02__1": "absent",
  "25003111018071003__S4__2026-07-06__1": "absent",
  "25003111018071003__S4__2026-07-07__1": "absent",
  "25003111018071003__S4__2026-07-08__1": "absent",
  "25003111018071004__S4__2026-07-01__1": "absent",
  "25003111018071004__S4__2026-07-02__1": "absent",
  "25003111018071004__S4__2026-07-06__1": "absent",
  "25003111018071004__S4__2026-07-07__1": "absent",
  "25003111018071004__S4__2026-07-08__1": "absent",
  "25003111018071005__S4__2026-07-01__1": "absent",
  "25003111018071005__S4__2026-07-02__1": "absent",
  "25003111018071005__S4__2026-07-06__1": "absent",
  "25003111018071005__S4__2026-07-07__1": "absent",
  "25003111018071005__S4__2026-07-08__1": "absent",
  "25003111018071006__S4__2026-07-01__1": "absent",
  "25003111018071006__S4__2026-07-02__1": "absent",
  "25003111018071006__S4__2026-07-06__1": "absent",
  "25003111018071006__S4__2026-07-07__1": "absent",
  "25003111018071006__S4__2026-07-08__1": "absent",
  "25003111018071007__S4__2026-07-01__1": "absent",
  "25003111018071007__S4__2026-07-02__1": "absent",
  "25003111018071007__S4__2026-07-06__1": "absent",
  "25003111018071007__S4__2026-07-07__1": "present",
  "25003111018071007__S4__2026-07-08__1": "present",
  "25003111018071008__S4__2026-07-01__1": "absent",
  "25003111018071008__S4__2026-07-02__1": "absent",
  "25003111018071008__S4__2026-07-06__1": "absent",
  "25003111018071008__S4__2026-07-07__1": "absent",
  "25003111018071008__S4__2026-07-08__1": "absent",
  "25003111018071009__S4__2026-07-01__1": "absent",
  "25003111018071009__S4__2026-07-02__1": "absent",
  "25003111018071009__S4__2026-07-06__1": "absent",
  "25003111018071009__S4__2026-07-07__1": "present",
  "25003111018071009__S4__2026-07-08__1": "present",
  "25003111018071010__S4__2026-07-01__1": "absent",
  "25003111018071010__S4__2026-07-02__1": "absent",
  "25003111018071010__S4__2026-07-06__1": "absent",
  "25003111018071010__S4__2026-07-07__1": "absent",
  "25003111018071010__S4__2026-07-08__1": "absent",
  "25003111025081006__S5__2026-07-01__1": "present",
  "25003111025081006__S5__2026-07-02__1": "present",
  "25003111025081006__S5__2026-07-06__1": "present",
  "25003111025081006__S5__2026-07-08__1": "absent",
  "25003111025081007__S5__2026-07-01__1": "absent",
  "25003111025081007__S5__2026-07-02__1": "absent",
  "25003111025081007__S5__2026-07-06__1": "present",
  "25003111025081007__S5__2026-07-08__1": "present",
  "25003111025081015__S5__2026-07-01__1": "absent",
  "25003111025081015__S5__2026-07-02__1": "absent",
  "25003111025081015__S5__2026-07-06__1": "present",
  "25003111025081015__S5__2026-07-08__1": "present",
  "25003111025081020__S5__2026-07-01__1": "absent",
  "25003111025081020__S5__2026-07-02__1": "absent",
  "25003111025081020__S5__2026-07-06__1": "absent",
  "25003111025081020__S5__2026-07-08__1": "absent",
  "25003111025081024__S5__2026-07-01__1": "present",
  "25003111025081024__S5__2026-07-02__1": "present",
  "25003111025081024__S5__2026-07-06__1": "present",
  "25003111025081024__S5__2026-07-08__1": "absent",
  "25003111025081025__S5__2026-07-01__1": "present",
  "25003111025081025__S5__2026-07-02__1": "present",
  "25003111025081025__S5__2026-07-06__1": "present",
  "25003111025081025__S5__2026-07-08__1": "absent",
  "25003111025081038__S5__2026-07-01__1": "absent",
  "25003111025081038__S5__2026-07-02__1": "absent",
  "25003111025081038__S5__2026-07-06__1": "present",
  "25003111025081038__S5__2026-07-08__1": "present",
  "25003111025081046__S5__2026-07-01__1": "absent",
  "25003111025081046__S5__2026-07-02__1": "absent",
  "25003111025081046__S5__2026-07-06__1": "absent",
  "25003111025081046__S5__2026-07-08__1": "absent",
  "25003111025081047__S5__2026-07-01__1": "absent",
  "25003111025081047__S5__2026-07-02__1": "absent",
  "25003111025081047__S5__2026-07-06__1": "absent",
  "25003111025081047__S5__2026-07-08__1": "absent",
  "25003111025081048__S5__2026-07-01__1": "absent",
  "25003111025081048__S5__2026-07-02__1": "absent",
  "25003111025081048__S5__2026-07-06__1": "absent",
  "25003111025081048__S5__2026-07-08__1": "absent",
  "25003111025081049__S5__2026-07-01__1": "absent",
  "25003111025081049__S5__2026-07-02__1": "absent",
  "25003111025081049__S5__2026-07-06__1": "present",
  "25003111025081049__S5__2026-07-08__1": "absent",
  "25003111018071001__S5__2026-07-01__1": "absent",
  "25003111018071001__S5__2026-07-02__1": "absent",
  "25003111018071001__S5__2026-07-06__1": "absent",
  "25003111018071001__S5__2026-07-08__1": "absent",
  "25003111018071003__S5__2026-07-01__1": "absent",
  "25003111018071003__S5__2026-07-02__1": "absent",
  "25003111018071003__S5__2026-07-06__1": "absent",
  "25003111018071003__S5__2026-07-08__1": "absent",
  "25003111018071004__S5__2026-07-01__1": "absent",
  "25003111018071004__S5__2026-07-02__1": "absent",
  "25003111018071004__S5__2026-07-06__1": "absent",
  "25003111018071004__S5__2026-07-08__1": "absent",
  "25003111018071005__S5__2026-07-01__1": "absent",
  "25003111018071005__S5__2026-07-02__1": "absent",
  "25003111018071005__S5__2026-07-06__1": "absent",
  "25003111018071005__S5__2026-07-08__1": "absent",
  "25003111018071006__S5__2026-07-01__1": "absent",
  "25003111018071006__S5__2026-07-02__1": "absent",
  "25003111018071006__S5__2026-07-06__1": "absent",
  "25003111018071006__S5__2026-07-08__1": "absent",
  "25003111018071007__S5__2026-07-01__1": "absent",
  "25003111018071007__S5__2026-07-02__1": "absent",
  "25003111018071007__S5__2026-07-06__1": "absent",
  "25003111018071007__S5__2026-07-08__1": "absent",
  "25003111018071008__S5__2026-07-01__1": "absent",
  "25003111018071008__S5__2026-07-02__1": "absent",
  "25003111018071008__S5__2026-07-06__1": "absent",
  "25003111018071008__S5__2026-07-08__1": "absent",
  "25003111018071009__S5__2026-07-01__1": "absent",
  "25003111018071009__S5__2026-07-02__1": "absent",
  "25003111018071009__S5__2026-07-06__1": "absent",
  "25003111018071009__S5__2026-07-08__1": "absent",
  "25003111018071010__S5__2026-07-01__1": "absent",
  "25003111018071010__S5__2026-07-02__1": "absent",
  "25003111018071010__S5__2026-07-06__1": "absent",
  "25003111018071010__S5__2026-07-08__1": "absent",
  "25003111025081001__S6__2026-07-01__1": "absent",
  "25003111025081001__S6__2026-07-06__1": "present",
  "25003111025081001__S6__2026-07-07__1": "absent",
  "25003111025081001__S6__2026-07-08__1": "absent",
  "25003111025081008__S6__2026-07-01__1": "absent",
  "25003111025081008__S6__2026-07-06__1": "absent",
  "25003111025081008__S6__2026-07-07__1": "absent",
  "25003111025081008__S6__2026-07-08__1": "present",
  "25003111025081009__S6__2026-07-01__1": "absent",
  "25003111025081009__S6__2026-07-06__1": "present",
  "25003111025081009__S6__2026-07-07__1": "absent",
  "25003111025081009__S6__2026-07-08__1": "present",
  "25003111025081010__S6__2026-07-01__1": "absent",
  "25003111025081010__S6__2026-07-06__1": "present",
  "25003111025081010__S6__2026-07-07__1": "absent",
  "25003111025081010__S6__2026-07-08__1": "absent",
  "25003111025081018__S6__2026-07-01__1": "absent",
  "25003111025081018__S6__2026-07-06__1": "absent",
  "25003111025081018__S6__2026-07-07__1": "absent",
  "25003111025081018__S6__2026-07-08__1": "absent",
  "25003111025081023__S6__2026-07-01__1": "absent",
  "25003111025081023__S6__2026-07-06__1": "absent",
  "25003111025081023__S6__2026-07-07__1": "absent",
  "25003111025081023__S6__2026-07-08__1": "absent",
  "25003111025081026__S6__2026-07-01__1": "absent",
  "25003111025081026__S6__2026-07-06__1": "absent",
  "25003111025081026__S6__2026-07-07__1": "absent",
  "25003111025081026__S6__2026-07-08__1": "absent",
  "25003111025081028__S6__2026-07-01__1": "absent",
  "25003111025081028__S6__2026-07-06__1": "absent",
  "25003111025081028__S6__2026-07-07__1": "present",
  "25003111025081028__S6__2026-07-08__1": "present",
  "25003111025081039__S6__2026-07-01__1": "absent",
  "25003111025081039__S6__2026-07-06__1": "absent",
  "25003111025081039__S6__2026-07-07__1": "absent",
  "25003111025081039__S6__2026-07-08__1": "absent",
  "25003111025081043__S6__2026-07-01__1": "absent",
  "25003111025081043__S6__2026-07-06__1": "present",
  "25003111025081043__S6__2026-07-07__1": "present",
  "25003111025081043__S6__2026-07-08__1": "absent",
  "25003111025081051__S6__2026-07-01__1": "absent",
  "25003111025081051__S6__2026-07-06__1": "absent",
  "25003111025081051__S6__2026-07-07__1": "absent",
  "25003111025081051__S6__2026-07-08__1": "absent",
  "25003111025081055__S6__2026-07-01__1": "present",
  "25003111025081055__S6__2026-07-06__1": "present",
  "25003111025081055__S6__2026-07-07__1": "present",
  "25003111025081055__S6__2026-07-08__1": "present",
  "25003311025081001__S6__2026-07-01__1": "absent",
  "25003311025081001__S6__2026-07-06__1": "present",
  "25003311025081001__S6__2026-07-07__1": "present",
  "25003311025081001__S6__2026-07-08__1": "present",
  "25003311025081002__S6__2026-07-01__1": "absent",
  "25003311025081002__S6__2026-07-06__1": "present",
  "25003311025081002__S6__2026-07-07__1": "present",
  "25003311025081002__S6__2026-07-08__1": "present",
  "25003311025081003__S6__2026-07-01__1": "absent",
  "25003311025081003__S6__2026-07-06__1": "absent",
  "25003311025081003__S6__2026-07-07__1": "absent",
  "25003311025081003__S6__2026-07-08__1": "absent",
  "25003311025081004__S6__2026-07-01__1": "absent",
  "25003311025081004__S6__2026-07-06__1": "present",
  "25003311025081004__S6__2026-07-07__1": "present",
  "25003311025081004__S6__2026-07-08__1": "absent",
  "25003311025081005__S6__2026-07-01__1": "absent",
  "25003311025081005__S6__2026-07-06__1": "present",
  "25003311025081005__S6__2026-07-07__1": "present",
  "25003311025081005__S6__2026-07-08__1": "present",
  "25003111018071001__S6__2026-07-01__1": "absent",
  "25003111018071001__S6__2026-07-06__1": "absent",
  "25003111018071001__S6__2026-07-07__1": "absent",
  "25003111018071001__S6__2026-07-08__1": "absent",
  "25003111018071003__S6__2026-07-01__1": "absent",
  "25003111018071003__S6__2026-07-06__1": "absent",
  "25003111018071003__S6__2026-07-07__1": "absent",
  "25003111018071003__S6__2026-07-08__1": "absent",
  "25003111018071004__S6__2026-07-01__1": "absent",
  "25003111018071004__S6__2026-07-06__1": "absent",
  "25003111018071004__S6__2026-07-07__1": "absent",
  "25003111018071004__S6__2026-07-08__1": "absent",
  "25003111018071005__S6__2026-07-01__1": "absent",
  "25003111018071005__S6__2026-07-06__1": "absent",
  "25003111018071005__S6__2026-07-07__1": "absent",
  "25003111018071005__S6__2026-07-08__1": "absent",
  "25003111018071006__S6__2026-07-01__1": "absent",
  "25003111018071006__S6__2026-07-06__1": "absent",
  "25003111018071006__S6__2026-07-07__1": "absent",
  "25003111018071006__S6__2026-07-08__1": "absent",
  "25003111018071007__S6__2026-07-01__1": "absent",
  "25003111018071007__S6__2026-07-06__1": "absent",
  "25003111018071007__S6__2026-07-07__1": "absent",
  "25003111018071007__S6__2026-07-08__1": "absent",
  "25003111018071008__S6__2026-07-01__1": "absent",
  "25003111018071008__S6__2026-07-06__1": "absent",
  "25003111018071008__S6__2026-07-07__1": "absent",
  "25003111018071008__S6__2026-07-08__1": "present",
  "25003111018071009__S6__2026-07-01__1": "absent",
  "25003111018071009__S6__2026-07-06__1": "absent",
  "25003111018071009__S6__2026-07-07__1": "absent",
  "25003111018071009__S6__2026-07-08__1": "absent",
  "25003111018071010__S6__2026-07-01__1": "absent",
  "25003111018071010__S6__2026-07-06__1": "present",
  "25003111018071010__S6__2026-07-07__1": "absent",
  "25003111018071010__S6__2026-07-08__1": "present",
  "25003111025081001__S7__2026-07-06__1": "present",
  "25003111025081001__S7__2026-07-06__2": "absent",
  "25003111025081001__S7__2026-07-07__1": "absent",
  "25003111025081001__S7__2026-07-08__1": "absent",
  "25003111025081002__S7__2026-07-06__1": "absent",
  "25003111025081002__S7__2026-07-06__2": "absent",
  "25003111025081002__S7__2026-07-07__1": "absent",
  "25003111025081002__S7__2026-07-08__1": "absent",
  "25003111025081004__S7__2026-07-06__1": "present",
  "25003111025081004__S7__2026-07-06__2": "present",
  "25003111025081004__S7__2026-07-07__1": "present",
  "25003111025081004__S7__2026-07-08__1": "present",
  "25003111025081005__S7__2026-07-06__1": "absent",
  "25003111025081005__S7__2026-07-06__2": "absent",
  "25003111025081005__S7__2026-07-07__1": "absent",
  "25003111025081005__S7__2026-07-08__1": "absent",
  "25003111025081006__S7__2026-07-06__1": "present",
  "25003111025081006__S7__2026-07-06__2": "present",
  "25003111025081006__S7__2026-07-07__1": "present",
  "25003111025081006__S7__2026-07-08__1": "absent",
  "25003111025081007__S7__2026-07-06__1": "present",
  "25003111025081007__S7__2026-07-06__2": "absent",
  "25003111025081007__S7__2026-07-07__1": "present",
  "25003111025081007__S7__2026-07-08__1": "present",
  "25003111025081008__S7__2026-07-06__1": "absent",
  "25003111025081008__S7__2026-07-06__2": "absent",
  "25003111025081008__S7__2026-07-07__1": "absent",
  "25003111025081008__S7__2026-07-08__1": "present",
  "25003111025081009__S7__2026-07-06__1": "present",
  "25003111025081009__S7__2026-07-06__2": "absent",
  "25003111025081009__S7__2026-07-07__1": "absent",
  "25003111025081009__S7__2026-07-08__1": "present",
  "25003111025081010__S7__2026-07-06__1": "present",
  "25003111025081010__S7__2026-07-06__2": "absent",
  "25003111025081010__S7__2026-07-07__1": "absent",
  "25003111025081010__S7__2026-07-08__1": "absent",
  "25003111025081011__S7__2026-07-06__1": "present",
  "25003111025081011__S7__2026-07-06__2": "present",
  "25003111025081011__S7__2026-07-07__1": "present",
  "25003111025081011__S7__2026-07-08__1": "present",
  "25003111025081013__S7__2026-07-06__1": "absent",
  "25003111025081013__S7__2026-07-06__2": "absent",
  "25003111025081013__S7__2026-07-07__1": "absent",
  "25003111025081013__S7__2026-07-08__1": "absent",
  "25003111025081014__S7__2026-07-06__1": "present",
  "25003111025081014__S7__2026-07-06__2": "present",
  "25003111025081014__S7__2026-07-07__1": "present",
  "25003111025081014__S7__2026-07-08__1": "present",
  "25003111025081015__S7__2026-07-06__1": "present",
  "25003111025081015__S7__2026-07-06__2": "absent",
  "25003111025081015__S7__2026-07-07__1": "present",
  "25003111025081015__S7__2026-07-08__1": "present",
  "25003111025081016__S7__2026-07-06__1": "absent",
  "25003111025081016__S7__2026-07-06__2": "absent",
  "25003111025081016__S7__2026-07-07__1": "absent",
  "25003111025081016__S7__2026-07-08__1": "absent",
  "25003111025081017__S7__2026-07-06__1": "present",
  "25003111025081017__S7__2026-07-06__2": "present",
  "25003111025081017__S7__2026-07-07__1": "present",
  "25003111025081017__S7__2026-07-08__1": "absent",
  "25003111025081018__S7__2026-07-06__1": "absent",
  "25003111025081018__S7__2026-07-06__2": "absent",
  "25003111025081018__S7__2026-07-07__1": "absent",
  "25003111025081018__S7__2026-07-08__1": "absent",
  "25003111025081019__S7__2026-07-06__1": "absent",
  "25003111025081019__S7__2026-07-06__2": "absent",
  "25003111025081019__S7__2026-07-07__1": "present",
  "25003111025081019__S7__2026-07-08__1": "present",
  "25003111025081020__S7__2026-07-06__1": "absent",
  "25003111025081020__S7__2026-07-06__2": "absent",
  "25003111025081020__S7__2026-07-07__1": "absent",
  "25003111025081020__S7__2026-07-08__1": "absent",
  "25003111025081021__S7__2026-07-06__1": "absent",
  "25003111025081021__S7__2026-07-06__2": "absent",
  "25003111025081021__S7__2026-07-07__1": "absent",
  "25003111025081021__S7__2026-07-08__1": "present",
  "25003111025081023__S7__2026-07-06__1": "absent",
  "25003111025081023__S7__2026-07-06__2": "absent",
  "25003111025081023__S7__2026-07-07__1": "absent",
  "25003111025081023__S7__2026-07-08__1": "absent",
  "25003111025081024__S7__2026-07-06__1": "present",
  "25003111025081024__S7__2026-07-06__2": "present",
  "25003111025081024__S7__2026-07-07__1": "present",
  "25003111025081024__S7__2026-07-08__1": "present",
  "25003111025081025__S7__2026-07-06__1": "present",
  "25003111025081025__S7__2026-07-06__2": "present",
  "25003111025081025__S7__2026-07-07__1": "present",
  "25003111025081025__S7__2026-07-08__1": "present",
  "25003111025081026__S7__2026-07-06__1": "absent",
  "25003111025081026__S7__2026-07-06__2": "absent",
  "25003111025081026__S7__2026-07-07__1": "absent",
  "25003111025081026__S7__2026-07-08__1": "absent",
  "25003111025081027__S7__2026-07-06__1": "absent",
  "25003111025081027__S7__2026-07-06__2": "absent",
  "25003111025081027__S7__2026-07-07__1": "absent",
  "25003111025081027__S7__2026-07-08__1": "absent",
  "25003111025081028__S7__2026-07-06__1": "absent",
  "25003111025081028__S7__2026-07-06__2": "absent",
  "25003111025081028__S7__2026-07-07__1": "present",
  "25003111025081028__S7__2026-07-08__1": "present",
  "25003111025081029__S7__2026-07-06__1": "absent",
  "25003111025081029__S7__2026-07-06__2": "absent",
  "25003111025081029__S7__2026-07-07__1": "absent",
  "25003111025081029__S7__2026-07-08__1": "absent",
  "25003111025081030__S7__2026-07-06__1": "present",
  "25003111025081030__S7__2026-07-06__2": "absent",
  "25003111025081030__S7__2026-07-07__1": "present",
  "25003111025081030__S7__2026-07-08__1": "absent",
  "25003111025081031__S7__2026-07-06__1": "absent",
  "25003111025081031__S7__2026-07-06__2": "absent",
  "25003111025081031__S7__2026-07-07__1": "absent",
  "25003111025081031__S7__2026-07-08__1": "absent",
  "25003111025081032__S7__2026-07-06__1": "absent",
  "25003111025081032__S7__2026-07-06__2": "absent",
  "25003111025081032__S7__2026-07-07__1": "absent",
  "25003111025081032__S7__2026-07-08__1": "absent",
  "25003111025081034__S7__2026-07-06__1": "present",
  "25003111025081034__S7__2026-07-06__2": "absent",
  "25003111025081034__S7__2026-07-07__1": "absent",
  "25003111025081034__S7__2026-07-08__1": "present",
  "25003111025081035__S7__2026-07-06__1": "present",
  "25003111025081035__S7__2026-07-06__2": "present",
  "25003111025081035__S7__2026-07-07__1": "present",
  "25003111025081035__S7__2026-07-08__1": "present",
  "25003111025081036__S7__2026-07-06__1": "present",
  "25003111025081036__S7__2026-07-06__2": "present",
  "25003111025081036__S7__2026-07-07__1": "present",
  "25003111025081036__S7__2026-07-08__1": "present",
  "25003111025081038__S7__2026-07-06__1": "present",
  "25003111025081038__S7__2026-07-06__2": "absent",
  "25003111025081038__S7__2026-07-07__1": "present",
  "25003111025081038__S7__2026-07-08__1": "present",
  "25003111025081039__S7__2026-07-06__1": "absent",
  "25003111025081039__S7__2026-07-06__2": "absent",
  "25003111025081039__S7__2026-07-07__1": "absent",
  "25003111025081039__S7__2026-07-08__1": "absent",
  "25003111025081040__S7__2026-07-06__1": "present",
  "25003111025081040__S7__2026-07-06__2": "present",
  "25003111025081040__S7__2026-07-07__1": "present",
  "25003111025081040__S7__2026-07-08__1": "present",
  "25003111025081041__S7__2026-07-06__1": "present",
  "25003111025081041__S7__2026-07-06__2": "present",
  "25003111025081041__S7__2026-07-07__1": "present",
  "25003111025081041__S7__2026-07-08__1": "present",
  "25003111025081043__S7__2026-07-06__1": "present",
  "25003111025081043__S7__2026-07-06__2": "absent",
  "25003111025081043__S7__2026-07-07__1": "present",
  "25003111025081043__S7__2026-07-08__1": "absent",
  "25003111025081044__S7__2026-07-06__1": "absent",
  "25003111025081044__S7__2026-07-06__2": "absent",
  "25003111025081044__S7__2026-07-07__1": "absent",
  "25003111025081044__S7__2026-07-08__1": "absent",
  "25003111025081046__S7__2026-07-06__1": "absent",
  "25003111025081046__S7__2026-07-06__2": "absent",
  "25003111025081046__S7__2026-07-07__1": "absent",
  "25003111025081046__S7__2026-07-08__1": "absent",
  "25003111025081047__S7__2026-07-06__1": "absent",
  "25003111025081047__S7__2026-07-06__2": "absent",
  "25003111025081047__S7__2026-07-07__1": "absent",
  "25003111025081047__S7__2026-07-08__1": "absent",
  "25003111025081048__S7__2026-07-06__1": "absent",
  "25003111025081048__S7__2026-07-06__2": "absent",
  "25003111025081048__S7__2026-07-07__1": "absent",
  "25003111025081048__S7__2026-07-08__1": "absent",
  "25003111025081049__S7__2026-07-06__1": "present",
  "25003111025081049__S7__2026-07-06__2": "present",
  "25003111025081049__S7__2026-07-07__1": "present",
  "25003111025081049__S7__2026-07-08__1": "present",
  "25003111025081050__S7__2026-07-06__1": "present",
  "25003111025081050__S7__2026-07-06__2": "present",
  "25003111025081050__S7__2026-07-07__1": "present",
  "25003111025081050__S7__2026-07-08__1": "present",
  "25003111025081051__S7__2026-07-06__1": "absent",
  "25003111025081051__S7__2026-07-06__2": "absent",
  "25003111025081051__S7__2026-07-07__1": "absent",
  "25003111025081051__S7__2026-07-08__1": "absent",
  "25003111025081052__S7__2026-07-06__1": "present",
  "25003111025081052__S7__2026-07-06__2": "present",
  "25003111025081052__S7__2026-07-07__1": "present",
  "25003111025081052__S7__2026-07-08__1": "present",
  "25003111025081053__S7__2026-07-06__1": "absent",
  "25003111025081053__S7__2026-07-06__2": "absent",
  "25003111025081053__S7__2026-07-07__1": "absent",
  "25003111025081053__S7__2026-07-08__1": "present",
  "25003111025081055__S7__2026-07-06__1": "present",
  "25003111025081055__S7__2026-07-06__2": "present",
  "25003111025081055__S7__2026-07-07__1": "present",
  "25003111025081055__S7__2026-07-08__1": "present",
  "25003111025081056__S7__2026-07-06__1": "absent",
  "25003111025081056__S7__2026-07-06__2": "present",
  "25003111025081056__S7__2026-07-07__1": "present",
  "25003111025081056__S7__2026-07-08__1": "present",
  "25003311025081001__S7__2026-07-06__1": "present",
  "25003311025081001__S7__2026-07-06__2": "present",
  "25003311025081001__S7__2026-07-07__1": "present",
  "25003311025081001__S7__2026-07-08__1": "present",
  "25003311025081002__S7__2026-07-06__1": "present",
  "25003311025081002__S7__2026-07-06__2": "present",
  "25003311025081002__S7__2026-07-07__1": "present",
  "25003311025081002__S7__2026-07-08__1": "present",
  "25003311025081003__S7__2026-07-06__1": "absent",
  "25003311025081003__S7__2026-07-06__2": "absent",
  "25003311025081003__S7__2026-07-07__1": "absent",
  "25003311025081003__S7__2026-07-08__1": "absent",
  "25003311025081004__S7__2026-07-06__1": "present",
  "25003311025081004__S7__2026-07-06__2": "present",
  "25003311025081004__S7__2026-07-07__1": "present",
  "25003311025081004__S7__2026-07-08__1": "absent",
  "25003311025081005__S7__2026-07-06__1": "present",
  "25003311025081005__S7__2026-07-06__2": "present",
  "25003311025081005__S7__2026-07-07__1": "present",
  "25003311025081005__S7__2026-07-08__1": "present",
  "25003111018071001__S7__2026-07-06__1": "absent",
  "25003111018071001__S7__2026-07-06__2": "absent",
  "25003111018071001__S7__2026-07-07__1": "present",
  "25003111018071001__S7__2026-07-08__1": "present",
  "25003111018071003__S7__2026-07-06__1": "absent",
  "25003111018071003__S7__2026-07-06__2": "absent",
  "25003111018071003__S7__2026-07-07__1": "absent",
  "25003111018071003__S7__2026-07-08__1": "absent",
  "25003111018071004__S7__2026-07-06__1": "absent",
  "25003111018071004__S7__2026-07-06__2": "absent",
  "25003111018071004__S7__2026-07-07__1": "absent",
  "25003111018071004__S7__2026-07-08__1": "absent",
  "25003111018071005__S7__2026-07-06__1": "absent",
  "25003111018071005__S7__2026-07-06__2": "absent",
  "25003111018071005__S7__2026-07-07__1": "absent",
  "25003111018071005__S7__2026-07-08__1": "absent",
  "25003111018071006__S7__2026-07-06__1": "absent",
  "25003111018071006__S7__2026-07-06__2": "absent",
  "25003111018071006__S7__2026-07-07__1": "absent",
  "25003111018071006__S7__2026-07-08__1": "absent",
  "25003111018071007__S7__2026-07-06__1": "absent",
  "25003111018071007__S7__2026-07-06__2": "absent",
  "25003111018071007__S7__2026-07-07__1": "present",
  "25003111018071007__S7__2026-07-08__1": "present",
  "25003111018071008__S7__2026-07-06__1": "absent",
  "25003111018071008__S7__2026-07-06__2": "absent",
  "25003111018071008__S7__2026-07-07__1": "absent",
  "25003111018071008__S7__2026-07-08__1": "present",
  "25003111018071009__S7__2026-07-06__1": "absent",
  "25003111018071009__S7__2026-07-06__2": "absent",
  "25003111018071009__S7__2026-07-07__1": "present",
  "25003111018071009__S7__2026-07-08__1": "present",
  "25003111018071010__S7__2026-07-06__1": "absent",
  "25003111018071010__S7__2026-07-06__2": "absent",
  "25003111018071010__S7__2026-07-07__1": "absent",
  "25003111018071010__S7__2026-07-08__1": "present",
  "25003111025081004__S8__2026-07-01__1": "absent",
  "25003111025081004__S8__2026-07-02__1": "absent",
  "25003111025081004__S8__2026-07-06__1": "present",
  "25003111025081004__S8__2026-07-07__1": "present",
  "25003111025081004__S8__2026-07-07__2": "present",
  "25003111025081005__S8__2026-07-01__1": "absent",
  "25003111025081005__S8__2026-07-02__1": "absent",
  "25003111025081005__S8__2026-07-06__1": "absent",
  "25003111025081005__S8__2026-07-07__1": "absent",
  "25003111025081005__S8__2026-07-07__2": "absent",
  "25003111025081006__S8__2026-07-01__1": "present",
  "25003111025081006__S8__2026-07-02__1": "present",
  "25003111025081006__S8__2026-07-06__1": "present",
  "25003111025081006__S8__2026-07-07__1": "present",
  "25003111025081006__S8__2026-07-07__2": "absent",
  "25003111025081008__S8__2026-07-01__1": "absent",
  "25003111025081008__S8__2026-07-02__1": "absent",
  "25003111025081008__S8__2026-07-06__1": "absent",
  "25003111025081008__S8__2026-07-07__1": "absent",
  "25003111025081008__S8__2026-07-07__2": "absent",
  "25003111025081009__S8__2026-07-01__1": "absent",
  "25003111025081009__S8__2026-07-02__1": "absent",
  "25003111025081009__S8__2026-07-06__1": "absent",
  "25003111025081009__S8__2026-07-07__1": "absent",
  "25003111025081009__S8__2026-07-07__2": "absent",
  "25003111025081010__S8__2026-07-01__1": "absent",
  "25003111025081010__S8__2026-07-02__1": "absent",
  "25003111025081010__S8__2026-07-06__1": "present",
  "25003111025081010__S8__2026-07-07__1": "absent",
  "25003111025081010__S8__2026-07-07__2": "absent",
  "25003111025081017__S8__2026-07-01__1": "present",
  "25003111025081017__S8__2026-07-02__1": "present",
  "25003111025081017__S8__2026-07-06__1": "present",
  "25003111025081017__S8__2026-07-07__1": "present",
  "25003111025081017__S8__2026-07-07__2": "absent",
  "25003111025081018__S8__2026-07-01__1": "absent",
  "25003111025081018__S8__2026-07-02__1": "absent",
  "25003111025081018__S8__2026-07-06__1": "absent",
  "25003111025081018__S8__2026-07-07__1": "absent",
  "25003111025081018__S8__2026-07-07__2": "absent",
  "25003111025081019__S8__2026-07-01__1": "present",
  "25003111025081019__S8__2026-07-02__1": "present",
  "25003111025081019__S8__2026-07-06__1": "absent",
  "25003111025081019__S8__2026-07-07__1": "present",
  "25003111025081019__S8__2026-07-07__2": "present",
  "25003111025081020__S8__2026-07-01__1": "absent",
  "25003111025081020__S8__2026-07-02__1": "absent",
  "25003111025081020__S8__2026-07-06__1": "absent",
  "25003111025081020__S8__2026-07-07__1": "absent",
  "25003111025081020__S8__2026-07-07__2": "absent",
  "25003111025081021__S8__2026-07-01__1": "absent",
  "25003111025081021__S8__2026-07-02__1": "absent",
  "25003111025081021__S8__2026-07-06__1": "absent",
  "25003111025081021__S8__2026-07-07__1": "absent",
  "25003111025081021__S8__2026-07-07__2": "present",
  "25003111025081023__S8__2026-07-01__1": "absent",
  "25003111025081023__S8__2026-07-02__1": "absent",
  "25003111025081023__S8__2026-07-06__1": "absent",
  "25003111025081023__S8__2026-07-07__1": "absent",
  "25003111025081023__S8__2026-07-07__2": "absent",
  "25003111025081024__S8__2026-07-01__1": "present",
  "25003111025081024__S8__2026-07-02__1": "present",
  "25003111025081024__S8__2026-07-06__1": "present",
  "25003111025081024__S8__2026-07-07__1": "present",
  "25003111025081024__S8__2026-07-07__2": "present",
  "25003111025081025__S8__2026-07-01__1": "present",
  "25003111025081025__S8__2026-07-02__1": "present",
  "25003111025081025__S8__2026-07-06__1": "present",
  "25003111025081025__S8__2026-07-07__1": "present",
  "25003111025081025__S8__2026-07-07__2": "present",
  "25003111025081026__S8__2026-07-01__1": "absent",
  "25003111025081026__S8__2026-07-02__1": "absent",
  "25003111025081026__S8__2026-07-06__1": "absent",
  "25003111025081026__S8__2026-07-07__1": "absent",
  "25003111025081026__S8__2026-07-07__2": "absent",
  "25003111025081028__S8__2026-07-01__1": "absent",
  "25003111025081028__S8__2026-07-02__1": "absent",
  "25003111025081028__S8__2026-07-06__1": "absent",
  "25003111025081028__S8__2026-07-07__1": "present",
  "25003111025081028__S8__2026-07-07__2": "absent",
  "25003111025081030__S8__2026-07-01__1": "absent",
  "25003111025081030__S8__2026-07-02__1": "absent",
  "25003111025081030__S8__2026-07-06__1": "present",
  "25003111025081030__S8__2026-07-07__1": "present",
  "25003111025081030__S8__2026-07-07__2": "present",
  "25003111025081032__S8__2026-07-01__1": "absent",
  "25003111025081032__S8__2026-07-02__1": "absent",
  "25003111025081032__S8__2026-07-06__1": "absent",
  "25003111025081032__S8__2026-07-07__1": "absent",
  "25003111025081032__S8__2026-07-07__2": "absent",
  "25003111025081034__S8__2026-07-01__1": "absent",
  "25003111025081034__S8__2026-07-02__1": "absent",
  "25003111025081034__S8__2026-07-06__1": "present",
  "25003111025081034__S8__2026-07-07__1": "absent",
  "25003111025081034__S8__2026-07-07__2": "present",
  "25003111025081035__S8__2026-07-01__1": "present",
  "25003111025081035__S8__2026-07-02__1": "present",
  "25003111025081035__S8__2026-07-06__1": "present",
  "25003111025081035__S8__2026-07-07__1": "present",
  "25003111025081035__S8__2026-07-07__2": "present",
  "25003111025081036__S8__2026-07-01__1": "absent",
  "25003111025081036__S8__2026-07-02__1": "absent",
  "25003111025081036__S8__2026-07-06__1": "present",
  "25003111025081036__S8__2026-07-07__1": "present",
  "25003111025081036__S8__2026-07-07__2": "present",
  "25003111025081040__S8__2026-07-01__1": "present",
  "25003111025081040__S8__2026-07-02__1": "absent",
  "25003111025081040__S8__2026-07-06__1": "present",
  "25003111025081040__S8__2026-07-07__1": "present",
  "25003111025081040__S8__2026-07-07__2": "present",
  "25003111025081041__S8__2026-07-01__1": "present",
  "25003111025081041__S8__2026-07-02__1": "absent",
  "25003111025081041__S8__2026-07-06__1": "present",
  "25003111025081041__S8__2026-07-07__1": "present",
  "25003111025081041__S8__2026-07-07__2": "present",
  "25003111025081043__S8__2026-07-01__1": "absent",
  "25003111025081043__S8__2026-07-02__1": "absent",
  "25003111025081043__S8__2026-07-06__1": "absent",
  "25003111025081043__S8__2026-07-07__1": "present",
  "25003111025081043__S8__2026-07-07__2": "absent",
  "25003111025081046__S8__2026-07-01__1": "absent",
  "25003111025081046__S8__2026-07-02__1": "absent",
  "25003111025081046__S8__2026-07-06__1": "absent",
  "25003111025081046__S8__2026-07-07__1": "absent",
  "25003111025081046__S8__2026-07-07__2": "absent",
  "25003111025081048__S8__2026-07-01__1": "absent",
  "25003111025081048__S8__2026-07-02__1": "absent",
  "25003111025081048__S8__2026-07-06__1": "absent",
  "25003111025081048__S8__2026-07-07__1": "absent",
  "25003111025081048__S8__2026-07-07__2": "absent",
  "25003111025081049__S8__2026-07-01__1": "absent",
  "25003111025081049__S8__2026-07-02__1": "absent",
  "25003111025081049__S8__2026-07-06__1": "present",
  "25003111025081049__S8__2026-07-07__1": "present",
  "25003111025081049__S8__2026-07-07__2": "present",
  "25003111025081050__S8__2026-07-01__1": "present",
  "25003111025081050__S8__2026-07-02__1": "absent",
  "25003111025081050__S8__2026-07-06__1": "present",
  "25003111025081050__S8__2026-07-07__1": "present",
  "25003111025081050__S8__2026-07-07__2": "present",
  "25003111025081051__S8__2026-07-01__1": "absent",
  "25003111025081051__S8__2026-07-02__1": "absent",
  "25003111025081051__S8__2026-07-06__1": "absent",
  "25003111025081051__S8__2026-07-07__1": "absent",
  "25003111025081051__S8__2026-07-07__2": "absent",
  "25003111025081052__S8__2026-07-01__1": "present",
  "25003111025081052__S8__2026-07-02__1": "present",
  "25003111025081052__S8__2026-07-06__1": "present",
  "25003111025081052__S8__2026-07-07__1": "present",
  "25003111025081052__S8__2026-07-07__2": "present",
  "25003111025081055__S8__2026-07-01__1": "present",
  "25003111025081055__S8__2026-07-02__1": "present",
  "25003111025081055__S8__2026-07-06__1": "present",
  "25003111025081055__S8__2026-07-07__1": "present",
  "25003111025081055__S8__2026-07-07__2": "present",
  "25003111025081056__S8__2026-07-01__1": "present",
  "25003111025081056__S8__2026-07-02__1": "present",
  "25003111025081056__S8__2026-07-06__1": "present",
  "25003111025081056__S8__2026-07-07__1": "present",
  "25003111025081056__S8__2026-07-07__2": "present",
  "25003311025081001__S8__2026-07-01__1": "absent",
  "25003311025081001__S8__2026-07-02__1": "absent",
  "25003311025081001__S8__2026-07-06__1": "present",
  "25003311025081001__S8__2026-07-07__1": "present",
  "25003311025081001__S8__2026-07-07__2": "present",
  "25003311025081002__S8__2026-07-01__1": "absent",
  "25003311025081002__S8__2026-07-02__1": "absent",
  "25003311025081002__S8__2026-07-06__1": "present",
  "25003311025081002__S8__2026-07-07__1": "present",
  "25003311025081002__S8__2026-07-07__2": "present",
  "25003311025081003__S8__2026-07-01__1": "absent",
  "25003311025081003__S8__2026-07-02__1": "absent",
  "25003311025081003__S8__2026-07-06__1": "absent",
  "25003311025081003__S8__2026-07-07__1": "absent",
  "25003311025081003__S8__2026-07-07__2": "absent",
  "25003311025081004__S8__2026-07-01__1": "absent",
  "25003311025081004__S8__2026-07-02__1": "absent",
  "25003311025081004__S8__2026-07-06__1": "present",
  "25003311025081004__S8__2026-07-07__1": "present",
  "25003311025081004__S8__2026-07-07__2": "absent",
  "25003311025081005__S8__2026-07-01__1": "absent",
  "25003311025081005__S8__2026-07-02__1": "absent",
  "25003311025081005__S8__2026-07-06__1": "present",
  "25003311025081005__S8__2026-07-07__1": "present",
  "25003311025081005__S8__2026-07-07__2": "present",
  "25003111018071001__S8__2026-07-01__1": "absent",
  "25003111018071001__S8__2026-07-02__1": "absent",
  "25003111018071001__S8__2026-07-06__1": "absent",
  "25003111018071001__S8__2026-07-07__1": "present",
  "25003111018071001__S8__2026-07-07__2": "present",
  "25003111018071003__S8__2026-07-01__1": "absent",
  "25003111018071003__S8__2026-07-02__1": "absent",
  "25003111018071003__S8__2026-07-06__1": "absent",
  "25003111018071003__S8__2026-07-07__1": "absent",
  "25003111018071003__S8__2026-07-07__2": "absent",
  "25003111018071004__S8__2026-07-01__1": "absent",
  "25003111018071004__S8__2026-07-02__1": "absent",
  "25003111018071004__S8__2026-07-06__1": "absent",
  "25003111018071004__S8__2026-07-07__1": "absent",
  "25003111018071004__S8__2026-07-07__2": "absent",
  "25003111018071005__S8__2026-07-01__1": "absent",
  "25003111018071005__S8__2026-07-02__1": "absent",
  "25003111018071005__S8__2026-07-06__1": "absent",
  "25003111018071005__S8__2026-07-07__1": "absent",
  "25003111018071005__S8__2026-07-07__2": "absent",
  "25003111018071006__S8__2026-07-01__1": "absent",
  "25003111018071006__S8__2026-07-02__1": "absent",
  "25003111018071006__S8__2026-07-06__1": "absent",
  "25003111018071006__S8__2026-07-07__1": "absent",
  "25003111018071006__S8__2026-07-07__2": "absent",
  "25003111018071007__S8__2026-07-01__1": "absent",
  "25003111018071007__S8__2026-07-02__1": "absent",
  "25003111018071007__S8__2026-07-06__1": "absent",
  "25003111018071007__S8__2026-07-07__1": "present",
  "25003111018071007__S8__2026-07-07__2": "present",
  "25003111018071008__S8__2026-07-01__1": "absent",
  "25003111018071008__S8__2026-07-02__1": "absent",
  "25003111018071008__S8__2026-07-06__1": "absent",
  "25003111018071008__S8__2026-07-07__1": "absent",
  "25003111018071008__S8__2026-07-07__2": "absent",
  "25003111018071009__S8__2026-07-01__1": "absent",
  "25003111018071009__S8__2026-07-02__1": "absent",
  "25003111018071009__S8__2026-07-06__1": "absent",
  "25003111018071009__S8__2026-07-07__1": "present",
  "25003111018071009__S8__2026-07-07__2": "present",
  "25003111018071010__S8__2026-07-01__1": "absent",
  "25003111018071010__S8__2026-07-02__1": "absent",
  "25003111018071010__S8__2026-07-06__1": "present",
  "25003111018071010__S8__2026-07-07__1": "absent",
  "25003111018071010__S8__2026-07-07__2": "absent",
};

const REAL_SEED_SESSIONS = {
  "S1__2026-07-01__1": { time: "9:15 - 10:00", remark: "" },
  "S1__2026-07-02__1": { time: "2:00 - 3:30", remark: "" },
  "S1__2026-07-06__1": { time: "10:45 - 11:30", remark: "" },
  "S1__2026-07-07__1": { time: "9:15 - 10:00", remark: "" },
  "S2__2026-07-01__1": { time: "10:00 - 10:45", remark: "" },
  "S2__2026-07-02__1": { time: "11:30 - 12:15", remark: "" },
  "S2__2026-07-06__1": { time: "3:30 - 4:15", remark: "" },
  "S2__2026-07-07__1": { time: "2:45 - 4:15", remark: "" },
  "S2__2026-07-08__1": { time: "9:15 - 10:00", remark: "" },
  "S3__2026-07-01__1": { time: "10:45 - 11:30", remark: "" },
  "S3__2026-07-02__1": { time: "10:45 - 11:30", remark: "" },
  "S3__2026-07-06__1": { time: "11:30 - 12:15", remark: "" },
  "S3__2026-07-08__1": { time: "9:15 - 10:00", remark: "" },
  "S4__2026-07-01__1": { time: "11:30 - 12:15", remark: "" },
  "S4__2026-07-02__1": { time: "9:15 - 10:45", remark: "" },
  "S4__2026-07-06__1": { time: "10:00 - 10:45", remark: "" },
  "S4__2026-07-07__1": { time: "2:00 - 2:45", remark: "" },
  "S4__2026-07-08__1": { time: "11:30 - 12:15", remark: "" },
  "S5__2026-07-01__1": { time: "11:30 - 12:15", remark: "" },
  "S5__2026-07-02__1": { time: "9:15 - 10:45", remark: "" },
  "S5__2026-07-06__1": { time: "10:00 - 10:45", remark: "" },
  "S5__2026-07-08__1": { time: "11:30 - 12:15", remark: "" },
  "S6__2026-07-01__1": { time: "11:30 - 12:15", remark: "" },
  "S6__2026-07-06__1": { time: "10:00 - 10:45", remark: "" },
  "S6__2026-07-07__1": { time: "10:45 - 11:30", remark: "" },
  "S6__2026-07-08__1": { time: "11:30 - 12:15", remark: "" },
  "S7__2026-07-06__1": { time: "9:15 - 10:00", remark: "" },
  "S7__2026-07-06__2": { time: "2:00 - 3:30", remark: "" },
  "S7__2026-07-07__1": { time: "10:00 - 10:45", remark: "" },
  "S7__2026-07-08__1": { time: "12:15 - 1:00", remark: "" },
  "S8__2026-07-01__1": { time: "2:00 - 3:30", remark: "" },
  "S8__2026-07-02__1": { time: "12:15 - 1:00", remark: "" },
  "S8__2026-07-06__1": { time: "10:45 - 11:30", remark: "" },
  "S8__2026-07-07__1": { time: "10:45 - 11:30", remark: "" },
  "S8__2026-07-07__2": { time: "2:00 - 3:30", remark: "" },
  "S9__2026-07-02__1": { time: "12:15 - 1:00", remark: "" },
  "S9__2026-07-06__1": { time: "10:45 - 11:30", remark: "" },
  "S9__2026-07-08__1": { time: "", remark: "" },
};

function seedAttendance() {
  return { ...REAL_SEED_ATTENDANCE };
}

function seedSessions() {
  return { ...REAL_SEED_SESSIONS };
}

const TABS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutGrid },
  { id: "students", label: "Students", icon: Users },
  { id: "subjects", label: "Subjects", icon: BookOpen },
  { id: "mark", label: "Mark Attendance", icon: CalendarCheck },
  { id: "all", label: "All Attendance", icon: Table2 },
  { id: "reports", label: "Reports", icon: FileBarChart },
  { id: "audit", label: "Audit Log", icon: History },
  { id: "users", label: "Manage Users", icon: ShieldCheck, adminOnly: true },
];

// All data now lives in a shared database (via /api/state, backed by
// Upstash Redis) so every faculty member sees the same students, subjects,
// and attendance regardless of which browser/device they're on.
const SAVE_DEBOUNCE_MS = 150;
const POLL_INTERVAL_MS = 1200;

export default function AttendancePortal() {
  const [subjects, setSubjects] = useState(seedSubjects);
  const [students, setStudents] = useState(seedStudents);
  const [attendance, setAttendance] = useState({});
  const [sessions, setSessions] = useState({});
  // Audit trail: who last marked/changed each attendance record, and when.
  // Kept separate from `attendance` so all the existing status === "present"
  // checks elsewhere don't need to change. Never included in Excel exports.
  const [attendanceMeta, setAttendanceMeta] = useState({});
  const [tab, setTab] = useState("dashboard");
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [downloadSubjectId, setDownloadSubjectId] = useState("all");
  const [importingAttendance, setImportingAttendance] = useState(false);
  const attendanceFileInputRef = useRef(null);
  const [hasLoadedOnce, setHasLoadedOnce] = useState(false);
  const [loadError, setLoadError] = useState(false);
  const [syncState, setSyncState] = useState("idle"); // idle | saving | saved | error
  const idCounter = useRef(100);
  const loadedRef = useRef(false);
  const saveTimer = useRef(null);
  // Keys explicitly removed since the last successful save — sent alongside
  // the merge payload so the server knows to actually delete them (a plain
  // key-union merge alone can't tell "never existed" apart from "removed").
  const pendingDeletesRef = useRef({ attendance: new Set(), sessions: new Set(), subjects: new Set(), students: new Set() });
  // Bumped on every local edit (see the save-triggering effect below) so an
  // in-flight poll can tell whether it's become stale mid-flight.
  const localVersionRef = useRef(0);
  // Snapshot of exactly what's known-saved as of the last successful sync.
  // Every save diffs current state against this to send ONLY what actually
  // changed — critical for deletes: sending the *entire* local state on
  // every save (the previous approach) means any client with a slightly
  // stale snapshot (another tab, another device, or even this same tab a
  // moment later) can silently re-upsert a record someone else just
  // deleted, since a plain "upsert everything I know about" can't tell the
  // difference between "unchanged" and "please restore this."
  const lastSyncedRef = useRef({ attendance: {}, sessions: {}, attendanceMeta: {}, subjects: [], students: [] });

  // ---------- authentication ----------
  // Session is remembered per-browser via localStorage (just who's logged
  // in — not shared data, which still lives entirely in the database).
  const [currentUser, setCurrentUser] = useState(() => {
    try {
      const raw = localStorage.getItem("attendance-portal:session");
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  });
  const login = (user) => {
    setCurrentUser(user);
    try {
      localStorage.setItem("attendance-portal:session", JSON.stringify(user));
    } catch {
      /* ignore storage errors (e.g. private browsing) */
    }
  };
  const logout = () => {
    setCurrentUser(null);
    try {
      localStorage.removeItem("attendance-portal:session");
    } catch {
      /* ignore */
    }
  };

  const applyServerData = (data) => {
    if (data) {
      const finalSubjects = data.subjects && data.subjects.length ? data.subjects : seedSubjects;
      const finalStudents = data.students && data.students.length ? data.students : seedStudents;
      setSubjects(finalSubjects);
      setStudents(finalStudents);
      // Migrate any pre-existing "leave"/"proxy" records (from before the
      // app was simplified to just Present/Absent) into Absent, and any
      // pre-existing keys from before multi-session support (which had no
      // "slot" segment) into slot "1", so old saved data keeps working.
      const rawAttendance = data.attendance || {};
      const cleanAttendance = {};
      Object.entries(rawAttendance).forEach(([key, status]) => {
        const parts = key.split("__");
        const migratedKey = parts.length === 3 ? `${key}__1` : key;
        cleanAttendance[migratedKey] = status === "present" ? "present" : "absent";
      });
      setAttendance(cleanAttendance);

      const rawSessions = data.sessions || {};
      const cleanSessions = {};
      Object.entries(rawSessions).forEach(([key, val]) => {
        const parts = key.split("__");
        const migratedKey = parts.length === 2 ? `${key}__1` : key;
        cleanSessions[migratedKey] = val;
      });
      setSessions(cleanSessions);
      const finalMeta = data.attendanceMeta || {};
      setAttendanceMeta(finalMeta);

      // This is now the authoritative "known saved" baseline. Every future
      // save diffs against this rather than re-sending everything, so a
      // record someone else deletes can't come back just because this tab
      // happened to save something unrelated afterward.
      lastSyncedRef.current = {
        attendance: cleanAttendance,
        sessions: cleanSessions,
        attendanceMeta: finalMeta,
        subjects: finalSubjects,
        students: finalStudents,
      };
    } else {
      // nothing saved on the server yet — seed it with the real one-time
      // import. Nothing is "synced" yet, so reset the baseline to empty —
      // the next save will correctly treat all of this seed data as new
      // and upload every bit of it.
      const seedAtt = seedAttendance();
      const seedSess = seedSessions();
      setAttendance(seedAtt);
      setSessions(seedSess);
      setAttendanceMeta({});
      lastSyncedRef.current = { attendance: {}, sessions: {}, attendanceMeta: {}, subjects: [], students: [] };
    }
  };

  const loadFromServer = async () => {
    const versionBeforeFetch = localVersionRef.current;
    try {
      const res = await fetch("/api/state");
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      // If a local edit happened WHILE this fetch was in flight, this
      // response no longer reflects the latest local state — applying it
      // now would silently revert whatever the user just clicked (this is
      // what caused marks to visibly "undo themselves" a couple of times
      // before finally sticking). Discard it; the debounced save already
      // in progress will persist the newer edit, and the next poll will
      // pick up the merged result once that lands.
      if (localVersionRef.current !== versionBeforeFetch) return;
      applyServerData(data); // data === null here means the DB is genuinely empty (server said so) — safe to seed and save
      setSyncState("saved");
      // Only flip this on once we've *confirmed* what the server actually
      // has. Setting it on failure too (as a previous version of this code
      // did) meant a flaky connection on a new device would fall back to
      // local sample data and then happily auto-save that sample data over
      // whatever real data was already in the database — silently wiping
      // it. Now, on failure, we leave whatever's currently on screen alone
      // and simply don't allow saving until a load actually succeeds.
      loadedRef.current = true;
      setHasLoadedOnce(true);
      setLoadError(false);
    } catch (err) {
      console.error("Failed to load attendance data:", err);
      setSyncState("error");
      setLoadError(true);
      // Deliberately do NOT seed fallback data here, and do NOT enable
      // saving — we don't actually know what's in the database right now,
      // so guessing would risk overwriting it once a save eventually
      // fires. The polling loop keeps retrying in the background; once a
      // fetch succeeds, the real data loads in and saving turns on normally.
    }
  };

  // initial load — if older pre-update data still exists, the server
  // automatically restores it in place of any placeholder data before
  // responding, so this single fetch is all that's needed.
  React.useEffect(() => {
    loadFromServer();
  }, []);


  // lightweight polling so faculty see each other's edits without a manual
  // refresh — paused while this tab has unsaved local edits in flight
  React.useEffect(() => {
    const interval = setInterval(() => {
      if (document.visibilityState === "visible" && syncState !== "saving") {
        loadFromServer();
      }
    }, POLL_INTERVAL_MS);
    return () => clearInterval(interval);
  }, [syncState]);

  // debounced save to the server whenever data changes
  React.useEffect(() => {
    if (!loadedRef.current) return;
    localVersionRef.current += 1;
    setSyncState("saving");
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      // snapshot which deletions we're about to flush, so we only clear
      // the ones actually sent (not any added mid-flight by the user)
      const deletedAttendanceKeys = Array.from(pendingDeletesRef.current.attendance);
      const deletedSessionKeys = Array.from(pendingDeletesRef.current.sessions);
      const deletedSubjectIds = Array.from(pendingDeletesRef.current.subjects);
      const deletedStudentIds = Array.from(pendingDeletesRef.current.students);

      // Diff against the last known-saved snapshot — only send what's
      // actually different. This is the fix for deletions (or anything
      // else) getting silently undone: previously every save re-uploaded
      // the ENTIRE local attendance/sessions/students/subjects state, so a
      // tab with a slightly stale snapshot would re-create a record
      // someone else had just deleted, simply by saving something
      // unrelated. Diffing means an unchanged record is never re-sent, so
      // there's nothing for a stale tab to accidentally resurrect.
      const baseline = lastSyncedRef.current;
      const diffAttendance = {};
      Object.entries(attendance).forEach(([k, v]) => {
        if (baseline.attendance[k] !== v) diffAttendance[k] = v;
      });
      const diffSessions = {};
      Object.entries(sessions).forEach(([k, v]) => {
        if (JSON.stringify(baseline.sessions[k]) !== JSON.stringify(v)) diffSessions[k] = v;
      });
      const diffMeta = {};
      Object.entries(attendanceMeta).forEach(([k, v]) => {
        if (JSON.stringify(baseline.attendanceMeta[k]) !== JSON.stringify(v)) diffMeta[k] = v;
      });
      const diffSubjects = subjects.filter((s) => {
        const prev = baseline.subjects.find((p) => p.id === s.id);
        return !prev || JSON.stringify(prev) !== JSON.stringify(s);
      });
      const diffStudents = students.filter((s) => {
        const prev = baseline.students.find((p) => p.id === s.id);
        return !prev || JSON.stringify(prev) !== JSON.stringify(s);
      });

      const nothingToSend =
        Object.keys(diffAttendance).length === 0 &&
        Object.keys(diffSessions).length === 0 &&
        Object.keys(diffMeta).length === 0 &&
        diffSubjects.length === 0 &&
        diffStudents.length === 0 &&
        deletedAttendanceKeys.length === 0 &&
        deletedSessionKeys.length === 0 &&
        deletedSubjectIds.length === 0 &&
        deletedStudentIds.length === 0;
      if (nothingToSend) {
        setSyncState("saved");
        return;
      }

      try {
        const res = await fetch("/api/state", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            subjects: diffSubjects,
            students: diffStudents,
            attendance: diffAttendance,
            sessions: diffSessions,
            attendanceMeta: diffMeta,
            deletedAttendanceKeys,
            deletedSessionKeys,
            deletedSubjectIds,
            deletedStudentIds,
          }),
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        deletedAttendanceKeys.forEach((k) => pendingDeletesRef.current.attendance.delete(k));
        deletedSessionKeys.forEach((k) => pendingDeletesRef.current.sessions.delete(k));
        deletedSubjectIds.forEach((k) => pendingDeletesRef.current.subjects.delete(k));
        deletedStudentIds.forEach((k) => pendingDeletesRef.current.students.delete(k));
        // this save succeeded — the full current state (including whatever
        // was just sent) is now the new known-saved baseline
        lastSyncedRef.current = {
          attendance: { ...attendance },
          sessions: { ...sessions },
          attendanceMeta: { ...attendanceMeta },
          subjects: [...subjects],
          students: [...students],
        };
        setSyncState("saved");
      } catch (err) {
        console.error("Failed to save attendance data:", err);
        setSyncState("error");
      }
    }, SAVE_DEBOUNCE_MS);
    return () => clearTimeout(saveTimer.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [subjects, students, attendance, sessions, attendanceMeta]);

  const resetToSampleData = async () => {
    if (!window.confirm("This clears everything for EVERYONE using this app and restores the original sample data. Continue?")) return;
    try {
      await fetch("/api/state", { method: "DELETE" });
    } catch (err) {
      console.error("Failed to clear server data:", err);
    }
    pendingDeletesRef.current = { attendance: new Set(), sessions: new Set(), subjects: new Set(), students: new Set() };
    // the database was just wiped, so nothing is "synced" anymore — reset
    // the diff baseline to empty, otherwise the next save would only send
    // whatever differs from the OLD (now-deleted) data instead of everything
    lastSyncedRef.current = { attendance: {}, sessions: {}, attendanceMeta: {}, subjects: [], students: [] };
    setSubjects(seedSubjects);
    setStudents(seedStudents);
    setAttendance(seedAttendance());
    setSessions(seedSessions());
    setAttendanceMeta({});
  };

  // ---------- derived helpers ----------
  const subjectsFor = (student) =>
    subjects.filter((s) => s.type === "core" || student.electives.includes(s.id));

  const studentsFor = (subject) =>
    subject.type === "core" ? students : students.filter((st) => st.electives.includes(subject.id));

  const subjectStats = useMemo(() => {
    return subjects.map((sub) => {
      const roster = studentsFor(sub);
      let present = 0,
        total = 0;
      roster.forEach((st) => {
        Object.keys(attendance).forEach((key) => {
          const [sid, subId] = key.split("__");
          if (sid === st.id && subId === sub.id) {
            total++;
            if (attendance[key] === "present") present++;
          }
        });
      });
      const pct = total ? Math.round((present / total) * 100) : 0;
      return { ...sub, roster: roster.length, pct, total, present };
    });
  }, [subjects, students, attendance]);

  const studentStats = useMemo(() => {
    return students.map((st) => {
      const subs = subjectsFor(st);
      const bySubject = subs.map((sub) => {
        let present = 0,
          total = 0;
        Object.keys(attendance).forEach((key) => {
          const [sid, subId] = key.split("__");
          if (sid === st.id && subId === sub.id) {
            total++;
            if (attendance[key] === "present") present++;
          }
        });
        return { subject: sub, present, total, pct: total ? Math.round((present / total) * 100) : 0 };
      });
      const totalPresent = bySubject.reduce((a, b) => a + b.present, 0);
      const totalAll = bySubject.reduce((a, b) => a + b.total, 0);
      return { ...st, bySubject, overallPct: totalAll ? Math.round((totalPresent / totalAll) * 100) : 0 };
    });
  }, [students, subjects, attendance]);

  // IMPORTANT: these must come after every hook above (useState/useMemo/etc.)
  // so the same hooks run in the same order on every render — an early
  // return placed before a hook call causes "Rendered more hooks than
  // during the previous render" once state flips (loading, auth, etc).
  if (!currentUser) {
    return <LoginScreen onLogin={login} />;
  }

  if (!hasLoadedOnce) {
    return (
      <div
        style={{
          minHeight: "100vh",
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          justifyContent: "center",
          gap: 14,
          background: COLORS.parchment,
          fontFamily: "Inter, sans-serif",
          color: COLORS.ink,
          padding: 24,
          textAlign: "center",
        }}
      >
        <style>{FONT_IMPORT + RESPONSIVE_CSS + ANIMATION_CSS}</style>
        {loadError ? (
          <>
            <div style={{ width: 34, height: 34, borderRadius: "50%", background: COLORS.absentSoft, display: "flex", alignItems: "center", justifyContent: "center" }}>
              <span style={{ color: COLORS.absent, fontWeight: 700 }}>!</span>
            </div>
            <div style={{ fontSize: 14, fontWeight: 600 }}>Couldn't reach the shared database</div>
            <div style={{ fontSize: 12, color: COLORS.slate, maxWidth: 320 }}>
              To make sure nothing gets overwritten, the app won't show or save any data until it can confirm what's actually stored. Retrying
              automatically in the background — or try now:
            </div>
            <button onClick={loadFromServer} style={btnPrimary}>
              Retry now
            </button>
          </>
        ) : (
          <>
            <div
              className="rp-spin"
              style={{ width: 34, height: 34, borderRadius: "50%", border: `3px solid ${COLORS.line}`, borderTopColor: COLORS.brass }}
            />
            <div style={{ fontSize: 13, color: COLORS.slate }}>Loading attendance data…</div>
          </>
        )}
      </div>
    );
  }

  // ---------- actions ----------
  const addStudent = (data) => {
    setStudents((prev) => [...prev, { ...data, id: data.id || `E24${String(idCounter.current++).padStart(3, "0")}` }]);
  };
  const editStudent = (oldId, data) => {
    const newId = data.id || oldId;
    setStudents((prev) => prev.map((s) => (s.id === oldId ? { ...s, ...data, id: newId } : s)));
    if (newId !== oldId) {
      // migrate any existing attendance records (and their audit trail) so
      // they still point at this student — the old keys must be explicitly
      // marked deleted, or a server-side merge would resurrect them
      const remapKey = (key) => {
        const [sid, ...rest] = key.split("__");
        const finalSid = sid === oldId ? newId : sid;
        return [finalSid, ...rest].join("__");
      };
      setAttendance((prev) => {
        const next = {};
        Object.entries(prev).forEach(([key, val]) => {
          const newKey = remapKey(key);
          if (newKey !== key) pendingDeletesRef.current.attendance.add(key);
          next[newKey] = val;
        });
        return next;
      });
      setAttendanceMeta((prev) => {
        const next = {};
        Object.entries(prev).forEach(([key, val]) => {
          next[remapKey(key)] = val;
        });
        return next;
      });
    }
  };
  const removeStudent = (id) => {
    setStudents((prev) => prev.filter((s) => s.id !== id));
    pendingDeletesRef.current.students.add(id);
  };
  // Bulk import from an uploaded Excel/CSV roster: upserts by enrollment
  // number — existing students get their details refreshed, new ones get
  // added. Doesn't touch students already present but absent from the file.
  const importStudents = (parsedRows) => {
    const byId = new Map(students.map((s) => [s.id, s]));
    let added = 0,
      updated = 0,
      skipped = 0;
    parsedRows.forEach((r) => {
      if (!r.id || !r.name) {
        skipped++;
        return;
      }
      const existing = byId.get(r.id);
      if (existing) {
        byId.set(r.id, {
          ...existing,
          name: r.name || existing.name,
          email: r.email || existing.email,
          program: r.program || existing.program,
          electives: r.electives && r.electives.length ? r.electives : existing.electives,
        });
        updated++;
      } else {
        byId.set(r.id, {
          id: r.id,
          name: r.name,
          email: r.email || "",
          program: r.program || "MSc",
          electives: r.electives || [],
        });
        added++;
      }
    });
    setStudents(Array.from(byId.values()));
    return { added, updated, skipped };
  };
  // Bulk-merges a parsed multi-sheet attendance workbook (see
  // parseAttendanceWorkbook) into the shared attendance/sessions state,
  // same as if every record had been marked by hand — including the audit
  // trail, so it's clear this batch came from an import rather than being
  // silently indistinguishable from manual entry.
  const importAttendanceWorkbook = (parsed) => {
    const idSet = new Set(students.map((s) => s.id));
    const metaStamp = { editedBy: currentUser?.username ? `${currentUser.username} (import)` : "import", editedAt: new Date().toISOString() };
    const newAttendance = {};
    const newMeta = {};
    const unmatchedStudents = new Set();
    let matched = 0;
    parsed.records.forEach((r) => {
      if (!idSet.has(r.enroll)) {
        unmatchedStudents.add(r.enroll);
        return;
      }
      const key = `${r.enroll}__${r.subjectId}__${r.date}__${r.slot}`;
      newAttendance[key] = r.status;
      newMeta[key] = metaStamp;
      matched++;
    });
    setAttendance((prev) => ({ ...prev, ...newAttendance }));
    setAttendanceMeta((prev) => ({ ...prev, ...newMeta }));

    const newSessions = {};
    parsed.sessions.forEach((s) => {
      newSessions[`${s.subjectId}__${s.date}__${s.slot}`] = { time: s.time, remark: "" };
    });
    setSessions((prev) => {
      const merged = { ...prev };
      Object.entries(newSessions).forEach(([key, val]) => {
        merged[key] = { ...(merged[key] || { time: "", remark: "" }), ...val, remark: merged[key]?.remark || val.remark };
      });
      return merged;
    });

    return {
      matched,
      unmatchedStudents: Array.from(unmatchedStudents),
      matchedSheets: parsed.matchedSheets,
      unmatchedSheets: parsed.unmatchedSheets,
    };
  };
  const handleAttendanceFileImport = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setImportingAttendance(true);
    try {
      const parsed = await parseAttendanceWorkbook(file, subjects);
      if (parsed.matchedSheets.length === 0) {
        window.alert(
          `Couldn't match any sheet names to your subjects.\n\nSheets found: ${parsed.unmatchedSheets.join(", ")}\n\nMake sure sheet names match a subject's code (e.g. "RM", "OS&MF") or full name.`
        );
        return;
      }
      const result = importAttendanceWorkbook(parsed);
      let msg = `Imported ${result.matched} attendance records from: ${result.matchedSheets.join(", ")}.`;
      if (result.unmatchedSheets.length) msg += `\n\nSkipped sheets (couldn't match to a subject): ${result.unmatchedSheets.join(", ")}.`;
      if (result.unmatchedStudents.length) msg += `\n\n${result.unmatchedStudents.length} enrollment number(s) in the file don't match any current student and were skipped.`;
      window.alert(msg);
    } catch (err) {
      console.error("Attendance import failed:", err);
      window.alert("Couldn't read that file. Make sure it's a valid .xlsx workbook.");
    } finally {
      setImportingAttendance(false);
    }
  };
  const addSubject = (data) => {
    setSubjects((prev) => [...prev, { ...data, id: `S${idCounter.current++}` }]);
  };
  const editSubject = (id, data) => {
    setSubjects((prev) => prev.map((s) => (s.id === id ? { ...s, ...data, id } : s)));
  };
  const removeSubject = (id) => {
    setSubjects((prev) => prev.filter((s) => s.id !== id));
    pendingDeletesRef.current.subjects.add(id);
  };
  // Switch a student from one elective to another WITHIN THE SAME GROUP,
  // carrying their existing attendance history over to the new subject
  // (rather than losing it) — e.g. moving from OS&MF to SCS.
  const transferElective = (studentId, fromSubjectId, toSubjectId) => {
    setStudents((prev) =>
      prev.map((s) =>
        s.id === studentId ? { ...s, electives: s.electives.map((eid) => (eid === fromSubjectId ? toSubjectId : eid)) } : s
      )
    );
    const remapKey = (key) => {
      const [sid, subId, ...rest] = key.split("__");
      return [studentId, toSubjectId, ...rest].join("__");
    };
    setAttendance((prev) => {
      const next = { ...prev };
      Object.keys(prev).forEach((key) => {
        const [sid, subId] = key.split("__");
        if (sid === studentId && subId === fromSubjectId) {
          next[remapKey(key)] = next[key];
          delete next[key];
          pendingDeletesRef.current.attendance.add(key);
        }
      });
      return next;
    });
    setAttendanceMeta((prev) => {
      const next = { ...prev };
      Object.keys(prev).forEach((key) => {
        const [sid, subId] = key.split("__");
        if (sid === studentId && subId === fromSubjectId) {
          next[remapKey(key)] = next[key];
          delete next[key];
        }
      });
      return next;
    });
  };
  // Per-student status (present/absent) for one lecture. `slot` distinguishes
  // multiple sessions of the same subject on the same date (e.g. a morning
  // and an afternoon lecture) — defaults to "1" for the common single-session case.
  const setMark = (studentId, subjectId, date, status, slot = "1") => {
    const key = `${studentId}__${subjectId}__${date}__${slot}`;
    setAttendance((prev) => ({ ...prev, [key]: status }));
    setAttendanceMeta((prev) => ({
      ...prev,
      [key]: { editedBy: currentUser?.username || "unknown", editedAt: new Date().toISOString() },
    }));
    pendingDeletesRef.current.attendance.delete(key); // re-marking un-deletes it, if it was pending
  };
  const deleteMark = (studentId, subjectId, date, slot = "1") => {
    const key = `${studentId}__${subjectId}__${date}__${slot}`;
    setAttendance((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
    setAttendanceMeta((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
    pendingDeletesRef.current.attendance.add(key);
  };
  // Wipes attendance for EVERY student for one specific session (subject +
  // date + slot) in one go — e.g. the class was cancelled and shouldn't
  // have been marked.
  const deleteSession = (subjectId, date, slot = "1") => {
    const suffix = `__${subjectId}__${date}__${slot}`;
    setAttendance((prev) => {
      const next = {};
      Object.entries(prev).forEach(([key, val]) => {
        if (key.endsWith(suffix)) {
          pendingDeletesRef.current.attendance.add(key);
        } else {
          next[key] = val;
        }
      });
      return next;
    });
    setAttendanceMeta((prev) => {
      const next = {};
      Object.entries(prev).forEach(([key, val]) => {
        if (!key.endsWith(suffix)) next[key] = val;
      });
      return next;
    });
    const sessionKey = `${subjectId}__${date}__${slot}`;
    setSessions((prev) => {
      const next = { ...prev };
      delete next[sessionKey];
      return next;
    });
    pendingDeletesRef.current.sessions.add(sessionKey);
  };
  // Session-level time + remark for one lecture (subject + date + slot) —
  // shared by every student in that lecture, e.g. "faculty on leave, proxy conducted".
  const getSession = (subjectId, date, slot = "1") => sessions[`${subjectId}__${date}__${slot}`] || { time: "", remark: "" };
  const setSession = (subjectId, date, patch, slot = "1") => {
    const key = `${subjectId}__${date}__${slot}`;
    setSessions((prev) => ({
      ...prev,
      [key]: { time: "", remark: "", ...prev[key], ...patch },
    }));
  };
  // All sessions (slots) recorded for a subject on a given date, sorted by
  // slot number — e.g. [{slot:"1", time:"09:15 - 10:00"}, {slot:"2", time:"02:00 - 02:45"}]
  const slotsForDate = (subjectId, date) => {
    const prefix = `${subjectId}__${date}__`;
    return Object.keys(sessions)
      .filter((k) => k.startsWith(prefix))
      .map((k) => k.slice(prefix.length))
      .sort((a, b) => Number(a) - Number(b))
      .map((slot) => ({ slot, ...sessions[`${prefix}${slot}`] }));
  };

  // Excel fill/font colors for the three attendance-percentage tiers
  const TIER_STYLE = {
    green: { fill: "FFC6EFCE", font: "FF006100" }, // >= 75
    orange: { fill: "FFFFE8B0", font: "FF9C5700" }, // 60 - 74
    red: { fill: "FFFFC7CE", font: "FF9C0006" }, // < 60
  };
  const tierFor = (pct) => (pct >= 75 ? "green" : pct >= 60 ? "orange" : "red");

  const colorPctCell = (cell, pct) => {
    const tier = TIER_STYLE[tierFor(pct)];
    cell.value = `${pct}%`;
    cell.alignment = { horizontal: "center" };
    cell.font = { color: { argb: tier.font }, bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: tier.fill } };
  };

  const styleHeader = (row) => {
    row.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFF7F4EA" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1B2A4A" } };
      cell.alignment = { vertical: "middle" };
    });
    row.height = 20;
  };

  const styleHeaderCell = (cell, align = "center") => {
    cell.font = { bold: true, color: { argb: "FFF7F4EA" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1B2A4A" } };
    cell.alignment = { vertical: "middle", horizontal: align, wrapText: true };
  };

  // One detailed sheet per subject: one column per SESSION (a date can have
  // more than one — e.g. a morning and an afternoon lecture — each becomes
  // its own column, labeled with the date and that session's time).
  const buildSubjectSheet = (wb, subject, usedNames) => {
    const roster = studentsFor(subject);
    const sessionSet = new Set(); // "date__slot" keys
    Object.keys(attendance).forEach((key) => {
      const [sid, subId, date, slot] = key.split("__");
      if (subId === subject.id && roster.some((r) => r.id === sid)) sessionSet.add(`${date}__${slot}`);
    });
    const cols = Array.from(sessionSet)
      .map((k) => {
        const [date, slot] = k.split("__");
        return { date, slot };
      })
      .sort((a, b) => (a.date === b.date ? Number(a.slot) - Number(b.slot) : a.date < b.date ? -1 : 1));

    let baseName = subject.code.replace(/[\\/*?:[\]]/g, "-").slice(0, 28) || subject.id;
    let sheetName = baseName;
    let n = 2;
    while (usedNames.has(sheetName)) sheetName = `${baseName}-${n++}`.slice(0, 31);
    usedNames.add(sheetName);

    const ws = wb.addWorksheet(sheetName);
    const LEFT = ["Sr No", "Enrollment No", "Name", "Program"];
    const RIGHT = ["Classes Held", "Present", "Attendance %"];
    const totalCols = LEFT.length + cols.length + RIGHT.length;

    ws.getColumn(1).width = 7;
    ws.getColumn(2).width = 16;
    ws.getColumn(3).width = 24;
    ws.getColumn(4).width = 10;
    cols.forEach((_, i) => {
      ws.getColumn(LEFT.length + i + 1).width = 10;
    });
    RIGHT.forEach((_, i) => {
      ws.getColumn(LEFT.length + cols.length + i + 1).width = 12;
    });

    // Row 1 — title
    ws.mergeCells(1, 1, 1, Math.max(totalCols, 5));
    const title = ws.getCell(1, 1);
    title.value = `${subject.code} · ${subject.name} — Attendance Register`;
    title.font = { bold: true, size: 13, color: { argb: "FFF7F4EA" } };
    title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1B2A4A" } };
    title.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(1).height = 22;

    // Rows 2-3 — headers (fixed + summary columns span both rows; each
    // session is one column showing the date and, underneath, its time —
    // a date with two sessions gets two adjacent columns)
    LEFT.forEach((label, i) => {
      ws.mergeCells(2, i + 1, 3, i + 1);
      const cell = ws.getCell(2, i + 1);
      cell.value = label;
      styleHeaderCell(cell);
    });
    cols.forEach(({ date, slot }, i) => {
      const col = LEFT.length + i + 1;
      const session = sessions[`${subject.id}__${date}__${slot}`] || { time: "", remark: "" };
      ws.mergeCells(2, col, 3, col);
      const dateCell = ws.getCell(2, col);
      const dObj = new Date(date);
      const dateLabel = dObj.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
      dateCell.value = session.time ? `${dateLabel}\n${session.time}` : dateLabel;
      styleHeaderCell(dateCell);
    });
    RIGHT.forEach((label, i) => {
      const col = LEFT.length + cols.length + i + 1;
      ws.mergeCells(2, col, 3, col);
      styleHeaderCell(Object.assign(ws.getCell(2, col), { value: label }));
    });
    ws.getRow(2).height = 28;
    ws.getRow(3).height = 4;

    // Data rows
    roster.forEach((st, idx) => {
      const r = 4 + idx;
      ws.getCell(r, 1).value = idx + 1;
      ws.getCell(r, 2).value = st.id;
      ws.getCell(r, 3).value = st.name;
      ws.getCell(r, 4).value = st.program;
      [1, 2, 3, 4].forEach((c) => {
        ws.getCell(r, c).alignment = { vertical: "middle" };
        if (idx % 2 === 1) ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF7F4EA" } };
      });

      let present = 0,
        total = 0;
      cols.forEach(({ date, slot }, i) => {
        const col = LEFT.length + i + 1;
        const status = attendance[`${st.id}__${subject.id}__${date}__${slot}`];
        const statusCell = ws.getCell(r, col);
        if (status) {
          total++;
          if (status === "present") present++;
          const meta = STATUS_META[status];
          statusCell.value = meta.abbr;
          statusCell.alignment = { horizontal: "center" };
          statusCell.font = { bold: true, color: { argb: meta.font } };
          statusCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: meta.argb } };
        } else {
          statusCell.value = "";
          statusCell.alignment = { horizontal: "center" };
        }
      });

      const pct = total ? Math.round((present / total) * 100) : 0;
      const rightBase = LEFT.length + cols.length;
      ws.getCell(r, rightBase + 1).value = total;
      ws.getCell(r, rightBase + 1).alignment = { horizontal: "center" };
      ws.getCell(r, rightBase + 2).value = present;
      ws.getCell(r, rightBase + 2).alignment = { horizontal: "center" };
      colorPctCell(ws.getCell(r, rightBase + 3), pct);
    });

    // Legend + any session-level remarks, as footnotes rather than columns
    let footRow = 4 + roster.length + 1;
    ws.getCell(footRow, 1).value = "Legend:";
    ws.getCell(footRow, 1).font = { bold: true };
    STATUS_ORDER.forEach((s, i) => {
      const meta = STATUS_META[s];
      const cell = ws.getCell(footRow, 2 + i);
      cell.value = `${meta.abbr} = ${meta.label}`;
      cell.font = { bold: true, color: { argb: meta.font } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: meta.argb } };
    });
    ws.getCell(footRow, 2 + STATUS_ORDER.length).value = "Blank = not recorded";
    ws.getCell(footRow, 2 + STATUS_ORDER.length).font = { italic: true, size: 10, color: { argb: "FF6B7280" } };

    const remarkCols = cols.filter(({ date, slot }) => (sessions[`${subject.id}__${date}__${slot}`]?.remark || "").trim());
    if (remarkCols.length > 0) {
      footRow += 1;
      ws.getCell(footRow, 1).value = "Session notes:";
      ws.getCell(footRow, 1).font = { bold: true };
      remarkCols.forEach(({ date, slot }) => {
        footRow += 1;
        const dObj = new Date(date);
        const label = dObj.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
        const session = sessions[`${subject.id}__${date}__${slot}`];
        const cell = ws.getCell(footRow, 1);
        ws.mergeCells(footRow, 1, footRow, Math.min(6, totalCols));
        cell.value = `${label}${session.time ? ` (${session.time})` : ""}: ${session.remark}`;
        cell.font = { italic: true, size: 10, color: { argb: "FF6B7280" } };
      });
    }

    ws.views = [{ state: "frozen", xSplit: 4, ySplit: 3 }];
  };

  const exportExcel = async (scopeSubjectId = "all") => {
    const wb = new ExcelJS.Workbook();
    wb.creator = "Attendance Register";
    wb.created = new Date();
    const single = scopeSubjectId !== "all";
    const scopedSubjects = single ? subjects.filter((s) => s.id === scopeSubjectId) : subjects;
    const scopedStudents = single ? studentsFor(scopedSubjects[0]) : students;

    // ---- Student Detail ----
    const wsStudents = wb.addWorksheet("Student Detail");
    wsStudents.columns = [
      { header: "Sr No", key: "sr", width: 7 },
      { header: "Enrollment No", key: "id", width: 16 },
      { header: "Name", key: "name", width: 24 },
      { header: "Email", key: "email", width: 28 },
      { header: "Program", key: "program", width: 10 },
      { header: "Electives", key: "electives", width: 40 },
    ];
    scopedStudents.forEach((s, i) =>
      wsStudents.addRow({
        sr: i + 1,
        id: s.id,
        name: s.name,
        email: s.email,
        program: s.program,
        electives: subjects.filter((sub) => s.electives.includes(sub.id)).map((s2) => s2.name).join(", "),
      })
    );
    styleHeader(wsStudents.getRow(1));

    // ---- One detailed date-grid sheet per subject (or just the chosen one) ----
    const usedNames = new Set(["Student Detail"]);
    scopedSubjects.forEach((sub) => buildSubjectSheet(wb, sub, usedNames));

    if (single) {
      // single-subject export stops here — just the roster + that subject's register
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/octet-stream" });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${scopedSubjects[0].code.replace(/[\\/*?:[\]]/g, "-")}_attendance_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
      return;
    }

    // ---- Overall attendance per student (color-coded %) ----
    const wsOverall = wb.addWorksheet("Overall Attendance");
    wsOverall.columns = [
      { header: "Enrollment No", key: "id", width: 16 },
      { header: "Name", key: "name", width: 24 },
      { header: "Email", key: "email", width: 28 },
      { header: "Program", key: "program", width: 10 },
      { header: "Subjects Taken", key: "subjectCount", width: 14 },
      { header: "Overall Attendance %", key: "pct", width: 18 },
    ];
    studentStats.forEach((st) => {
      const row = wsOverall.addRow({
        id: st.id,
        name: st.name,
        email: st.email,
        program: st.program,
        subjectCount: st.bySubject.length,
        pct: null,
      });
      colorPctCell(row.getCell("pct"), st.overallPct);
    });
    styleHeader(wsOverall.getRow(1));
    wsOverall.addRow([]);
    const legendStart = wsOverall.rowCount + 1;
    wsOverall.getCell(`A${legendStart}`).value = "Legend:";
    wsOverall.getCell(`A${legendStart}`).font = { bold: true };
    [
      ["green", "75% and above"],
      ["orange", "60% – 74%"],
      ["red", "Below 60%"],
    ].forEach(([tier, label], i) => {
      const r = legendStart + 1 + i;
      const cell = wsOverall.getCell(`A${r}`);
      cell.value = label;
      cell.font = { color: { argb: TIER_STYLE[tier].font }, bold: true };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: TIER_STYLE[tier].fill } };
    });

    // ---- Subject-wise average (color-coded %) ----
    const wsSubjAvg = wb.addWorksheet("Subjects");
    wsSubjAvg.columns = [
      { header: "Code", key: "code", width: 14 },
      { header: "Subject", key: "subject", width: 26 },
      { header: "Type", key: "type", width: 10 },
      { header: "Program", key: "program", width: 10 },
      { header: "Roster Size", key: "roster", width: 12 },
      { header: "Average Attendance %", key: "pct", width: 18 },
    ];
    subjectStats.forEach((s) => {
      const row = wsSubjAvg.addRow({
        code: s.code,
        subject: s.name,
        type: s.type,
        program: s.program,
        roster: s.roster,
        pct: null,
      });
      colorPctCell(row.getCell("pct"), s.pct);
    });
    styleHeader(wsSubjAvg.getRow(1));

    // ---- Cross-subject summary (per student, per subject, one row each) ----
    const wsSummary = wb.addWorksheet("Attendance Summary");
    wsSummary.columns = [
      { header: "Enrollment No", key: "id", width: 16 },
      { header: "Name", key: "name", width: 24 },
      { header: "Program", key: "program", width: 10 },
      { header: "Subject Code", key: "code", width: 14 },
      { header: "Subject", key: "subject", width: 26 },
      { header: "Present", key: "present", width: 10 },
      { header: "Total Classes", key: "total", width: 14 },
      { header: "Attendance %", key: "pct", width: 14 },
    ];
    studentStats.forEach((st) => {
      st.bySubject.forEach((b) => {
        const row = wsSummary.addRow({
          id: st.id,
          name: st.name,
          program: st.program,
          code: b.subject.code,
          subject: b.subject.name,
          present: b.present,
          total: b.total,
          pct: null,
        });
        colorPctCell(row.getCell("pct"), b.pct);
      });
    });
    styleHeader(wsSummary.getRow(1));

    // ---- Raw daily log (every record, with time + remark) ----
    const wsLog = wb.addWorksheet("Daily Log");
    wsLog.columns = [
      { header: "Date", key: "date", width: 14 },
      { header: "Enrollment No", key: "id", width: 16 },
      { header: "Name", key: "name", width: 24 },
      { header: "Subject", key: "subject", width: 26 },
      { header: "Time", key: "time", width: 16 },
      { header: "Status", key: "status", width: 12 },
      { header: "Remark", key: "remark", width: 30 },
    ];
    Object.entries(attendance)
      .map(([key, status]) => {
        const [sid, subId, date, slot] = key.split("__");
        const st = students.find((s) => s.id === sid);
        const sub = subjects.find((s) => s.id === subId);
        const session = sessions[`${subId}__${date}__${slot}`] || { time: "", remark: "" };
        return {
          date,
          id: sid,
          name: st ? st.name : sid,
          subject: sub ? sub.name : subId,
          time: session.time,
          status: STATUS_META[status]?.label || status,
          remark: session.remark,
          _statusKey: status,
        };
      })
      .sort((a, b) => (a.date < b.date ? 1 : -1))
      .forEach((r) => {
        const row = wsLog.addRow(r);
        const meta = STATUS_META[r._statusKey];
        const cell = row.getCell("status");
        if (meta) {
          cell.font = { color: { argb: meta.font }, bold: true };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: meta.argb } };
        }
      });
    styleHeader(wsLog.getRow(1));

    // ---- trigger download ----
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/octet-stream" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `attendance_register_${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  };

  return (
    <div style={{ fontFamily: "Inter, sans-serif", background: COLORS.parchment, minHeight: "100vh", color: COLORS.ink }}>
      <style>{FONT_IMPORT + RESPONSIVE_CSS + ANIMATION_CSS}</style>

      {/* Mobile-only top bar with hamburger toggle */}
      <div className="rp-topbar">
        <button
          onClick={() => setMobileMenuOpen((v) => !v)}
          style={{ background: "transparent", border: "none", color: COLORS.parchment, cursor: "pointer", padding: 4, display: "flex" }}
          aria-label="Toggle menu"
        >
          <Menu size={22} />
        </button>
        <div style={{ display: "flex", alignItems: "center", gap: 8, fontFamily: "Fraunces, serif", fontWeight: 700, fontSize: 16 }}>
          <Stamp size={18} color={COLORS.brass} />
          Attendance Register
        </div>
        <div style={{ width: 22 }} />
      </div>

      {/* Backdrop, closes the drawer on tap outside it (mobile only) */}
      {mobileMenuOpen && (
        <div
          className="rp-backdrop"
          onClick={() => setMobileMenuOpen(false)}
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 150 }}
        />
      )}

      <div style={{ minHeight: "100vh", display: "flex" }}>
        {/* Sidebar */}
        <aside
          className={`rp-sidebar${mobileMenuOpen ? " rp-open" : ""}`}
          style={{ width: 220, background: COLORS.ink, color: COLORS.parchment, flexShrink: 0, display: "flex", flexDirection: "column" }}
        >
          <div style={{ padding: "24px 20px", borderBottom: `1px solid ${COLORS.inkSoft}` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <Stamp size={22} color={COLORS.brass} />
              <span style={{ fontFamily: "Fraunces, serif", fontWeight: 700, fontSize: 18 }}>Attendance Register</span>
            </div>
            <div style={{ fontSize: 11, color: COLORS.brassSoft, marginTop: 6, letterSpacing: 1, textTransform: "uppercase" }}>
              MSc &amp; PGD · Batch 2024
            </div>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 14 }}>
              <div style={{ fontSize: 12 }}>
                <div style={{ fontWeight: 600 }}>{currentUser.username}</div>
                <div style={{ fontSize: 10, color: COLORS.brassSoft, textTransform: "uppercase", letterSpacing: 0.5 }}>
                  {currentUser.role === "admin" ? "Administrator" : "Faculty"}
                </div>
              </div>
              <button
                onClick={logout}
                style={{ background: "transparent", border: `1px solid ${COLORS.inkSoft}`, color: COLORS.brassSoft, fontSize: 11, padding: "4px 10px", borderRadius: 6, cursor: "pointer" }}
              >
                Log out
              </button>
            </div>
          </div>
          <nav style={{ padding: "12px 10px", flex: 1, overflowY: "auto" }}>
            {TABS.filter((t) => !t.adminOnly || currentUser.role === "admin").map((t) => {
              const Icon = t.icon;
              const active = tab === t.id;
              return (
                <button
                  key={t.id}
                  className="rp-nav-btn"
                  data-active={active}
                  onClick={() => {
                    setTab(t.id);
                    setMobileMenuOpen(false);
                  }}
                  style={{
                    width: "100%",
                    display: "flex",
                    alignItems: "center",
                    gap: 10,
                    padding: "10px 12px",
                    marginBottom: 4,
                    borderRadius: 8,
                    background: active ? COLORS.brass : "transparent",
                    color: active ? COLORS.ink : COLORS.parchment,
                    fontWeight: active ? 600 : 500,
                    fontSize: 14,
                    border: "none",
                    cursor: "pointer",
                    textAlign: "left",
                  }}
                >
                  <Icon size={16} />
                  {t.label}
                </button>
              );
            })}
          </nav>
          <div style={{ padding: 16, borderTop: `1px solid ${COLORS.inkSoft}` }}>
            <select
              value={downloadSubjectId}
              onChange={(e) => setDownloadSubjectId(e.target.value)}
              style={{
                width: "100%",
                marginBottom: 8,
                padding: "8px 10px",
                borderRadius: 8,
                border: `1px solid ${COLORS.inkSoft}`,
                background: COLORS.inkSoft,
                color: COLORS.parchment,
                fontSize: 12,
              }}
            >
              <option value="all">All subjects (full workbook)</option>
              {subjects.map((s) => (
                <option key={s.id} value={s.id}>
                  {s.code} only
                </option>
              ))}
            </select>
            <button
              onClick={() => exportExcel(downloadSubjectId)}
              style={{
                width: "100%",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                gap: 8,
                padding: "10px 12px",
                borderRadius: 8,
                background: "transparent",
                border: `1px solid ${COLORS.brass}`,
                color: COLORS.brass,
                fontWeight: 600,
                fontSize: 13,
                cursor: "pointer",
              }}
            >
              <Download size={15} />
              Download Excel
            </button>
            <div
              style={{
                fontSize: 10,
                textAlign: "center",
                marginTop: 10,
                opacity: 0.9,
                color: syncState === "error" ? "#E8A19A" : COLORS.brassSoft,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                gap: 5,
              }}
            >
              <span
                className={syncState === "saving" ? "rp-sync-dot-saving" : ""}
                style={{
                  width: 6,
                  height: 6,
                  borderRadius: "50%",
                  background: syncState === "error" ? "#E8A19A" : syncState === "saving" ? COLORS.brass : "#8FBF9F",
                  display: "inline-block",
                }}
              />
              {syncState === "saving" && "Saving to shared database…"}
              {syncState === "saved" && "Synced — shared with all faculty"}
              {syncState === "error" && "Couldn't reach database (check setup)"}
              {syncState === "idle" && "—"}
            </div>
            <input
              ref={attendanceFileInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleAttendanceFileImport}
              style={{ display: "none" }}
            />
            <button
              onClick={() => attendanceFileInputRef.current?.click()}
              disabled={importingAttendance}
              style={{
                width: "100%",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                gap: 8,
                padding: "9px 12px",
                marginTop: 10,
                borderRadius: 8,
                background: "transparent",
                border: `1px solid ${COLORS.inkSoft}`,
                color: COLORS.parchment,
                fontWeight: 500,
                fontSize: 12,
                cursor: importingAttendance ? "default" : "pointer",
                opacity: importingAttendance ? 0.6 : 1,
              }}
              title='Upload a multi-sheet Excel workbook — one sheet per subject, named by subject code (e.g. "RM", "OS&MF")'
            >
              <Upload size={14} />
              {importingAttendance ? "Importing…" : "Import Attendance (Excel)"}
            </button>
            <button
              onClick={resetToSampleData}
              style={{
                width: "100%",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                gap: 8,
                padding: "8px 12px",
                marginTop: 8,
                borderRadius: 8,
                background: "transparent",
                border: "none",
                color: COLORS.brassSoft,
                fontWeight: 500,
                fontSize: 11,
                textDecoration: "underline",
                cursor: "pointer",
              }}
            >
              Reset to sample data
            </button>
          </div>
        </aside>

        {/* Main */}
        <main className="rp-main" style={{ flex: 1, padding: "28px 36px", overflow: "auto", minWidth: 0 }}>
          <div key={tab} className="rp-tab-content">
          {tab === "dashboard" && <Dashboard subjectStats={subjectStats} students={students} />}
          {tab === "students" && (
            <StudentsTab
              students={students}
              subjects={subjects}
              attendance={attendance}
              addStudent={addStudent}
              editStudent={editStudent}
              removeStudent={removeStudent}
              transferElective={transferElective}
              importStudents={importStudents}
            />
          )}
          {tab === "subjects" && (
            <SubjectsTab
              subjects={subjects}
              addSubject={addSubject}
              editSubject={editSubject}
              removeSubject={removeSubject}
              studentsFor={studentsFor}
            />
          )}
          {tab === "mark" && (
            <MarkTab
              subjects={subjects}
              studentsFor={studentsFor}
              attendance={attendance}
              setMark={setMark}
              getSession={getSession}
              setSession={setSession}
              slotsForDate={slotsForDate}
            />
          )}
          {tab === "all" && (
            <AllAttendanceTab
              students={students}
              subjects={subjects}
              attendance={attendance}
              sessions={sessions}
              studentsFor={studentsFor}
              setMark={setMark}
              setSession={setSession}
              deleteMark={deleteMark}
              deleteSession={deleteSession}
            />
          )}
          {tab === "reports" && <ReportsTab studentStats={studentStats} />}
          {tab === "audit" && (
            <AuditLogTab
              students={students}
              subjects={subjects}
              attendance={attendance}
              attendanceMeta={attendanceMeta}
              sessions={sessions}
            />
          )}
          {tab === "users" && currentUser.role === "admin" && <UsersTab />}
          </div>
        </main>
      </div>
    </div>
  );
}

// ============================================================
function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    setBusy(true);
    try {
      const res = await fetch("/api/users", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "login", username, password }),
      });
      const data = await res.json();
      if (data.ok) {
        onLogin({ username: data.username, role: data.role });
      } else {
        setError(data.error || "Invalid username or password.");
      }
    } catch (err) {
      setError("Couldn't reach the server. Check your connection and try again.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: COLORS.ink,
        fontFamily: "Inter, sans-serif",
      }}
    >
      <style>{FONT_IMPORT + RESPONSIVE_CSS + ANIMATION_CSS}</style>
      <form
        onSubmit={submit}
        className="rp-login-card"
        style={{
          background: COLORS.parchment,
          borderRadius: 16,
          padding: 36,
          width: 360,
          boxShadow: "0 20px 60px rgba(0,0,0,0.35)",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
          <Stamp size={24} color={COLORS.brass} />
          <div style={{ fontFamily: "Fraunces, serif", fontWeight: 700, fontSize: 20, color: COLORS.ink }}>Attendance Register</div>
        </div>
        <div style={{ fontSize: 12, color: COLORS.slate, marginBottom: 24 }}>MSc &amp; PGD · Batch 2024</div>

        <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6, color: COLORS.ink }}>Username</div>
        <input
          autoFocus
          value={username}
          onChange={(e) => setUsername(e.target.value)}
          style={{ ...inputStyle, marginBottom: 14 }}
          placeholder="Username"
        />
        <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6, color: COLORS.ink }}>Password</div>
        <input
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          style={{ ...inputStyle, marginBottom: 18 }}
          placeholder="Password"
        />

        {error && (
          <div className="rp-banner" style={{ fontSize: 12, color: COLORS.absent, marginBottom: 14, fontWeight: 600 }}>{error}</div>
        )}

        <button type="submit" disabled={busy} style={{ ...btnPrimary, width: "100%", justifyContent: "center", opacity: busy ? 0.6 : 1 }}>
          {busy ? "Signing in…" : "Sign In"}
        </button>
      </form>
    </div>
  );
}

// ============================================================
function SectionTitle({ eyebrow, title, action }) {
  return (
    <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", marginBottom: 22 }}>
      <div>
        <div style={{ fontSize: 11, letterSpacing: 1.5, textTransform: "uppercase", color: COLORS.brass, fontWeight: 600, marginBottom: 4 }}>
          {eyebrow}
        </div>
        <h1 style={{ fontFamily: "Fraunces, serif", fontSize: 28, fontWeight: 700, margin: 0 }}>{title}</h1>
      </div>
      {action}
    </div>
  );
}

// ============================================================
function Dashboard({ subjectStats, students }) {
  const overallAvg = subjectStats.length
    ? Math.round(subjectStats.reduce((a, b) => a + b.pct, 0) / subjectStats.length)
    : 0;
  const maxPct = Math.max(...subjectStats.map((s) => s.pct), 1);

  return (
    <div>
      <SectionTitle eyebrow="Overview" title="All Subjects · Average Attendance" />

      <div className="rp-grid-3" style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16, marginBottom: 30 }}>
        <StatCard label="Enrolled Students" value={students.length} />
        <StatCard label="Subjects Running" value={subjectStats.length} />
        <StatCard label="Overall Average Attendance" value={`${overallAvg}%`} accent />
      </div>

      <div
        style={{
          background: "#fff",
          border: `1px solid ${COLORS.line}`,
          borderRadius: 12,
          padding: 24,
        }}
      >
        <div style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 16, marginBottom: 18 }}>
          Average attendance per subject
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {subjectStats.map((s) => (
            <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 14 }}>
              <div style={{ width: 190, fontSize: 13, fontWeight: 500 }}>
                {s.name}
                <div style={{ fontSize: 11, color: COLORS.slate, fontFamily: "JetBrains Mono, monospace" }}>
                  {s.code} · {s.type === "core" ? "Core" : "Elective"} · {s.roster} students
                </div>
              </div>
              <div style={{ flex: 1, background: COLORS.parchmentDark, borderRadius: 6, height: 14, position: "relative" }}>
                <div
                  className="rp-bar-fill"
                  style={{
                    width: `${(s.pct / 100) * 100}%`,
                    background: pctColor(s.pct),
                    height: "100%",
                    borderRadius: 6,
                  }}
                />
              </div>
              <div style={{ width: 46, textAlign: "right", fontFamily: "JetBrains Mono, monospace", fontWeight: 600, fontSize: 13 }}>
                {s.pct}%
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function StatCard({ label, value, accent }) {
  return (
    <div
      className="rp-card-hover"
      style={{
        background: accent ? COLORS.ink : "#fff",
        color: accent ? COLORS.parchment : COLORS.ink,
        border: `1px solid ${COLORS.line}`,
        borderRadius: 12,
        padding: "18px 20px",
      }}
    >
      <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: 1, color: accent ? COLORS.brassSoft : COLORS.slate, marginBottom: 8 }}>
        {label}
      </div>
      <div style={{ fontFamily: "Fraunces, serif", fontSize: 30, fontWeight: 700 }}>{value}</div>
    </div>
  );
}

// ============================================================
function StudentsTab({ students, subjects, attendance, addStudent, editStudent, removeStudent, transferElective, importStudents }) {
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null); // null = adding a new student
  const emptyForm = { name: "", email: "", id: "", program: "MSc", electiveByGroup: { G1: "", G2: "" } };
  const [form, setForm] = useState(emptyForm);
  const [search, setSearch] = useState("");
  const [transferStudent, setTransferStudent] = useState(null);
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState(null); // { added, updated, skipped } or { error }
  const fileInputRef = useRef(null);

  const handleImportFile = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = ""; // reset so re-selecting the same file re-triggers onChange
    if (!file) return;
    setImporting(true);
    setImportResult(null);
    try {
      const rows = await parseImportFile(file, subjects);
      if (rows.length === 0) {
        setImportResult({ error: "No rows found — check the file has a header row with Enrollment No / Name columns." });
        return;
      }
      const result = importStudents(rows);
      setImportResult(result);
    } catch (err) {
      console.error("Import failed:", err);
      setImportResult({ error: "Couldn't read that file. Make sure it's a valid .xlsx or .csv." });
    } finally {
      setImporting(false);
    }
  };

  const optionsFor = (groupId) => subjects.filter((s) => s.type === "elective" && s.group === groupId);

  const electivesToForm = (student) => {
    const byGroup = { G1: "", G2: "" };
    ELECTIVE_GROUPS.forEach((g) => {
      const pick = student.electives.find((eid) => subjects.find((s) => s.id === eid && s.group === g.id));
      if (pick) byGroup[g.id] = pick;
    });
    return byGroup;
  };

  const startAdd = () => {
    setEditingId(null);
    setForm(emptyForm);
    setShowForm(true);
  };
  const startEdit = (student) => {
    setEditingId(student.id);
    setForm({
      name: student.name,
      email: student.email,
      id: student.id,
      program: student.program,
      electiveByGroup: electivesToForm(student),
    });
    setShowForm(true);
  };

  const submit = (e) => {
    e.preventDefault();
    if (!form.name || !form.email) return;
    const electives = ELECTIVE_GROUPS.map((g) => form.electiveByGroup[g.id]).filter(Boolean);
    if (editingId) {
      editStudent(editingId, { name: form.name, email: form.email, id: form.id || editingId, program: form.program, electives });
    } else {
      addStudent({ name: form.name, email: form.email, id: form.id, program: form.program, electives });
    }
    setForm(emptyForm);
    setEditingId(null);
    setShowForm(false);
  };

  const cancelForm = () => {
    setForm(emptyForm);
    setEditingId(null);
    setShowForm(false);
  };

  const filtered = students.filter(
    (s) => s.name.toLowerCase().includes(search.toLowerCase()) || s.id.toLowerCase().includes(search.toLowerCase())
  );

  const presentStats = (studentId) => {
    let present = 0,
      total = 0;
    Object.entries(attendance).forEach(([key, status]) => {
      if (key.startsWith(`${studentId}__`)) {
        total++;
        if (status === "present") present++;
      }
    });
    return { present, total };
  };

  const coreCount = subjects.filter((s) => s.type === "core").length;

  return (
    <div>
      <SectionTitle
        eyebrow="Batch 2024"
        title="Students"
        action={
          <div style={{ display: "flex", gap: 8 }}>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleImportFile} style={{ display: "none" }} />
            <button
              onClick={() => fileInputRef.current?.click()}
              disabled={importing}
              style={{ ...btnGhost, borderColor: COLORS.line, color: COLORS.ink, opacity: importing ? 0.6 : 1 }}
            >
              <Upload size={14} style={{ marginRight: 6 }} />
              {importing ? "Importing…" : "Import Excel/CSV"}
            </button>
            <button onClick={showForm && !editingId ? cancelForm : startAdd} style={btnPrimary}>
              <Plus size={15} /> Add Student
            </button>
          </div>
        }
      />

      {importResult && (
        <div
          className="rp-banner"
          style={{
            background: importResult.error ? COLORS.absentSoft : COLORS.presentSoft,
            color: importResult.error ? COLORS.absent : COLORS.present,
            fontSize: 13,
            fontWeight: 600,
            padding: "10px 14px",
            borderRadius: 8,
            marginBottom: 16,
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            gap: 12,
          }}
        >
          <span>
            {importResult.error
              ? importResult.error
              : `Imported: ${importResult.added} added, ${importResult.updated} updated${importResult.skipped ? `, ${importResult.skipped} skipped (missing name/enrollment no.)` : ""}.`}
          </span>
          <button onClick={() => setImportResult(null)} style={{ ...iconBtn, color: "inherit" }}>
            <X size={14} />
          </button>
        </div>
      )}

      {showForm &&
        createPortal(
          <div
            onClick={cancelForm}
            className="rp-modal-backdrop"
            style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.35)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 20 }}
          >
        <form
          onSubmit={submit}
          onClick={(e) => e.stopPropagation()}
          className="rp-modal-card"
          style={{ ...cardStyle, width: "min(480px, 92vw)", margin: 0, maxHeight: "88vh", overflowY: "auto" }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <div style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 15 }}>
              {editingId ? "Edit Student" : "New Student"}
            </div>
            <button type="button" onClick={cancelForm} style={iconBtn}>
              <X size={16} />
            </button>
          </div>
          <div className="rp-form-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
            <input required placeholder="Full name" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} style={inputStyle} />
            <input required type="email" placeholder="Email" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} style={inputStyle} />
            <input placeholder="Enrollment no. (auto if blank)" value={form.id} onChange={(e) => setForm({ ...form, id: e.target.value })} style={inputStyle} />
            <select value={form.program} onChange={(e) => setForm({ ...form, program: e.target.value })} style={inputStyle}>
              <option value="MSc">MSc</option>
              <option value="PGD">PGD</option>
            </select>
          </div>

          <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.slate, marginBottom: 10 }}>
            {coreCount} core subjects apply automatically to every student.
            {ELECTIVE_GROUPS.length > 0 && " Pick one elective from each group below — this also updates their roster in the Subjects tab."}
          </div>
          {ELECTIVE_GROUPS.length > 0 && (
            <div className="rp-form-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 16 }}>
              {ELECTIVE_GROUPS.map((g) => (
                <div key={g.id}>
                  <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6 }}>
                    {g.label} <span style={{ color: COLORS.slate, fontWeight: 400 }}>({g.pick})</span>
                  </div>
                  <select
                    required
                    value={form.electiveByGroup[g.id]}
                    onChange={(e) =>
                      setForm((f) => ({ ...f, electiveByGroup: { ...f.electiveByGroup, [g.id]: e.target.value } }))
                    }
                    style={inputStyle}
                  >
                    <option value="" disabled>
                      Select subject…
                    </option>
                    {optionsFor(g.id).map((el) => (
                      <option key={el.id} value={el.id}>
                        {el.code} · {el.name}
                      </option>
                    ))}
                  </select>
                </div>
              ))}
            </div>
          )}

          <div style={{ display: "flex", gap: 10 }}>
            <button type="submit" style={btnPrimary}>
              <Check size={15} /> {editingId ? "Save Changes" : "Save Student"}
            </button>
            <button type="button" onClick={cancelForm} style={{ ...btnGhost, borderColor: COLORS.line, color: COLORS.slate }}>
              Cancel
            </button>
          </div>
        </form>
          </div>,
          document.body
        )}

      <div style={{ position: "relative", marginBottom: 14, maxWidth: 320 }}>
        <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: COLORS.slate }} />
        <input
          placeholder="Search by name or enrollment no."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{ ...inputStyle, paddingLeft: 32 }}
        />
      </div>

      <div className="rp-scroll-x" style={{ background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 12, overflowX: "auto", overflowY: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: COLORS.parchmentDark, textAlign: "left" }}>
              <th style={thStyle}>Enrollment No.</th>
              <th style={thStyle}>Name</th>
              <th style={thStyle}>Email</th>
              <th style={thStyle}>Program</th>
              <th style={thStyle}>Electives</th>
              <th style={{ ...thStyle, textAlign: "center" }}>Total Present</th>
              <th style={thStyle}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((s, idx) => {
              const { present, total } = presentStats(s.id);
              const prevProgram = idx > 0 ? filtered[idx - 1].program : null;
              const showDivider = s.program !== prevProgram;
              return (
              <React.Fragment key={s.id}>
                {showDivider && (
                  <tr>
                    <td
                      colSpan={7}
                      style={{
                        background: COLORS.brassSoft,
                        color: "#6B4F22",
                        padding: "6px 16px",
                        fontSize: 11,
                        fontWeight: 700,
                        textTransform: "uppercase",
                        letterSpacing: 1,
                      }}
                    >
                      {s.program} Students
                    </td>
                  </tr>
                )}
              <tr style={{ borderTop: `1px solid ${COLORS.line}`, background: editingId === s.id ? COLORS.parchmentDark : "transparent" }}>
                <td style={{ ...tdStyle, fontFamily: "JetBrains Mono, monospace" }}>{s.id}</td>
                <td style={tdStyle}>{s.name}</td>
                <td style={{ ...tdStyle, color: COLORS.slate }}>{s.email}</td>
                <td style={tdStyle}>
                  <Badge>{s.program}</Badge>
                </td>
                <td style={tdStyle}>
                  {subjects.filter((sub) => s.electives.includes(sub.id)).map((sub) => sub.name).join(", ") || "—"}
                </td>
                <td style={{ ...tdStyle, textAlign: "center" }}>
                  <span style={{ fontWeight: 700, color: pctColor(total ? Math.round((present / total) * 100) : 0) }}>{present}</span>
                  <span style={{ color: COLORS.slate }}> / {total}</span>
                </td>
                <td style={tdStyle}>
                  <div style={{ display: "flex", gap: 4 }}>
                    <button onClick={() => startEdit(s)} style={iconBtn} title="Edit student">
                      <Pencil size={14} />
                    </button>
                    {ELECTIVE_GROUPS.length > 0 && s.electives.length > 0 && (
                      <button onClick={() => setTransferStudent(s)} style={iconBtn} title="Transfer elective">
                        <ArrowRightLeft size={14} />
                      </button>
                    )}
                    <button onClick={() => removeStudent(s.id)} style={iconBtn} title="Remove student">
                      <X size={14} />
                    </button>
                  </div>
                </td>
              </tr>
              </React.Fragment>
              );
            })}
            {filtered.length === 0 && (
              <tr>
                <td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: COLORS.slate, padding: 24 }}>
                  No students found.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {transferStudent && (
        <TransferElectivePopover
          student={transferStudent}
          subjects={subjects}
          transferElective={transferElective}
          onClose={() => setTransferStudent(null)}
        />
      )}
    </div>
  );
}

function TransferElectivePopover({ student, subjects, transferElective, onClose }) {
  const [picks, setPicks] = useState({}); // { groupId: newSubjectId }
  const [done, setDone] = useState([]); // groupIds already transferred, for confirmation feedback

  const groupInfo = ELECTIVE_GROUPS.map((g) => {
    const options = subjects.filter((s) => s.type === "elective" && s.group === g.id);
    const currentId = student.electives.find((eid) => options.some((o) => o.id === eid));
    const current = options.find((o) => o.id === currentId);
    const alternatives = options.filter((o) => o.id !== currentId);
    return { ...g, current, alternatives };
  }).filter((g) => g.current); // only show groups this student is actually enrolled in

  const doTransfer = (groupId, fromId) => {
    const toId = picks[groupId];
    if (!toId || toId === fromId) return;
    transferElective(student.id, fromId, toId);
    setDone((d) => [...d, groupId]);
  };

  return createPortal(
    <div
      onClick={onClose}
      className="rp-modal-backdrop"
      style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.35)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}
    >
      <div onClick={(e) => e.stopPropagation()} className="rp-modal-card" style={{ ...cardStyle, width: "min(420px, 92vw)", margin: 0 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <div style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 15 }}>Transfer Elective</div>
          <button onClick={onClose} style={iconBtn}>
            <X size={16} />
          </button>
        </div>
        <div style={{ fontSize: 12, color: COLORS.slate, marginBottom: 16 }}>
          {student.name} · {student.id} — moving to a new subject within the same group carries their existing attendance history over.
        </div>

        {groupInfo.map((g) => (
          <div key={g.id} style={{ marginBottom: 16, paddingBottom: 16, borderBottom: `1px solid ${COLORS.line}` }}>
            <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6 }}>{g.label}</div>
            <div style={{ fontSize: 12, color: COLORS.slate, marginBottom: 8 }}>
              Currently: <strong style={{ color: COLORS.ink }}>{g.current.code} · {g.current.name}</strong>
            </div>
            {done.includes(g.id) ? (
              <div style={{ fontSize: 12, color: COLORS.present, fontWeight: 600 }}>✓ Transferred — attendance history moved.</div>
            ) : (
              <div style={{ display: "flex", gap: 8 }}>
                <select
                  value={picks[g.id] || ""}
                  onChange={(e) => setPicks((p) => ({ ...p, [g.id]: e.target.value }))}
                  style={{ ...inputStyle, flex: 1 }}
                >
                  <option value="" disabled>
                    Transfer to…
                  </option>
                  {g.alternatives.map((o) => (
                    <option key={o.id} value={o.id}>
                      {o.code} · {o.name}
                    </option>
                  ))}
                </select>
                <button
                  onClick={() => doTransfer(g.id, g.current.id)}
                  disabled={!picks[g.id]}
                  style={{ ...btnPrimary, opacity: picks[g.id] ? 1 : 0.5, cursor: picks[g.id] ? "pointer" : "not-allowed" }}
                >
                  Transfer
                </button>
              </div>
            )}
          </div>
        ))}

        <button onClick={onClose} style={{ ...btnGhost, borderColor: COLORS.line, color: COLORS.slate, width: "100%", justifyContent: "center" }}>
          Done
        </button>
      </div>
    </div>,
    document.body
  );
}

// ============================================================
function SubjectsTab({ subjects, addSubject, editSubject, removeSubject, studentsFor }) {
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const emptyForm = { code: "", name: "", type: "core", program: "Both", group: "G1", faculty: "", defaultTime: "" };
  const [form, setForm] = useState(emptyForm);

  const startAdd = () => {
    setEditingId(null);
    setForm(emptyForm);
    setShowForm(true);
  };
  const startEdit = (s) => {
    setEditingId(s.id);
    setForm({
      code: s.code,
      name: s.name,
      type: s.type,
      program: s.program,
      group: s.group || "G1",
      faculty: s.faculty || "",
      defaultTime: s.defaultTime || "",
    });
    setShowForm(true);
  };
  const cancelForm = () => {
    setForm(emptyForm);
    setEditingId(null);
    setShowForm(false);
  };

  const submit = (e) => {
    e.preventDefault();
    if (!form.code || !form.name) return;
    const payload = form.type === "core" ? { ...form, group: null } : form;
    if (editingId) editSubject(editingId, payload);
    else addSubject(payload);
    setForm(emptyForm);
    setEditingId(null);
    setShowForm(false);
  };

  const core = subjects.filter((s) => s.type === "core");
  const groups = ELECTIVE_GROUPS.map((g) => ({
    ...g,
    subjects: subjects.filter((s) => s.type === "elective" && s.group === g.id),
  }));

  const SubjectCard = (s) => (
    <div key={s.id} className="rp-card-hover" style={{ ...cardStyle, margin: 0, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
      <div>
        <div style={{ fontFamily: "JetBrains Mono, monospace", fontSize: 11, color: COLORS.brass, marginBottom: 4 }}>{s.code}</div>
        <div style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 16, marginBottom: 6 }}>{s.name}</div>
        {s.faculty && <div style={{ fontSize: 12, color: COLORS.slate, marginBottom: 6 }}>{s.faculty}</div>}
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          <Badge tone={s.type === "core" ? "ink" : "brass"}>{s.type === "core" ? "Core" : "Elective"}</Badge>
          <Badge>{s.program}</Badge>
          <Badge tone="slate">{studentsFor(s).length} students</Badge>
          {s.defaultTime && <Badge tone="slate">{s.defaultTime}</Badge>}
        </div>
      </div>
      <div style={{ display: "flex", gap: 4 }}>
        <button onClick={() => startEdit(s)} style={iconBtn} title="Edit subject">
          <Pencil size={14} />
        </button>
        <button onClick={() => removeSubject(s.id)} style={iconBtn} title="Remove subject">
          <X size={14} />
        </button>
      </div>
    </div>
  );

  return (
    <div>
      <SectionTitle
        eyebrow={`Curriculum · ${subjects.length} subjects total`}
        title="Subjects"
        action={
          <button onClick={showForm && !editingId ? cancelForm : startAdd} style={btnPrimary}>
            <Plus size={15} /> Add Subject
          </button>
        }
      />

      {showForm &&
        createPortal(
          <div
            onClick={cancelForm}
            className="rp-modal-backdrop"
            style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.35)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 20 }}
          >
        <form
          onSubmit={submit}
          onClick={(e) => e.stopPropagation()}
          className="rp-modal-card"
          style={{ ...cardStyle, width: "min(560px, 92vw)", margin: 0, maxHeight: "88vh", overflowY: "auto" }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <div style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 15 }}>
              {editingId ? "Edit Subject" : "New Subject"}
            </div>
            <button type="button" onClick={cancelForm} style={iconBtn}>
              <X size={16} />
            </button>
          </div>
          <div className="rp-form-grid" style={{ display: "grid", gridTemplateColumns: "1fr 2fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
            <input required placeholder="Code e.g. EL-401" value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} style={inputStyle} />
            <input required placeholder="Subject name" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} style={inputStyle} />
            <select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })} style={inputStyle}>
              <option value="core">Core</option>
              <option value="elective">Elective</option>
            </select>
            {form.type === "elective" ? (
              <select value={form.group} onChange={(e) => setForm({ ...form, group: e.target.value })} style={inputStyle}>
                {ELECTIVE_GROUPS.map((g) => (
                  <option key={g.id} value={g.id}>
                    {g.label}
                  </option>
                ))}
              </select>
            ) : (
              <div />
            )}
          </div>
          <div className="rp-form-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 16 }}>
            <select value={form.program} onChange={(e) => setForm({ ...form, program: e.target.value })} style={inputStyle}>
              <option value="Both">Both (MSc + PGD)</option>
              <option value="MSc">MSc only</option>
              <option value="PGD">PGD only</option>
            </select>
            <input placeholder="Faculty name (optional)" value={form.faculty} onChange={(e) => setForm({ ...form, faculty: e.target.value })} style={inputStyle} />
            <select value={form.defaultTime} onChange={(e) => setForm({ ...form, defaultTime: e.target.value })} style={inputStyle}>
              <option value="">Default time slot (optional)</option>
              {TIME_SLOTS.map((slot) => (
                <option key={slot} value={slot}>
                  {slot}
                </option>
              ))}
            </select>
          </div>
          <div style={{ display: "flex", gap: 10 }}>
            <button type="submit" style={btnPrimary}>
              <Check size={15} /> {editingId ? "Save Changes" : "Save Subject"}
            </button>
            <button type="button" onClick={cancelForm} style={{ ...btnGhost, borderColor: COLORS.line, color: COLORS.slate }}>
              Cancel
            </button>
          </div>
        </form>
          </div>,
          document.body
        )}

      <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.slate, marginBottom: 10, textTransform: "uppercase", letterSpacing: 1 }}>
        Core subjects ({core.length}) · compulsory for everyone
      </div>
      <div className="rp-grid-2" style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 14, marginBottom: 28 }}>
        {core.map(SubjectCard)}
      </div>

      {groups.map((g) => (
        <div key={g.id} style={{ marginBottom: 28 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.slate, marginBottom: 10, textTransform: "uppercase", letterSpacing: 1 }}>
            {g.label} ({g.subjects.length}) · {g.pick}
          </div>
          <div className="rp-grid-2" style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 14 }}>{g.subjects.map(SubjectCard)}</div>
        </div>
      ))}
    </div>
  );
}

// ============================================================
function MarkTab({ subjects, studentsFor, attendance, setMark, getSession, setSession, slotsForDate }) {
  const [subjectId, setSubjectId] = useState(subjects[0]?.id);
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [slot, setSlot] = useState("1");
  const subject = subjects.find((s) => s.id === subjectId);
  const roster = subject ? studentsFor(subject) : [];
  const existingSlots = subject ? slotsForDate(subject.id, date) : [];
  const session = getSession(subjectId, date, slot);
  const time = session.time || (existingSlots.length === 0 ? subject?.defaultTime : "") || "";

  // when subject or date changes, default back to the first existing
  // session for that date (or slot "1" if this date has none yet)
  React.useEffect(() => {
    setSlot(existingSlots.length > 0 ? existingSlots[0].slot : "1");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [subjectId, date]);

  const addNewSession = () => {
    const nextSlot = existingSlots.length ? String(Math.max(...existingSlots.map((s) => Number(s.slot))) + 1) : "1";
    setSlot(nextSlot);
  };

  const markAll = (status) => {
    roster.forEach((st) => setMark(st.id, subjectId, date, status, slot));
  };

  const presentCount = roster.filter((st) => attendance[`${st.id}__${subjectId}__${date}__${slot}`] === "present").length;

  return (
    <div>
      <SectionTitle
        eyebrow="Daily Roll Call"
        title="Mark Attendance"
        action={
          roster.length > 0 && (
            <div style={{ textAlign: "right" }}>
              <div style={{ fontFamily: "Fraunces, serif", fontWeight: 700, fontSize: 20, color: pctColor(Math.round((presentCount / roster.length) * 100)) }}>
                {presentCount} / {roster.length}
              </div>
              <div style={{ fontSize: 11, color: COLORS.slate }}>Total Present</div>
            </div>
          )
        }
      />

      <div style={{ display: "flex", gap: 12, marginBottom: 14, alignItems: "center", flexWrap: "wrap" }}>
        <select value={subjectId} onChange={(e) => setSubjectId(e.target.value)} style={{ ...inputStyle, maxWidth: 280 }}>
          {subjects.map((s) => (
            <option key={s.id} value={s.id}>
              {s.code} · {s.name}
            </option>
          ))}
        </select>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={{ ...inputStyle, maxWidth: 170 }} />

        {existingSlots.length > 0 && (
          <select value={slot} onChange={(e) => setSlot(e.target.value)} style={{ ...inputStyle, maxWidth: 200 }}>
            {existingSlots.map((s) => (
              <option key={s.slot} value={s.slot}>
                Session {s.slot}{s.time ? ` — ${s.time}` : ""}
              </option>
            ))}
            {!existingSlots.some((s) => s.slot === slot) && <option value={slot}>New session</option>}
          </select>
        )}
        <button
          type="button"
          onClick={addNewSession}
          style={{ ...btnGhost, fontSize: 12, padding: "8px 12px" }}
          title="Add another lecture for this subject on this date (e.g. a second session later in the day)"
        >
          <Plus size={13} style={{ marginRight: 4 }} />
          Add Session
        </button>

        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <Clock size={14} color={COLORS.slate} />
          <select
            value={time}
            onChange={(e) => setSession(subjectId, date, { time: e.target.value }, slot)}
            style={{ ...inputStyle, maxWidth: 170 }}
          >
            <option value="" disabled>
              Select time slot…
            </option>
            {!TIME_SLOTS.includes(time) && time && <option value={time}>{time} (custom)</option>}
            {TIME_SLOTS.map((t) => (
              <option key={t} value={t}>
                {t}
              </option>
            ))}
          </select>
        </div>
        <div style={{ flex: 1 }} />
        <button onClick={() => markAll("present")} style={btnGhost}>
          Mark all present
        </button>
        <button onClick={() => markAll("absent")} style={btnGhostRed}>
          Mark all absent
        </button>
      </div>

      {existingSlots.length > 1 && (
        <div style={{ fontSize: 12, color: COLORS.slate, marginBottom: 14, marginTop: -6 }}>
          This date has {existingSlots.length} sessions for {subject?.code} — make sure you're marking the right one above.
        </div>
      )}

      <div style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6 }}>
          Remark for this session <span style={{ color: COLORS.slate, fontWeight: 400 }}>(applies to the whole session, not one student — e.g. "faculty on leave, proxy conducted")</span>
        </div>
        <input
          placeholder="Optional note about this session…"
          value={session.remark}
          onChange={(e) => setSession(subjectId, date, { remark: e.target.value }, slot)}
          style={inputStyle}
        />
      </div>

      <div className="rp-scroll-x" style={{ background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 12, overflowX: "auto", overflowY: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: COLORS.parchmentDark, textAlign: "left" }}>
              <th style={thStyle}>Enrollment No.</th>
              <th style={thStyle}>Name</th>
              <th style={{ ...thStyle, textAlign: "center" }}>Status</th>
            </tr>
          </thead>
          <tbody>
            {roster.map((st, idx) => {
              const key = `${st.id}__${subjectId}__${date}__${slot}`;
              const status = attendance[key]; // undefined = blank/unmarked
              const prevProgram = idx > 0 ? roster[idx - 1].program : null;
              const showDivider = st.program !== prevProgram;
              return (
                <React.Fragment key={st.id}>
                  {showDivider && (
                    <tr>
                      <td
                        colSpan={3}
                        style={{
                          background: COLORS.brassSoft,
                          color: "#6B4F22",
                          padding: "5px 16px",
                          fontSize: 11,
                          fontWeight: 700,
                          textTransform: "uppercase",
                          letterSpacing: 1,
                        }}
                      >
                        {st.program} Students
                      </td>
                    </tr>
                  )}
                <tr style={{ borderTop: `1px solid ${COLORS.line}` }}>
                  <td style={{ ...tdStyle, fontFamily: "JetBrains Mono, monospace" }}>{st.id}</td>
                  <td style={tdStyle}>{st.name}</td>
                  <td style={{ ...tdStyle, textAlign: "center" }}>
                    <div style={{ display: "inline-flex", gap: 4, flexWrap: "wrap", justifyContent: "center" }}>
                      {STATUS_ORDER.map((s) => {
                        const meta = STATUS_META[s];
                        const active = status === s;
                        return (
                          <button
                            key={s}
                            className="rp-status-btn"
                            onClick={() => setMark(st.id, subjectId, date, s, slot)}
                            style={{
                              ...pillBtn,
                              background: active ? meta.color : "#fff",
                              color: active ? "#fff" : meta.color,
                              borderColor: meta.color,
                            }}
                          >
                            {meta.label}
                          </button>
                        );
                      })}
                      {!status && <span style={{ fontSize: 11, color: COLORS.slate, alignSelf: "center", marginLeft: 4 }}>unmarked</span>}
                    </div>
                  </td>
                </tr>
                </React.Fragment>
              );
            })}
            {roster.length === 0 && (
              <tr>
                <td colSpan={3} style={{ ...tdStyle, textAlign: "center", color: COLORS.slate, padding: 24 }}>
                  No students enrolled in this subject.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================
function AllAttendanceTab({ students, subjects, attendance, sessions, studentsFor, setMark, setSession, deleteMark, deleteSession }) {
  const [subjectId, setSubjectId] = useState(""); // empty = nothing chosen yet, nothing renders
  const [search, setSearch] = useState("");
  const [editingSession, setEditingSession] = useState(null); // {date, slot} currently open in the popover
  const [addingDate, setAddingDate] = useState(""); // date picker value for the "add session" mini-form
  const [addingTime, setAddingTime] = useState("");

  const subject = subjects.find((s) => s.id === subjectId);
  const roster = useMemo(() => (subject ? studentsFor(subject) : []), [subject, studentsFor]);

  // Each column is one SESSION (date + slot) — a date can have more than one,
  // e.g. a morning and an afternoon lecture of the same subject.
  const cols = useMemo(() => {
    if (!subject) return [];
    const set = new Set();
    const prefix = `${subject.id}__`;
    Object.keys(sessions).forEach((key) => {
      if (key.startsWith(prefix)) set.add(key.slice(prefix.length)); // "date__slot"
    });
    Object.keys(attendance).forEach((key) => {
      const [sid, subId, date, slot] = key.split("__");
      if (subId === subject.id && roster.some((r) => r.id === sid)) set.add(`${date}__${slot}`);
    });
    return Array.from(set)
      .map((k) => {
        const [date, slot] = k.split("__");
        return { date, slot };
      })
      .sort((a, b) => (a.date === b.date ? Number(a.slot) - Number(b.slot) : a.date < b.date ? -1 : 1));
  }, [sessions, attendance, subject, roster]);

  const filteredRoster = roster.filter(
    (s) => !search || s.name.toLowerCase().includes(search.toLowerCase()) || s.id.toLowerCase().includes(search.toLowerCase())
  );

  const formatDate = (d) => {
    const dObj = new Date(d);
    return dObj.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
  };

  const cycleStatus = (current) => {
    const i = STATUS_ORDER.indexOf(current);
    return STATUS_ORDER[(i + 1) % STATUS_ORDER.length];
  };

  const nextSlotFor = (date) => {
    const existing = cols.filter((c) => c.date === date).map((c) => Number(c.slot));
    return existing.length ? String(Math.max(...existing) + 1) : "1";
  };

  const addSession = () => {
    if (!addingDate) return;
    const slot = nextSlotFor(addingDate);
    setSession(subject.id, addingDate, { time: addingTime, remark: "" }, slot);
    setAddingDate("");
    setAddingTime("");
  };

  return (
    <div>
      <SectionTitle eyebrow="Attendance Register" title="All Attendance" />

      <div style={{ ...cardStyle, marginBottom: subject ? 20 : 0 }}>
        <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 8 }}>Choose a subject to view its attendance register</div>
        <select
          value={subjectId}
          onChange={(e) => {
            setSubjectId(e.target.value);
            setSearch("");
            setEditingSession(null);
          }}
          style={{ ...inputStyle, maxWidth: 380 }}
        >
          <option value="">Select a subject…</option>
          {subjects.map((s) => (
            <option key={s.id} value={s.id}>
              {s.code} · {s.name}
            </option>
          ))}
        </select>
      </div>

      {!subject && (
        <div style={{ color: COLORS.slate, fontSize: 13, textAlign: "center", padding: "40px 0" }}>
          Pick a subject above to see every student's attendance, session by session — just like the exported Excel sheet.
        </div>
      )}

      {subject && (
        <>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 12 }}>
            <div>
              <div style={{ fontFamily: "Fraunces, serif", fontWeight: 700, fontSize: 18 }}>
                {subject.code} · {subject.name}
              </div>
              <div style={{ fontSize: 12, color: COLORS.slate }}>
                {roster.length} students · {cols.length} session{cols.length === 1 ? "" : "s"} recorded
              </div>
            </div>
            <div style={{ position: "relative", maxWidth: 260, flex: "1 1 200px" }}>
              <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: COLORS.slate }} />
              <input
                placeholder="Search students…"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                style={{ ...inputStyle, paddingLeft: 32 }}
              />
            </div>
          </div>

          <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 16, flexWrap: "wrap", background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 10, padding: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.slate }}>Add another session:</div>
            <input type="date" value={addingDate} onChange={(e) => setAddingDate(e.target.value)} style={{ ...inputStyle, maxWidth: 160 }} />
            <select value={addingTime} onChange={(e) => setAddingTime(e.target.value)} style={{ ...inputStyle, maxWidth: 170 }}>
              <option value="">Select time slot…</option>
              {TIME_SLOTS.map((slot) => (
                <option key={slot} value={slot}>
                  {slot}
                </option>
              ))}
            </select>
            <button onClick={addSession} disabled={!addingDate} style={{ ...btnPrimary, opacity: addingDate ? 1 : 0.5, cursor: addingDate ? "pointer" : "not-allowed" }}>
              <Plus size={15} /> Add Session
            </button>
            {addingDate && cols.some((c) => c.date === addingDate) && (
              <div style={{ fontSize: 11, color: COLORS.slate }}>
                This date already has {cols.filter((c) => c.date === addingDate).length} session(s) — this adds another one.
              </div>
            )}
          </div>

          {cols.length === 0 ? (
            <div style={{ color: COLORS.slate, fontSize: 13, textAlign: "center", padding: "30px 0" }}>
              No sessions recorded for this subject yet — mark attendance in the "Mark Attendance" tab, or add a session above.
            </div>
          ) : (
            <div className="rp-scroll-x" style={{ background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 12, overflow: "auto" }}>
              <table style={{ borderCollapse: "collapse", fontSize: 12.5, minWidth: "100%" }}>
                <thead>
                  <tr style={{ background: COLORS.parchmentDark, textAlign: "left" }}>
                    <th style={{ ...thStyle, position: "sticky", left: 0, background: COLORS.parchmentDark, zIndex: 1 }}>Enrollment No.</th>
                    <th style={{ ...thStyle, position: "sticky", left: 130, background: COLORS.parchmentDark, zIndex: 1 }}>Name</th>
                    {cols.map(({ date, slot }) => {
                      const session = sessions[`${subject.id}__${date}__${slot}`] || { time: "", remark: "" };
                      const colKey = `${date}__${slot}`;
                      return (
                        <th
                          key={colKey}
                          style={{ ...thStyle, textAlign: "center", minWidth: 68, cursor: "pointer" }}
                          onClick={() => setEditingSession({ date, slot })}
                        >
                          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 2 }}>
                            <span>{formatDate(date)}</span>
                            <span style={{ fontWeight: 400, fontSize: 10, color: COLORS.slate }}>{session.time || "set time"}</span>
                            {session.remark && (
                              <span
                                title={session.remark}
                                style={{ width: 5, height: 5, borderRadius: "50%", background: COLORS.warn, display: "inline-block" }}
                              />
                            )}
                          </div>
                        </th>
                      );
                    })}
                    <th style={{ ...thStyle, textAlign: "center" }}>Present</th>
                    <th style={{ ...thStyle, textAlign: "center" }}>Total</th>
                    <th style={{ ...thStyle, textAlign: "center" }}>%</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRoster.map((st, idx) => {
                    let present = 0,
                      total = 0;
                    const prevProgram = idx > 0 ? filteredRoster[idx - 1].program : null;
                    const showDivider = st.program !== prevProgram;
                    return (
                      <React.Fragment key={st.id}>
                        {showDivider && (
                          <tr>
                            <td
                              colSpan={cols.length + 5}
                              style={{
                                background: COLORS.brassSoft,
                                color: "#6B4F22",
                                padding: "5px 16px",
                                fontSize: 11,
                                fontWeight: 700,
                                textTransform: "uppercase",
                                letterSpacing: 1,
                                position: "sticky",
                                left: 0,
                              }}
                            >
                              {st.program} Students
                            </td>
                          </tr>
                        )}
                      <tr style={{ borderTop: `1px solid ${COLORS.line}`, background: idx % 2 === 1 ? COLORS.parchment : "#fff" }}>
                        <td style={{ ...tdStyle, fontFamily: "JetBrains Mono, monospace", position: "sticky", left: 0, background: "inherit" }}>{st.id}</td>
                        <td style={{ ...tdStyle, position: "sticky", left: 130, background: "inherit", whiteSpace: "nowrap" }}>{st.name}</td>
                        {cols.map(({ date, slot }) => {
                          const key = `${st.id}__${subject.id}__${date}__${slot}`;
                          const status = attendance[key];
                          if (status) {
                            total++;
                            if (status === "present") present++;
                          }
                          const meta = status ? STATUS_META[status] : null;
                          return (
                            <td key={`${date}__${slot}`} style={{ ...tdStyle, textAlign: "center", padding: 4 }}>
                              <button
                                onClick={() => setMark(st.id, subject.id, date, status ? cycleStatus(status) : "present", slot)}
                                title={meta ? `${meta.label} — click to change` : "No record — click to mark present"}
                                style={{
                                  width: 30,
                                  height: 24,
                                  borderRadius: 6,
                                  border: `1px solid ${meta ? meta.color : COLORS.line}`,
                                  background: meta ? meta.color : "#fff",
                                  color: meta ? "#fff" : COLORS.slate,
                                  fontWeight: 700,
                                  fontSize: 11,
                                  cursor: "pointer",
                                }}
                              >
                                {meta ? meta.abbr : "—"}
                              </button>
                            </td>
                          );
                        })}
                        <td style={{ ...tdStyle, textAlign: "center" }}>{present}</td>
                        <td style={{ ...tdStyle, textAlign: "center" }}>{total}</td>
                        <td style={{ ...tdStyle, textAlign: "center", fontWeight: 700, color: pctColor(total ? Math.round((present / total) * 100) : 0) }}>
                          {total ? Math.round((present / total) * 100) : 0}%
                        </td>
                      </tr>
                      </React.Fragment>
                    );
                  })}
                  {filteredRoster.length === 0 && (
                    <tr>
                      <td colSpan={cols.length + 5} style={{ ...tdStyle, textAlign: "center", color: COLORS.slate, padding: 24 }}>
                        No students match your search.
                      </td>
                    </tr>
                  )}
                </tbody>
                {filteredRoster.length > 0 && (
                  <tfoot>
                    <tr style={{ borderTop: `2px solid ${COLORS.ink}`, background: COLORS.parchmentDark, fontWeight: 700 }}>
                      <td style={{ ...tdStyle, position: "sticky", left: 0, background: COLORS.parchmentDark }} colSpan={2}>
                        Total Present (that session)
                      </td>
                      {cols.map(({ date, slot }) => {
                        const presentThatSession = filteredRoster.filter(
                          (st) => attendance[`${st.id}__${subject.id}__${date}__${slot}`] === "present"
                        ).length;
                        return (
                          <td key={`${date}__${slot}`} style={{ ...tdStyle, textAlign: "center", color: pctColor(Math.round((presentThatSession / filteredRoster.length) * 100)) }}>
                            {presentThatSession}/{filteredRoster.length}
                          </td>
                        );
                      })}
                      <td colSpan={3} style={{ ...tdStyle, textAlign: "center", color: COLORS.slate, fontWeight: 400 }}>
                        ← present per session
                      </td>
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>
          )}

          <div style={{ fontSize: 11, color: COLORS.slate, marginTop: 10 }}>
            Click any status cell to cycle Present → Absent. Click a session's header to edit its time and remark, or delete it. A date with more than
            one lecture (e.g. morning + afternoon) shows as separate columns.
          </div>
        </>
      )}

      {editingSession && subject && (
        <SessionEditPopover
          subject={subject}
          date={editingSession.date}
          slot={editingSession.slot}
          session={sessions[`${subject.id}__${editingSession.date}__${editingSession.slot}`] || { time: "", remark: "" }}
          onSave={(patch) => setSession(subject.id, editingSession.date, patch, editingSession.slot)}
          onDelete={() => deleteSession(subject.id, editingSession.date, editingSession.slot)}
          onClose={() => setEditingSession(null)}
        />
      )}
    </div>
  );
}

function SessionEditPopover({ subject, date, slot, session, onSave, onDelete, onClose }) {
  const [time, setTime] = useState(session.time);
  const [remark, setRemark] = useState(session.remark);

  const save = () => {
    onSave({ time, remark });
    onClose();
  };

  const handleDelete = () => {
    if (
      !window.confirm(
        `Delete ALL attendance for ${subject.code} on ${new Date(date).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })}${time ? ` (${time})` : ""}?\n\nThis removes every student's status for this specific session and can't be undone.`
      )
    )
      return;
    onDelete();
    onClose();
  };

  return createPortal(
    <div
      onClick={onClose}
      className="rp-modal-backdrop"
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(27,42,74,0.35)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 50,
      }}
    >
      <div onClick={(e) => e.stopPropagation()} className="rp-modal-card" style={{ ...cardStyle, width: "min(360px, 92vw)", margin: 0 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <div style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 15 }}>
            {subject.code} — {new Date(date).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" })}
            {Number(slot) > 1 && <span style={{ color: COLORS.slate, fontWeight: 400 }}> (session {slot})</span>}
          </div>
          <button onClick={onClose} style={iconBtn}>
            <X size={16} />
          </button>
        </div>

        <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6 }}>Lecture time</div>
        <select value={time} onChange={(e) => setTime(e.target.value)} style={{ ...inputStyle, marginBottom: 14 }}>
          <option value="" disabled>
            Select time slot…
          </option>
          {!TIME_SLOTS.includes(time) && time && <option value={time}>{time} (custom)</option>}
          {TIME_SLOTS.map((s) => (
            <option key={s} value={s}>
              {s}
            </option>
          ))}
        </select>

        <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6 }}>
          Remark <span style={{ color: COLORS.slate, fontWeight: 400 }}>(for this session, not one student)</span>
        </div>
        <input
          placeholder="e.g. faculty on leave, proxy conducted"
          value={remark}
          onChange={(e) => setRemark(e.target.value)}
          style={{ ...inputStyle, marginBottom: 16 }}
        />

        <div style={{ display: "flex", gap: 10, justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", gap: 10 }}>
            <button onClick={save} style={btnPrimary}>
              <Check size={15} /> Save
            </button>
            <button onClick={onClose} style={{ ...btnGhost, borderColor: COLORS.line, color: COLORS.slate }}>
              Cancel
            </button>
          </div>
          <button
            onClick={handleDelete}
            style={{ ...iconBtn, color: COLORS.absent, display: "flex", alignItems: "center", gap: 4, fontSize: 12, fontWeight: 600 }}
            title="Delete all attendance for this session"
          >
            <X size={14} /> Delete session
          </button>
        </div>
      </div>
    </div>,
    document.body
  );
}

// ============================================================
function AuditLogTab({ students, subjects, attendance, attendanceMeta, sessions }) {
  const [subjectFilter, setSubjectFilter] = useState("all");
  const [markedByFilter, setMarkedByFilter] = useState("all");
  const [search, setSearch] = useState("");

  const editors = useMemo(() => {
    const set = new Set();
    Object.values(attendanceMeta).forEach((m) => m?.editedBy && set.add(m.editedBy));
    return Array.from(set).sort();
  }, [attendanceMeta]);

  const rows = useMemo(() => {
    return Object.entries(attendance)
      .map(([key, status]) => {
        const [sid, subId, date, slot] = key.split("__");
        const student = students.find((s) => s.id === sid);
        const subject = subjects.find((s) => s.id === subId);
        const session = sessions[`${subId}__${date}__${slot}`] || { time: "" };
        const meta = attendanceMeta[key];
        return {
          key,
          sid,
          date,
          slot,
          time: session.time,
          student,
          subject,
          status,
          editedBy: meta?.editedBy || "—",
          editedAt: meta?.editedAt || null,
        };
      })
      .filter((r) => r.student && r.subject)
      .filter((r) => subjectFilter === "all" || r.subject.id === subjectFilter)
      .filter((r) => markedByFilter === "all" || r.editedBy === markedByFilter)
      .filter(
        (r) =>
          !search ||
          r.student.name.toLowerCase().includes(search.toLowerCase()) ||
          r.student.id.toLowerCase().includes(search.toLowerCase())
      )
      .sort((a, b) => {
        if (!a.editedAt && !b.editedAt) return 0;
        if (!a.editedAt) return 1;
        if (!b.editedAt) return -1;
        return b.editedAt.localeCompare(a.editedAt); // most recently edited first
      });
  }, [attendance, attendanceMeta, sessions, students, subjects, subjectFilter, markedByFilter, search]);

  const formatWhen = (iso) => {
    if (!iso) return "—";
    const d = new Date(iso);
    return d.toLocaleString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
  };

  return (
    <div>
      <SectionTitle eyebrow={`${rows.length} records`} title="Audit Log" />
      <div style={{ fontSize: 12, color: COLORS.slate, marginBottom: 14, marginTop: -10 }}>
        Who marked or last changed each attendance record, and when. This information is for in-app reference only — it is never included in
        the exported Excel files.
      </div>

      <div style={{ display: "flex", gap: 12, marginBottom: 16, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: "1 1 240px", maxWidth: 300 }}>
          <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: COLORS.slate }} />
          <input
            placeholder="Search by name or enrollment no."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={{ ...inputStyle, paddingLeft: 32 }}
          />
        </div>
        <select value={subjectFilter} onChange={(e) => setSubjectFilter(e.target.value)} style={{ ...inputStyle, maxWidth: 240 }}>
          <option value="all">All subjects</option>
          {subjects.map((s) => (
            <option key={s.id} value={s.id}>
              {s.code} · {s.name}
            </option>
          ))}
        </select>
        <select value={markedByFilter} onChange={(e) => setMarkedByFilter(e.target.value)} style={{ ...inputStyle, maxWidth: 200 }}>
          <option value="all">All editors</option>
          {editors.map((e) => (
            <option key={e} value={e}>
              {e}
            </option>
          ))}
        </select>
      </div>

      <div className="rp-scroll-x" style={{ background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 12, overflowX: "auto", overflowY: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: COLORS.parchmentDark, textAlign: "left" }}>
              <th style={thStyle}>Date</th>
              <th style={thStyle}>Time</th>
              <th style={thStyle}>Enrollment No.</th>
              <th style={thStyle}>Name</th>
              <th style={thStyle}>Subject</th>
              <th style={thStyle}>Status</th>
              <th style={thStyle}>Marked By</th>
              <th style={thStyle}>Last Updated</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.key} style={{ borderTop: `1px solid ${COLORS.line}` }}>
                <td style={{ ...tdStyle, fontFamily: "JetBrains Mono, monospace", whiteSpace: "nowrap" }}>{r.date}</td>
                <td style={{ ...tdStyle, color: COLORS.slate, fontSize: 12 }}>{r.time || "—"}</td>
                <td style={{ ...tdStyle, fontFamily: "JetBrains Mono, monospace" }}>{r.sid}</td>
                <td style={tdStyle}>{r.student.name}</td>
                <td style={tdStyle}>
                  {r.subject.code}
                  {Number(r.slot) > 1 && <span style={{ color: COLORS.slate }}> (session {r.slot})</span>}
                </td>
                <td style={tdStyle}>
                  <Badge tone={r.status === "present" ? "brass" : "slate"}>{STATUS_META[r.status]?.label || r.status}</Badge>
                </td>
                <td style={{ ...tdStyle, fontWeight: 600 }}>{r.editedBy}</td>
                <td style={{ ...tdStyle, color: COLORS.slate, fontSize: 12, whiteSpace: "nowrap" }}>{formatWhen(r.editedAt)}</td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr>
                <td colSpan={8} style={{ ...tdStyle, textAlign: "center", color: COLORS.slate, padding: 24 }}>
                  No matching records.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================
function ReportsTab({ studentStats }) {
  const [openId, setOpenId] = useState(studentStats[0]?.id || null);
  const open = studentStats.find((s) => s.id === openId);

  return (
    <div>
      <SectionTitle eyebrow="Per Student" title="Attendance Reports" />
      <div className="rp-form-grid" style={{ display: "grid", gridTemplateColumns: "300px 1fr", gap: 18 }}>
        <div className="rp-scroll-x" style={{ background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 12, overflowX: "auto", overflowY: "hidden", alignSelf: "start" }}>
          {studentStats.map((s) => (
            <button
              key={s.id}
              onClick={() => setOpenId(s.id)}
              style={{
                width: "100%",
                textAlign: "left",
                padding: "12px 16px",
                border: "none",
                borderBottom: `1px solid ${COLORS.line}`,
                background: openId === s.id ? COLORS.parchmentDark : "#fff",
                cursor: "pointer",
              }}
            >
              <div style={{ fontWeight: 600, fontSize: 13 }}>{s.name}</div>
              <div style={{ fontSize: 11, color: COLORS.slate, fontFamily: "JetBrains Mono, monospace" }}>{s.id}</div>
              <div style={{ fontSize: 11, marginTop: 4, color: pctColor(s.overallPct), fontWeight: 600 }}>
                {s.overallPct}% overall
              </div>
            </button>
          ))}
        </div>

        {open && (
          <div style={{ background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 12, padding: 22 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 18 }}>
              <div>
                <div style={{ fontFamily: "Fraunces, serif", fontWeight: 700, fontSize: 20 }}>{open.name}</div>
                <div style={{ fontSize: 12, color: COLORS.slate }}>
                  {open.id} · {open.email} · {open.program}
                </div>
              </div>
              <StatCard label="Overall Attendance" value={`${open.overallPct}%`} accent />
            </div>

            <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.slate, marginBottom: 10, textTransform: "uppercase", letterSpacing: 1 }}>
              Subject-wise breakdown
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {open.bySubject.map((b) => (
                <div key={b.subject.id} style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <div style={{ width: 200, fontSize: 13 }}>{b.subject.name}</div>
                  <div style={{ flex: 1, background: COLORS.parchmentDark, borderRadius: 6, height: 10 }}>
                    <div
                      className="rp-bar-fill"
                      style={{
                        width: `${b.pct}%`,
                        background: pctColor(b.pct),
                        height: "100%",
                        borderRadius: 6,
                      }}
                    />
                  </div>
                  <div style={{ width: 90, fontSize: 12, textAlign: "right", color: COLORS.slate }}>
                    {b.present}/{b.total} days
                  </div>
                  <div style={{ width: 40, fontFamily: "JetBrains Mono, monospace", fontWeight: 600, fontSize: 13, textAlign: "right" }}>
                    {b.pct}%
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ============================================================
function UsersTab() {
  const [users, setUsers] = useState(null); // null = loading
  const [error, setError] = useState("");
  const [newUsername, setNewUsername] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [adding, setAdding] = useState(false);
  const [passwordEdits, setPasswordEdits] = useState({}); // { username: newPasswordDraft }
  const [busyUser, setBusyUser] = useState(null);
  const [toast, setToast] = useState("");

  const adminAuth = { adminUser: ADMIN_USERNAME, adminPass: ADMIN_PASSWORD };

  const loadUsers = async () => {
    try {
      const params = new URLSearchParams(adminAuth);
      const res = await fetch(`/api/users?${params.toString()}`);
      const data = await res.json();
      if (data.ok) {
        setUsers(data.users);
        setError("");
      } else {
        setError(data.error || "Failed to load users.");
      }
    } catch {
      setError("Couldn't reach the server.");
    }
  };

  React.useEffect(() => {
    loadUsers();
  }, []);

  const flashToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(""), 2500);
  };

  const addUser = async (e) => {
    e.preventDefault();
    if (!newUsername || !newPassword) return;
    setAdding(true);
    setError("");
    try {
      const res = await fetch("/api/users", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "add", ...adminAuth, username: newUsername, password: newPassword }),
      });
      const data = await res.json();
      if (data.ok) {
        setNewUsername("");
        setNewPassword("");
        flashToast(`User "${newUsername}" created.`);
        loadUsers();
      } else {
        setError(data.error || "Failed to add user.");
      }
    } catch {
      setError("Couldn't reach the server.");
    } finally {
      setAdding(false);
    }
  };

  const changePassword = async (username) => {
    const newPw = passwordEdits[username];
    if (!newPw) return;
    setBusyUser(username);
    setError("");
    try {
      const res = await fetch("/api/users", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "changePassword", ...adminAuth, username, newPassword: newPw }),
      });
      const data = await res.json();
      if (data.ok) {
        setPasswordEdits((p) => ({ ...p, [username]: "" }));
        flashToast(`Password updated for "${username}".`);
      } else {
        setError(data.error || "Failed to change password.");
      }
    } catch {
      setError("Couldn't reach the server.");
    } finally {
      setBusyUser(null);
    }
  };

  const deleteUser = async (username) => {
    if (!window.confirm(`Remove login access for "${username}"? They won't be able to sign in anymore.`)) return;
    setBusyUser(username);
    try {
      const res = await fetch("/api/users", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "delete", ...adminAuth, username }),
      });
      const data = await res.json();
      if (data.ok) {
        flashToast(`User "${username}" removed.`);
        loadUsers();
      } else {
        setError(data.error || "Failed to remove user.");
      }
    } catch {
      setError("Couldn't reach the server.");
    } finally {
      setBusyUser(null);
    }
  };

  return (
    <div>
      <SectionTitle eyebrow="Admin Control Panel" title="Manage Users" />

      {toast && (
        <div className="rp-banner" style={{ background: COLORS.presentSoft, color: COLORS.present, fontSize: 13, fontWeight: 600, padding: "10px 14px", borderRadius: 8, marginBottom: 16 }}>
          {toast}
        </div>
      )}
      {error && (
        <div className="rp-banner" style={{ background: COLORS.absentSoft, color: COLORS.absent, fontSize: 13, fontWeight: 600, padding: "10px 14px", borderRadius: 8, marginBottom: 16 }}>
          {error}
        </div>
      )}

      <form onSubmit={addUser} style={cardStyle}>
        <div style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 15, marginBottom: 12 }}>Add a new user</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: 12 }}>
          <input required placeholder="Username" value={newUsername} onChange={(e) => setNewUsername(e.target.value)} style={inputStyle} />
          <input required type="text" placeholder="Password" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} style={inputStyle} />
          <button type="submit" disabled={adding} style={{ ...btnPrimary, opacity: adding ? 0.6 : 1 }}>
            <Plus size={15} /> {adding ? "Adding…" : "Add User"}
          </button>
        </div>
      </form>

      <div className="rp-scroll-x" style={{ background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 12, overflowX: "auto", overflowY: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: COLORS.parchmentDark, textAlign: "left" }}>
              <th style={thStyle}>Username</th>
              <th style={thStyle}>Role</th>
              <th style={thStyle}>Created</th>
              <th style={thStyle}>New Password</th>
              <th style={thStyle}></th>
            </tr>
          </thead>
          <tbody>
            <tr style={{ borderTop: `1px solid ${COLORS.line}`, background: COLORS.parchmentDark }}>
              <td style={tdStyle}>
                <strong>{ADMIN_USERNAME}</strong>
              </td>
              <td style={tdStyle}>
                <Badge tone="ink">Administrator</Badge>
              </td>
              <td style={{ ...tdStyle, color: COLORS.slate }}>—</td>
              <td style={{ ...tdStyle, color: COLORS.slate, fontSize: 12 }}>Fixed account, not editable here</td>
              <td style={tdStyle}></td>
            </tr>
            {users === null && (
              <tr>
                <td colSpan={5} style={{ ...tdStyle, textAlign: "center", color: COLORS.slate, padding: 24 }}>
                  Loading users…
                </td>
              </tr>
            )}
            {users &&
              users.map((u) => (
                <tr key={u.username} style={{ borderTop: `1px solid ${COLORS.line}` }}>
                  <td style={tdStyle}>{u.username}</td>
                  <td style={tdStyle}>
                    <Badge>Faculty</Badge>
                  </td>
                  <td style={{ ...tdStyle, color: COLORS.slate, fontSize: 12 }}>
                    {u.createdAt ? new Date(u.createdAt).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "—"}
                  </td>
                  <td style={tdStyle}>
                    <input
                      placeholder="New password…"
                      value={passwordEdits[u.username] || ""}
                      onChange={(e) => setPasswordEdits((p) => ({ ...p, [u.username]: e.target.value }))}
                      style={{ ...inputStyle, fontSize: 12, width: 160 }}
                    />
                  </td>
                  <td style={tdStyle}>
                    <div style={{ display: "flex", gap: 6 }}>
                      <button
                        onClick={() => changePassword(u.username)}
                        disabled={!passwordEdits[u.username] || busyUser === u.username}
                        style={{ ...btnGhost, fontSize: 11, padding: "6px 10px", opacity: passwordEdits[u.username] ? 1 : 0.5 }}
                      >
                        Update
                      </button>
                      <button onClick={() => deleteUser(u.username)} style={iconBtn} title="Remove user">
                        <X size={14} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            {users && users.length === 0 && (
              <tr>
                <td colSpan={5} style={{ ...tdStyle, textAlign: "center", color: COLORS.slate, padding: 24 }}>
                  No additional users yet — add faculty accounts above.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------- shared styles ----------
function Badge({ children, tone = "brass" }) {
  const map = {
    brass: { bg: COLORS.brassSoft, color: "#6B4F22" },
    ink: { bg: COLORS.ink, color: COLORS.parchment },
    slate: { bg: "#EEF0F3", color: COLORS.slate },
  };
  const c = map[tone] || map.brass;
  return (
    <span style={{ background: c.bg, color: c.color, fontSize: 11, fontWeight: 600, padding: "3px 8px", borderRadius: 20 }}>
      {children}
    </span>
  );
}

const cardStyle = {
  background: "#fff",
  border: `1px solid ${COLORS.line}`,
  borderRadius: 12,
  padding: 20,
  marginBottom: 20,
};

const inputStyle = {
  width: "100%",
  padding: "9px 12px",
  borderRadius: 8,
  border: `1px solid ${COLORS.line}`,
  fontSize: 13,
  fontFamily: "Inter, sans-serif",
  outline: "none",
  background: "#fff",
  color: COLORS.ink,
};

const thStyle = { padding: "10px 16px", fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, color: COLORS.slate };
const tdStyle = { padding: "10px 16px" };

const btnPrimary = {
  display: "inline-flex",
  alignItems: "center",
  gap: 6,
  background: COLORS.ink,
  color: COLORS.parchment,
  border: "none",
  padding: "9px 16px",
  borderRadius: 8,
  fontSize: 13,
  fontWeight: 600,
  cursor: "pointer",
};

const btnGhost = {
  background: "#fff",
  border: `1px solid ${COLORS.present}`,
  color: COLORS.present,
  padding: "8px 12px",
  borderRadius: 8,
  fontSize: 12,
  fontWeight: 600,
  cursor: "pointer",
};

const btnGhostRed = { ...btnGhost, border: `1px solid ${COLORS.absent}`, color: COLORS.absent };

const iconBtn = {
  background: "transparent",
  border: "none",
  cursor: "pointer",
  color: COLORS.slate,
  padding: 4,
};

const pillBtn = {
  border: "1px solid",
  padding: "5px 10px",
  borderRadius: 20,
  fontSize: 11,
  fontWeight: 600,
  cursor: "pointer",
};
