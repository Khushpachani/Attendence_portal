import React, { useState, useMemo, useRef } from "react";
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

// The four ways a lecture can be marked. "Present" is the only status that
// counts toward attendance % — leave and proxy are recorded for the audit
// trail (who took leave, who was proxied) but don't grant attendance credit.
const STATUS_META = {
  present: { label: "Present", abbr: "P", color: "#3F6B4F", argb: "FFC6EFCE", font: "FF006100" },
  absent: { label: "Absent", abbr: "A", color: "#A34A3D", argb: "FFFFC7CE", font: "FF9C0006" },
  leave: { label: "On Leave", abbr: "L", color: "#C98A2C", argb: "FFFFE8B0", font: "FF9C5700" },
  proxy: { label: "Proxy", abbr: "PR", color: "#7B5EA7", argb: "FFE3D4F5", font: "FF5B3E96" },
};
const STATUS_ORDER = ["present", "absent", "leave", "proxy"];

// attendance color tiers: >=75 green, 60-74 orange, <60 red
function pctColor(pct) {
  if (pct >= 75) return COLORS.present;
  if (pct >= 60) return COLORS.warn;
  return COLORS.absent;
}

const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');`;

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
  { id: "25003111025081001", name: "Aachal Ashok Godse", email: "25mcsdf001@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081002", name: "Adit Sonone", email: "25mcsdf002@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081004", name: "Anuj Deepak Rathod", email: "25mcsdf004@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081005", name: "BHARGAVI RAMESH BHAI RATHVA", email: "25mcsdf005@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081006", name: "Dhruv Mukeshbhai Kachhadiya", email: "25mcsdf006@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111025081007", name: "DRASHTI VISHALKUMAR PATEL", email: "25mcsdf007@student.rru.ac.in", program: "MSc", electives: ["S9", "S6"] },
  { id: "25003111025081008", name: "Jaykumar Hareshbhai Patel", email: "25mcsdf008@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081009", name: "Kashish Choudhary", email: "25mcsdf009@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081010", name: "Kathiya Mayurdhvajsinh Kirtisinh", email: "25mcsdf010@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081011", name: "KHUSHI BHARAT PATEL", email: "25mcsdf011@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081013", name: "Manish Dipak Dhaygude", email: "25mcsdf013@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111025081014", name: "Naman Nilesh Gaur", email: "25mcsdf014@student.rru.ac.in", program: "MSc", electives: ["S9", "S6"] },
  { id: "25003111025081015", name: "Nishka Bijalkumar Bhatt", email: "25mcsdf015@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081016", name: "Parth Sanjaykumar Darji", email: "25mcsdf016@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081017", name: "PATEL AYUSH BAKULBHAI", email: "25mcsdf017@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081018", name: "Raj Karamata Lakhamanbhai", email: "25mcsdf018@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081019", name: "SMITKUMAR HARESHBHAI CHUADHARY", email: "25mcsdf019@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111025081020", name: "Souharda Patra", email: "25mcsdf020@student.rru.ac.in", program: "MSc", electives: ["S9", "S6"] },
  { id: "25003111025081021", name: "SUMIT Kumar Khuman Bhai Solanki", email: "25mcsdf021@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081023", name: "Taksh shah", email: "25mcsdf023@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081024", name: "USHVI NIRMALBHAI SHAH", email: "25mcsdf024@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081025", name: "VASOYA KESHVI RAJESHKUMAR", email: "25mcsdf025@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081026", name: "Adesh Dipeshbhai Teraiya", email: "25mcsdf026@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111025081027", name: "Ajay Gohali", email: "25mcsdf027@student.rru.ac.in", program: "MSc", electives: ["S9", "S6"] },
  { id: "25003111025081028", name: "Aksha Dharmesh Bhai Chudasama", email: "25mcsdf028@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081029", name: "AKSHITA BHAVSAR", email: "25mcsdf029@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081030", name: "Ashwani Kaushik", email: "25mcsdf030@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081031", name: "Garima", email: "25mcsdf031@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081032", name: "HARSH SANJAY SAWANT", email: "25mcsdf032@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111025081034", name: "Jay Jigneshbhai Mangroliya", email: "25mcsdf034@student.rru.ac.in", program: "MSc", electives: ["S9", "S6"] },
  { id: "25003111025081035", name: "Keval Jagdishbhai Vaghani", email: "25mcsdf035@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081036", name: "Keval Kinnarkumar Patel", email: "25mcsdf036@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081038", name: "Mitaliraj Ghosh", email: "25mcsdf038@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081039", name: "Mukta Bhausaheb Kodagapatil", email: "25mcsdf039@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081040", name: "Nitin Dinesh Bhai Prajapati", email: "25mcsdf040@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111025081041", name: "Parth Ghanshyambhai Jethva", email: "25mcsdf041@student.rru.ac.in", program: "MSc", electives: ["S9", "S6"] },
  { id: "25003111025081043", name: "Renuka Devi Palli", email: "25mcsdf043@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081044", name: "Riya Pradip Dalvi", email: "25mcsdf044@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081046", name: "Sourav Das", email: "25mcsdf046@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081047", name: "Swarnava Pal", email: "25mcsdf047@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081048", name: "Tiyasha Ray", email: "25mcsdf048@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111025081049", name: "Vrushti Paragkumar MEHTA", email: "25mcsdf049@student.rru.ac.in", program: "MSc", electives: ["S9", "S6"] },
  { id: "25003111025081050", name: "Vishvaraj Upendrasinh chauhan", email: "25mcsdf050@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003111025081051", name: "Aayush Arun Nishad", email: "25mcsdf051@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003111025081052", name: "Pachani Khush Bhavesbhai", email: "25mcsdf052@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003111025081053", name: "Pranav Prakash Nair", email: "25mcsdf053@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003111025081055", name: "Unnat Patel", email: "25mcsdf055@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111025081056", name: "Chaudhary suresh premaram", email: "25mcsdf056@student.rru.ac.in", program: "MSc", electives: ["S9", "S6"] },
  { id: "25003311025081001", name: "KATSANDE HUDSON WALTER", email: "25mcsdf001-f@student.rru.ac.in", program: "MSc", electives: ["S8", "S4"] },
  { id: "25003311025081002", name: "MUTUKU GRASIANO TERERAI", email: "25mcsdf002-f@student.rru.ac.in", program: "MSc", electives: ["S9", "S5"] },
  { id: "25003311025081003", name: "Mohamed Suma", email: "25mcsdf003-f@student.rru.ac.in", program: "MSc", electives: ["S8", "S6"] },
  { id: "25003311025081004", name: "Maxwell Chukwuekezie Onyebueke", email: "25mcsdf004-f@student.rru.ac.in", program: "MSc", electives: ["S9", "S4"] },
  { id: "25003311025081005", name: "Opeyemi Olaolu Ayegbo", email: "25mcsdf005-f@student.rru.ac.in", program: "MSc", electives: ["S8", "S5"] },
  { id: "25003111018071001", name: "Harshkumar Jagdishbhai Patel", email: "25003111018071001@student.rru.ac.in", program: "PGD", electives: ["S9", "S6"] },
  { id: "25003111018071003", name: "Radhika Utpal Yagnik", email: "25003111018071003@student.rru.ac.in", program: "PGD", electives: ["S8", "S4"] },
  { id: "25003111018071004", name: "Sanskruti Sandeep Patil", email: "25003111018071004@student.rru.ac.in", program: "PGD", electives: ["S9", "S5"] },
  { id: "25003111018071005", name: "SHIVAM DWIVEDI", email: "25003111018071005@student.rru.ac.in", program: "PGD", electives: ["S8", "S6"] },
  { id: "25003111018071006", name: "Abhishek Bharat Parmar", email: "25003111018071006@student.rru.ac.in", program: "PGD", electives: ["S9", "S4"] },
  { id: "25003111018071007", name: "Anshraj Navdeepbhai Dodiya", email: "25003111018071007@student.rru.ac.in", program: "PGD", electives: ["S8", "S5"] },
  { id: "25003111018071008", name: "Jesika khurshed Turel", email: "25003111018071008@student.rru.ac.in", program: "PGD", electives: ["S9", "S6"] },
  { id: "25003111018071009", name: "Krunal Pareshkumar Mewada", email: "25003111018071009@student.rru.ac.in", program: "PGD", electives: ["S8", "S4"] },
  { id: "25003111018071010", name: "Riya jagatbahi Sojitra", email: "25003111018071010@student.rru.ac.in", program: "PGD", electives: ["S9", "S5"] },
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

// Attendance holds ONE status per student per lecture. Time and remark
// belong to the *session* (subject + date) as a whole — e.g. "faculty on
// leave, proxy lecture taken by Mr. X" or "class rescheduled to lab" — not
// to any individual student, so they live in a separate `sessions` map.
function seedAttendance(dates) {
  const rec = {};
  seedStudents.forEach((st) => {
    const subs = [...seedSubjects.filter((s) => s.type === "core"), ...seedSubjects.filter((s) => st.electives.includes(s.id))];
    subs.forEach((sub) => {
      dates.forEach((date) => {
        const key = `${st.id}__${sub.id}__${date}`;
        const roll = Math.random();
        let status = "present";
        if (roll < 0.06) status = "leave";
        else if (roll < 0.1) status = "proxy";
        else if (roll < 0.2) status = "absent";
        rec[key] = status;
      });
    });
  });
  return rec;
}

const SESSION_REMARKS = [
  "Faculty on leave — proxy lecture conducted",
  "Class rescheduled to lab session",
  "Guest lecture in place of regular class",
  "College event — attendance affected",
];

function seedSessions(dates) {
  const rec = {};
  seedSubjects.forEach((sub) => {
    dates.forEach((date, i) => {
      const key = `${sub.id}__${date}`;
      // sprinkle in a handful of session-level remarks for demo purposes
      const remark = i % 7 === 0 ? SESSION_REMARKS[(sub.id.charCodeAt(1) + i) % SESSION_REMARKS.length] : "";
      rec[key] = { time: sub.defaultTime || "", remark };
    });
  });
  return rec;
}

const TABS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutGrid },
  { id: "students", label: "Students", icon: Users },
  { id: "subjects", label: "Subjects", icon: BookOpen },
  { id: "mark", label: "Mark Attendance", icon: CalendarCheck },
  { id: "all", label: "All Attendance", icon: Table2 },
  { id: "reports", label: "Reports", icon: FileBarChart },
];

// All data now lives in a shared database (via /api/state, backed by
// Upstash Redis) so every faculty member sees the same students, subjects,
// and attendance regardless of which browser/device they're on.
const SAVE_DEBOUNCE_MS = 900;
const POLL_INTERVAL_MS = 20000;

export default function AttendancePortal() {
  const [subjects, setSubjects] = useState(seedSubjects);
  const [students, setStudents] = useState(seedStudents);
  const [attendance, setAttendance] = useState({});
  const [sessions, setSessions] = useState({});
  const [tab, setTab] = useState("dashboard");
  const [loading, setLoading] = useState(true);
  const [syncState, setSyncState] = useState("idle"); // idle | saving | saved | error
  const idCounter = useRef(100);
  const loadedRef = useRef(false);
  const saveTimer = useRef(null);

  const applyServerData = (data) => {
    if (data) {
      setSubjects(data.subjects && data.subjects.length ? data.subjects : seedSubjects);
      setStudents(data.students && data.students.length ? data.students : seedStudents);
      setAttendance(data.attendance || {});
      setSessions(data.sessions || {});
    } else {
      // nothing saved on the server yet — seed it with sample data
      const dates = pastDates(10);
      setAttendance(seedAttendance(dates));
      setSessions(seedSessions(dates));
    }
  };

  const loadFromServer = async () => {
    try {
      const res = await fetch("/api/state");
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      applyServerData(data);
      setSyncState("saved");
    } catch (err) {
      console.error("Failed to load attendance data:", err);
      applyServerData(null);
      setSyncState("error");
    } finally {
      loadedRef.current = true;
      setLoading(false);
    }
  };

  // initial load
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
    setSyncState("saving");
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        const res = await fetch("/api/state", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ subjects, students, attendance, sessions }),
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        setSyncState("saved");
      } catch (err) {
        console.error("Failed to save attendance data:", err);
        setSyncState("error");
      }
    }, SAVE_DEBOUNCE_MS);
    return () => clearTimeout(saveTimer.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [subjects, students, attendance, sessions]);

  const resetToSampleData = async () => {
    if (!window.confirm("This clears everything for EVERYONE using this app and restores the original sample data. Continue?")) return;
    setSubjects(seedSubjects);
    setStudents(seedStudents);
    const dates = pastDates(10);
    setAttendance(seedAttendance(dates));
    setSessions(seedSessions(dates));
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

  // IMPORTANT: this must come after every hook above (useState/useMemo/etc.)
  // so the same hooks run in the same order on every render — an early
  // return placed before a hook call causes "Rendered more hooks than
  // during the previous render" once `loading` flips from true to false.
  if (loading) {
    return (
      <div
        style={{
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          background: COLORS.parchment,
          fontFamily: "Inter, sans-serif",
          color: COLORS.ink,
        }}
      >
        <style>{FONT_IMPORT}</style>
        Loading attendance data…
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
      // migrate any existing attendance records so they still point at this student
      setAttendance((prev) => {
        const next = {};
        Object.entries(prev).forEach(([key, val]) => {
          const [sid, subId, date] = key.split("__");
          const finalSid = sid === oldId ? newId : sid;
          next[`${finalSid}__${subId}__${date}`] = val;
        });
        return next;
      });
    }
  };
  const removeStudent = (id) => {
    setStudents((prev) => prev.filter((s) => s.id !== id));
  };
  const addSubject = (data) => {
    setSubjects((prev) => [...prev, { ...data, id: `S${idCounter.current++}` }]);
  };
  const removeSubject = (id) => {
    setSubjects((prev) => prev.filter((s) => s.id !== id));
  };
  // Per-student status (present/absent/leave/proxy) for one lecture.
  const setMark = (studentId, subjectId, date, status) => {
    const key = `${studentId}__${subjectId}__${date}`;
    setAttendance((prev) => ({ ...prev, [key]: status }));
  };
  const deleteMark = (studentId, subjectId, date) => {
    const key = `${studentId}__${subjectId}__${date}`;
    setAttendance((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
  };
  // Session-level time + remark for one lecture (subject + date) — shared by
  // every student in that lecture, e.g. "faculty on leave, proxy conducted".
  const getSession = (subjectId, date) => sessions[`${subjectId}__${date}`] || { time: "", remark: "" };
  const setSession = (subjectId, date, patch) => {
    const key = `${subjectId}__${date}`;
    setSessions((prev) => ({
      ...prev,
      [key]: { time: "", remark: "", ...prev[key], ...patch },
    }));
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

  // One detailed sheet per subject: dates as columns (Status per student +
  // one Remark per date, merged down the column since the remark describes
  // that lecture session — not any individual student).
  const buildSubjectSheet = (wb, subject, usedNames) => {
    const roster = studentsFor(subject);
    const dateSet = new Set();
    Object.keys(attendance).forEach((key) => {
      const [sid, subId, date] = key.split("__");
      if (subId === subject.id && roster.some((r) => r.id === sid)) dateSet.add(date);
    });
    const dates = Array.from(dateSet).sort();

    let baseName = subject.code.replace(/[\\/*?:[\]]/g, "-").slice(0, 28) || subject.id;
    let sheetName = baseName;
    let n = 2;
    while (usedNames.has(sheetName)) sheetName = `${baseName}-${n++}`.slice(0, 31);
    usedNames.add(sheetName);

    const ws = wb.addWorksheet(sheetName);
    const LEFT = ["Sr No", "Enrollment No", "Name", "Program"];
    const RIGHT = ["Classes Held", "Present", "Attendance %"];
    const totalCols = LEFT.length + dates.length * 2 + RIGHT.length;

    ws.getColumn(1).width = 7;
    ws.getColumn(2).width = 16;
    ws.getColumn(3).width = 24;
    ws.getColumn(4).width = 10;
    dates.forEach((d, i) => {
      const base = LEFT.length + i * 2;
      ws.getColumn(base + 1).width = 9;
      ws.getColumn(base + 2).width = 24;
    });
    RIGHT.forEach((_, i) => {
      ws.getColumn(LEFT.length + dates.length * 2 + i + 1).width = 12;
    });

    // Row 1 — title
    ws.mergeCells(1, 1, 1, Math.max(totalCols, 5));
    const title = ws.getCell(1, 1);
    title.value = `${subject.code} · ${subject.name} — Attendance Register`;
    title.font = { bold: true, size: 13, color: { argb: "FFF7F4EA" } };
    title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1B2A4A" } };
    title.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(1).height = 22;

    // Rows 2-3 — headers (fixed columns span both rows, each date spans 2 columns)
    LEFT.forEach((label, i) => {
      ws.mergeCells(2, i + 1, 3, i + 1);
      const cell = ws.getCell(2, i + 1);
      cell.value = label;
      styleHeaderCell(cell);
    });
    dates.forEach((d, i) => {
      const base = LEFT.length + i * 2;
      const session = sessions[`${subject.id}__${d}`] || { time: "", remark: "" };
      ws.mergeCells(2, base + 1, 2, base + 2);
      const dateCell = ws.getCell(2, base + 1);
      const dObj = new Date(d);
      const dateLabel = dObj.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
      dateCell.value = session.time ? `${dateLabel}\n${session.time}` : dateLabel;
      styleHeaderCell(dateCell);
      styleHeaderCell(Object.assign(ws.getCell(3, base + 1), { value: "Status" }));
      styleHeaderCell(Object.assign(ws.getCell(3, base + 2), { value: "Remark" }), "left");
    });
    RIGHT.forEach((label, i) => {
      const col = LEFT.length + dates.length * 2 + i + 1;
      ws.mergeCells(2, col, 3, col);
      styleHeaderCell(Object.assign(ws.getCell(2, col), { value: label }));
    });
    ws.getRow(2).height = 30;
    ws.getRow(3).height = 16;

    // Data rows — Status is per student; Remark is per date (merged below).
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
      dates.forEach((d, i) => {
        const base = LEFT.length + i * 2;
        const status = attendance[`${st.id}__${subject.id}__${d}`];
        const statusCell = ws.getCell(r, base + 1);
        if (status) {
          total++;
          if (status === "present") present++;
          const meta = STATUS_META[status];
          statusCell.value = meta.abbr;
          statusCell.alignment = { horizontal: "center" };
          statusCell.font = { bold: true, color: { argb: meta.font } };
          statusCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: meta.argb } };
        } else {
          statusCell.value = "—";
          statusCell.alignment = { horizontal: "center" };
          statusCell.font = { color: { argb: "FFB0B0B0" } };
        }
      });

      const pct = total ? Math.round((present / total) * 100) : 0;
      const rightBase = LEFT.length + dates.length * 2;
      ws.getCell(r, rightBase + 1).value = total;
      ws.getCell(r, rightBase + 1).alignment = { horizontal: "center" };
      ws.getCell(r, rightBase + 2).value = present;
      ws.getCell(r, rightBase + 2).alignment = { horizontal: "center" };
      colorPctCell(ws.getCell(r, rightBase + 3), pct);
    });

    // One remark per date, merged down the whole student column for that
    // date — makes it visually obvious the note applies to the session,
    // not to any single student.
    if (roster.length > 0) {
      const lastRow = 4 + roster.length - 1;
      dates.forEach((d, i) => {
        const base = LEFT.length + i * 2;
        const session = sessions[`${subject.id}__${d}`] || { time: "", remark: "" };
        ws.mergeCells(4, base + 2, lastRow, base + 2);
        const cell = ws.getCell(4, base + 2);
        cell.value = session.remark || "";
        cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
        cell.font = { italic: true, size: 10, color: { argb: "FF6B7280" } };
        if (session.remark) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFBF6E8" } };
        }
      });
    }

    // Legend
    const legendRow = 4 + roster.length + 1;
    ws.getCell(legendRow, 1).value = "Legend:";
    ws.getCell(legendRow, 1).font = { bold: true };
    STATUS_ORDER.forEach((s, i) => {
      const meta = STATUS_META[s];
      const cell = ws.getCell(legendRow, 2 + i);
      cell.value = `${meta.abbr} = ${meta.label}`;
      cell.font = { bold: true, color: { argb: meta.font } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: meta.argb } };
    });
    ws.getCell(legendRow + 1, 1).value = "Note: the Remark column is per lecture date, not per student.";
    ws.getCell(legendRow + 1, 1).font = { italic: true, size: 10, color: { argb: "FF6B7280" } };

    ws.views = [{ state: "frozen", xSplit: 4, ySplit: 3 }];
  };

  const exportExcel = async () => {
    const wb = new ExcelJS.Workbook();
    wb.creator = "Attendance Register";
    wb.created = new Date();

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
    students.forEach((s, i) =>
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

    // ---- One detailed date-grid sheet per subject ----
    const usedNames = new Set(["Student Detail"]);
    subjects.forEach((sub) => buildSubjectSheet(wb, sub, usedNames));

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
        const [sid, subId, date] = key.split("__");
        const st = students.find((s) => s.id === sid);
        const sub = subjects.find((s) => s.id === subId);
        const session = sessions[`${subId}__${date}`] || { time: "", remark: "" };
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
      <style>{FONT_IMPORT}</style>
      <div style={{ minHeight: "100vh", display: "flex" }}>
        {/* Sidebar */}
        <aside style={{ width: 220, background: COLORS.ink, color: COLORS.parchment, flexShrink: 0, display: "flex", flexDirection: "column" }}>
          <div style={{ padding: "24px 20px", borderBottom: `1px solid ${COLORS.inkSoft}` }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <Stamp size={22} color={COLORS.brass} />
              <span style={{ fontFamily: "Fraunces, serif", fontWeight: 700, fontSize: 18 }}>Attendance Register</span>
            </div>
            <div style={{ fontSize: 11, color: COLORS.brassSoft, marginTop: 6, letterSpacing: 1, textTransform: "uppercase" }}>
              MSc &amp; PGD · Batch 2024
            </div>
          </div>
          <nav style={{ padding: "12px 10px", flex: 1 }}>
            {TABS.map((t) => {
              const Icon = t.icon;
              const active = tab === t.id;
              return (
                <button
                  key={t.id}
                  onClick={() => setTab(t.id)}
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
            <button
              onClick={exportExcel}
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
        <main style={{ flex: 1, padding: "28px 36px", overflow: "auto" }}>
          {tab === "dashboard" && <Dashboard subjectStats={subjectStats} students={students} />}
          {tab === "students" && (
            <StudentsTab
              students={students}
              subjects={subjects}
              addStudent={addStudent}
              editStudent={editStudent}
              removeStudent={removeStudent}
            />
          )}
          {tab === "subjects" && (
            <SubjectsTab subjects={subjects} addSubject={addSubject} removeSubject={removeSubject} studentsFor={studentsFor} />
          )}
          {tab === "mark" && (
            <MarkTab
              subjects={subjects}
              studentsFor={studentsFor}
              attendance={attendance}
              setMark={setMark}
              getSession={getSession}
              setSession={setSession}
            />
          )}
          {tab === "all" && (
            <AllAttendanceTab
              students={students}
              subjects={subjects}
              attendance={attendance}
              sessions={sessions}
              setMark={setMark}
              setSession={setSession}
              deleteMark={deleteMark}
            />
          )}
          {tab === "reports" && <ReportsTab studentStats={studentStats} />}
        </main>
      </div>
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

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16, marginBottom: 30 }}>
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
                  style={{
                    width: `${(s.pct / 100) * 100}%`,
                    background: pctColor(s.pct),
                    height: "100%",
                    borderRadius: 6,
                    transition: "width .3s",
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
function StudentsTab({ students, subjects, addStudent, editStudent, removeStudent }) {
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null); // null = adding a new student
  const emptyForm = { name: "", email: "", id: "", program: "MSc", electiveByGroup: { G1: "", G2: "" } };
  const [form, setForm] = useState(emptyForm);
  const [search, setSearch] = useState("");

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

  const coreCount = subjects.filter((s) => s.type === "core").length;

  return (
    <div>
      <SectionTitle
        eyebrow="Batch 2024"
        title="Students"
        action={
          <button onClick={showForm && !editingId ? cancelForm : startAdd} style={btnPrimary}>
            <Plus size={15} /> Add Student
          </button>
        }
      />

      {showForm && (
        <form onSubmit={submit} style={cardStyle}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <div style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 15 }}>
              {editingId ? "Edit Student" : "New Student"}
            </div>
            <button type="button" onClick={cancelForm} style={iconBtn}>
              <X size={16} />
            </button>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
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
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 16 }}>
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

      <div style={{ background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 12, overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: COLORS.parchmentDark, textAlign: "left" }}>
              <th style={thStyle}>Enrollment No.</th>
              <th style={thStyle}>Name</th>
              <th style={thStyle}>Email</th>
              <th style={thStyle}>Program</th>
              <th style={thStyle}>Electives</th>
              <th style={thStyle}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((s) => (
              <tr key={s.id} style={{ borderTop: `1px solid ${COLORS.line}`, background: editingId === s.id ? COLORS.parchmentDark : "transparent" }}>
                <td style={{ ...tdStyle, fontFamily: "JetBrains Mono, monospace" }}>{s.id}</td>
                <td style={tdStyle}>{s.name}</td>
                <td style={{ ...tdStyle, color: COLORS.slate }}>{s.email}</td>
                <td style={tdStyle}>
                  <Badge>{s.program}</Badge>
                </td>
                <td style={tdStyle}>
                  {subjects.filter((sub) => s.electives.includes(sub.id)).map((sub) => sub.name).join(", ") || "—"}
                </td>
                <td style={tdStyle}>
                  <div style={{ display: "flex", gap: 4 }}>
                    <button onClick={() => startEdit(s)} style={iconBtn} title="Edit student">
                      <Pencil size={14} />
                    </button>
                    <button onClick={() => removeStudent(s.id)} style={iconBtn} title="Remove student">
                      <X size={14} />
                    </button>
                  </div>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr>
                <td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: COLORS.slate, padding: 24 }}>
                  No students found.
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
function SubjectsTab({ subjects, addSubject, removeSubject, studentsFor }) {
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ code: "", name: "", type: "core", program: "Both", group: "G1" });

  const submit = (e) => {
    e.preventDefault();
    if (!form.code || !form.name) return;
    addSubject(form.type === "core" ? { ...form, group: null } : form);
    setForm({ code: "", name: "", type: "core", program: "Both", group: "G1" });
    setShowForm(false);
  };

  const core = subjects.filter((s) => s.type === "core");
  const groups = ELECTIVE_GROUPS.map((g) => ({
    ...g,
    subjects: subjects.filter((s) => s.type === "elective" && s.group === g.id),
  }));

  const SubjectCard = (s) => (
    <div key={s.id} style={{ ...cardStyle, margin: 0, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
      <div>
        <div style={{ fontFamily: "JetBrains Mono, monospace", fontSize: 11, color: COLORS.brass, marginBottom: 4 }}>{s.code}</div>
        <div style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 16, marginBottom: 6 }}>{s.name}</div>
        <div style={{ display: "flex", gap: 6 }}>
          <Badge tone={s.type === "core" ? "ink" : "brass"}>{s.type === "core" ? "Core" : "Elective"}</Badge>
          <Badge>{s.program}</Badge>
          <Badge tone="slate">{studentsFor(s).length} students</Badge>
        </div>
      </div>
      <button onClick={() => removeSubject(s.id)} style={iconBtn}>
        <X size={14} />
      </button>
    </div>
  );

  return (
    <div>
      <SectionTitle
        eyebrow="Curriculum · 9 subjects total"
        title="Subjects"
        action={
          <button onClick={() => setShowForm((v) => !v)} style={btnPrimary}>
            <Plus size={15} /> Add Subject
          </button>
        }
      />

      {showForm && (
        <form onSubmit={submit} style={cardStyle}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr 1fr 1fr 1fr", gap: 12, marginBottom: 16 }}>
            <input required placeholder="Code e.g. EL-401" value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} style={inputStyle} />
            <input required placeholder="Subject name" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} style={inputStyle} />
            <select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })} style={inputStyle}>
              <option value="core">Core</option>
              <option value="elective">Elective</option>
            </select>
            {form.type === "elective" && (
              <select value={form.group} onChange={(e) => setForm({ ...form, group: e.target.value })} style={inputStyle}>
                {ELECTIVE_GROUPS.map((g) => (
                  <option key={g.id} value={g.id}>
                    {g.label}
                  </option>
                ))}
              </select>
            )}
            <select value={form.program} onChange={(e) => setForm({ ...form, program: e.target.value })} style={inputStyle}>
              <option value="Both">Both (MSc + PGD)</option>
              <option value="MSc">MSc only</option>
              <option value="PGD">PGD only</option>
            </select>
          </div>
          <button type="submit" style={btnPrimary}>
            <Check size={15} /> Save Subject
          </button>
        </form>
      )}

      <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.slate, marginBottom: 10, textTransform: "uppercase", letterSpacing: 1 }}>
        Core subjects ({core.length}) · compulsory for everyone
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 14, marginBottom: 28 }}>
        {core.map(SubjectCard)}
      </div>

      {groups.map((g) => (
        <div key={g.id} style={{ marginBottom: 28 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.slate, marginBottom: 10, textTransform: "uppercase", letterSpacing: 1 }}>
            {g.label} ({g.subjects.length}) · {g.pick}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 14 }}>{g.subjects.map(SubjectCard)}</div>
        </div>
      ))}
    </div>
  );
}

// ============================================================
function MarkTab({ subjects, studentsFor, attendance, setMark, getSession, setSession }) {
  const [subjectId, setSubjectId] = useState(subjects[0]?.id);
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const subject = subjects.find((s) => s.id === subjectId);
  const roster = subject ? studentsFor(subject) : [];
  const session = getSession(subjectId, date);
  const time = session.time || subject?.defaultTime || "";

  const markAll = (status) => {
    roster.forEach((st) => setMark(st.id, subjectId, date, status));
  };

  return (
    <div>
      <SectionTitle eyebrow="Daily Roll Call" title="Mark Attendance" />

      <div style={{ display: "flex", gap: 12, marginBottom: 14, alignItems: "center", flexWrap: "wrap" }}>
        <select value={subjectId} onChange={(e) => setSubjectId(e.target.value)} style={{ ...inputStyle, maxWidth: 280 }}>
          {subjects.map((s) => (
            <option key={s.id} value={s.id}>
              {s.code} · {s.name}
            </option>
          ))}
        </select>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={{ ...inputStyle, maxWidth: 170 }} />
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <Clock size={14} color={COLORS.slate} />
          <input
            placeholder="e.g. 09:15 - 10:00"
            value={time}
            onChange={(e) => setSession(subjectId, date, { time: e.target.value })}
            style={{ ...inputStyle, maxWidth: 160 }}
          />
        </div>
        <div style={{ flex: 1 }} />
        <button onClick={() => markAll("present")} style={btnGhost}>
          Mark all present
        </button>
        <button onClick={() => markAll("absent")} style={btnGhostRed}>
          Mark all absent
        </button>
      </div>

      <div style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6 }}>
          Remark for this lecture <span style={{ color: COLORS.slate, fontWeight: 400 }}>(applies to the whole session, not one student — e.g. "faculty on leave, proxy conducted")</span>
        </div>
        <input
          placeholder="Optional note about this date's lecture…"
          value={session.remark}
          onChange={(e) => setSession(subjectId, date, { remark: e.target.value })}
          style={inputStyle}
        />
      </div>

      <div style={{ background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 12, overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: COLORS.parchmentDark, textAlign: "left" }}>
              <th style={thStyle}>Enrollment No.</th>
              <th style={thStyle}>Name</th>
              <th style={{ ...thStyle, textAlign: "center" }}>Status</th>
            </tr>
          </thead>
          <tbody>
            {roster.map((st) => {
              const key = `${st.id}__${subjectId}__${date}`;
              const status = attendance[key] || "present";
              return (
                <tr key={st.id} style={{ borderTop: `1px solid ${COLORS.line}` }}>
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
                            onClick={() => setMark(st.id, subjectId, date, s)}
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
                    </div>
                  </td>
                </tr>
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
function AllAttendanceTab({ students, subjects, attendance, sessions, setMark, setSession, deleteMark }) {
  const [subjectFilter, setSubjectFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const [dateFilter, setDateFilter] = useState("all");
  const [search, setSearch] = useState("");

  const availableDates = useMemo(() => {
    const dates = new Set();
    Object.keys(attendance).forEach((key) => dates.add(key.split("__")[2]));
    return Array.from(dates).sort((a, b) => (a < b ? 1 : -1)); // newest first
  }, [attendance]);

  const rows = useMemo(() => {
    return Object.entries(attendance)
      .map(([key, status]) => {
        const [sid, subId, date] = key.split("__");
        const student = students.find((s) => s.id === sid);
        const subject = subjects.find((s) => s.id === subId);
        const session = sessions[`${subId}__${date}`] || { time: "", remark: "" };
        return { key, sid, subId, date, student, subject, status, time: session.time, remark: session.remark };
      })
      .filter((r) => r.student && r.subject)
      .filter((r) => subjectFilter === "all" || r.subId === subjectFilter)
      .filter((r) => statusFilter === "all" || r.status === statusFilter)
      .filter((r) => dateFilter === "all" || r.date === dateFilter)
      .filter(
        (r) =>
          !search ||
          r.student.name.toLowerCase().includes(search.toLowerCase()) ||
          r.student.id.toLowerCase().includes(search.toLowerCase())
      )
      .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : a.student.name.localeCompare(b.student.name)));
  }, [attendance, sessions, students, subjects, subjectFilter, statusFilter, dateFilter, search]);

  const formatDate = (d) => {
    const dObj = new Date(d);
    return dObj.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
  };

  return (
    <div>
      <SectionTitle eyebrow={`${rows.length} records`} title="All Attendance" />
      <div style={{ fontSize: 12, color: COLORS.slate, marginBottom: 14, marginTop: -10 }}>
        Time and Remark apply to the whole lecture (subject + date) — editing either one updates it for every student in that session.
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
        <select value={dateFilter} onChange={(e) => setDateFilter(e.target.value)} style={{ ...inputStyle, maxWidth: 200 }}>
          <option value="all">All dates</option>
          {availableDates.map((d) => (
            <option key={d} value={d}>
              {formatDate(d)}
            </option>
          ))}
        </select>
        <select value={subjectFilter} onChange={(e) => setSubjectFilter(e.target.value)} style={{ ...inputStyle, maxWidth: 240 }}>
          <option value="all">All subjects</option>
          {subjects.map((s) => (
            <option key={s.id} value={s.id}>
              {s.code} · {s.name}
            </option>
          ))}
        </select>
        <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} style={{ ...inputStyle, maxWidth: 160 }}>
          <option value="all">All statuses</option>
          {STATUS_ORDER.map((s) => (
            <option key={s} value={s}>
              {STATUS_META[s].label}
            </option>
          ))}
        </select>
        {(dateFilter !== "all" || subjectFilter !== "all" || statusFilter !== "all" || search) && (
          <button
            onClick={() => {
              setDateFilter("all");
              setSubjectFilter("all");
              setStatusFilter("all");
              setSearch("");
            }}
            style={{ ...btnGhost, borderColor: COLORS.line, color: COLORS.slate }}
          >
            Clear filters
          </button>
        )}
      </div>

      <div style={{ background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 12, overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: COLORS.parchmentDark, textAlign: "left" }}>
              <th style={thStyle}>Date</th>
              <th style={thStyle}>Enrollment No.</th>
              <th style={thStyle}>Name</th>
              <th style={thStyle}>Subject</th>
              <th style={thStyle}>Time (session)</th>
              <th style={thStyle}>Status (student)</th>
              <th style={thStyle}>Remark (session)</th>
              <th style={thStyle}></th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.key} style={{ borderTop: `1px solid ${COLORS.line}` }}>
                <td style={{ ...tdStyle, fontFamily: "JetBrains Mono, monospace", whiteSpace: "nowrap" }}>{r.date}</td>
                <td style={{ ...tdStyle, fontFamily: "JetBrains Mono, monospace" }}>{r.sid}</td>
                <td style={tdStyle}>{r.student.name}</td>
                <td style={tdStyle}>
                  {r.subject.code}
                  <div style={{ fontSize: 11, color: COLORS.slate }}>{r.subject.name}</div>
                </td>
                <td style={tdStyle}>
                  <input
                    value={r.time}
                    onChange={(e) => setSession(r.subId, r.date, { time: e.target.value })}
                    style={{ ...inputStyle, fontSize: 12, width: 120 }}
                  />
                </td>
                <td style={tdStyle}>
                  <select
                    value={r.status}
                    onChange={(e) => setMark(r.sid, r.subId, r.date, e.target.value)}
                    style={{
                      ...inputStyle,
                      width: 118,
                      fontSize: 12,
                      fontWeight: 600,
                      color: STATUS_META[r.status].color,
                      borderColor: STATUS_META[r.status].color,
                    }}
                  >
                    {STATUS_ORDER.map((s) => (
                      <option key={s} value={s}>
                        {STATUS_META[s].label}
                      </option>
                    ))}
                  </select>
                </td>
                <td style={tdStyle}>
                  <input
                    value={r.remark}
                    placeholder="Add a note about this lecture…"
                    onChange={(e) => setSession(r.subId, r.date, { remark: e.target.value })}
                    style={{ ...inputStyle, fontSize: 12, minWidth: 200 }}
                  />
                </td>
                <td style={tdStyle}>
                  <button onClick={() => deleteMark(r.sid, r.subId, r.date)} style={iconBtn} title="Delete this student's record">
                    <X size={14} />
                  </button>
                </td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr>
                <td colSpan={8} style={{ ...tdStyle, textAlign: "center", color: COLORS.slate, padding: 24 }}>
                  No matching attendance records.
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
      <div style={{ display: "grid", gridTemplateColumns: "300px 1fr", gap: 18 }}>
        <div style={{ background: "#fff", border: `1px solid ${COLORS.line}`, borderRadius: 12, overflow: "hidden", alignSelf: "start" }}>
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
