// Vercel serverless function — shared storage for all faculty using the app.
// Backed by Upstash Redis (installed from the Vercel Marketplace as
// "Upstash for Redis"), which injects KV_REST_API_URL / KV_REST_API_TOKEN
// (or UPSTASH_REDIS_REST_URL / UPSTASH_REDIS_REST_TOKEN, depending on how
// it was connected) into this project's environment variables.
import { Redis } from "@upstash/redis";

const redis = new Redis({
  url: process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL,
  token: process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN,
});

// Each piece of data lives in its own Redis HASH, keyed by record id.
// This is the important part for multi-user correctness: HSET on a hash
// field is a single atomic Redis command. Two faculty saving at the exact
// same moment, editing DIFFERENT records, now literally cannot conflict —
// there is no "read the whole thing, merge in memory, write the whole
// thing back" step for Redis to race on. That in-between step is what
// caused records to occasionally go missing under concurrent use, even
// with the previous merge-on-read approach. Only two people editing the
// very same field at the very same instant still resolve as last-write-wins
// for that one field, which is normal and expected.
const H = {
  subjects: "attendance-portal:h:subjects", // field = subject.id,  value = subject object
  students: "attendance-portal:h:students", // field = student.id,  value = student object
  attendance: "attendance-portal:h:attendance", // field = "studentId__subjectId__date__slot", value = "present" | "absent"
  sessions: "attendance-portal:h:sessions", // field = "subjectId__date__slot", value = {time, remark}
  attendanceMeta: "attendance-portal:h:attendanceMeta", // field = same as attendance, value = {editedBy, editedAt}
};

// @upstash/redis auto-serializes on write and auto-deserializes on read for
// most values, but be defensive in case a value ever comes back as a raw
// JSON string (depends on SDK version / how it was originally written).
function coerce(value) {
  if (typeof value !== "string") return value;
  try {
    return JSON.parse(value);
  } catch {
    return value; // plain string like "present"/"absent" — not JSON, that's fine
  }
}

async function readAll() {
  const [subjectsRaw, studentsRaw, attendanceRaw, sessionsRaw, metaRaw] = await Promise.all([
    redis.hgetall(H.subjects),
    redis.hgetall(H.students),
    redis.hgetall(H.attendance),
    redis.hgetall(H.sessions),
    redis.hgetall(H.attendanceMeta),
  ]);
  const toObj = (raw) => {
    const out = {};
    Object.entries(raw || {}).forEach(([k, v]) => (out[k] = coerce(v)));
    return out;
  };
  const subjectsObj = toObj(subjectsRaw);
  const studentsObj = toObj(studentsRaw);
  const hasAny = [subjectsRaw, studentsRaw, attendanceRaw, sessionsRaw, metaRaw].some((r) => r && Object.keys(r).length > 0);
  if (!hasAny) return null; // signal "nothing saved yet" so the client seeds sample data
  return {
    subjects: Object.values(subjectsObj),
    students: Object.values(studentsObj),
    attendance: toObj(attendanceRaw),
    sessions: toObj(sessionsRaw),
    attendanceMeta: toObj(metaRaw),
  };
}

async function writeAll(body) {
  const ops = [];

  if (Array.isArray(body.subjects) && body.subjects.length > 0) {
    const fields = {};
    body.subjects.forEach((s) => s && s.id && (fields[s.id] = s));
    if (Object.keys(fields).length) ops.push(redis.hset(H.subjects, fields));
  }
  if (Array.isArray(body.students) && body.students.length > 0) {
    const fields = {};
    body.students.forEach((s) => s && s.id && (fields[s.id] = s));
    if (Object.keys(fields).length) ops.push(redis.hset(H.students, fields));
  }
  if (body.attendance && Object.keys(body.attendance).length > 0) {
    ops.push(redis.hset(H.attendance, body.attendance));
  }
  if (body.sessions && Object.keys(body.sessions).length > 0) {
    ops.push(redis.hset(H.sessions, body.sessions));
  }
  if (body.attendanceMeta && Object.keys(body.attendanceMeta).length > 0) {
    ops.push(redis.hset(H.attendanceMeta, body.attendanceMeta));
  }

  if (Array.isArray(body.deletedSubjectIds) && body.deletedSubjectIds.length) {
    ops.push(redis.hdel(H.subjects, ...body.deletedSubjectIds));
  }
  if (Array.isArray(body.deletedStudentIds) && body.deletedStudentIds.length) {
    ops.push(redis.hdel(H.students, ...body.deletedStudentIds));
  }
  if (Array.isArray(body.deletedAttendanceKeys) && body.deletedAttendanceKeys.length) {
    ops.push(redis.hdel(H.attendance, ...body.deletedAttendanceKeys));
    ops.push(redis.hdel(H.attendanceMeta, ...body.deletedAttendanceKeys));
  }
  if (Array.isArray(body.deletedSessionKeys) && body.deletedSessionKeys.length) {
    ops.push(redis.hdel(H.sessions, ...body.deletedSessionKeys));
  }

  await Promise.all(ops);
}

export default async function handler(req, res) {
  if (!process.env.KV_REST_API_URL && !process.env.UPSTASH_REDIS_REST_URL) {
    res.status(500).json({
      error: "No database connected. In the Vercel dashboard, add the 'Upstash for Redis' integration from the Marketplace, connect it to this project, then redeploy.",
    });
    return;
  }

  try {
    if (req.method === "GET") {
      const data = await readAll();
      res.status(200).json(data);
      return;
    }

    if (req.method === "PUT" || req.method === "POST") {
      const body = req.body && typeof req.body === "object" ? req.body : JSON.parse(req.body || "{}");
      if (!body || typeof body !== "object") {
        res.status(400).json({ error: "Request body must be a JSON object." });
        return;
      }
      await writeAll(body);
      res.status(200).json({ ok: true, savedAt: new Date().toISOString() });
      return;
    }

    if (req.method === "DELETE") {
      await redis.del(H.subjects, H.students, H.attendance, H.sessions, H.attendanceMeta);
      res.status(200).json({ ok: true });
      return;
    }

    res.setHeader("Allow", "GET, PUT, DELETE");
    res.status(405).json({ error: "Method not allowed" });
  } catch (err) {
    console.error("attendance-portal state API error:", err);
    res.status(500).json({ error: "Server error", detail: String(err) });
  }
}
