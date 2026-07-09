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

const STATE_KEY = "attendance-portal:state";

// Merge a client's local snapshot into whatever is currently stored, rather
// than blindly overwriting it. This matters because multiple faculty can
// have the app open at once: if two people save within moments of each
// other using a plain overwrite, whoever saves second wipes out whatever
// the first person just added. Instead:
//   - attendance / sessions / attendanceMeta are merged key-by-key (a union
//     of what's already stored and what this client knows), so an addition
//     from one client can never silently erase an addition from another.
//   - explicit deletions are still respected: the client tells us exactly
//     which keys it removed since its last successful save, and those are
//     removed from the merged result even though a plain union would have
//     kept them.
//   - subjects / students are small, infrequently-edited lists, so those
//     are simply replaced with the client's copy when provided.
function mergeState(existing, incoming) {
  const base = existing && typeof existing === "object" ? existing : {};
  const merged = {
    subjects: incoming.subjects !== undefined ? incoming.subjects : base.subjects || [],
    students: incoming.students !== undefined ? incoming.students : base.students || [],
    attendance: { ...(base.attendance || {}), ...(incoming.attendance || {}) },
    sessions: { ...(base.sessions || {}), ...(incoming.sessions || {}) },
    attendanceMeta: { ...(base.attendanceMeta || {}), ...(incoming.attendanceMeta || {}) },
  };
  (incoming.deletedAttendanceKeys || []).forEach((k) => {
    delete merged.attendance[k];
    delete merged.attendanceMeta[k];
  });
  (incoming.deletedSessionKeys || []).forEach((k) => {
    delete merged.sessions[k];
  });
  return merged;
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
      const data = await redis.get(STATE_KEY);
      res.status(200).json(data || null);
      return;
    }

    if (req.method === "PUT" || req.method === "POST") {
      const body = req.body && typeof req.body === "object" ? req.body : JSON.parse(req.body || "{}");
      if (!body || typeof body !== "object") {
        res.status(400).json({ error: "Request body must be a JSON object." });
        return;
      }
      const existing = await redis.get(STATE_KEY);
      const merged = mergeState(existing, body);
      await redis.set(STATE_KEY, merged);
      res.status(200).json({ ok: true, savedAt: new Date().toISOString() });
      return;
    }

    res.setHeader("Allow", "GET, PUT");
    res.status(405).json({ error: "Method not allowed" });
  } catch (err) {
    console.error("attendance-portal state API error:", err);
    res.status(500).json({ error: "Server error", detail: String(err) });
  }
}
