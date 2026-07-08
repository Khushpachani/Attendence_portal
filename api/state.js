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
      await redis.set(STATE_KEY, body);
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
