// Vercel serverless function — handles login and admin user-management,
// backed by the same Upstash Redis database as /api/state.js.
//
// There is exactly one hardcoded admin account (User_Admin / Pass_Admin).
// Every other account is a plain "user" role, created by the admin from
// the Users tab. Passwords are stored as SHA-256 hashes, never plaintext.
import { Redis } from "@upstash/redis";
import crypto from "crypto";

const redis = new Redis({
  url: process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL,
  token: process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN,
});

const USERS_KEY = "attendance-portal:users";
const ADMIN_USERNAME = "User_Admin";
const ADMIN_PASSWORD = "Pass_Admin";

const hash = (pw) => crypto.createHash("sha256").update(String(pw)).digest("hex");

async function getUsers() {
  const data = await redis.get(USERS_KEY);
  return data && typeof data === "object" ? data : {};
}

function isAdmin(username, password) {
  return username === ADMIN_USERNAME && password === ADMIN_PASSWORD;
}

export default async function handler(req, res) {
  if (!process.env.KV_REST_API_URL && !process.env.UPSTASH_REDIS_REST_URL) {
    res.status(500).json({
      error: "No database connected. Add the 'Upstash for Redis' integration in the Vercel dashboard and redeploy.",
    });
    return;
  }

  try {
    const body = req.method === "GET" ? req.query : req.body && typeof req.body === "object" ? req.body : JSON.parse(req.body || "{}");

    // ---- LOGIN ----
    if (req.method === "POST" && body.action === "login") {
      const { username, password } = body;
      if (!username || !password) {
        res.status(400).json({ ok: false, error: "Username and password required." });
        return;
      }
      if (isAdmin(username, password)) {
        res.status(200).json({ ok: true, username, role: "admin" });
        return;
      }
      const users = await getUsers();
      const record = users[username];
      if (record && record.passwordHash === hash(password)) {
        res.status(200).json({ ok: true, username, role: "user" });
        return;
      }
      res.status(401).json({ ok: false, error: "Invalid username or password." });
      return;
    }

    // Everything below requires admin credentials in the body
    const { adminUser, adminPass } = body;
    if (!isAdmin(adminUser, adminPass)) {
      res.status(403).json({ ok: false, error: "Admin credentials required." });
      return;
    }

    // ---- LIST USERS (admin) ----
    if (req.method === "GET") {
      const users = await getUsers();
      const list = Object.entries(users).map(([username, u]) => ({ username, role: u.role, createdAt: u.createdAt }));
      res.status(200).json({ ok: true, users: list });
      return;
    }

    // ---- ADD USER (admin) ----
    if (req.method === "POST" && body.action === "add") {
      const { username, password } = body;
      if (!username || !password) {
        res.status(400).json({ ok: false, error: "Username and password required." });
        return;
      }
      if (username === ADMIN_USERNAME) {
        res.status(400).json({ ok: false, error: "That username is reserved for the admin account." });
        return;
      }
      const users = await getUsers();
      if (users[username]) {
        res.status(409).json({ ok: false, error: "That username already exists." });
        return;
      }
      users[username] = { passwordHash: hash(password), role: "user", createdAt: new Date().toISOString() };
      await redis.set(USERS_KEY, users);
      res.status(200).json({ ok: true });
      return;
    }

    // ---- CHANGE PASSWORD (admin) ----
    if ((req.method === "PUT" || req.method === "POST") && body.action === "changePassword") {
      const { username, newPassword } = body;
      if (!username || !newPassword) {
        res.status(400).json({ ok: false, error: "Username and new password required." });
        return;
      }
      const users = await getUsers();
      if (!users[username]) {
        res.status(404).json({ ok: false, error: "User not found." });
        return;
      }
      users[username].passwordHash = hash(newPassword);
      await redis.set(USERS_KEY, users);
      res.status(200).json({ ok: true });
      return;
    }

    // ---- DELETE USER (admin) ----
    if (req.method === "DELETE" || (req.method === "POST" && body.action === "delete")) {
      const { username } = body;
      const users = await getUsers();
      delete users[username];
      await redis.set(USERS_KEY, users);
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ ok: false, error: "Method not allowed" });
  } catch (err) {
    console.error("attendance-portal users API error:", err);
    res.status(500).json({ ok: false, error: "Server error", detail: String(err) });
  }
}
