# Attendance Register — MSc & PGD

A university attendance portal built with React + Vite. Tracks students across
a shared MSc/PGD batch, handles electives (different students, different
subjects), marks daily attendance, and exports a color-coded Excel workbook.

## Features
- Students: name, email, enrollment number, program (MSc/PGD), electives
- Subjects: core (everyone) and electives (only enrolled students)
- Daily roll call per subject, present/absent per student
- Dashboard: average attendance per subject, across all subjects
- Reports: per-student, per-subject breakdown + overall attendance
- Excel export with color-coded attendance:
  - 🟩 Green — 75% and above
  - 🟧 Orange — 60% to 74%
  - 🟥 Red — below 60%

## Run locally

You need [Node.js](https://nodejs.org) 18+ installed.

```bash
npm install
npm run dev
```

Open the URL it prints (usually `http://localhost:5173`).

To build a production bundle locally:

```bash
npm run build
npm run preview
```

## Deploy to Vercel

### Option A — Vercel CLI (fastest)

1. Install the CLI once: `npm install -g vercel`
2. From this project folder, run:
   ```bash
   vercel
   ```
3. Answer the prompts (link to a new project, keep defaults — Vercel
   auto-detects Vite). It'll give you a live preview URL.
4. To push to production: `vercel --prod`

### Option B — GitHub + Vercel dashboard (recommended for ongoing use)

1. Push this folder to a new GitHub repository:
   ```bash
   git init
   git add .
   git commit -m "Attendance portal"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git push -u origin main
   ```
2. Go to [vercel.com](https://vercel.com) and sign in (GitHub login is
   easiest).
3. Click **Add New → Project**, then select your repository.
4. Vercel will auto-detect the **Vite** framework preset:
   - Build command: `npm run build`
   - Output directory: `dist`
   (already set in `vercel.json`, so you don't need to change anything)
5. Click **Deploy**. In under a minute you'll get a live URL like
   `https://your-project.vercel.app`.
6. Any future `git push` to `main` automatically redeploys.

## Shared database (for multiple faculty)

The app stores all data (students, subjects, attendance, session notes) in a
shared database via `/api/state.js`, so everyone using the deployed site sees
the same data — not just whoever's browser entered it.

**One-time setup on Vercel:**

1. Open your project on [vercel.com](https://vercel.com) → **Storage** tab
   (or **Integrations** → Marketplace).
2. Install **"Upstash for Redis"** and connect it to this project. Vercel
   will provision a free Redis database and automatically add the
   `KV_REST_API_URL` / `KV_REST_API_TOKEN` environment variables to your
   project — you don't need to copy/paste anything manually.
3. Redeploy the project (Vercel usually prompts you to; if not, go to
   **Deployments** → **Redeploy**) so the new environment variables are
   picked up.
4. Open the site — the sidebar shows a small status dot: green "Synced" once
   it's talking to the database. If it shows a red "Couldn't reach database"
   message, double check step 2–3.

That's it — anyone who opens the deployed URL now reads and writes the same
data. Local development (`npm run dev`) won't have access to the database
unless you also run `vercel env pull .env.local` and use `vercel dev`
instead of `vite dev`, since `/api` functions only run on Vercel's servers
(or via the Vercel CLI), not in Vite's local dev server.

**Good to know:**
- Saves are debounced (~1 second after you stop typing/clicking) and there's
  light polling (~every 20s) so other open tabs pick up changes without a
  manual refresh.
- Concurrent edits use "last write wins" — there's no per-field merge, so if
  two people edit the same record at the exact same moment, the later save
  overwrites the earlier one. Fine for a small faculty team; would need more
  work (e.g. per-record writes, optimistic locking) at larger scale.
- The "Reset to sample data" link in the sidebar now resets the **shared**
  database for everyone, not just your browser — it asks for confirmation
  first.

## Notes
- The Excel export uses `exceljs` in the browser, so colors and formatting
  come through in the downloaded `.xlsx` file (not just on screen).
- If the "Upstash for Redis" integration is ever removed or its environment
  variables go missing, `/api/state` returns a clear error message rather
  than failing silently, and the app falls back to the built-in sample data.
