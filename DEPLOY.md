# Deploying Vineyard Ops — Supabase + GitHub + Netlify

Three parts, in this order:

1. **Supabase** — the shared database (so every phone sees the same data)
2. **GitHub** — where the code lives
3. **Netlify** — builds the code and serves the website

Total time: about 20 minutes, once.

---

## Part 1 — Supabase (the database)

**1.1** Go to [supabase.com](https://supabase.com) → **Start your project** → sign in with GitHub or email.

**1.2** Click **New project**.
- **Name:** `loveblock-vineyard`
- **Database Password:** click Generate, then **save it somewhere safe** (you won't need it for the app, but you'll want it one day)
- **Region:** `Southeast Asia (Singapore)` or `Oceania (Sydney)` — Sydney is closest to NZ
- Click **Create new project** and wait ~2 minutes while it builds.

**1.3** Create the table. In the left sidebar click **SQL Editor** → **New query**. Paste this in and click **Run**:

```sql
create table if not exists public.kv (
  key        text primary key,
  value      jsonb,
  updated_at timestamptz default now()
);

alter table public.kv enable row level security;

drop policy if exists "app read"  on public.kv;
drop policy if exists "app write" on public.kv;

create policy "app read"  on public.kv for select using (true);
create policy "app write" on public.kv for all    using (true) with check (true);

-- live updates: pushes an operator's "Done" to your screen instantly
alter table public.kv replica identity full;

do $$
begin
  alter publication supabase_realtime add table public.kv;
exception
  when duplicate_object then null;
end $$;
```

You should see **Success. No rows returned**. (This SQL is also in the project as `supabase-setup.sql`.)

**1.4** Get your two keys. Left sidebar → **Project Settings** (gear) → **API**. Copy these two, paste them in a note for a minute:

- **Project URL** — looks like `https://abcdefgh.supabase.co`
- **anon public** key — a long string starting `eyJ...`

> The `anon` key is designed to be public and safe in a website. Never use the `service_role` key here.

---

## Part 2 — GitHub (the code)

**2.1** Unzip `vineyard-ops.zip`. You'll get a `vineyard-ops` folder.

**2.2** Go to [github.com](https://github.com) → **+** (top right) → **New repository**.
- **Name:** `vineyard-ops`
- Set it **Private**
- Do **not** tick "Add a README"
- **Create repository**

**2.3** Upload the code. Easiest way, no command line:

On the new empty repo page, click **uploading an existing file**. Open your `vineyard-ops` folder, select everything **inside** it (not the folder itself), and drag it into the browser.

> Important: the files (`package.json`, `index.html`, the `src` folder…) must sit at the **top level** of the repo, not inside another `vineyard-ops` folder. That's the #1 cause of failed deploys.

Then click **Commit changes**.

*If you prefer the terminal:*
```bash
cd vineyard-ops
git init
git add .
git commit -m "Vineyard Ops"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/vineyard-ops.git
git push -u origin main
```

---

## Part 3 — Netlify (the website)

**3.1** Go to [netlify.com](https://netlify.com) → sign up / log in **with GitHub**.

**3.2** Click **Add new site** → **Import an existing project** → **GitHub** → authorise → pick your `vineyard-ops` repo.

**3.3** Netlify reads `netlify.toml` from the project, so the build settings should already be filled in:
- Build command: `npm run build`
- Publish directory: `dist`

Leave them as they are.

**3.4** **Before clicking Deploy**, click **Add environment variables** (or "Show advanced" → New variable) and add these two — this is what connects the app to your database:

| Key | Value |
|---|---|
| `VITE_SUPABASE_URL` | your Project URL from step 1.4 |
| `VITE_SUPABASE_ANON_KEY` | your anon public key from step 1.4 |

The names must match exactly, including `VITE_`.

**3.5** Click **Deploy**. It takes 1–2 minutes. When it's done you get a URL like `https://random-name-123.netlify.app`.

**3.6** Rename it (optional): **Site configuration** → **Change site name** → e.g. `loveblock-vineyard` → your URL becomes `https://loveblock-vineyard.netlify.app`.

---

## Check it worked

1. Open the site. Log in as manager: **0000**
2. The header should read **"Manager console · synced"**. If it says "this device only", the env vars didn't take — see troubleshooting.
3. Schedule a job in Work.
4. Open the same URL on your phone, log in as an operator (**1234**), and you should see that job.

That's the whole point of Supabase — same data everywhere.

---

## Day-to-day use

**On phones:** open the URL in Chrome (Android) or Safari (iPhone) → menu → **Add to Home Screen**. It then behaves like an installed app.

**Updating the app later:** change the code on GitHub (or upload new files) → Netlify rebuilds automatically within a couple of minutes. **Your data is never touched by a deploy** — it lives in Supabase, completely separately.

---

## Live updates

When an operator taps **Done**, your board updates within about a second —
no refreshing. Same the other way: work you schedule appears on their phones
straight away, and a job dragged from one operator to another moves on
everyone's screen.

This is on automatically once step 1.3 has been run. If the signal drops, the
app falls back to checking every 20 seconds, and catches up as soon as it
reconnects.

To confirm it's working: open the site on two devices side by side, tick a job
done on one, and watch the other.

If updates only ever arrive after ~20 seconds, part 3 of the SQL didn't run —
go back to the SQL Editor and run the snippet again (it's safe to re-run).

---

## Troubleshooting

**Header says "this device only"**
The env vars are missing or misspelled. Netlify → Site configuration → Environment variables. Check both names are exactly `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY`. After fixing, you must **redeploy**: Deploys → Trigger deploy → **Clear cache and deploy site**. (Vite bakes these in at build time, so a redeploy is required — editing them alone does nothing.)

**Build failed**
Open the deploy log and read the last few lines. Most common: the files ended up nested one folder too deep in GitHub. The repo root must contain `package.json`.

**Site loads but data doesn't save**
The SQL in step 1.3 probably didn't run. Supabase → Table Editor — you should see a `kv` table. If not, run the SQL again.

**Blank white page**
Usually a stale service worker. Hard-refresh (Ctrl/Cmd + Shift + R).

---

## A note on security

Anyone with your site URL can reach the data — the operator/manager codes are the only gate, and the anon key is visible in the page source (that's normal and expected for this kind of app). That's fine for a private vineyard tool nobody's advertising.

If you'd like it locked down properly later — real logins per person, or rules so operators can only edit their own rows — that's a straightforward next step. Just ask.
