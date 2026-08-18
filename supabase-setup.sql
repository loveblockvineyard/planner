-- Run this once in Supabase → SQL Editor → New query → Run.

-- 1. the table that holds all app data (one row per key)
create table if not exists public.kv (
  key        text primary key,
  value      jsonb,
  updated_at timestamptz default now()
);

-- 2. permissions: the app has no per-person login, so the anon key
--    is allowed to read/write. The operator/manager codes are the gate.
alter table public.kv enable row level security;

drop policy if exists "app read"  on public.kv;
drop policy if exists "app write" on public.kv;

create policy "app read"  on public.kv for select using (true);
create policy "app write" on public.kv for all    using (true) with check (true);

-- 3. LIVE UPDATES — this is what pushes an operator's "Done" to the
--    manager's screen instantly. Without it the app still works, but
--    only refreshes every 20 seconds instead of immediately.
alter table public.kv replica identity full;

do $$
begin
  alter publication supabase_realtime add table public.kv;
exception
  when duplicate_object then null;   -- already added, nothing to do
end $$;
