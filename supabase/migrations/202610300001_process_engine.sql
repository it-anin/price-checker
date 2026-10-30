-- Run in the CURRENT project's SQL editor (hzxlulphlvisromrniat).
-- This creates new tables only. It does not read or alter products/prices.
begin;

create table if not exists public.grab_files (
  name text primary key,
  content text not null,
  updated_at timestamptz not null default now(),
  constraint grab_files_supported_name check (name ~ '^GRAB_(SRC|SSS|KKL)\.CSV$')
);

-- Keep the original app_settings / pending_skus JSON format to preserve
-- per-SKU names, branch membership, and any legacy metadata during migration.
create table if not exists public.app_settings (
  key text primary key,
  value text not null,
  updated_at timestamptz not null default now()
);

alter table public.grab_files enable row level security;
alter table public.app_settings enable row level security;

grant select on public.grab_files to anon, authenticated;
grant insert, update on public.grab_files to authenticated;
grant select on public.app_settings to anon, authenticated;
grant insert, update on public.app_settings to authenticated;

create policy "engine_read_templates" on public.grab_files
  for select to anon, authenticated using (true);
create policy "engine_authenticated_insert_templates" on public.grab_files
  for insert to authenticated with check ((select auth.uid()) is not null);
create policy "engine_authenticated_update_templates" on public.grab_files
  for update to authenticated using ((select auth.uid()) is not null)
  with check ((select auth.uid()) is not null);

create policy "engine_read_pending_skus" on public.app_settings
  for select to anon, authenticated using (key = 'pending_skus');
create policy "engine_authenticated_insert_pending" on public.app_settings
  for insert to authenticated with check (key = 'pending_skus' and (select auth.uid()) is not null);
create policy "engine_authenticated_update_pending" on public.app_settings
  for update to authenticated using (key = 'pending_skus' and (select auth.uid()) is not null)
  with check (key = 'pending_skus' and (select auth.uid()) is not null);

-- Do not seed [] here: a missing record blocks RUN until the old pending list
-- has been copied and verified. An intentionally empty list must be explicit.
notify pgrst, 'reload schema';
commit;
