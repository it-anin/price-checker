-- Run in the CURRENT project's SQL editor.
-- New tables only. Filled by tools/stock-sync (service-role key, runs in the office);
-- the web app only reads them with the anon key.
begin;

create table if not exists public.stock_onhand (
  branch text not null,
  sku text not null,
  qty numeric not null,
  synced_at timestamptz not null default now(),
  primary key (branch, sku),
  constraint stock_onhand_branch check (branch in ('src', 'sss', 'kkl'))
);

create table if not exists public.stock_sync_status (
  branch text primary key,
  synced_at timestamptz not null default now(),
  row_count integer not null default 0,
  ok boolean not null default true,
  message text,
  constraint stock_sync_status_branch check (branch in ('src', 'sss', 'kkl'))
);

alter table public.stock_onhand enable row level security;
alter table public.stock_sync_status enable row level security;

grant select on public.stock_onhand to anon, authenticated;
grant select on public.stock_sync_status to anon, authenticated;

create policy "stock_read_onhand" on public.stock_onhand
  for select to anon, authenticated using (true);
create policy "stock_read_sync_status" on public.stock_sync_status
  for select to anon, authenticated using (true);

-- No insert/update/delete policies: only the service-role key (bypasses RLS) may write.

commit;
