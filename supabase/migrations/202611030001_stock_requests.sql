-- Run in the CURRENT project's SQL editor, after 202611020001_stock_onhand.sql.
-- On-demand live stock: the web app asks (RPC request_stock_sync), the office agent
-- (tools/stock-sync --serve) sees the request, queries MSSQL and writes stock_onhand.
-- New objects only. Nothing existing is read or altered.
begin;

create table if not exists public.stock_requests (
  id uuid primary key default gen_random_uuid(),
  requested_by uuid,
  requested_at timestamptz not null default now(),
  status text not null default 'pending',
  started_at timestamptz,
  finished_at timestamptz,
  message text,
  constraint stock_requests_status check (status in ('pending', 'running', 'done', 'error'))
);

-- At most one open request at a time: double clicks and several users share one run.
create unique index if not exists stock_requests_one_open
  on public.stock_requests ((true)) where status in ('pending', 'running');

create index if not exists stock_requests_finished_idx
  on public.stock_requests (finished_at desc) where status = 'done';

-- Heartbeat written by the agent so the web app can tell "agent offline" immediately.
create table if not exists public.stock_agent_status (
  id integer primary key,
  last_seen timestamptz not null default now(),
  info text,
  constraint stock_agent_status_single check (id = 1)
);

alter table public.stock_requests enable row level security;
alter table public.stock_agent_status enable row level security;

-- Supabase grants new public tables to anon/authenticated by default; take that back
-- so RLS is not the only thing standing between a browser and these tables.
revoke all on public.stock_requests from anon, authenticated;
revoke all on public.stock_agent_status from anon, authenticated;
grant select on public.stock_requests to authenticated;
grant select on public.stock_agent_status to authenticated;

drop policy if exists "stock_requests_read" on public.stock_requests;
create policy "stock_requests_read" on public.stock_requests
  for select to authenticated using (true);
drop policy if exists "stock_agent_status_read" on public.stock_agent_status;
create policy "stock_agent_status_read" on public.stock_agent_status
  for select to authenticated using (true);
-- No insert/update/delete policies: clients write only through the RPC below,
-- the agent writes with the service-role key (bypasses RLS).

create or replace function public.request_stock_sync()
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  v_id uuid;
begin
  if (select auth.uid()) is null then
    raise exception 'ต้องเข้าสู่ระบบก่อนดึงสต็อกสด' using errcode = '28000';
  end if;

  -- A request nobody answered (agent down/crashed) must not block everyone forever.
  update public.stock_requests
     set status = 'error', finished_at = now(), message = 'หมดเวลา: agent ไม่ตอบสนอง'
   where status in ('pending', 'running') and requested_at < now() - interval '2 minutes';

  -- Someone already asked and it is still in progress: share that run.
  select id into v_id from public.stock_requests
   where status in ('pending', 'running') order by requested_at desc limit 1;
  if v_id is not null then return v_id; end if;

  -- A run finished a moment ago: its data is fresh enough, do not hit MSSQL again.
  select id into v_id from public.stock_requests
   where status = 'done' and finished_at > now() - interval '10 seconds'
   order by finished_at desc limit 1;
  if v_id is not null then return v_id; end if;

  begin
    insert into public.stock_requests (requested_by) values ((select auth.uid()))
    returning id into v_id;
  exception when unique_violation then
    -- Lost a race with another caller: use the request that won.
    select id into v_id from public.stock_requests
     where status in ('pending', 'running') order by requested_at desc limit 1;
  end;
  return v_id;
end;
$$;

revoke all on function public.request_stock_sync() from public, anon;
grant execute on function public.request_stock_sync() to authenticated;

commit;
