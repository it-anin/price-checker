# MSSQL live bridge: architecture and handoff guide

How a web app hosted in the cloud (Vercel) reads **live** data from an **on-premises MSSQL** server, with no inbound
ports, no VPN and no API route. It is built and running for on-hand stock (report R01.102, `CF_QUANTITY`) in
`it-anin/price-checker`. This guide is written so an agent in a **different** repository can reuse the pattern for other
tables. UI strings are Thai and stay Thai; everything else is English.

> **If this repository is public:** remove section 11 (environment facts) first. It holds internal names, not secrets.

Contents: 1 When to use · 2 Architecture · 3 Request lifecycle · 4 File map · 5 Security · 6 Data contract ·
7 Add a new job · 8 Generalizing · 9 Pitfalls learned · 10 Windows runbook · 11 Environment facts · 12 Verification ·
Appendix A generalized SQL · Appendix B reading ProMaxx `.REP` files · Appendix C language-neutral agent contract

---

## 1. When to use / when not

Use it when **all** are true: the source is MSSQL on a LAN you cannot expose; the web app is cloud-hosted; each result
is a **small–medium table (up to a few tens of thousands of rows)**; a few seconds of delay after a button press is fine;
read-only access is enough.

Do **not** use it for: hundreds of thousands of rows or more (mirroring the whole result into Supabase stops being sensible:
aggregate in SQL, restrict by date range, or send a computed report instead); sub-second latency; writing back to MSSQL;
anonymous visitors (this design requires a signed-in user).

## 2. Architecture

```
 Browser (signed in)          Supabase (cloud)                 Office PC (BigyaMainPC)
 ───────────────────          ────────────────                 ───────────────────────
 press button ──RPC────────►  request_stock_sync()
                              └ inserts stock_requests(pending)
                                                    ◄──poll 2 s── agent (sync.mjs --serve)
                                                                   claims → running
                                                                   runs FIXED read-only SQL ──► MSSQL
                                                                   upserts stock_onhand ───────►│
 poll request row ◄───────────  closes request: done | error ◄─────┘ (service-role key, office only)
 read stock_onhand (paged) ◄──  stock_onhand, stock_sync_status
```

- Vercel cannot reach the LAN, and opening MSSQL (port 1433) to the internet is not acceptable. The agent only makes
  **outbound** HTTPS calls to Supabase, which acts as a mailbox.
- No Next.js API route is needed (this repository's `AGENTS.md` forbids creating API routes unless asked). The browser talks to
  Supabase through the shared client in `lib/supabase.js` only.
- MSSQL is queried **only when someone presses the button**. There is no timer. (An earlier version synced every 15 minutes;
  it was removed because live data was required.)
- If the agent is off, the page says so within seconds and offers the last stored data with its age.

## 3. Request lifecycle

States of a `stock_requests` row: `pending → running → done | error`.

| Step | Who | What |
|---|---|---|
| 1 | Web `handleEngineLoadLive()` | not signed in → show login form, send nothing. Agent heartbeat older than 90 s → offer last stored data. |
| 2 | RPC `request_stock_sync()` | rejects `auth.uid()` null; expires stale requests; returns the **open** request if one exists; returns a request that **finished < 10 s ago**; otherwise inserts a new one. |
| 3 | Agent `claimNextRequest()` | oldest `pending` → `running`. The update filters on `status='pending'`, so a second process cannot claim it twice. |
| 4 | Agent `syncBranch()` ×3 in parallel | read MSSQL → `toStockRows` → upsert → delete vanished rows → write `stock_sync_status`. |
| 5 | Agent `finishRequest()` | `done` with a short summary (`src 5461, sss 3287, kkl 4007`) or `error` with the reason (≤ 300 chars). |
| 6 | Web `waitForStockRequest()` | polls the row every 1 s until done/error or 90 s → then `loadStockFromDb()`. |

Constants (all in code; change the code, not this table): agent poll `POLL_MS = 2000` and heartbeat `HEARTBEAT_MS = 30000`
(`tools/stock-sync/serve.mjs`); `AGENT_ONLINE_SECONDS = 90`, `LIVE_TIMEOUT_MS = 90_000`, `LIVE_POLL_MS = 1_000`
(`lib/stock-live.ts`); request expiry **2 minutes** and reuse window **10 seconds** (`supabase/migrations/202611030001_stock_requests.sql`);
upsert `CHUNK = 1000` with `CONCURRENCY = 4` (`tools/stock-sync/sync.mjs`); read page `PAGE_SIZE = 1000` (`lib/stock-source.ts`).
End-to-end latency is **not benchmarked here**: the agent logs seconds per request in `serve.log`; measure on the real server.

The heartbeat has its **own timer**, so the agent still looks online while a long request runs.
On start the agent closes requests left `pending/running` by a crash with `error` ("agent รีสตาร์ทระหว่างดึงข้อมูล").

## 4. File map (this repository)

| File | Role | Reuse |
|---|---|---|
| `tools/stock-sync/sync.mjs` | agent entry: `--serve`, `--dry-run`, `--compare`, one-shot; `syncBranch()`, `writeBranch()` | copy, replace stock specifics |
| `tools/stock-sync/serve.mjs` | pure/thin helpers: `summarizeResults`, `claimNextRequest`, `finishRequest`, `recoverStuckRequests`, `sendHeartbeat` | copy as is |
| `tools/stock-sync/query.mjs` | `buildConnectionConfig(env)` (named-instance safe), `parseBranchMap`, `chunkRows`, stock SQL `buildStockQuery`, `toStockRows` | copy first and last helpers; the SQL is stock-only |
| `tools/stock-sync/compare.mjs` | CSV reader + `compareStock` for equivalence checks against an existing report | copy the idea; adapt columns |
| `tools/stock-sync/run-serve.cmd`, `env.example`, `README.md`, `package.json` | Windows launcher (`%~dp0`), env template, ops notes; deps `mssql`, `@supabase/supabase-js` | copy |
| `supabase/migrations/202611020001_stock_onhand.sql` | result tables + read-only RLS | stock-only; use as a template |
| `supabase/migrations/202611030001_stock_requests.sql` | mailbox tables, heartbeat table, RPC `request_stock_sync()` | superseded by Appendix A for new work |
| `lib/stock-live.ts` | pure: `agentOnline`, `requestStatusText`, `pollStockRequest` (injectable `sleep`/`now`) | copy as is |
| `lib/stock-source.ts` | Supabase calls: `requestLiveStock`, `waitForStockRequest`, `loadAgentLastSeen`, `loadMssqlStock` (paged) | copy, rename |
| `lib/process-engine.ts` | `stockReportFromRows` (rows → domain object), `stockSyncWarning` (stored-data age) | pattern only |
| `app/page.tsx` | `handleEngineLoadLive()` and `loadStockFromDb()`; button text `ดึงจากฐานข้อมูล` | pattern only |
| `tests/stock-sync.test.mjs`, `tests/stock-live.test.cjs` | helper tests; fake Supabase for claim/finish | copy and adapt |

## 5. Security model

1. **MSSQL login is read-only and minimal.** `SELECT` on only the tables the query needs; no write roles; not `sysadmin`.
   Prove it, do not assume it:
   ```sql
   -- must return no rows
   SELECT r.name FROM sys.server_role_members m
   JOIN sys.server_principals r ON r.principal_id = m.role_principal_id
   JOIN sys.server_principals l ON l.principal_id = m.member_principal_id WHERE l.name = '<login>';
   -- must fail with Msg 229 "permission was denied" (it changes no rows anyway)
   USE [<db>]; EXECUTE AS USER = '<login>';
   UPDATE dbo.<AnyTable> SET <AnyColumn> = <AnyColumn> WHERE 1 = 0;
   REVERT;
   ```
2. **The browser never sends SQL.** A request carries only a job name from an allowlist. The agent runs fixed queries from code.
3. **Secrets live only on the office PC** (`.env`: MSSQL password, Supabase **service-role** key). Never in the repository, Vercel, the
   browser, chat messages or the docs. `.gitignore` covers `.env*`. The browser uses the **anon** key and the user's session.
4. **Supabase hardening.** New public tables are granted to `anon`/`authenticated` by default: the migrations `revoke all` and grant
   only `select` to `authenticated`. Clients cannot insert/update/delete these tables; they go through the `security definer` RPC
   (`set search_path = public`, `revoke … from public, anon`). The agent writes with the service-role key (bypasses RLS).
5. **Signed-in users only.** The RPC raises `28000` when `auth.uid()` is null. Also **turn off public sign-ups** in Supabase Auth,
   otherwise anyone can register and press the button.
6. **Load protection for the ERP.** One open request per job, a 10 s reuse window, 2 min expiry. Repeated presses cannot hammer MSSQL.
7. **Empty result is an error**, never "delete everything". Vanished rows are deleted only after every chunk was upserted.

## 6. Data contract

`fixed SQL → rows → pure mapper (toStockRows) → upsert in chunks → delete vanished → status row → web reads paged → pure builder (stockReportFromRows)`

- **Mapper normalizes** keys (trim, upper-case), sums duplicates, rounds to 4 decimals, skips null/non-finite. Pure and unit-tested.
- **Write** (`writeBranch`): `upsert … onConflict: 'branch,sku'` in 1000-row chunks, 4 at a time; then
  `delete … where branch = X and synced_at < <this run's timestamp>`. All rows of a run share one `synced_at`.
- **Status row** per dataset (`stock_sync_status`: `synced_at`, `row_count`, `ok`, `message`). On failure it records `ok=false` and the reason.
- **Read** (`loadMssqlStock`): Supabase returns at most 1000 rows per request, so page with `.range()` **and a stable `.order()`**
  (without an order, pages can overlap or skip rows).
- **Builder** converts rows to the app's domain object; it throws a Thai message when empty or non-numeric.
- Stored data keeps working as a fallback: `stockSyncWarning` warns when older than 60 minutes or `ok=false`.

## 7. Add a new job (checklist)

Assume small–medium tables and the **same Supabase project** as an existing agent. Order matters.

1. **Define the contract.** Which MSSQL tables/columns, filters, key column, expected row count. If an ERP report already produces
   this data, read its SQL (PowerBuilder `.REP`/`.FRM` files contain it, see Appendix B) and treat that report as the source of truth.
   Note that ProMaxx **injects the user's filters at run time** (warehouse, until-date, item): the SQL you read is a template, so decide
   every filter explicitly (a fixed value, or an allowlisted parameter) and check the result against an export made with the same filters.
2. **Grant read access.** The DBA grants `SELECT` on exactly the new tables to the agent's login (or creates a separate login per
   project). The current `web_readonly` user can read only `BRANCH, ITEMS, TRANDETAILITEM, TRANEXTRAINFO, TRANS, VOUCHER, WAREHOUSE`.
3. **Write the fixed query** in a pure module (`buildXQuery(params)`) with **parameters for ids, never string concatenation** of
   request data. Rewrite correlated subqueries as `GROUP BY` joins when the source report uses them (equivalent, far faster).
4. **Write the pure mapper** and unit-test it (duplicates, nulls, rounding, whitespace, ordering).
5. **Create the result table** (migration): primary key = natural key, `synced_at`, RLS on, `revoke all` from `anon, authenticated`,
   `grant select … to authenticated` (or to `anon` only if the data is public, which is rarely true).
6. **Register the job** in the agent (Section 8) and make it return `{ ok, count }` so `summarizeResults` can report it.
7. **Web side.** Reuse `requestLiveStock`/`waitForStockRequest`/`pollStockRequest` (rename, parameterize by job). Add a loader with
   paging + stable order, a pure builder, and a handler that: requires login, checks agent heartbeat, requests, waits, loads,
   and on any failure shows the Thai reason and **leaves already-loaded data untouched**.
8. **Prove equivalence** with an existing ERP export (Section 12). Do this before telling anyone it works.
9. **Document**: update the repository's agent guide and this file's environment facts; add tests to the test command.

**Same agent or a new agent?** Same Supabase project → add the job to the existing agent (one Task Scheduler entry, one `.env`).
Different Supabase project → run a **second agent instance** in its own folder with its own `.env` and Task name; never share a
service-role key between projects.

## 8. Generalizing (recommended, not yet built)

Today everything is named `stock_*` and the RPC takes no argument. For several jobs use the Appendix A schema: an allowlist table
`bridge_jobs`, one `bridge_requests` table with a `job` column, **one open request per job**, per-job cooldown and expiry, and
`request_bridge_job(p_job text)`. On the agent side keep a registry of fixed jobs:

```js
// Sketch (NOT run anywhere): the agent only runs jobs it knows. Unknown names are closed as errors.
const jobs = {
  stock:    { run: (ctx) => syncStock(ctx) },     // each returns [{ branch|key, ok, count?, error? }]
  products: { run: (ctx) => syncProducts(ctx) },
}
// claim → const job = jobs[request.job]
// if (!job) await finishRequest(supabase, id, { status: 'error', message: 'agent ไม่รู้จักงานนี้' })
```

The agent still handles one request at a time; different jobs queue naturally. Appendix A was tested **only on the SQL side**
(Postgres 16); the registry sketch and web changes are untested design.

## 9. Pitfalls learned (all hit in this project)

| Symptom / trap | Cause and fix |
|---|---|
| Cannot connect to `BigyaMainPC\SQLSERVER` although the port setting looks right | tedious rejects `port` together with `instanceName`. `buildConnectionConfig` sends only `instanceName` when `MSSQL_INSTANCE` is set. A named instance also needs the **SQL Server Browser** service (UDP 1434) or a fixed TCP port. Running the agent on the SQL Server PC with `MSSQL_SERVER=localhost` avoids the firewall. |
| Totals about 4.5× larger than the ERP report for one branch (242,472.65 vs 54,398.65) | The branch has several warehouses (SRC: 0 `Warehouse`, 1 `Front Store`); the report was exported for one. The `.REP` explains why: its `<fsql>` block ends with `<filter>AND!<warehouse><item></filter>`, i.e. ProMaxx appends the warehouse/item the user picked. Set `warehouseIds` per branch and **pin them for every branch**, so a new warehouse cannot silently change numbers. |
| Report offers an "until date" filter but our SQL has none | In R01.102 the stock subquery contains `<filter>and!<untildateexcol>H.TRANDATE!…</filter>`. Our fixed query has no `H.TRANDATE` condition, so it returns stock **as of now**, including future-dated documents. It matched the real exports on 2026-11-02; if users ever export "until a past date" it will not. Decide this on purpose (fixed "now", or an allowlisted date parameter) and record it. |
| `Invalid supabaseUrl: Must be a valid HTTP or HTTPS URL` | `SUPABASE_URL` was missing `https://`, or a key was pasted into it. Format: `https://<project-ref>.supabase.co`, no quotes, no `/rest/v1`. |
| `Cannot find module …\sync.mjs` | Command run from the wrong folder (repo root instead of `tools\stock-sync`). `--env-file=.env` is also relative to the **current folder**; `run-serve.cmd` does `cd /d "%~dp0"` first. |
| `Register-ScheduledTask` / `schtasks`: **Access is denied** (0x80070005) | Registering a task as `SYSTEM` needs an **elevated** (Run as administrator) PowerShell. |
| `Move-Item` of a folder with `.git` leaves a half-moved mess | `.git` files are read-only and a terminal was still inside the folder. Do not retry: back up the `.env` files, `git clone` fresh at the destination, restore `.env`, then `rmdir /s /q` the old folder. |
| Pasted command fails with strange `>>` output | The `PS C:\…>` prompt text was copied with the command. Paste only the command. |
| Secrets sync to the cloud | Do not keep the project under Desktop/Documents (OneDrive). The agent folder holds `.env`. |
| New Supabase tables readable by `anon` | Supabase grants new public tables to API roles by default. `revoke all … from anon, authenticated` then grant only what is needed; test as `anon`. |
| Pages overlap or miss rows when loading | Missing `.order()` on a paged read. Always order by the key. |
| CSV export from the ERP shows garbage Thai | The export is UTF-8 or **TIS-620 / Windows-874**. Decode strict UTF-8, fall back to `windows-874` (`decodeBytes`). |
| `--env-file` not recognized | Node must be ≥ 20.6 (v24 works). |
| Request stays `pending` | Agent not running, wrong `.env`, or heartbeat older than 90 s. Read `serve.log`; the web page already tells the user "agent ไม่ตอบสนอง". |

## 10. Windows runbook (agent on the office PC)

1. Install Node ≥ 20.6 and Git. Clone the repository to a short path without spaces, **outside OneDrive folders**
   (this project: `C:\Users\AninMainPC\Automation\price-checker`).
2. `cd tools\stock-sync`, `npm install`, `copy env.example .env`, fill it in (`notepad .env`; check with `dir /a` that it is `.env`, not `.env.txt`).
3. Prove the connection without writing: `node --env-file=.env sync.mjs --dry-run`.
4. Prove the data against the ERP: `node --env-file=.env sync.mjs --compare src.csv sss.csv kkl.csv` (Section 12).
5. Test listening by hand: `node --env-file=.env sync.mjs --serve`, press the button in the web app, read the `[serve] done … วินาที` line.
6. Keep it running: in **PowerShell as Administrator**, register the `Stock Listener` task (script in `tools/stock-sync/README.md`):
   trigger *At startup*, account `SYSTEM`, restart every 1 minute on failure, no time limit, do not start a second instance.
   Launcher `run-serve.cmd` appends output to `serve.log`.
7. Check: `Get-ScheduledTask 'Stock Listener' | Get-ScheduledTaskInfo` (result `267009` = running), `type serve.log`,
   and `select * from public.stock_agent_status` (`last_seen` within a minute).
8. Update code: `git pull`, then restart the task (`Stop-ScheduledTask` / `Start-ScheduledTask`). `.env` is untracked and survives.
9. Rotate a password: edit `.env`, restart the task. Stop/remove: `Stop-ScheduledTask 'Stock Listener'`, `Unregister-ScheduledTask 'Stock Listener' -Confirm:$false`.

## 11. Environment facts (ProMaxx at the office; no secrets)

- SQL Server **`BigyaMainPC\SQLSERVER`** (named instance), database **`MSMAXX`**, read-only user **`web_readonly`** (password is never recorded in docs).
- Branch ids (`SELECT SYSBRANCHID, NAME FROM BRANCH`): 0 อนิน สาขาแยกชากค้อ · 1 อ่อนนุช · 2 ตลาดเวิลด์มาร์เก็ต · 3 กรุงเทพกรีฑา ·
  4 อนิน สาขาเก้ากิโล · 5 อนิน สาขาสวนเสือศรีราชา. Mapping used by the app: **SRC = 0 (warehouse 1 only), SSS = 5, KKL = 4**.
- Report **R01.102** (`d_01102_classic_p22112018`), `CF_QUANTITY = SUM(D.BASEQUANTITY * V.FSTOCKMAIN)` over `TRANS H, TRANDETAILITEM D, VOUCHER V`
  with `H.FTaxInvoiceReplacement=0, H.FVoidTaxInvoice=0, V.FSTOCKMAIN<>0, D.FPROCESS=1, H.FCANCEL=0, H.FAPPROVE=1`, items
  `FEFFECTSTOCK=1, FSHOWONREPORT=1, FUSED=1, ITEMID NOT IN ('DEPOSIT','DEPOSIT-NV')`; rows with NULL quantity are dropped (zero is kept).
  The rewritten `GROUP BY` query in `tools/stock-sync/query.mjs` matches the exported CSV on every SKU for all three branches.
- Verified totals on 2026-11-02 for orientation only (stock moves): SRC 5,461 SKUs / 54,398.65 · SSS 3,287 / 25,650 · KKL 4,007 / 20,156.
- The R01.102 `.REP` header says `<filter>wh-ud-it</filter>`: the report's filter dialog offers **w**are**h**ouse, **u**ntil-**d**ate and **it**em (our reading; not confirmed against the ProMaxx screen). That is why the real exports were restricted to one warehouse per branch.
- Grab template store ids and file names are in `docs/process-engine.md`.

## 12. Verification recipe

1. **Unit tests** for every pure helper (mapper, config builder, summarizer, poller with fake clock, claim/finish with a fake Supabase).
   `node --test tests/*.cjs tests/*.mjs`, plus `npx tsc --noEmit` and `npm run build`.
2. **SQL tests on real Postgres.** Create a throwaway Postgres 16 cluster, add stand-in roles `anon`, `authenticated`,
   `service_role` (bypassrls), a schema `auth` with `auth.uid()` reading `current_setting('request.jwt.claim.sub')`, and Supabase's
   default privileges (`alter default privileges … grant all … to anon, authenticated, service_role`). Then assert: anon cannot execute
   the RPC; no user id is rejected; unknown/disabled/null job rejected; repeated calls share one request; cooldown reuse; expiry;
   index blocks a second open row; browsers cannot insert/update/delete or read the allowlist; the migration re-runs safely.
3. **Equivalence with the ERP.** Export the existing report right before running
   `node --env-file=.env sync.mjs --compare <file>` and require `✔ ตรงกันทุก SKU` (counts, totals, per-key values). Differences:
   few → data moved, re-export; many only in MSSQL → the export used fewer warehouses/filters; many only in the file → wrong key/branch.
4. **Live check** on the deployed app: signed out (login form, no request), agent stopped (offline message + stored data offered),
   double press (one `[serve] done` line), request error (reason shown, loaded data kept).

---

## Appendix A: generalized mailbox SQL (tested on Postgres 16)

Tested only as SQL with stand-in Supabase roles (all assertions listed in Section 12.2, plus per-job independence and per-job cooldown/expiry).
Not yet used with a real agent. Idempotent.

```sql
-- Generalized request mailbox: ONE set of tables serves many jobs (stock, products, prices, ...).
-- Run after creating your own result tables. Idempotent.
begin;

-- Allowlist of job names the browser may ask for. Only migrations / service role edit it.
create table if not exists public.bridge_jobs (
  job text primary key,
  cooldown_seconds integer not null default 10 check (cooldown_seconds >= 0),
  expiry_seconds integer not null default 120 check (expiry_seconds >= 10),
  enabled boolean not null default true
);

create table if not exists public.bridge_requests (
  id uuid primary key default gen_random_uuid(),
  job text not null references public.bridge_jobs (job),
  requested_by uuid,
  requested_at timestamptz not null default now(),
  status text not null default 'pending',
  started_at timestamptz,
  finished_at timestamptz,
  message text,
  constraint bridge_requests_status check (status in ('pending', 'running', 'done', 'error'))
);

-- At most one open request PER JOB: presses and users share one run of that job.
create unique index if not exists bridge_requests_one_open_per_job
  on public.bridge_requests (job) where status in ('pending', 'running');
create index if not exists bridge_requests_done_idx
  on public.bridge_requests (job, finished_at desc) where status = 'done';

create table if not exists public.bridge_agent_status (
  id integer primary key,
  last_seen timestamptz not null default now(),
  info text,
  constraint bridge_agent_status_single check (id = 1)
);

alter table public.bridge_jobs enable row level security;
alter table public.bridge_requests enable row level security;
alter table public.bridge_agent_status enable row level security;

-- Supabase grants new public tables to anon/authenticated by default: take that back.
revoke all on public.bridge_jobs, public.bridge_requests, public.bridge_agent_status from anon, authenticated;
grant select on public.bridge_requests, public.bridge_agent_status to authenticated;

drop policy if exists "bridge_requests_read" on public.bridge_requests;
create policy "bridge_requests_read" on public.bridge_requests for select to authenticated using (true);
drop policy if exists "bridge_agent_status_read" on public.bridge_agent_status;
create policy "bridge_agent_status_read" on public.bridge_agent_status for select to authenticated using (true);
-- bridge_jobs has RLS on and no policy for clients: invisible to the browser.

create or replace function public.request_bridge_job(p_job text)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  v_job public.bridge_jobs;
  v_id uuid;
begin
  if (select auth.uid()) is null then
    raise exception 'ต้องเข้าสู่ระบบก่อน' using errcode = '28000';
  end if;

  select * into v_job from public.bridge_jobs where job = p_job and enabled;
  if not found then
    raise exception 'ไม่รู้จักงาน: %', coalesce(p_job, '(null)') using errcode = '22023';
  end if;

  -- A request nobody answered (agent down/crashed) must not block this job forever.
  update public.bridge_requests
     set status = 'error', finished_at = now(), message = 'หมดเวลา: agent ไม่ตอบสนอง'
   where job = p_job and status in ('pending', 'running')
     and requested_at < now() - make_interval(secs => v_job.expiry_seconds);

  select id into v_id from public.bridge_requests
   where job = p_job and status in ('pending', 'running') order by requested_at desc limit 1;
  if v_id is not null then return v_id; end if;

  select id into v_id from public.bridge_requests
   where job = p_job and status = 'done'
     and finished_at > now() - make_interval(secs => v_job.cooldown_seconds)
   order by finished_at desc limit 1;
  if v_id is not null then return v_id; end if;

  begin
    insert into public.bridge_requests (job, requested_by) values (p_job, (select auth.uid()))
    returning id into v_id;
  exception when unique_violation then
    select id into v_id from public.bridge_requests
     where job = p_job and status in ('pending', 'running') order by requested_at desc limit 1;
  end;
  return v_id;
end;
$$;

revoke all on function public.request_bridge_job(text) from public, anon;
grant execute on function public.request_bridge_job(text) to authenticated;

insert into public.bridge_jobs (job) values ('stock') on conflict (job) do nothing;

commit;
```

Seed one row per job in `bridge_jobs`, e.g. `insert into public.bridge_jobs (job, cooldown_seconds, expiry_seconds) values ('products', 10, 120);`.
Call from the browser: `supabase.rpc('request_bridge_job', { p_job: 'products' })`; error codes: `28000` not signed in, `22023` unknown/disabled job.


---

## Appendix B: reading a ProMaxx `.REP` file

A `.REP` is a PowerBuilder DataWindow export. It holds the **real SQL** of a ProMaxx report, so it is the best specification for a new
job. Facts below were checked against `R01.102.REP` (report "R01.102 Onhand inventory by warehouse", DataWindow `d_01102_classic_p22112018`).
Keep `.REP` files in a **private** repository: they expose internal table and column names.

| Part of the file | What it is |
|---|---|
| encoding | UTF-8 **with BOM**, CRLF line endings, Thai text. Strip the BOM, accept CR/LF. |
| `<SYNTAX>` … `</SYNTAX>` | outer wrapper; starts with `<Header><Vertion>9</Vertion></Header>` |
| `<fconfig>` | metadata: `<type>REPORT</type>`, titles in three languages (`<langth>`, `<langen>`, `<langch>`) for the group and the report, `<datawindow>` name, `<filter>wh-ud-it</filter>`, page margins |
| `<dw>` … `</dw>` | the DataWindow source: first line `$PBExportHeader$<name>.srd`, then `release 12.5;` |
| `datawindow(…)`, `header(…)`, `summary(…)`, `footer(…)`, `detail(…)` | layout bands (ignore for data work) |
| `table(column=(type=… name=… dbname=…) … retrieve="SELECT …" filter="…" sort="…")` | **column list with types** (e.g. `decimal(4)`, `char(1020)`, `long`) and the **`retrieve` SQL** (a double-quoted string over many lines) |
| `group(level=n … by=(…))`, `compute(… expression="sum( x for group n )" …)`, `text(…)`, `column(…)`, `line(…)` | grouping, **client-side calculations**, labels, fields. Totals and sums here are computed by ProMaxx, **not by SQL** |
| `<fsql>` … `</fsql>` (after `</dw>`) | a second copy of the SQL that is **identical except for ProMaxx filter hooks** such as `<filter>and!<untildateexcol>H.TRANDATE!<PLANNO>1</PLANNO></filter>` and `<filter>AND!<warehouse><item></filter>` |

How to use it:

1. Prefer the `<fsql>` copy to see **which filters the user can apply** and where they are injected; use `retrieve=` plus the `column=(…)` list for
   **column names and types** (`dbname="CF_QUANTITY"` is the alias, `dbname="WAREHOUSE.SYSBRANCHID"` is table-qualified).
2. Remove PowerBuilder comments that start with `//` (they appear inside the SQL, e.g. `//( V.VOUCHERTYPE <> 'TW' ) AND`). T-SQL does not accept them.
3. The `filter="not isnull( cf_quantity )"` and `sort="…"` after `retrieve` run **in the client after the query**. Reproduce them yourself
   (here: an inner join/`GROUP BY` drops items with no movement; zero stays; sort is up to you).
4. Replace each filter hook with an explicit decision (see Section 9). Never splice user text into the SQL: bind typed parameters for ids and dates.
5. Rewrite correlated subqueries that run per item × warehouse as one `GROUP BY` join, then prove equivalence with an export (Section 12).
6. **Stop and tell the owner** if a `.REP` contains a password or connection string. Do not copy it anywhere.

Confirmed in this file: every row of the table above. **Not present in this file** (general PowerBuilder behaviour, check when you meet it):
retrieval arguments (`arguments=(…)` with `:name` placeholders in the SQL) and `~` escape characters inside the SQL string.
Our interpretation of the `<untildateexcol>`, `<warehouse>`, `<item>` hooks and of `wh-ud-it` comes from their names and from the exports that matched;
it is not documented by the vendor, so confirm it on the report's filter screen.

---

## Appendix C: language-neutral agent contract

The web app and the agent only meet in Supabase (Appendix A or the `stock_*` tables), so the agent can be written in any language. The reference
is Node (`tools/stock-sync`). An agent in another language (for example Go) must behave exactly like this. Items 1-8 mirror `serve.mjs` and `sync.mjs`;
the REST notes at the end come from the PostgREST documentation and were **not run** in this project.

1. **Start-up recovery.** Close every request still `pending` or `running` as `error` with the message `agent รีสตาร์ทระหว่างดึงข้อมูล`.
2. **Heartbeat.** Upsert the single row `id = 1` (`last_seen = now()`, `info`) once at start and then every 30 s from a timer that does **not** wait for
   the request loop, so a long request still looks online. The web app treats the agent as offline after 90 s.
3. **Poll every 2 s.** Select the oldest `pending` request (`order by requested_at limit 1`).
4. **Claim atomically.** `update … set status='running', started_at=now() where id=<id> and status='pending'` and read how many rows changed.
   Zero rows means another process claimed it: do nothing. Never read-then-write.
5. **Run the job** from a registry of fixed jobs compiled into the agent. The request carries **no SQL and no parameters**; an unknown job is closed as `error`.
   Run the units of a job (for example branches) concurrently. Each unit returns `{ ok, count | error }` and **never throws**.
   Per unit: run the fixed read-only query → map (trim, upper-case keys, merge duplicate keys, round to 4 decimals, skip null/non-finite) →
   **empty result = error, never delete** → upsert in chunks of 1000 (up to 4 at once) on the natural key with one `synced_at` for the whole run →
   only after all chunks succeeded `delete where <scope> and synced_at < <this run's synced_at>` → write the status row
   (`synced_at`, `row_count`, `ok=true`, `message=null`). On failure write `ok=false`, `row_count=0`, the reason (≤ 200 chars) and delete nothing.
6. **Close the request.** `done` with a short summary such as `src 5461, sss 3287, kkl 4007`, or `error` with the reasons joined by ` | ` (≤ 300 chars),
   and set `finished_at`. Any error status must reach the web app: it shows the message in Thai.
7. **MSSQL connection failure.** Close the request as `error` (`เชื่อมต่อ MSSQL ไม่สำเร็จ: …`), drop the connection, and reconnect on the next request.
8. **Never die.** Wrap every loop iteration; log and continue after Supabase or network errors; sleep 2 s between iterations.
9. **Secrets.** MSSQL password and the service-role key come from `.env` or the environment; never log them.
10. **Tests.** Cover the mapper, the claim/close logic with a fake Supabase, the empty-result guard, chunking, and the connection-config builder
    (named instance = no port). Then run the equivalence check against a real ERP export (Section 12).

PostgREST notes (service-role key in both `apikey` and `Authorization: Bearer` headers; **untested here**):

| Need | Request |
|---|---|
| Upsert rows | `POST /rest/v1/<table>?on_conflict=<key columns>` with header `Prefer: resolution=merge-duplicates` and a JSON array body |
| Atomic claim | `PATCH /rest/v1/<requests>?id=eq.<id>&status=eq.pending` with `Prefer: return=representation`; an empty array means "not claimed" |
| Delete vanished rows | `DELETE /rest/v1/<table>?<scope>=eq.<x>&synced_at=lt.<timestamp>` |
| Call the web-side function (not needed by the agent) | `POST /rest/v1/rpc/<function>` |
