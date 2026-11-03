# Starter prompt: "ProMaxx data search" project (driven by `.REP` report files)

A ready-to-paste prompt for the **next project**: a web page that **searches data from ProMaxx (MSSQL)**, refreshed live by a button.
It is a specialised version of `docs/mssql-bridge-prompt.md`. The decisions already taken are written into the prompt:

| Decision | Value |
|---|---|
| Source of truth for the data | the ProMaxx **`.REP` report files** (the tables are not known yet) |
| What users do | **search** and read the data (anything else is undecided) |
| Web stack | Next.js + React + TypeScript + Supabase JS, hosted on Vercel |
| Supabase | a **separate** Supabase project (not the one of `price-checker`) |
| Office agent | Node (the tested reference) first; a **Go variant** is included below and is **not tested** |

## Before you paste it

1. Create the new project/repository (private) and copy these files from `price-checker` (keep the paths):
   - `docs/mssql-bridge.md` (read it first: it is self-contained; Appendix B explains `.REP` files, Appendix C is the agent contract)
   - `tools/stock-sync/` (whole folder **except** `.env`, `node_modules`, `*.log`, `compare-*.csv`)
   - `lib/stock-live.ts`, `lib/stock-source.ts`
   - `supabase/migrations/202611030001_stock_requests.sql` and `202611020001_stock_onhand.sql` (as templates; the guide's Appendix A is the generalized version)
   - `tests/stock-sync.test.mjs`, `tests/stock-live.test.cjs`
2. Put the `.REP` files you want to use in a folder named `reports/` of the new project. They contain internal table and column names: **private repository only**.
   If you open one in Notepad and see a password or a connection string, tell the agent and do not commit it.
3. Never put `.env`, passwords or keys in the project files or in the chat.
4. Paste the block between the lines below into the new agent together with the files.

---

```text
You are a senior engineer joining my project. Goal: build a web page that lets people SEARCH data from our on-premises MSSQL
(ProMaxx), using the "MSSQL live bridge" pattern in the attached docs/mssql-bridge.md. Read that file completely first, including
Appendix B (how to read ProMaxx .REP files) and Appendix C (the agent contract). It was built and verified in another repository;
the attached files are the reference implementation.

ABOUT ME
- I work in IT support and am new to programming. Reply to me in Thai, step by step, and explain the CAUSE of any problem (not only
  the fixed code). When there is more than one way, give short pros and cons. Keep code, identifiers and file names in English;
  keep user-facing UI text in Thai.

THE PRODUCT
- A search page over ProMaxx data. People type a SKU, a code or a name and see matching rows. Anything beyond search is undecided:
  do not invent extra features; ask me.
- The data is COPIED into this project's Supabase and the search runs against that copy (fast, no MSSQL call per keystroke).
  A button "ดึงสดจาก MSSQL" refreshes the copy by asking the office agent to query MSSQL at that moment (the bridge pattern).
  There is no timer: MSSQL is queried only when someone presses the button.
- Empty copy (first use): show a clear Thai message such as "ยังไม่มีข้อมูล กดดึงสดจาก MSSQL" instead of an empty table.
- Always show "ข้อมูล ณ <date time>" and warn when the copy is old or the last refresh failed.

WHERE THE DATA COMES FROM (important)
- I do NOT know the table names. I have ProMaxx report files (.REP) in the folder reports/. Each one contains the real SQL that ProMaxx
  uses for that report. Treat these files as the specification.
- Read them as described in Appendix B: use the <fsql> copy to see the filters ProMaxx injects at run time (for example warehouse,
  until-date, item), use retrieve= and the column=(...) list for column names and types, remove the PowerBuilder // comments,
  and separate what ProMaxx computes on the client (filter=, sort=, group(...), compute(...)) from what SQL returns.
- If any .REP contains a password or a connection string, STOP and tell me. Do not copy it anywhere.
- Every injected filter must become an explicit decision: a fixed value, or an allowlisted parameter. Never build SQL from text.
  Ask me for each one (for example: which branch ids / warehouses, and whether "until date" means "now").

STACK (decided)
- Web: Next.js (App Router) + React + TypeScript + the Supabase JS client, hosted on Vercel. Reuse lib/stock-live.ts and
  lib/stock-source.ts as the starting point (rename "stock" to the real job name).
- Supabase: a NEW, separate project (not the one used by price-checker). Users sign in with Supabase Auth; only signed-in users may
  press the refresh button. Remind me to turn off public sign-ups.
- Office agent: Node, from the attached tools/stock-sync (reference implementation), running on the SQL Server PC
  (BigyaMainPC) as its own folder, its own .env and its own Task Scheduler task. Never share a service-role key between projects.
  (A Go variant is described at the end of this prompt; do not use it unless I say so.)
- SQL Server: BigyaMainPC\SQLSERVER (named instance), database MSMAXX; the agent runs on that PC, so MSSQL_SERVER=localhost with
  MSSQL_INSTANCE=SQLSERVER. The current read-only user can read only BRANCH, ITEMS, TRANDETAILITEM, TRANEXTRAINFO, TRANS, VOUCHER,
  WAREHOUSE. For any other table the DBA must GRANT SELECT first: list exactly which tables you need. I will type the password into
  the agent's .env myself.
- The web app is on Vercel and cannot reach the office LAN. The agent only makes outbound calls to Supabase.

SEARCH DESIGN
- The pattern fits small to medium tables (up to a few tens of thousands of rows). Ask me for the real row counts. If a report returns
  more, STOP and ask me before choosing another approach (aggregate in SQL, limit by date, or leave that report out).
- Search against the Supabase copy: case-insensitive partial match (ilike) with a pg_trgm index for the searched text columns, debounce
  typing, server-side pagination (.range() with a stable .order()), a sensible result limit, and a clear "no result" state.
- Keep codes (SKU, item id) as TEXT so leading zeros survive; trim and upper-case keys the same way in the agent and in the web search.
- Thai text: do not strip or "normalise" Thai characters without telling me; test with real Thai names.

HARD RULES
1. Read-only against MSSQL. Fixed queries in code; ids and dates passed as typed parameters; the browser can never send SQL or free
   parameters. A refresh request carries only a job name from an allowlist.
2. Never ask me to paste passwords or keys in chat, never write them into committed files, never print .env values. .env stays on the
   office PC; the Supabase service-role key exists only there.
3. Follow this project's own rules first (read its AGENTS.md / CLAUDE.md if present): do not create API routes unless they allow it.
4. New Supabase tables: RLS on, revoke all from anon and authenticated, grant only what is needed, then test as anon. Test the SQL on a
   real throwaway Postgres with stand-in roles as in section 12 of the guide. Never claim a migration works without running it.
5. Pure functions (the .REP reader, mappers, config builders, pollers) get unit tests. Do not say "done" while tests fail. Report test
   output and everything you could not verify (anything that needs the real MSSQL or the real Supabase) honestly.
6. Do not run commands against my production MSSQL or Supabase yourself. Give me exact commands (one per fenced bash block) and ask
   me to run them and paste back the NON-secret output.
7. Keep it simple: reuse the reference files, build one report first, and add the generalized multi-job mailbox (Appendix A of the
   guide) only when a second report needs it (propose it to me first).

HOW TO WORK (stop where it says WAIT)
1. Read the guide and the attached files. Then write a small, tested, PURE reader for .REP files (put it in tools/ with unit tests) that
   prints, per report: titles, DataWindow name, the filters offered, the column list with types, the table names found in the SQL, the SQL with
   // comments removed, and the parts ProMaxx computes on the client. Run it on reports/ (I will run commands for you if you cannot).
2. Give me a table, one row per report: purpose, tables used, filters, columns, client-side parts, and your guess of the row count.
   Ask me: which report first, which branches/warehouses, what "until date" should mean, which tables the DBA must GRANT SELECT on.
   WAIT for my answers.
3. Propose a short plan: fixed SQL (rewriting correlated subqueries as GROUP BY joins where useful), the Supabase tables with a natural key
   and the pg_trgm index, the agent job, the search page, the tests, the rollout order and the risks. WAIT for my approval.
4. Implement in small commits on a feature branch: query + mapper + tests; migration + SQL tests on real Postgres; agent job; search page.
5. Give me the rollout commands: run the migration; fill the agent's .env; node ... --dry-run; the equivalence check; start --serve; then
   test the page signed out, signed in, with the agent stopped, and with a double press. The equivalence check compares the agent's data
   with an export of THE SAME ProMaxx report made with THE SAME filters (same date, warehouse, item) and must report no differences.
6. After my results, fix problems, update this project's agent guide and docs, and tell me exactly what is verified and what is not.

DONE MEANS
- Searching is fast and correct on the copy; the refresh button returns fresh data in a few seconds while the agent runs and shows a clear
  Thai message when it does not; only signed-in users can refresh; an empty or partial dataset is never loaded silently; and the data
  matches the ProMaxx report row by row.

GO AGENT VARIANT (use only if I ask for it; it has NOT been tested)
- Same contract as Appendix C of the guide, implemented in Go. Differences you must handle:
  - SQL Server driver: github.com/microsoft/go-mssqldb. A named instance is written host\instance in the connection string and needs the
    SQL Server Browser service (UDP 1434); do not send a port together with an instance name. Verify the exact connection-string
    format against the driver's documentation before using it.
  - Supabase: use plain net/http against the REST API (PostgREST), not a community client, so behaviour is explicit. Send the service-role key
    in both apikey and Authorization: Bearer headers. Use the request shapes listed at the end of Appendix C and check them against the
    PostgREST documentation.
  - Deliver a single Windows .exe plus a Task Scheduler (or Windows service) setup; read .env values from the environment; never log secrets.
  - Rewrite everything the Node reference does: heartbeat on its own timer, atomic claim, concurrent branches, chunked upsert, delete-after-write,
    empty-result guard, summaries, start-up recovery, never-die loop; plus tests for each (fake HTTP server for Supabase).
  - Keep the web app unchanged: it only talks to Supabase tables and the RPC.
  - Do not call it finished until it passes the same equivalence check as the Node agent.
```

---

## Notes for you

- **Same Supabase or a new one?** This prompt assumes a **new** one, so the agent is a second instance. If you change your mind, replace the "Supabase" and "Office agent"
  lines with: "reuse the existing Supabase project and extend the existing agent with a second job (guide, Sections 7 and 8)".
- **Grants first.** Ask the database administrator for `GRANT SELECT` on exactly the tables the agent lists after step 2 (preferably to a separate login for this project),
  then run the read-only proof from section 5 of the guide.
- **Go or Node?** Node first: it is the only agent proven end to end here. Go is a good fit later if you want a single `.exe` without installing Node on the server.
- **Latency is unmeasured.** The agent prints the seconds per request in `serve.log`; ask the new agent to report what you measure.
- **Never paste** the contents of `.env` into a chat, even for debugging. Paste only error messages.
