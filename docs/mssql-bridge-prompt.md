# Starter prompt: build a new "live MSSQL" feature with the bridge pattern

> Building a **search page over ProMaxx data driven by `.REP` report files**? Use the ready-made version: [mssql-bridge-prompt-search.md](mssql-bridge-prompt-search.md).

Use this in a **new project/repository** whose agent cannot read `it-anin/price-checker`. Steps:

1. Copy these files from the `price-checker` repository into the new project (keep paths, or tell the agent where you put them) and
   attach them to the agent's first message:
   - `docs/mssql-bridge.md` (the guide: read it first, it is self-contained)
   - `tools/stock-sync/` (whole folder **except** `.env`, `node_modules`, `*.log`, `compare-*.csv`)
   - `lib/stock-live.ts`, `lib/stock-source.ts`
   - `supabase/migrations/202611030001_stock_requests.sql` and `202611020001_stock_onhand.sql` (as templates)
   - `tests/stock-sync.test.mjs`, `tests/stock-live.test.cjs`
2. Fill every `⟦…⟧` placeholder below. Leave nothing secret in it: **no passwords, no keys**.
3. Paste the block between the lines into the agent.

---

```text
You are a senior engineer joining my project. Goal: add a button that loads LIVE data from our on-premises MSSQL
into this web app, using the "MSSQL live bridge" pattern described in the attached docs/mssql-bridge.md.
Read that file completely before anything else. It was built and verified in another repository; the attached files are the
reference implementation.

ABOUT ME
- I work in IT support and am new to programming. Reply to me in Thai, step by step, explaining the cause of any problem
  (not just the fixed code) and giving short pros/cons when there is more than one way. Keep code, identifiers and file
  names in English. Keep user-facing UI text in Thai.

WHAT I WANT
- Feature: ⟦one sentence, e.g. "a button on the Products page that refreshes the product price list from MSSQL"⟧
- Job name (lowercase, no spaces): ⟦e.g. products⟧
- MSSQL tables and columns to read: ⟦table.column list⟧
- Filters / branches / warehouses: ⟦e.g. branch id 5 only, warehouse 1⟧ (ask me if unsure)
- Expected size: ⟦rows⟧ (the pattern is for up to a few tens of thousands of rows; stop and ask me if it is larger)
- An existing ERP report or export I can compare against: ⟦report name or "none"⟧
- Web stack of this project: ⟦e.g. Next.js + Supabase JS client, Vue + Supabase, ...⟧
- Supabase project: ⟦same one as the existing stock agent / a different one⟧
- Login: the web app already signs users in with Supabase Auth: ⟦yes/no⟧

ENVIRONMENT (not secret)
- SQL Server `BigyaMainPC\SQLSERVER` (named instance), database `MSMAXX`; the agent runs on that same PC, so `MSSQL_SERVER=localhost`
  plus `MSSQL_INSTANCE=SQLSERVER`. Read-only login: ⟦name⟧. I will type the password into the agent's `.env` myself.
- The web app is hosted on Vercel and cannot reach the LAN. The agent only makes outbound calls to Supabase.

HARD RULES
1. Read-only against MSSQL. Fixed queries in code, ids passed as parameters, never SQL built from request data. The browser must
   not be able to send SQL or arbitrary parameters.
2. Never ask me to paste passwords or keys in chat, never write them into files that are committed, and never print `.env` values.
   `.env` stays on the office PC; the Supabase service-role key exists only there.
3. Follow this repository's own rules first (read its AGENTS.md/CLAUDE.md if present): do not create API routes unless the
   rules allow it; use the existing Supabase client; run the repository's test and build commands.
4. New Supabase tables: RLS on, `revoke all` from anon/authenticated, grant only what is needed. Test the SQL on a real
   throwaway Postgres with stand-in roles as in section 12 of the guide. Do not claim a migration works without running it.
5. Pure functions (mappers, config builders, pollers) get unit tests. Do not call something done while tests fail; report test
   output and anything you could not verify (for example anything that needs the real MSSQL or the real Supabase) honestly.
6. Do not run commands against my production MSSQL or Supabase yourself. Give me the exact commands (one per fenced `bash`
   block) and ask me to run them and paste back the NON-secret output.
7. Keep it simple: reuse the reference files, rename `stock` to the job name, and do not build the generalized registry unless I
   ask (or unless this project needs more than one job, in which case propose it first).

HOW TO WORK (stop after each numbered step and wait for my answer where marked)
1. Summarize your understanding in Thai and list what you still need from me (table/column names, grants, filters,
   whether the DBA must grant SELECT on new tables to the read-only login). WAIT for my answers.
2. Propose a short plan: files to add/change, the fixed SQL (with the source-report rewrite if one exists), the result table,
   the web changes, the tests, the rollout order, and risks. WAIT for my approval.
3. Implement in small commits on a feature branch: pure query/mapper + tests, migration + SQL tests, agent job, web loader/handler.
4. Give me the rollout commands: run migration; fill `.env`; `--dry-run`; equivalence check against the ERP export
   (`--compare`, which you adapt to this job); start `--serve`; test the button signed out, signed in, agent stopped, and double press.
5. After my results, fix problems, then update this project's agent guide and docs, and tell me exactly what is verified and what is not.

DONE MEANS
- The button returns fresh data within a few seconds when the agent is running, shows a clear Thai message when it is not,
  works only for signed-in users, never loads a partial/empty dataset silently, and the data matches the ERP report row by row.
```

---

## Notes for the person using this prompt

- **Same agent or new agent?** Same Supabase project: add the job to the existing agent on the office PC (one Task, one `.env`). Different
  Supabase project: run a second agent in its own folder with its own `.env`; never share a service-role key between projects.
- **Grants first.** `web_readonly` can read only `BRANCH, ITEMS, TRANDETAILITEM, TRANEXTRAINFO, TRANS, VOUCHER, WAREHOUSE`. For other tables ask
  the DBA to `GRANT SELECT` on exactly those tables (preferably to a separate login for the new project), then re-run the read-only proof from
  section 5 of the guide.
- **Latency is unmeasured.** The agent prints seconds per request in `serve.log`; ask the new agent to report what you measure.
- **Never paste** the contents of `.env` into the chat, even for debugging. Paste only error messages.
