# Stock Sync Agent

Vercel cannot reach the office MSSQL, so this small Node program runs **in the office**,
reads on-hand stock (the `CF_QUANTITY` column of report R01.102) and copies it to Supabase
(`stock_onhand`, `stock_sync_status`). The web app's **"ดึงจาก MSSQL"** button reads those tables.

## Setup
1. Run `supabase/migrations/202611020001_stock_onhand.sql` in the Supabase SQL editor.
2. Create a **read-only** MSSQL login (db_datareader on the ProMaxx database).
3. `cd tools/stock-sync && npm install`
4. `cp env.example .env` and fill it in. Ask ProMaxx admin for each branch's `SYSBRANCHID`
   (`SELECT SYSBRANCHID, NAME FROM BRANCH`) and warehouse ids (`SELECT * FROM WAREHOUSE`).

## Run
```bash
node --env-file=.env sync.mjs --dry-run   # print SKU count / totals, write nothing
node --env-file=.env sync.mjs             # one sync (use Task Scheduler / cron)
node --env-file=.env sync.mjs --watch     # loop every SYNC_INTERVAL_MINUTES (default 15)
```

## Verify before trusting it
Export R01.102 from ProMaxx for one branch at the same moment, then compare per-SKU quantity and the
grand total with the `--dry-run` output. They must match; if not, check `warehouseIds`.

## Safety
- An empty result is treated as an error and never deletes existing rows.
- SKUs that disappear are deleted only after all chunks were upserted successfully.
- A failed run sets `stock_sync_status.ok=false`; the app warns when data is older than 60 minutes or the last sync failed.
