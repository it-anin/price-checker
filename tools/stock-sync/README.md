# Stock Sync Agent

Vercel cannot reach the office MSSQL, so this small Node program runs **in the office**,
reads on-hand stock (the `CF_QUANTITY` column of report R01.102) and copies it to Supabase
(`stock_onhand`, `stock_sync_status`). The web app's **"ดึงจาก MSSQL"** button reads those tables.

## Setup
1. Run `supabase/migrations/202611020001_stock_onhand.sql` in the Supabase SQL editor.
2. Use a **read-only** MSSQL login: `SELECT` only on TRANS, TRANDETAILITEM, VOUCHER, ITEMS, WAREHOUSE; no write roles; not `sysadmin`.
3. `cd tools/stock-sync && npm install`
4. `cp env.example .env` and fill it in. Ask ProMaxx admin for each branch's `SYSBRANCHID`
   (`SELECT SYSBRANCHID, NAME FROM BRANCH`) and warehouse ids (`SELECT * FROM WAREHOUSE`).

## Named instance (e.g. `BigyaMainPC\SQLSERVER`)
Set `MSSQL_SERVER=BigyaMainPC` and `MSSQL_INSTANCE=SQLSERVER` (do not set a port; it is ignored).
Either start the **SQL Server Browser** service and allow UDP 1434 + the instance's TCP port in the firewall,
or give the instance a fixed port in SQL Server Configuration Manager and use `MSSQL_PORT` without `MSSQL_INSTANCE`.
Running the agent on the SQL Server machine itself (`MSSQL_SERVER=localhost`) avoids the firewall entirely.

## Run
```bash
node --env-file=.env sync.mjs --dry-run   # print SKU count / totals, write nothing
node --env-file=.env sync.mjs             # one sync (use Task Scheduler / cron)
node --env-file=.env sync.mjs --watch     # loop every SYNC_INTERVAL_MINUTES (default 15)
```

## Verify before trusting it
Export R01.102 from ProMaxx for one branch at the same moment, then compare per-SKU quantity and the
grand total with the `--dry-run` output. They must match; if not, check `warehouseIds`.

## Check the login is really read-only
```sql
-- must return no rows (no sysadmin / other server roles)
SELECT r.name FROM sys.server_role_members m
JOIN sys.server_principals r ON r.principal_id = m.role_principal_id
JOIN sys.server_principals l ON l.principal_id = m.member_principal_id WHERE l.name = 'web_readonly';
-- must fail with "permission was denied" (changes no rows anyway)
USE [MSMAXX]; EXECUTE AS USER = 'web_readonly';
UPDATE dbo.ITEMS SET ITEMNAME = ITEMNAME WHERE 1 = 0;
REVERT;
```

## Safety
- An empty result is treated as an error and never deletes existing rows.
- SKUs that disappear are deleted only after all chunks were upserted successfully.
- A failed run sets `stock_sync_status.ok=false`; the app warns when data is older than 60 minutes or the last sync failed.
