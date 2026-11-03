# Stock Sync Agent

Vercel cannot reach the office MSSQL, so this small Node program runs **in the office** and answers the web app:

1. A signed-in user presses **"ดึงสดจาก MSSQL"** in PROCESS ENGINE.
2. The web app asks Supabase (`request_stock_sync()`), which creates one `stock_requests` row.
3. This agent (`--serve`) sees the request within ~2 s, queries MSSQL **right then** (stock of report R01.102,
   column `CF_QUANTITY`), writes `stock_onhand` / `stock_sync_status` and closes the request.
4. The web app reads the fresh rows. Nothing queries MSSQL on a timer: only a button press does.

## Setup
1. Run both migrations in the Supabase SQL editor, in order: `supabase/migrations/202611020001_stock_onhand.sql`, then `202611030001_stock_requests.sql`.
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
node --env-file=.env sync.mjs --serve     # normal mode: stay running, sync live when the web app asks
node --env-file=.env sync.mjs --dry-run   # print SKU count / totals, write nothing
node --env-file=.env sync.mjs             # one sync of every branch, then exit
```
`--serve` prints one line per request (`[serve] done 4.2 วินาที · src 5461, sss 3287, kkl 4007`) and writes a
heartbeat every 30 s so the web app can tell when the agent is offline (no beat for 90 s).

## Keep it running on Windows (Task Scheduler, starts at boot, restarts on failure)
`run-serve.cmd` starts the listener and appends its output to `serve.log`. If `node.exe` is not in
`C:\Program Files\nodejs`, edit the path inside it (`where node` shows it). Then, in **PowerShell as Administrator**
(change `$folder` to where this folder really is):

```powershell
$folder = 'C:\Users\AninMainPC\Automation\price-checker\tools\stock-sync'
$action = New-ScheduledTaskAction -Execute "$folder\run-serve.cmd"
$trigger = New-ScheduledTaskTrigger -AtStartup
$settings = New-ScheduledTaskSettingsSet -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName 'Stock Listener' -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force
Start-ScheduledTask -TaskName 'Stock Listener'
```
Check: `Get-ScheduledTask 'Stock Listener' | Get-ScheduledTaskInfo` and `type serve.log`.
Stop / remove: `Stop-ScheduledTask 'Stock Listener'`, `Unregister-ScheduledTask 'Stock Listener' -Confirm:$false`.
If you used the old 15-minute task, remove it: `schtasks /Delete /TN "Stock Sync" /F`.

## Verify before trusting it
Export R01.102 from ProMaxx for each branch **right before** comparing (stock keeps moving), save them as
`src.csv`, `sss.csv`, `kkl.csv`, then run (read-only, writes nothing to Supabase):

```bash
node --env-file=.env sync.mjs --compare src.csv sss.csv kkl.csv
```

It reads each CSV exactly like the web app (SKU = column E, quantity = column G, TIS-620/UTF-8 auto-detect)
and prints, per branch: SKU counts, grand totals, matched SKUs, SKUs whose qty differ (`diff = MSSQL - CSV`),
and SKUs present on one side only. Differences are also saved to `compare-<branch>.csv` in the current folder.
Exit code is 0 only when every SKU matches.

| Result | Meaning |
|---|---|
| all branches `✔ ตรงกันทุก SKU` | the query is equivalent to R01.102 |
| few small diffs | stock moved between the export and the run: export again and re-run |
| many SKUs only in MSSQL (non-zero) | the CSV was exported for fewer warehouses: set `warehouseIds` in `STOCK_BRANCH_MAP` |
| many SKUs only in CSV | wrong `branchId`, or the CSV is from another branch |

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
- Only signed-in users can ask (`request_stock_sync()` rejects everyone else); the browser cannot write these tables directly.
- One request runs at a time: a second press (or another user) shares the running request, and a request that finished
  less than 10 s ago is reused, so repeated presses cannot hammer ProMaxx. Unanswered requests expire after 2 minutes.
- The agent runs only the fixed query in `query.mjs`; a request carries no SQL or parameters.
- An empty result is treated as an error and never deletes existing rows.
- SKUs that disappear are deleted only after all chunks were upserted successfully.
- A failed run sets `stock_sync_status.ok=false` and the request `status=error` with the reason shown in the web app.
- If the agent is offline the web app says so and offers the last stored data (with its age) instead of failing silently.
