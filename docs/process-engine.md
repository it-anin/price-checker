# PROCESS ENGINE

Main navigation opens the stock workspace without discarding in-memory reports when returning to the product view. The collapsed GRAB CSV MANAGER is available in the เครื่องมือทั้งหมด menu and stores approved templates in the existing Supabase project using the shared client in `lib/supabase.js`.

R01.102 inputs are named `src.csv`, `sss.csv`, and `kkl.csv` (case-insensitive). UTF-8 is decoded strictly, with TIS-620/Windows-874 fallback. SKU E and numeric quantity G build a per-SKU stock map; duplicate SKU quantities are summed and the branch reserve is subtracted once. Defaults are SRC=5, SSS=2, KKL=2. Nonpositive results become empty stock.

Ordinary CSV exports may contain unescaped inch marks inside unquoted product names, such as `UM Finger Cot Splint E01 (M/2-3") NEW 08-26`. The shared parser treats quotes inside an already-started unquoted field as literal text, including during delimiter detection. The original name and raw cell remain intact; no cell edits or encoding conversion are required. A quote at the start of a field still opens standard CSV quoting, with escaped quotes, embedded delimiters and multiline fields supported. Unterminated quoted fields and extra text after a closing quote still block processing. Stock upload header/column/StoreID checks continue to apply.

Stock templates use SKU **D=ITEMCODE**, E=CURRENTSTOCK, F=MAXSTOCK, P=ITEMSTATUS, and branch-specific StoreIDs. The actual migrated stock exports have C=ITEMNAME and three header/instruction rows; the parser supports those and price-bearing C=PRICE exports. This differs from the separate existing price-checker Grab menu import, which uses column I.

RUN re-reads `app_settings.value` where `key='pending_skus'` before reading each branch template sequentially. It understands the original JSON array of `{sku,name,branches}` records. Missing branches default to the three active branches; legacy WRD/ONN membership is retained. Missing, invalid or unreadable pending data stops the run rather than assuming no pending SKUs.

The PENDING card supports authenticated ADD/DELETE by branch and read-only price comparison for all visitors. It uses the existing JSON row with `updated_at` concurrency checks and invalidates prior stock outputs after membership edits. Price CSVs use SKU **I**, distinct from stock templates' SKU **D**. See [PENDING SKU UPDATE](pending-sku.md) for price sources, permissions, and persistence behavior.

Pending SKUs take precedence, then matching report quantities, then missing SKUs. Only raw CSV cells E/F/P change; other cells, quoting, multiline content, all headers and original row endings are preserved. Outputs have a UTF-8 BOM and Bangkok date suffix. DOWNLOAD ALL generates a single ZIP containing all successful CSV results; reports and results stay in memory only. CLEAR resets log/progress only.

GRAB CSV MANAGER accepts any source filename, including Thai names, spaces, arbitrary extensions, and files without an extension. The file picker does not filter names or extensions, and upload validation uses only the CSV contents. Non-CSV contents still fail the strict stock-column checks. It determines the branch from column A of the data and saves under the canonical name below, regardless of the source filename. Every nonblank data row must use that same StoreID; unknown, empty, or mixed StoreIDs are rejected before writing.

| StoreID (column A) | Saved / displayed filename |
|---|---|
| `3-C7MKLKLYL8LYAJ` | `GRAB_SSS.CSV` |
| `3-C3JVCCKYNU23TA` | `GRAB_SRC.CSV` |
| `3-C7AAJUMCGAE3RE` | `GRAB_KKL.CSV` |

New uploads must match the approved stock template's machine header **exactly**, including case, spelling, and column order:

```text
STOREID,ITEMID,ITEMNAME,ITEMCODE,CURRENTSTOCK,MAXSTOCK,RESTOCKLOCALTIME,RESTOCKFREQ,ISMONDAYRESTOCK,ISTUESDAYRESTOCK,ISWEDNESDAYRESTOCK,ISTHURSDAYRESTOCK,ISFRIDAYRESTOCK,ISSATURDAYRESTOCK,ISSUNDAYRESTOCK,ITEMSTATUS
```

The machine header and every nonblank data row must have exactly 16 columns (A–P). Human-readable headers and instruction rows before the machine header are preserved, as are all CSV contents. A rejected file leaves its stored template unchanged and does not stop other files. Upload and RUN share row/StoreID validation; RUN additionally retains compatibility with legacy stock/price-bearing templates already stored under canonical names. A zero-match result displays both an alert and a visible warning before download. PROCESSED and SKIPPED count branches; UPDATED counts rows whose E/F/P values changed.

Changed headers trigger a popup and a persistent warning inside GRAB CSV MANAGER. For each rejected file, the warning lists every changed column's letter, expected name, and actual name, including added/missing columns. Reordered columns appear at each affected position. Data rows with missing/extra columns report the row and actual count. Warnings remain after refreshing the file list and reset on the next upload attempt; rejected files never overwrite approved templates.

The live manager uses popup **04 / Side Notice** for data changes after successful uploads. Each file is compared with the previously stored content before saving; the right-side modal opens only after a changed version has been saved successfully. It shows per-file counts for edited, added, and removed rows, expandable before/after cells, and branch tabs when multiple files changed. Unchanged uploads show an inline success notice. A first upload without a baseline is reported as a new file and does not claim changes from a previous version.

`lib/grab-template-review.ts` compares decoded values in all 16 stock columns by stable ItemID, falling back to normalized SKU when ItemID is missing. Duplicate matching keys block the write. Row reordering, quoting, BOM and line-ending changes alone do not trigger a data-change notice. `lib/grab-template-source.ts` reads the current content/version and uses an `updated_at` filter for updates, retrying a concurrent change up to three times. A failed read or write never appears as a saved change. Select one file per branch in an upload batch; a second file for an already-saved branch in the same batch is rejected so it cannot silently replace that batch's comparison.

The drawer initially displays 50 changed rows, with “แสดงเพิ่ม” for more. The report export always includes every changed cell. Drawer CSV/report exports use the confirmed saved snapshot, with canonical filenames and one UTF-8 BOM. The file table's Export CSV button re-reads the latest stored file before downloading. “ดูผลอัปโหลดล่าสุด” reopens the latest changed-upload result from the current browser session. Native modal focus containment, Escape/close/acknowledge dismissal and restored focus are supported; on mobile the drawer fills the viewport width.

GRAB CSV MANAGER has no LOCK/UNLOCK switch or separate unlock password. The upload button and account sign-in form are available directly when the manager is expanded. Supabase Auth sign-in still controls authenticated-only table writes. The schema migration enables RLS; it does not grant anonymous write access or create auth users.

Database schema: `supabase/migrations/202610300001_process_engine.sql`. For Supabase SQL Editor, use `/workspace/price-checker-sql-editor.zip` and execute its five numbered SQL files separately. Each file is under 400 KB. Templates import with per-file transactions and checksums; the final step publishes pending data only after all three templates pass verification. The original combined SQL exceeds the Studio query limit. The bundle has not been applied to production by this task.

Validation:

```sh
node --test tests/*.test.*
npm run build
```

The migration snapshot contains 320 pending SKUs (SRC=54, SSS=187, KKL=60) and three approved templates, validated locally before packaging. No legacy database configuration or credentials are embedded in the app or migration bundle.
