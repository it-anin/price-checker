# PENDING SKU UPDATE

The PENDING card in PROCESS ENGINE manages products waiting for Grab approval. Read-only visitors can inspect entries and compare prices. ADD and branch closure use the same Supabase Auth session as GRAB CSV MANAGER.

## Pending column on the price-checking screen

The table in **ตรวจสอบราคา** shows Pending coverage beside each SKU. **รายการสินค้า** uses product-only Image Catalogue cards. A green check means the SKU is in PENDING for all three active branches (SRC, KKL and SSS); a red cross means one or more are missing, including SKUs absent from PENDING entirely. Coverage includes the union of manual and automatic membership. Legacy WRD/ONN branches do not count toward the three required branches, while older entries without a branch list keep the existing all-three default. SKU matching uses the same normalization as stock processing, preserving leading zeroes.

Click the status to open a small Studio popup listing the missing branches and each branch's current membership. It reads the existing PENDING snapshot and updates when that snapshot changes, including after ADD/closing work or automatic price resolution. Opening it does not edit PENDING or product data. Loading and failed refreshes use a neutral status rather than claiming incomplete coverage; errors can be retried from the popup. Unsaved automatic results are identified with the same notice as the PENDING workspace.

## Unified branch table and closing work

SRC/SSS/KKL use one tab set and one table. Each branch shows one row per active SKU, combining saved manual and automatic origins. The table shows DB, GRAB and Promaxx prices, with a reason badge for manually added rows, price differences or both. Automatic entries awaiting complete price data are labeled separately. Counts exclude closed records; legacy WRD/ONN membership remains intact. The selected branch stays selected when its last row is closed.

Select approval branches (all three initially) and paste SKUs separated by newlines, commas, semicolons, tabs or pipes. Enter submits; Shift+Enter adds a line. ADD merges branches without duplicates and explicitly reopens a selected closed branch. Names may be filled from approved stock templates after the save; lookup never changes prices.

Signed-in users can choose **นำออกจาก PENDING** on any row. It removes both manual and automatic membership for that SKU in the active branch, records the displayed GRAB/Promaxx pair in `resolvedPrices`, and preserves other branches and metadata. This releases that branch's stock hold; subsequent RUN uses normal stock rules there. Identical prices do not recreate the row after refresh, signing in again or RUN. A changed pair clears its acknowledgement; a current mismatch then reopens automatic membership. A verified match clears only automatic membership, retaining any manual hold. Missing values are stored as null, and comparisons use rounded integer cents, preserving zero.

Closed records remain in the existing `pending_skus` JSON with `branches: []`, empty source arrays and nonempty `resolvedPrices`, or with other still-active branches. The normalizer validates price pairs and prevents active/closed overlap. Stock sets, active counts, database price lookups and exports exclude empty-branch records; legacy empty branch arrays without this metadata retain the all-three default. The existing versioned save helper retries concurrent edits and preserves a peer closure when reconciling the same prices. Failed saves leave membership unchanged and require a successful reload before retrying.

## Export GM MME by branch

The card's **Export GM MME · SRC/SSS/KKL** button exports the active branch; the single branch tab set selects the export scope. Each click re-reads the saved `pending_skus` snapshot and looks up current `products` details through the existing read-only SKU API, in batches of 300 with at most four concurrent requests. It includes both saved membership sources, once per SKU, independently of product search/category/pagination. Anonymous automatic previews that have not been saved are excluded.

The workbook uses the sample's visible `TEMPLATE` layout: `*ประเภทสินค้า`, `*ชื่อสินค้า`, `*เลขที่ใบอนุญาตโฆษณา`, `*ราคาสินค้า`, `รหัสสินค้า`, `*รูปภาพสินค้า`, `หมวดหมู่รายการสินค้า`. Rows contain Supabase product data, base DB prices as numbers (including zero), text SKUs with leading zeroes, and original image links. Column F keeps the URL as its text and adds a native Excel hyperlink for HTTP(S) links, preserving Google Drive sharing/resource keys; empty URLs remain unlinked. Column widths, row heights and the two-decimal price format follow the sample. Instruction/example cells and unrelated hidden helper sheets are not exported as product data. The filename identifies the branch and Bangkok date.

Export does not write to Supabase or alter PENDING. Empty branches produce no file. Missing product records are excluded with a visible exported/total count and missing SKU warning; a failed snapshot or product batch prevents any partial download and permits retry. Export shares the Engine busy lock with edits and RUN.

## Automatic GRAB / Promaxx pending

Automatic membership compares persisted `src_price`, `sss_price` and `kkl_price` against `promaxx_price`, only for branches where the product is offered. It uses the same parser and rounded cents as the price-checking screen. Zero is valid. Missing or invalid prices do not create a new pending item and do not release an existing automatic pending item. An automatic branch is released when both valid prices match or an authenticated user closes that branch. Without an explicit closure, missing products or branches retain existing suppression until prices can be verified.

Price-checking product loads, price imports, opening PROCESS ENGINE and reloading PENDING trigger a full price scan. While the product-only Image Catalogue is active, scans are deferred until entering price checking, Tools or Engine so they do not compete with product images. The existing `GET /api/products?pendingPriceCheck=true` reads all SKUs in ordered pages, independently of the visible search, category and page. A failed page aborts the scan. RUN repeats this scan before calculating stock, and stops if either prices or pending entries cannot be read.

Authenticated scans persist automatic changes with the same RLS and CAS helper as manual edits. Read-only visitors see an explicitly marked in-memory result, which RUN also uses; signing in persists it automatically. The unified table shares these persisted GRAB/Promaxx price checks; PROCESS ENGINE has no separate price-file upload controls.

Existing entries without source fields remain manual. `manualBranches` and `autoPriceBranches` retain origins, and `branches` remains their union. A row can have both origins; the table and branch badges count it once. Closing work clears both origins only in that branch. Automatic resolution of a matched price preserves manual and legacy WRD/ONN membership. Names from product records never overwrite an existing PENDING name.

## Three price sources

- **GRAB:** persisted `src_price` / `sss_price` / `kkl_price` from the branch import tools in “เครื่องมือทั้งหมด”, read through `pendingAutoPrices` for the active branch. Only offered branches with valid prices are shown. Stock templates are not price files.
- **DATABASE:** current `products` selling price `*ราคาสินค้า`. Uses the existing read-only `POST /api/products` SKU lookup in parallel batches of at most 300. Loads on opening/reloading PENDING, after ADD/closing work, and when RUN refreshes the pending snapshot. Failed refreshes keep the previous prices with an error message beside the table.
- **PROMAX:** persisted `promaxx_price`, saved by the R05.103 comparison tool in “เครื่องมือทั้งหมด” and read through the same checks as GRAB. Update it through that tool rather than uploading again in PROCESS ENGINE.

Prices use the existing `parsePriceRobust` and compare rounded cents. `·` means prices are not loaded; `—` means no valid saved price for the SKU/branch; zero displays `0.00`. GRAB differences against either reference are marked with a red row and `≠`, with per-branch counts. The unified table uses the full price-scan refresh flow above; no local price maps, CLEAR buttons or separate DATABASE refresh card remain.

## Safe persistence and RUN

The existing `app_settings` row `key='pending_skus'` remains the source of truth. `loadPendingSnapshot()` returns normalized entries plus `updatedAt`. `savePendingMutation()` handles `add`, `resolveBranch`, legacy `removeBranch`, `enrichNames` and `syncPrices`: it re-reads current data and updates with an `updated_at` equality filter, advancing the timestamp beyond its prior value. On a version conflict it reloads/reapplies the same operation, up to three attempts. Every writer in this UI uses this path; external/manual writers should also advance `updated_at`.

Failed loads, malformed data and failed saves block subsequent edits until a successful reload. Input is retained when ADD fails. Membership edits invalidate previously generated stock outputs so they must be generated again; this also applies when an explicit reload observes another user's changes. Name/metadata-only updates do not invalidate stock results. Writes are serialized against RUN and template uploads.

RUN always reads current pending data, applies suppression only to the SKU's branch, and stops if the source cannot be read. Authenticated runs may fill missing names using the same mutation helper; a name-save failure is logged and does not discard otherwise valid stock outputs. If name enrichment observes changed active membership during RUN, outputs are withheld and the user is asked to run again. No schema migration or new API route is required.

## Verification

```sh
node --test tests/process-engine.test.cjs tests/pending-sku.test.cjs tests/pending-price-api.test.cjs tests/pending-export.test.mjs
npm run build
```

Tests cover unified manual/automatic membership, persistent branch closure, reopening after changed prices, closed-record stock/export exclusions, legacy metadata, CAS retries/failures, invalid sources, timestamp monotonicity, column-I price parsing, actual stock-template names, TIS-620, cent comparisons, full price-scan pagination/failures and branch-specific stock suppression. Browser QA exercises the UI with mocked API responses and performs no production writes.
