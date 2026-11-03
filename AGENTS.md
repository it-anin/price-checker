# AGENTS.md — Price Checker

Repository: `it-anin/price-checker`, branch `main`. Next.js 16, React 19, TypeScript, Tailwind CSS, SheetJS and Supabase.

## Required workflow

- Read this file before editing. Use scoped `rg` searches and read only the relevant code/docs below.
- **For every feature addition/change or change to an implementation approach, update `AGENTS.md` in the same commit.** Include current behavior, entry points, constraints and checks; update the detailed feature doc as needed.
- Replace stale information and link to details. Avoid duplicated docs, command logs and commit history.
- Verify code with relevant tests, `npm run build` and lint against its baseline. For docs-only changes, validate links/factual references and `git diff --check`.
- **Commit and push every verified change to `origin/main`** so Vercel deploys it. Integrate upstream first; never force-push. Report checks and commit/deployment status accurately.

## Constraints

- Views live in [app/page.tsx](app/page.tsx), intentionally large; edit in place unless asked to split it. Theme: [app/globals.css](app/globals.css).
- Do not create API routes unless explicitly asked. Existing routes are `app/api/*/route.ts`; never use or modify legacy `.js` files inside `app/api/`.
- Import the shared Supabase client from `lib/supabase.js`; do not create another client or embed credentials.
- `products.branch` is a comma-separated string (`src,kkl,sss`). PENDING JSON entries use uppercase branch arrays; these are different formats.
- Use `parsePriceRobust()` for UI price parsing/comparisons and rounded cents for pending price comparisons. Preserve valid zero prices and leading-zero SKUs.
- Strip BOM before CSV parsing. Reuse `toExportRows()` for standard product/GM MME exports; master/price exports have separate layouts. Call `preloadImages()` after product loads.
- New bulk operations use 300-item chunks with bounded `Promise.all`; retain existing limits where different.
- Mock API/storage responses for browser QA; do not modify production data to test UI.

## Read only the relevant feature

| Work area | Entry points / detailed guide |
|---|---|
| Tools, Studio Cards, navigation | `toolGroups` in `app/page.tsx`; [tools guide](docs/tools-studio-cards.md) |
| Popups, product/image editor | `StudioModal`, `notifyPopup` in `app/page.tsx`; [popup guide](docs/popup-theme.md) |
| Stock engine and CSV validation | `lib/process-engine.ts`; [engine guide](docs/process-engine.md) |
| Live MSSQL stock (button, request mailbox, office agent) | `lib/stock-source.ts`, `lib/stock-live.ts`, `tools/stock-sync/`; [bridge guide](docs/mssql-bridge.md), [starter prompt for new projects](docs/mssql-bridge-prompt.md) |
| GRAB template saves and change review | `lib/grab-template-source.ts`, `lib/grab-template-review.ts`; engine guide |
| Pending membership, automatic price checks, product Pending column | `lib/pending-sku.ts`, `lib/process-engine-source.ts`; [pending guide](docs/pending-sku.md) |
| Archived design studies | [archive instructions](design-previews/README.md) |

## Current UI and persistence

- Split Inspector shows DB, Promaxx and three GRAB prices. Tools use Studio Cards, with GRAB CSV MANAGER (`pc-tool-card`) below the four groups. Navigation uses outline icons and labels.
- Card-only product-list studies: [ten-layout HTML gallery](design-previews/product-cards/index.html), with shared mock products, search/category/status filters, branch switching, Pending details, local price-review and metadata edits, and standalone HTML exports. These previews stay outside `public`; integrate a selected layout separately. Verify all layouts on desktop/mobile and exports without external requests or production writes.
- Live MSSQL button uses **Data Orbit (05)** while `engineBusy === 'mssql'`: a dot orbits a database icon, with a stable button width, disabled repeat clicks and `aria-busy`. The existing request lifecycle clears it on completion/error; reduced motion shows a static indicator. Verify pending/running, completion/error, anonymous access and mobile with mocked API responses. The [ten-concept HTML gallery](design-previews/mssql-loading/index.html) remains an offline reference outside `public`.
- Product view omits the header subtitle, summary strip and automatic PENDING banner to give the table more vertical space. Product data/names use 14px, headers/SKUs use 12px across breakpoints, with a larger viewport-based scroll area. PENDING still runs; its product dialog and PROCESS ENGINE retain details and unsaved-result labels. Verify desktop/mobile layout and navigation with mocked responses.
- Product **New Promaxx** follows Promaxx and displays persisted `products.new_promaxx_price` (null = no calculation). `handlePriceCalcUpload()` saves R05.105 results through `POST /api/products` (`updateNewPromaxxPrice`, 300 SKUs/request, 10 concurrent updates); errors remain visible and zero-row writes are failures. Base DB prices still require explicit Update. Apply [its migration](supabase/migrations/20261102103640_add_new_promaxx_price.sql) before deploying; `tests/new-promaxx.test.mjs` covers persistence, zero/missing prices, partial failures and unchanged Promaxx/base prices. Browser QA must mock API/storage responses.
- Reuse `StudioModal` and theme tokens. Await `notifyPopup()` for notices/confirmations. Escape/close cancels deletion; preserve focus and scroll handling.
- Product **Pending** column: check = all SRC/KKL/SSS, cross = incomplete, neutral = loading/error. `pendingBranchCoverage()` reads manual + automatic membership, ignoring WRD/ONN. Clicking shows missing branches without writes. Label unsaved automatic results.

| Action | Persistence |
|---|---|
| Import GRAB SRC/KKL/SSS | Immediately syncs membership and `src_price` / `kkl_price` / `sss_price`. Use full menus: absent SKUs lose that branch. |
| Compare Promaxx | Immediately saves `promaxx_price`, not the base price. |
| Calculate/update DB price | Calculation saves `new_promaxx_price`; explicit Update writes the base price. Formula: `ceil(level0 × 0.95 × 1.20 × 1.07)`. |
| GM MME vs GRAB comparison | Browser-only comparison and exports. |
| PENDING price display | Manual and automatic sections share persisted GRAB/Promaxx checks from Tools imports; base prices load separately via SKU lookup. No local price-upload cards. |
| GRAB CSV MANAGER upload | Saves validated CSV content to `grab_files` under its canonical branch name. |

`products` uses SKU `รหัสสินค้า (SKU NUMBER)`, name `*ชื่อสินค้า (NAME)`, base price `*ราคาสินค้า`, and category `หมวดหมู่สินค้า (CATEGORIES)`. `product_master` stores master SKUs; `sku_reference` provides the legacy reference-price comparison.

## CSV contracts: stock files and price menus differ

| Input | Current reader contract |
|---|---|
| GRAB Menu prices | C=price, I=SKU, data from record 3. |
| Approved GRAB stock template | A=StoreID, C=ITEMNAME, D=ITEMCODE/SKU, E=CURRENTSTOCK, F=MAXSTOCK, P=ITEMSTATUS. No selling price. |
| R01.102 stock report | E=SKU, G=quantity; skip leading G=`CF_QUANTITY` header. Names `src.csv` / `sss.csv` / `kkl.csv`, case-insensitive. Alternative source: live MSSQL via `stockReportFromRows()` (same `StockReport`). |
| R05.103 Promaxx prices | B=SKU, E=`5` filter, G=price. |
| R05.105 calculation upload | `handlePriceCalcUpload()` reads raw CF headers: **B=CF_ITEMID (SKU), G=CF_ITEMPRICELEVEL_PRICE**, filtered by **D=CF_BASEMULTIPLE=1, F=CF_LEVELNO=0**. Detect columns by header name, not position. Also accepts the explicit compact A=SKU / E=`0` layout. Regression checks: `tests/price-calc-upload.test.mjs`. |
| GM MME Excel | A=type, B=name, C=license, D=price, E=SKU, F=image, G=category. |

- Engine/PENDING and R05.105 calculation readers reuse `decodeCsv()` / `parseEngineCsv()` for BOM, UTF-8/TIS-620, quoted fields and literal inch marks. Legacy menu importers have separate readers; verify the path being fixed.
- `parseEngineCsv()` preserves raw cells/endings and literal inch marks inside unquoted names; invalid quoted fields still fail. Do not remove name quotes.
- Manager accepts **any filename/extension** but requires the exact 16-column A–P header (case/order/spelling), data-row widths and one known StoreID. Route to canonical `GRAB_SRC.CSV` / `GRAB_SSS.CSV` / `GRAB_KKL.CSV` using the engine guide's StoreID mapping. Reject before writing; retain column warnings.

## Engine and PENDING invariants

- Manager has no LOCK/UNLOCK; Auth/RLS controls writes. `saveGrabTemplate()` compares before saving with `updated_at` concurrency checks. Side Notice opens only after changed content saves; unchanged/new files use inline status. Export canonical names with one UTF-8 BOM.
- `app_settings`, key `pending_skus`, is the PENDING source. Use `loadPendingSnapshot()` / `savePendingMutation()` for normalized reads and versioned retries. Preserve manual/automatic origins, metadata and WRD/ONN membership; `branches` is the union of `manualBranches` and `autoPriceBranches` when those fields exist.
- Automatic PENDING compares branch prices with `promaxx_price` for offered branches: add differences, release verified matches, retain missing-price/product entries. Authenticated scans persist; anonymous scans show an unsaved preview. Full `pendingPriceCheck=true` scans ignore table filters/pagination.
- RUN refreshes prices and PENDING, then reads templates. Invalid/unreadable sources stop processing. PENDING suppresses only its branch; membership changes invalidate previous outputs.
- R01.102 duplicate quantities sum; reserve defaults: SRC=5, SSS=2, KKL=2. Nonpositive/missing stock is empty. Change only raw E/F/P, preserving all other cells, headers and line endings. Export uses UTF-8 BOM and Bangkok date; DOWNLOAD ALL produces one ZIP.
- Keep retired `tools` and `grab-popups` previews outside `public` with no live links. Parser changes require regenerating preview modules with `node scripts/build-grab-manager-previews.mjs` and `node scripts/build-grab-popup-previews.mjs`; these must not republish retired galleries.

## Live MSSQL stock

- **ดึงสดจาก MSSQL** (PROCESS ENGINE, `handleEngineLoadLive()`) needs a **signed-in** user. It calls RPC `request_stock_sync()`; the office agent (`node sync.mjs --serve` on BigyaMainPC, Task `Stock Listener`) queries MSSQL at that moment, writes `stock_onhand` / `stock_sync_status`, and closes the `stock_requests` row; the page polls it (`waitForStockRequest`) then `loadStockFromDb()` -> `loadMssqlStock()` -> `stockReportFromRows()`. No timer, no API route.
- Protection: one open request, 10 s reuse, 2 min expiry, heartbeat (`stock_agent_status`, offline after 90 s -> offer stored data). A failed request keeps loaded reports untouched. CSV upload stays as fallback.
- Branch mapping SRC=0 (warehouse 1 only), SSS=5, KKL=4 lives in the agent's `.env` (`STOCK_BRANCH_MAP`). Secrets (MSSQL password, service-role key) exist only on the office PC; never commit `.env*`.
- Migrations `202611020001_stock_onhand.sql` and `202611030001_stock_requests.sql` must be run in Supabase before the button works. Tests: `tests/stock-sync.test.mjs`, `tests/stock-live.test.cjs`, `tests/process-engine.test.cjs`. Agent usage, `--compare` equivalence check and Task Scheduler: `tools/stock-sync/README.md`; architecture and reuse for other tables: [bridge guide](docs/mssql-bridge.md).

## Commands and environment

```sh
npm run dev
node --test tests/*.test.*
npm run build
npm run lint
```

The repository has existing lint violations; compare changed-file findings rather than claiming lint passed or fixing unrelated code.

`.env.local` provides `NEXT_PUBLIC_SUPABASE_URL`, `NEXT_PUBLIC_SUPABASE_ANON_KEY`, `GOOGLE_DRIVE_FOLDER_ID`, `GOOGLE_OAUTH_CLIENT_ID`, `GOOGLE_OAUTH_CLIENT_SECRET`, `GOOGLE_OAUTH_REFRESH_TOKEN`. Never copy their values into docs or output. Database setup is in [the migration](supabase/migrations/202610300001_process_engine.sql); consult the engine guide before touching it.
