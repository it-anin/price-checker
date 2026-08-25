# AGENTS.md — Price Checker

> **Project context file** for AI coding agents (Cursor, Copilot, Windsurf, Cline, Aider, etc.)  
> Keep this file at the project root so agents pick it up automatically.

---

## Project Stack

| Layer | Technology |
|---|---|
| Framework | Next.js 16 (App Router) |
| UI | React 19 |
| Database | Supabase (PostgreSQL) |
| Spreadsheet parsing | SheetJS |
| Language | TypeScript (`.tsx` / `.ts`) |

---

## Dev Commands

```bash
npm run dev      # Start dev server → http://localhost:3000
npm run build    # Production build
npm run lint     # ESLint
```

---

## Project Structure

```
app/
├── page.tsx              # Main UI — all views (~2,100 lines)
├── layout.tsx
├── globals.css
└── api/
    ├── products/route.ts # CRUD products + branch sync
    ├── stats/route.ts    # Product/category counts
    ├── prices/route.ts   # Price mismatch check + bulk update
    ├── master/route.ts   # Master SKU list management
    └── upload-drive/     # Google Drive upload (OAuth refresh token)
lib/
└── supabase.js           # Supabase client singleton
.env.local                # See Environment Variables section
```

> ⚠️ Any `.js` files inside `app/api/` are **legacy — do not use or modify them**.

---

## Environment Variables (`.env.local`)

```
NEXT_PUBLIC_SUPABASE_URL=
NEXT_PUBLIC_SUPABASE_ANON_KEY=
GOOGLE_DRIVE_FOLDER_ID=
GOOGLE_OAUTH_CLIENT_ID=
GOOGLE_OAUTH_CLIENT_SECRET=
GOOGLE_OAUTH_REFRESH_TOKEN=           # scope: drive (full), from OAuth consent flow
```

---

## Database Schema (Supabase)

### `products`
| Column | Type | Notes |
|---|---|---|
| `รหัสสินค้า (SKU NUMBER)` | text | **unique**, primary identifier |
| `*ราคาสินค้า` | numeric | selling price |
| `branch` | text | comma-separated e.g. `"src,kkl,sss"` |
| `หมวดหมู่สินค้า (CATEGORIES)` | text | category / sheet grouping |

### `sku_reference`
| Column | Notes |
|---|---|
| `sku` | reference SKU |
| `ref_price` | expected price for mismatch detection |

### `product_master`
| Column | Notes |
|---|---|
| `sku` | master SKU list |

### `branches`
| Column | Notes |
|---|---|
| `id`, `name` | branch metadata — **not used in main flow** |

> **Branch encoding:** The `branch` field stores presence as a comma-separated string (`"src,kkl,sss"`), not separate boolean columns.

---

## API Reference

### `GET/POST/PUT/PATCH/DELETE /api/products`

| Method | Params / Body | Behaviour |
|---|---|---|
| GET | `?sheet=X` | Fetch products by category (or all), aggregate branch per SKU |
| GET | `?sku=X` | Fuzzy search by SKU (`ilike`) |
| POST | `{ skus }` | Batch fetch by SKU array — chunked at 300/request |
| POST | `{ availabilityBranch, skus }` | Sync branch availability from GRAB CSV upload |
| PUT | bulk array | Bulk upsert products (Excel import) |
| PATCH | single product | Edit one product record |
| DELETE | `{ sheet, branch }` | Delete by category; `branch='all'` deletes all branches |

### `GET /api/stats`
Returns total product count + per-category breakdown.

### `GET/POST /api/prices`

| Method | Params / Body | Behaviour |
|---|---|---|
| GET | — | Compare `products` vs `sku_reference`, return mismatch list |
| POST | `{ items: [{sku, correctPrice}] }` | Bulk update prices |

### `GET/POST /api/master`

| Method | Params | Behaviour |
|---|---|---|
| GET | — | Total master SKU count |
| GET | `?compare=true` | Compare master SKUs vs products table |
| GET | `?compareBranch=true` | Per-SKU branch presence `{sku, src, kkl, sss}` |
| POST | — | Replace entire master SKU list |

### `POST /api/upload-drive`
- Accepts `multipart/form-data` with fields `file` + `filename`
- Uploads to configured Google Drive folder via OAuth refresh token (user-owned files)
- Sets `anyone with the link` permission
- Returns `{ success, link, id }`

---

## Key Business Logic

### Grab Price Formula
```
A = level0 × 0.95
B = A × 1.20
D = Math.ceil(B × 1.07)   ← final Grab selling price
```
`level0` = price from R05.105 CSV (col G), filtered where col D = 1 and col F = 0.

### Input File Formats

| File | Format | Key Columns |
|---|---|---|
| `Grab_menu` | CSV | Col C (index 2) = price, Col I (index 8) = SKU; data starts row 3 |
| `GM MME` | Excel (.xlsx) | A=type, B=name, C=license, D=price, E=SKU, F=image, G=category |
| `R05.105` | CSV | B=SKU, D=1 (sale unit filter), F=0 (level filter), G=level-0 price |
| Import Excel | Excel (multi-sheet) | Auto-detect header in first 10 rows |
| Product Master | Excel / CSV | Col A = SKU, Col B = product name (optional) |

### Export Format (GM MME Submission)

| Col | Header |
|---|---|
| A | `*ประเภทสินค้า` |
| B | `*ชื่อสินค้า` |
| C | `*เลขที่ใบอนุญาตโฆษณา` |
| D | `*ราคาสินค้า` |
| E | `รหัสสินค้า` |
| F | `*รูปภาพสินค้า` (with hyperlink) |
| G | `หมวดหมู่รายการสินค้า` |

All exports use `toExportRows()` except master/price exports.

---

## Important Functions (`page.tsx`)

### Data & Sync
| Function | Purpose |
|---|---|
| `handleGrabCheck(file, branch)` | Read GRAB CSV → sync branch → compare prices vs DB → show modal |
| `handlePriceCalcUpload(file)` | Read R05.105 CSV → calculate Grab price using formula above |
| `handlePriceCalcGrabUpload(file)` | Read GRAB CSV for price comparison inside price-calc modal |
| `handleMmeVsGrabCheck()` | Compare MME Excel vs GRAB CSV across all 3 branches |
| `handleImportXlsx(file)` | Read multi-sheet Excel, auto-detect headers |
| `confirmImport()` | Bulk upsert via `PUT /api/products` |
| `confirmUpdatePrices()` | Bulk update prices via `POST /api/prices` |
| `handleMasterUpload(file)` | Read col A=SKU, col B=name → populate `masterNameMap` → POST master |
| `handleMasterBranchCompare()` | Fetch branch presence per SKU → merge with `masterNameMap` → show modal |
| `handleCsvToUtf8(file)` | Convert TIS-620 CSV → UTF-8 BOM |
| `toExportRows(list)` | Format products as XLSX export rows (GM MME format) |

### Utilities
| Function | Purpose |
|---|---|
| `parseCSVRows(text)` | Parse CSV, auto-detect delimiter (`,` or `;`) |
| `parsePriceRobust(str)` | Parse price strings including Thai number format (`,` decimal) |
| `convertDriveLink(url)` | Convert Google Drive share URL → thumbnail URL |
| `renderBranchFlag(value)` | Render branch badge(s) for `src` / `kkl` / `sss` |

---

## Coding Patterns

### Chunked API Calls
Batch requests are chunked at **200–300 items** using `Promise.all`:
```ts
const chunks = chunkArray(skus, 300);
const results = await Promise.all(chunks.map(chunk => fetchBatch(chunk)));
```

### Modal Pattern
Modals are controlled by `show*Modal` boolean state. Complex modal content uses IIFEs:
```ts
const content = (() => {
  // build modal JSX
})();
```

### CSV Handling
Always strip BOM before parsing:
```ts
text.replace(/^\uFEFF/, '')
```

### Branch Sync Flow
1. User uploads GRAB CSV
2. `POST /api/products` with `{ availabilityBranch, skus }`
3. API updates `branch` field (comma-separated) in Supabase

### Image Preloading
Call `preloadImages()` after every product load to warm the image cache.

---

## State Groups (`page.tsx`)

| Group | Key State |
|---|---|
| Core | `products`, `stats`, `search`, `selectedSheet`, `status`, `loading` |
| GRAB check | `grabResults`, `showGrabModal`, `grabMismatchProducts` |
| Price calc | `priceCalcResults`, `priceCalcGrabMap`, `showPriceCalcModal` |
| Import | `showImportModal`, `importSheets`, `importLog` |
| Edit product | `editProduct` — `null` = modal closed |
| Master SKU | `masterSkuCount`, `masterMissingCount`, `masterNameMap` (sku → name) |
| Master vs Branch | `showMasterBranchModal`, `masterBranchResults`, `masterBranchLoading`, `masterBranchTab`, `masterBranchSearch` |
| MME vs GRAB | `mmeCheckResults` (`{src, kkl, sss}`), `mmeActiveTab`, `mmeChecking` |
| Selection | `selectedSkus` — `Set<string>` |

---

## Agent Instructions

- **Do not create new API route files** unless explicitly asked — all routes live under `app/api/`.
- **Do not touch `.js` files** in `app/api/` — they are dead code.
- **`page.tsx` is intentionally large** (~2,100 lines) — prefer editing in place over splitting unless asked.
- **Branch values** are always comma-separated strings, never booleans or arrays in the DB.
- **Supabase client** must be imported from `lib/supabase.js` — do not create additional clients.
- When adding new bulk operations, follow the **chunk pattern** (300 items/chunk, `Promise.all`).
- All price comparisons must use `parsePriceRobust()` — raw string comparisons will fail on Thai-format numbers.
