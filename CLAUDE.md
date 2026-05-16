# Price Checker — CLAUDE.md

## Overview

Next.js 16 (App Router) + React 19 + Supabase + SheetJS  
ระบบจัดการข้อมูลสินค้า, ตรวจสอบราคา GRAB, และเปรียบเทียบ GM MME

---

## Commands

```bash
npm run dev      # development server (http://localhost:3000)
npm run build    # production build
npm run lint     # ESLint
```

---

## โครงสร้างโปรเจกต์

```
app/
├── page.tsx              # UI หลักทั้งหมด (~2,100 บรรทัด)
├── layout.tsx
├── globals.css
└── api/
    ├── products/route.ts # GET/POST/PUT/PATCH/DELETE สินค้า + sync branch
    ├── stats/route.ts    # GET จำนวนสินค้า/หมวดหมู่
    ├── prices/route.ts   # GET ราคาผิดพลาด / POST อัพเดทราคา
    └── master/route.ts   # GET/POST master SKU list
lib/
└── supabase.js           # Supabase client
.env.local                # NEXT_PUBLIC_SUPABASE_URL, NEXT_PUBLIC_SUPABASE_ANON_KEY
```

> ไฟล์ `.js` ใน `app/api/` เป็น legacy ไม่ได้ใช้งาน

---

## Supabase Tables

| Table | คำอธิบาย | Key fields |
|---|---|---|
| `products` | ข้อมูลสินค้าหลัก | `รหัสสินค้า (SKU NUMBER)` (unique), `*ราคาสินค้า`, `branch`, `หมวดหมู่สินค้า (CATEGORIES)` |
| `sku_reference` | ราคาอ้างอิงสำหรับตรวจสอบ | `sku`, `ref_price` |
| `product_master` | Master SKU list | `sku` |
| `branches` | ข้อมูลสาขา (ไม่ได้ใช้ใน main flow) | `id`, `name` |

**Branch field:** ใช้ค่า comma-separated เช่น `"src,kkl,sss"` แทน boolean per branch

---

## API Routes

### `/api/products`
- **GET** `?sheet=X` — ดึงสินค้าตามหมวดหมู่ (หรือ all), aggregate branch per SKU
- **GET** `?sku=X` — ค้นหา SKU (ilike fuzzy search)
- **POST** `{ skus }` — batch fetch by SKU array (chunk 300 ต่อครั้ง)
- **POST** `{ availabilityBranch, skus }` — sync branch availability จาก GRAB CSV
- **PUT** — bulk upsert products (import Excel)
- **PATCH** — แก้ไขสินค้า 1 รายการ
- **DELETE** `{ sheet, branch }` — ลบหมวดหมู่ (branch='all' = ลบทุกสาขา)

### `/api/stats`
- **GET** — จำนวนรวม + breakdown ตามหมวดหมู่

### `/api/prices`
- **GET** — เปรียบเทียบ products กับ `sku_reference`, คืน mismatch list
- **POST** `{ items: [{sku, correctPrice}] }` — bulk update ราคา

### `/api/master`
- **GET** — จำนวน master SKU; `?compare=true` → เปรียบเทียบกับ products table; `?compareBranch=true` → คืน per-SKU branch presence `{sku, src, kkl, sss}`
- **POST** — replace master SKU list ทั้งหมด

### `/api/upload-drive`
- **POST** (multipart/form-data) `file`, `filename` — อัพโหลดรูปไป Google Drive folder ที่กำหนด, ตั้ง permission `anyone with the link`, คืน `{ success, link, id }`
- ใช้ Service Account JWT (RS256) → exchange เป็น access token → upload (multipart) → set permission
- ENV ที่ต้องตั้ง: `GOOGLE_DRIVE_FOLDER_ID`, `GOOGLE_SERVICE_ACCOUNT_EMAIL`, `GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY` (รองรับ `\n` literal)

---

## page.tsx — ฟีเจอร์หลัก

### State Groups

| กลุ่ม | State สำคัญ |
|---|---|
| Core data | `products`, `stats`, `search`, `selectedSheet`, `status`, `loading` |
| GRAB check | `grabResults`, `showGrabModal`, `grabMismatchProducts` |
| Price calc | `priceCalcResults`, `priceCalcGrabMap`, `showPriceCalcModal` |
| Import | `showImportModal`, `importSheets`, `importLog` |
| Edit product | `editProduct` (null = modal ปิด) |
| Master SKU | `masterSkuCount`, `masterMissingCount`, `masterNameMap` (sku→ชื่อ) |
| Master vs Branch | `showMasterBranchModal`, `masterBranchResults`, `masterBranchLoading`, `masterBranchTab`, `masterBranchSearch` |
| MME vs GRAB | `mmeCheckResults` (`{src,kkl,sss}`), `mmeActiveTab`, `mmeChecking` |
| Selection | `selectedSkus` (Set\<string\>) |

### ฟังก์ชันหลัก

| ฟังก์ชัน | คำอธิบาย |
|---|---|
| `handleGrabCheck(file, branch)` | อ่าน GRAB CSV → sync branch → เทียบราคากับ DB → modal |
| `handlePriceCalcUpload(file)` | อ่าน R05.105 CSV → คำนวณราคา Grab (×0.95×1.20×1.07 ปัดขึ้น) |
| `handlePriceCalcGrabUpload(file)` | อ่าน GRAB CSV สำหรับเทียบราคาใน price calc modal |
| `handleMmeVsGrabCheck()` | เทียบ MME Excel กับ GRAB CSV ทั้ง 3 สาขา |
| `handleImportXlsx(file)` | อ่าน Excel หลาย sheet, detect header อัตโนมัติ |
| `confirmImport()` | bulk upsert ผ่าน PUT /api/products |
| `confirmUpdatePrices()` | bulk update ราคา ผ่าน POST /api/prices |
| `handleMasterUpload(file)` | อ่าน ColA=SKU, ColB=ชื่อสินค้า → บันทึก masterNameMap → POST /api/master |
| `handleMasterBranchCompare()` | GET /api/master?compareBranch=true → merge กับ masterNameMap → modal |
| `handleCsvToUtf8(file)` | แปลง CSV TIS-620 → UTF-8 BOM |
| `toExportRows(list)` | format สินค้าเป็น row สำหรับ XLSX export |

### สูตรคำนวณราคา Grab
```
A = level0 × 0.95
B = A × 1.20
D = Math.ceil(B × 1.07)   ← ราคาขาย Grab
```
`level0` = ราคาระดับ 0 จากไฟล์ R05.105 (col G, filter col D=1, col F=0)

---

## รูปแบบไฟล์ Input

| ไฟล์ | Format | คอลัมน์สำคัญ |
|---|---|---|
| Grab_menu | CSV | Col C (index 2) = ราคา, Col I (index 8) = SKU, เริ่มจาก row 3 |
| GM MME | Excel (.xlsx) | Col A=ประเภท, B=ชื่อ, C=ใบอนุญาต, D=ราคา, E=SKU, F=รูป, G=หมวดหมู่ |
| R05.105 | CSV | Col B=SKU, Col D=1 (หน่วยขาย), Col F=0 (ระดับ), Col G=ราคาระดับ 0 |
| Import Excel | Excel (multi-sheet) | Header detection อัตโนมัติ 10 แถวแรก |
| Product Master | Excel (.xlsx/.xls) / CSV | Col A=SKU, Col B=ชื่อสินค้า (optional, ใช้แสดงใน branch compare modal) |

---

## Export Format (GM MME)

ไฟล์ XLSX ที่ export ใช้คอลัมน์ตาม format GM MME submission:

| Col | Header | หมายเหตุ |
|---|---|---|
| A | `*ประเภทสินค้า` | |
| B | `*ชื่อสินค้า` | |
| C | `*เลขที่ใบอนุญาตโฆษณา` | |
| D | `*ราคาสินค้า` | |
| E | `รหัสสินค้า` | |
| F | `*รูปภาพสินค้า` | มี hyperlink |
| G | `หมวดหมู่รายการสินค้า` | |

ฟังก์ชัน `toExportRows()` ใช้ format นี้ทุก export ยกเว้น master/ราคา

---

## Utility Functions

| ฟังก์ชัน | คำอธิบาย |
|---|---|
| `parseCSVRows(text)` | parse CSV auto-detect delimiter (`,` หรือ `;`) |
| `parsePriceRobust(str)` | parse ราคา รองรับ Thai format (,) |
| `convertDriveLink(url)` | แปลง Google Drive share URL → thumbnail URL |
| `renderBranchFlag(value)` | แสดง badge สาขา (src/kkl/sss) |

---

## Patterns ที่ใช้

- **Chunking:** batch API calls 200–300 items ต่อ chunk ด้วย `Promise.all`
- **Modal pattern:** `show*Modal` state + IIFE `(() => { ... })()` สำหรับ modal ที่ซับซ้อน
- **CSV BOM strip:** `text.replace(/^﻿/, '')` ทุก CSV
- **Image preload:** `preloadImages()` เรียกหลัง load products
- **Branch sync:** อัพโหลด GRAB CSV → POST `availabilityBranch` → อัพเดท `branch` field ใน Supabase
