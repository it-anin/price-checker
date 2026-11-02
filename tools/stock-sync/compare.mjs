// Pure helpers for `sync.mjs --compare`: read an R01.102 CSV the same way the web app does
// (lib/process-engine.ts readStockReport: SKU = column E, quantity = column G) and diff it
// against the stock read from MSSQL. No DB / network access here, so it is unit-testable.

const SKU_PATTERN = /^[A-Z0-9][A-Z0-9._/-]*$/
const NOT_SKU = new Set(['SKU', 'ITEMCODE', 'SKUNUMBER', 'TOTAL'])

export function branchFromFileName(name) {
  const base = String(name).split(/[\\/]/).pop()
  const match = /^(src|sss|kkl)\.csv$/i.exec(base)
  if (!match) throw new Error(`ชื่อไฟล์ ${base} ต้องเป็น src.csv, sss.csv หรือ kkl.csv`)
  return match[1].toLowerCase()
}

// ProMaxx exports are UTF-8 or TIS-620 (Windows-874), same fallback as the app's decodeCsv.
export function decodeBytes(bytes) {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, '')
  } catch {
    return new TextDecoder('windows-874').decode(bytes).replace(/^﻿/, '')
  }
}

export function normalizeSku(value) {
  return String(value ?? '').replace(/ /g, ' ').replace(/[​-‍﻿\r]/g, '').trim().toUpperCase()
}

export function parseQuantity(value) {
  const text = String(value ?? '').trim().replace(/,/g, '')
  if (!/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)$/.test(text)) return null
  const number = Number(text)
  return Number.isFinite(number) ? number : null
}

// Same rules as parseEngineCsv in lib/process-engine.ts: delimiter auto-detect, quoted
// multi-line fields, and bare inch marks (3") kept literally inside unquoted names.
export function parseCsv(input) {
  const text = input.replace(/^﻿/, '')
  const counts = { ',': 0, ';': 0, '\t': 0 }
  let quoted = false, fieldStarted = false
  for (let i = 0; i < Math.min(text.length, 16000); i++) {
    const ch = text[i]
    if (ch === '"' && (quoted || !fieldStarted)) {
      if (quoted && text[i + 1] === '"') { i++; continue }
      quoted = !quoted
      fieldStarted = true
    } else if (!quoted) {
      if (ch in counts) { counts[ch]++; fieldStarted = false }
      else if (ch === '\r' || ch === '\n') fieldStarted = false
      else if (ch.trim()) fieldStarted = true
    }
  }
  const delimiter = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0]
  const rows = []
  let row = [], cell = '', inQuotes = false, closed = false
  const finishCell = () => { row.push(cell); cell = ''; closed = false }
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (inQuotes) {
      if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++ }
      else if (ch === '"') { inQuotes = false; closed = true }
      else cell += ch
    } else if (ch === delimiter) {
      finishCell()
    } else if (ch === '\r' || ch === '\n') {
      finishCell()
      if (ch === '\r' && text[i + 1] === '\n') i++
      rows.push(row); row = []
    } else if (ch === '"') {
      if (closed) throw new Error(`CSV มีเครื่องหมายคำพูดไม่ถูกต้องที่แถว ${rows.length + 1}`)
      if (cell.trim()) cell += ch
      else { cell = ''; inQuotes = true }
    } else {
      if (closed && ch.trim()) throw new Error(`CSV มีข้อมูลหลังปิดเครื่องหมายคำพูดที่แถว ${rows.length + 1}`)
      if (!closed) cell += ch
    }
  }
  if (inQuotes) throw new Error('CSV มีเครื่องหมายคำพูดที่ยังไม่ปิด')
  if (cell || row.length) { finishCell(); rows.push(row) }
  return rows
}

// Returns Map<sku, qty>; duplicate SKUs are summed, like the app does.
export function readStockCsv(text) {
  const rows = parseCsv(text)
  const stock = new Map()
  let rowCount = 0
  rows.forEach((row, index) => {
    if (rowCount === 0 && row[6]?.trim().toUpperCase() === 'CF_QUANTITY') return
    const sku = normalizeSku(row[4])
    if (!SKU_PATTERN.test(sku) || NOT_SKU.has(sku)) return
    const qty = parseQuantity(row[6])
    if (qty === null) throw new Error(`CSV แถว ${index + 1}: จำนวนคอลัมน์ G ไม่ใช่ตัวเลข`)
    stock.set(sku, (stock.get(sku) ?? 0) + qty)
    rowCount++
  })
  if (!stock.size) throw new Error('ไม่พบข้อมูลใน CSV: SKU ต้องอยู่คอลัมน์ E และจำนวนเป็นตัวเลขในคอลัมน์ G')
  return stock
}

const sum = values => Number([...values].reduce((total, value) => total + value, 0).toFixed(4))

// csv, db: Map<sku, qty>. diff = db - csv, so a positive number means MSSQL has more.
export function compareStock(csv, db, tolerance = 0.0001) {
  const onlyCsv = [], onlyDb = [], differ = []
  let matched = 0
  for (const [sku, qty] of csv) {
    if (!db.has(sku)) { onlyCsv.push({ sku, csv: qty }); continue }
    const other = db.get(sku)
    if (Math.abs(other - qty) > tolerance) differ.push({ sku, csv: qty, db: other, diff: Number((other - qty).toFixed(4)) })
    else matched++
  }
  for (const [sku, qty] of db) if (!csv.has(sku)) onlyDb.push({ sku, db: qty })
  const byDiff = (a, b) => Math.abs(b.diff) - Math.abs(a.diff)
  differ.sort(byDiff)
  onlyCsv.sort((a, b) => Math.abs(b.csv) - Math.abs(a.csv))
  onlyDb.sort((a, b) => Math.abs(b.db) - Math.abs(a.db))
  return {
    csvCount: csv.size, dbCount: db.size, csvTotal: sum(csv.values()), dbTotal: sum(db.values()),
    matched, differ, onlyCsv, onlyDb,
    // SKUs that exist on one side only with qty 0 usually just mean a different filter, not wrong stock.
    onlyCsvNonZero: onlyCsv.filter(row => row.csv !== 0).length,
    onlyDbNonZero: onlyDb.filter(row => row.db !== 0).length,
    identical: !differ.length && !onlyCsv.length && !onlyDb.length,
  }
}

// Every difference as CSV (UTF-8 BOM so Excel opens it) for review outside the console.
export function differencesToCsv(result) {
  const lines = ['sku,status,csv_qty,mssql_qty,diff']
  for (const row of result.differ) lines.push(`${row.sku},qty_differs,${row.csv},${row.db},${row.diff}`)
  for (const row of result.onlyCsv) lines.push(`${row.sku},only_in_csv,${row.csv},,`)
  for (const row of result.onlyDb) lines.push(`${row.sku},only_in_mssql,,${row.db},`)
  return '﻿' + lines.join('\r\n') + '\r\n'
}
