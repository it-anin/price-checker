export const ENGINE_BRANCHES = ['src', 'sss', 'kkl'] as const
export type EngineBranch = typeof ENGINE_BRANCHES[number]

export const ENGINE_STORES: Record<EngineBranch, string> = {
  src: '3-C3JVCCKYNU23TA',
  sss: '3-C7MKLKLYL8LYAJ',
  kkl: '3-C7AAJUMCGAE3RE',
}

export interface CsvDocument {
  rows: string[][]
  rawRows: string[][]
  endings: string[]
  delimiter: string
}

export interface StockReport {
  name: string
  branch: EngineBranch
  stock: Map<string, number>
  rowCount: number
  encoding: string
}

export function reportBranch(name: string): EngineBranch {
  const match = /^(src|sss|kkl)\.csv$/i.exec(name)
  if (!match) throw new Error('ชื่อ R01.102 ต้องเป็น src.csv, sss.csv หรือ kkl.csv')
  return match[1].toLowerCase() as EngineBranch
}

export function templateBranch(name: string): EngineBranch {
  const match = /^GRAB_(SRC|SSS|KKL)\.CSV$/.exec(name)
  if (!match) throw new Error('ชื่อ Template ต้องเป็น GRAB_SRC.CSV, GRAB_SSS.CSV หรือ GRAB_KKL.CSV (ตัวพิมพ์ใหญ่)')
  return match[1].toLowerCase() as EngineBranch
}

export function templateName(branch: EngineBranch) {
  return `GRAB_${branch.toUpperCase()}.CSV`
}

export function decodeCsv(buffer: ArrayBuffer): { text: string; encoding: string } {
  const bytes = new Uint8Array(buffer)
  try {
    return { text: new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^\uFEFF/, ''), encoding: 'UTF-8' }
  } catch {
    return { text: new TextDecoder('windows-874', { fatal: true }).decode(bytes).replace(/^\uFEFF/, ''), encoding: 'TIS-620' }
  }
}

// Keep the original text of each cell, including quotes and multiline content.
// Export can then replace E/F/P without rewriting any other field or header.
export function parseEngineCsv(input: string): CsvDocument {
  const text = input.replace(/^\uFEFF/, '')
  const counts: Record<string, number> = { ',': 0, ';': 0, '\t': 0 }
  let quoted = false
  for (let i = 0; i < Math.min(text.length, 16000); i++) {
    if (text[i] === '"') {
      if (quoted && text[i + 1] === '"') { i++; continue }
      quoted = !quoted
    } else if (!quoted && text[i] in counts) counts[text[i]]++
  }
  const delimiter = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0]
  const rows: string[][] = [], rawRows: string[][] = [], endings: string[] = []
  let row: string[] = [], rawRow: string[] = [], cell = '', start = 0, inQuotes = false, closed = false
  const finishCell = (end: number) => {
    row.push(cell)
    rawRow.push(text.slice(start, end))
    cell = ''; closed = false
  }
  const finishRow = (ending: string) => {
    rows.push(row); rawRows.push(rawRow); endings.push(ending)
    row = []; rawRow = []
  }
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (inQuotes) {
      if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++ }
      else if (ch === '"') { inQuotes = false; closed = true }
      else cell += ch
    } else if (ch === delimiter) {
      finishCell(i); start = i + 1
    } else if (ch === '\r' || ch === '\n') {
      finishCell(i)
      const ending = ch === '\r' && text[i + 1] === '\n' ? '\r\n' : ch
      if (ending.length === 2) i++
      finishRow(ending); start = i + 1
    } else if (ch === '"') {
      if (closed || cell.trim()) throw new Error(`CSV มีเครื่องหมายคำพูดไม่ถูกต้องที่แถว ${rows.length + 1}`)
      cell = ''; inQuotes = true
    } else {
      if (closed && ch.trim()) throw new Error(`CSV มีข้อมูลหลังปิดเครื่องหมายคำพูดที่แถว ${rows.length + 1}`)
      if (!closed) cell += ch
    }
  }
  if (inQuotes) throw new Error('CSV มีเครื่องหมายคำพูดที่ยังไม่ปิด')
  if (start < text.length || row.length) { finishCell(text.length); finishRow('') }
  return { rows, rawRows, endings, delimiter }
}

export function stockNumber(value: string): number | null {
  const text = value.trim().replace(/,/g, '')
  if (!/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)$/.test(text)) return null
  const number = Number(text)
  return Number.isFinite(number) ? number : null
}

export function normalizeStockSku(value: unknown): string {
  return String(value ?? '').replace(/\u00A0/g, ' ').replace(/[\u200B-\u200D\uFEFF\r]/g, '').trim().toUpperCase()
}

function isStockSku(value: string) {
  return /^[A-Z0-9][A-Z0-9._/-]*$/.test(value) && !['SKU', 'ITEMCODE', 'SKUNUMBER', 'TOTAL'].includes(value)
}

export function readStockReport(name: string, text: string, encoding: string): StockReport {
  const branch = reportBranch(name)
  const { rows } = parseEngineCsv(text)
  const stock = new Map<string, number>()
  let rowCount = 0
  rows.forEach((row, index) => {
    const sku = normalizeStockSku(row[4])
    if (!isStockSku(sku)) return
    const quantity = stockNumber(row[6] ?? '')
    if (quantity === null) throw new Error(`R01.102 แถว ${index + 1}: จำนวนคอลัมน์ G ไม่ใช่ตัวเลข`)
    stock.set(sku, (stock.get(sku) ?? 0) + quantity)
    rowCount++
  })
  if (!stock.size) throw new Error('ไม่พบข้อมูล R01.102: SKU ต้องอยู่คอลัมน์ E และจำนวนเป็นตัวเลขในคอลัมน์ G')
  return { name, branch, stock, rowCount, encoding }
}

const headerKey = (value: string) => value.trim().toUpperCase().replace(/[\s_()\-.*]/g, '')

export function validateGrabTemplate(name: string, text: string) {
  const branch = templateBranch(name)
  const document = parseEngineCsv(text)
  // The real stock exports have C=ITEMNAME (no price column), with a human
  // header, instructions, and a machine header. Price-bearing exports are
  // supported too. Locate data by the item code, then use the final header
  // before those records, without mistaking the instruction row for data.
  const skuColumn = 3
  const dataStart = document.rows.findIndex(row => isStockSku(normalizeStockSku(row[skuColumn])) && Boolean(row[2]?.trim()))
  let headerIndex = -1
  for (let index = 0; index < (dataStart < 0 ? document.rows.length : dataStart); index++) {
    const row = document.rows[index]
    if (headerKey(row[3] ?? '') === 'ITEMCODE' && headerKey(row[4] ?? '') === 'CURRENTSTOCK' && headerKey(row[5] ?? '') === 'MAXSTOCK') headerIndex = index
  }
  if (headerIndex < 0) throw new Error('โครงสร้างไม่ถูกต้อง: D ต้องเป็น ITEMCODE, E เป็น CURRENTSTOCK, F เป็น MAXSTOCK')
  const header = document.rows[headerIndex]
  if (header.length < 16 || !header[15]?.trim()) throw new Error('ไม่พบหัวคอลัมน์สถานะสต็อก P')
  const priceColumn = ['PRICE', 'ITEMPRICE', 'ราคาสินค้า'].includes(headerKey(header[2] ?? ''))
  if (!priceColumn && !['ITEMNAME', 'NAME', 'ชื่อสินค้า'].includes(headerKey(header[2] ?? ''))) throw new Error('คอลัมน์ C ต้องเป็น ITEMNAME หรือ PRICE')
  if (dataStart < 0) throw new Error('ไม่พบแถวข้อมูล GRAB ที่มี SKU ในคอลัมน์ D')
  const dataIndices: number[] = []
  for (let index = dataStart; index < document.rows.length; index++) {
    const row = document.rows[index]
    if (row.every(cell => !cell.trim())) continue
    if (row.length < 16) throw new Error(`แถว ${index + 1}: คอลัมน์ไม่ครบถึง P`)
    if (row[0].trim() !== ENGINE_STORES[branch]) throw new Error(`แถว ${index + 1}: StoreID ไม่ตรงกับ ${branch.toUpperCase()} (${ENGINE_STORES[branch]})`)
    if (!isStockSku(normalizeStockSku(row[skuColumn])) || (priceColumn ? stockNumber(row[2] ?? '') === null : !row[2]?.trim())) throw new Error(`แถว ${index + 1}: SKU หรือข้อมูลคอลัมน์ C ไม่ถูกต้อง`)
    dataIndices.push(index)
  }
  // Check IDs before the first data row too, so malformed leading records cannot hide there.
  for (let index = headerIndex + 1; index < dataStart; index++) {
    const row = document.rows[index]
    if (row[0]?.trim().startsWith('3-') || isStockSku(normalizeStockSku(row[skuColumn]))) throw new Error(`แถว ${index + 1}: ข้อมูลสินค้าก่อนแถวเริ่มต้นไม่ถูกต้อง`)
  }
  return { branch, document, headerIndex, dataStart, dataIndices, skuColumn }
}

export function processGrabTemplate(name: string, text: string, report: StockReport, subtract: number, pending: Set<string>) {
  if (!Number.isFinite(subtract) || subtract < 0) throw new Error('SUBTRACT ต้องเป็นตัวเลขตั้งแต่ 0 ขึ้นไป')
  const template = validateGrabTemplate(name, text)
  if (template.branch !== report.branch) throw new Error('สาขาของ R01.102 และ Template ไม่ตรงกัน')
  const { document, dataIndices, skuColumn } = template
  const rawRows = document.rawRows.map(row => [...row])
  let matched = 0, updated = 0, inStock = 0, pendingCount = 0
  for (const index of dataIndices) {
    const row = document.rows[index], sku = normalizeStockSku(row[skuColumn])
    if (report.stock.has(sku)) matched++
    if (pending.has(sku)) pendingCount++
    const quantity = pending.has(sku) || !report.stock.has(sku) ? 0 : report.stock.get(sku)! - subtract
    const current = quantity > 0 ? String(Number(quantity.toFixed(6))) : ''
    const status = current ? 'in stock' : 'out-of-stock indefinitely'
    if (current) inStock++
    if (row[4] !== current || row[5] !== current || row[15] !== status) updated++
    rawRows[index][4] = current
    rawRows[index][5] = current
    rawRows[index][15] = status
  }
  const content = '\uFEFF' + rawRows.map((row, index) => row.join(document.delimiter) + document.endings[index]).join('')
  return { content, matched, updated, inStock, pendingCount, total: dataIndices.length }
}

export function engineDate(date = new Date()) {
  const parts = new Intl.DateTimeFormat('en-GB', { timeZone: 'Asia/Bangkok', day: '2-digit', month: '2-digit', year: 'numeric' }).formatToParts(date)
  return ['day', 'month', 'year'].map(type => parts.find(part => part.type === type)?.value).join('-')
}

export function pendingSkuSets(raw: unknown): Record<EngineBranch, Set<string>> {
  if (!Array.isArray(raw)) throw new Error('pending_skus ต้องเป็น JSON array')
  const result: Record<EngineBranch, Set<string>> = { src: new Set(), sss: new Set(), kkl: new Set() }
  for (const item of raw) {
    if (!item || typeof item !== 'object' || !String(item.sku ?? '').trim()) throw new Error('pending_skus มีรายการ SKU ว่างหรือรูปแบบไม่ถูกต้อง')
    const sku = normalizeStockSku(item.sku)
    if (item.branches !== undefined && !Array.isArray(item.branches)) throw new Error(`pending_skus: branches ของ ${sku} ต้องเป็น array`)
    // Legacy records without branches apply to all three active branches.
    // Preserve WRD/ONN membership without making those records active elsewhere.
    const branches: string[] = item.branches?.length ? item.branches.map((branch: unknown) => String(branch).trim().toLowerCase()) : [...ENGINE_BRANCHES]
    if (branches.some(branch => ![...ENGINE_BRANCHES, 'wrd', 'onn'].includes(branch))) throw new Error(`pending_skus: สาขาของ ${sku} ไม่ถูกต้อง`)
    for (const branch of ENGINE_BRANCHES) if (branches.includes(branch)) result[branch].add(sku)
  }
  return result
}

// A single ZIP avoids the browser blocking multiple automatic CSV downloads.
// Store (uncompressed) entries: no dependency and no changes to the UTF-8 CSV bytes.
export function createResultsZip(files: { name: string; content: string }[]): Uint8Array<ArrayBuffer> {
  const encoder = new TextEncoder()
  const chunks: Uint8Array[] = [], central: Uint8Array[] = []
  let offset = 0, centralSize = 0
  const crc32 = (data: Uint8Array) => {
    let crc = 0xffffffff
    for (const byte of data) {
      crc ^= byte
      for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1))
    }
    return (crc ^ 0xffffffff) >>> 0
  }
  files.forEach(file => {
    const name = encoder.encode(file.name), data = encoder.encode(file.content), crc = crc32(data)
    const local = new Uint8Array(30 + name.length), view = new DataView(local.buffer)
    view.setUint32(0, 0x04034b50, true); view.setUint16(4, 20, true); view.setUint16(6, 0x800, true)
    view.setUint16(12, 0x21, true); view.setUint32(14, crc, true); view.setUint32(18, data.length, true); view.setUint32(22, data.length, true); view.setUint16(26, name.length, true)
    local.set(name, 30)
    const entry = new Uint8Array(46 + name.length), entryView = new DataView(entry.buffer)
    entryView.setUint32(0, 0x02014b50, true); entryView.setUint16(4, 20, true); entryView.setUint16(6, 20, true); entryView.setUint16(8, 0x800, true)
    entryView.setUint16(14, 0x21, true); entryView.setUint32(16, crc, true); entryView.setUint32(20, data.length, true); entryView.setUint32(24, data.length, true); entryView.setUint16(28, name.length, true); entryView.setUint32(42, offset, true)
    entry.set(name, 46); chunks.push(local, data); central.push(entry)
    offset += local.length + data.length; centralSize += entry.length
  })
  const end = new Uint8Array(22), endView = new DataView(end.buffer)
  endView.setUint32(0, 0x06054b50, true); endView.setUint16(8, files.length, true); endView.setUint16(10, files.length, true); endView.setUint32(12, centralSize, true); endView.setUint32(16, offset, true)
  const result = new Uint8Array(offset + centralSize + end.length)
  let position = 0
  for (const chunk of [...chunks, ...central, end]) { result.set(chunk, position); position += chunk.length }
  return result
}
