import { ENGINE_BRANCHES, ENGINE_STORES, normalizeStockSku, parseEngineCsv, validateGrabTemplate, type EngineBranch } from './process-engine'

export const PENDING_BRANCHES = ['SRC', 'SSS', 'KKL', 'WRD', 'ONN'] as const
export type PendingBranch = typeof PENDING_BRANCHES[number]
export interface PendingSkuEntry {
  [key: string]: unknown
  sku: string
  name: string
  branches: PendingBranch[]
}
export interface PendingSnapshot { entries: PendingSkuEntry[]; updatedAt: string }
export type PendingMutation =
  | { kind: 'add'; skus: string[]; branches: EngineBranch[] }
  | { kind: 'removeBranch'; sku: string; branch: EngineBranch }
  | { kind: 'enrichNames'; names: Map<string, string> }
export type PendingPriceParser = (value: string) => number | null
export interface PendingPriceFile { name: string; prices: Map<string, number>; names: Map<string, string>; encoding?: string }

export function parsePendingInput(text: string): string[] {
  return [...new Set(text.split(/[\n\r,;|\t]+/).map(normalizeStockSku).filter(Boolean))]
}

export function normalizePendingEntries(value: unknown): PendingSkuEntry[] {
  let raw = value
  if (typeof raw === 'string') {
    try { raw = JSON.parse(raw) } catch { throw new Error('pending_skus ไม่ใช่ JSON ที่ถูกต้อง') }
  }
  if (!Array.isArray(raw)) throw new Error('pending_skus ต้องเป็น JSON array')
  const seen = new Set<string>()
  return raw.map(item => {
    if (!item || typeof item !== 'object' || !['string', 'number'].includes(typeof item.sku)) throw new Error('pending_skus มีรายการ SKU ไม่ถูกต้อง')
    const sku = normalizeStockSku(item.sku)
    if (!sku || seen.has(sku)) throw new Error('pending_skus มี SKU ว่างหรือซ้ำ กรุณาตรวจข้อมูลก่อนบันทึก')
    seen.add(sku)
    if (item.branches !== undefined && !Array.isArray(item.branches)) throw new Error(`branches ของ ${sku} ต้องเป็น array`)
    const branches: PendingBranch[] = item.branches?.length
      ? [...new Set<string>(item.branches.map((branch: unknown) => String(branch).trim().toUpperCase()))].map(branch => {
        if (!PENDING_BRANCHES.includes(branch as PendingBranch)) throw new Error(`สาขาของ ${sku} ไม่ถูกต้อง`)
        return branch as PendingBranch
      }) : ['SRC', 'SSS', 'KKL']
    return { ...item, sku, name: typeof item.name === 'string' ? item.name : '', branches }
  })
}

export function mutatePendingEntries(entries: PendingSkuEntry[], mutation: PendingMutation) {
  let changed = false
  let next = entries.map(entry => ({ ...entry, branches: [...entry.branches] }))
  if (mutation.kind === 'add') {
    if (!mutation.branches.length || mutation.branches.some(branch => !ENGINE_BRANCHES.includes(branch))) throw new Error('เลือกสาขาอย่างน้อย 1 สาขา')
    const skus = [...new Set(mutation.skus.map(normalizeStockSku).filter(Boolean))]
    if (!skus.length) throw new Error('กรุณากรอก SKU')
    for (const sku of skus) {
      const entry = next.find(item => item.sku === sku)
      const selected = mutation.branches.map(branch => branch.toUpperCase() as PendingBranch)
      if (entry) {
        const combined = PENDING_BRANCHES.filter(branch => entry.branches.includes(branch) || selected.includes(branch))
        if (combined.some(branch => !entry.branches.includes(branch))) { entry.branches = combined; changed = true }
      } else { next.push({ sku, name: '', branches: selected }); changed = true }
    }
  } else if (mutation.kind === 'removeBranch') {
    const sku = normalizeStockSku(mutation.sku), branch = mutation.branch.toUpperCase() as PendingBranch
    if (!ENGINE_BRANCHES.includes(mutation.branch)) throw new Error('สาขาสำหรับ DELETE ไม่ถูกต้อง')
    next = next.flatMap(entry => {
      if (entry.sku !== sku || !entry.branches.includes(branch)) return [entry]
      changed = true
      const branches = entry.branches.filter(value => value !== branch)
      return branches.length ? [{ ...entry, branches }] : []
    })
  } else {
    next = next.map(entry => {
      const name = mutation.names.get(entry.sku)?.trim()
      if (entry.name.trim() || !name) return entry
      changed = true
      return { ...entry, name }
    })
  }
  return { entries: next, changed }
}

export function nextPendingTimestamp(previous: string, now = Date.now()): string {
  const timestamp = Date.parse(previous)
  if (!Number.isFinite(timestamp)) throw new Error('updated_at ของ PENDING ไม่ถูกต้อง กรุณาโหลดใหม่ก่อนบันทึก')
  return new Date(Math.max(now, timestamp + 1)).toISOString()
}

export function samePendingStockMembership(left: PendingSkuEntry[], right: PendingSkuEntry[]): boolean {
  const active = new Set<string>(ENGINE_BRANCHES.map(branch => branch.toUpperCase()))
  const membership = (entries: PendingSkuEntry[]) => new Set(entries.flatMap(entry => entry.branches.filter(branch => active.has(branch)).map(branch => `${branch}:${entry.sku}`)))
  const before = membership(left), after = membership(right)
  return before.size === after.size && [...before].every(key => after.has(key))
}

const key = (value: string) => value.trim().toUpperCase().replace(/[\s_()*./-]/g, '')
const nameHeaders = ['NAME', 'ITEMNAME', 'PRODUCTNAME', 'ชื่อสินค้า', 'ชื่อสินค้าNAME']
const priceHeaders = ['PRICE', 'ITEMPRICE', 'SELLINGPRICE', 'ราคาสินค้า']
const skuHeaders = ['SKU', 'SKUNUMBER', 'ITEMCODE', 'BARCODE', 'รหัสสินค้า', 'รหัสสินค้าSKUNUMBER']
function finitePrice(value: string, parse: PendingPriceParser) {
  const price = parse(value)
  return price !== null && Number.isFinite(price) ? price : null
}

export function readPendingGrabPrices(name: string, text: string, parse: PendingPriceParser) {
  const match = /^GRAB_(SRC|SSS|KKL)\.CSV$/i.exec(name)
  if (!match) throw new Error('ไฟล์ราคาต้องชื่อ GRAB_SRC.CSV, GRAB_SSS.CSV หรือ GRAB_KKL.CSV')
  const branch = match[1].toLowerCase() as EngineBranch
  const { rows } = parseEngineCsv(text)
  const headerIndex = rows.findIndex(row => nameHeaders.includes(key(row[1] ?? '')) && priceHeaders.includes(key(row[2] ?? '')) && skuHeaders.includes(key(row[8] ?? '')))
  if (headerIndex < 0) throw new Error('หัวตารางไฟล์ราคาไม่ถูกต้อง: B=NAME, C=PRICE, I=SKU (ไฟล์ Template สต็อกใช้แทนไฟล์ราคาไม่ได้)')
  const checkStore = key(rows[headerIndex][0] ?? '') === 'STOREID' || rows.slice(headerIndex + 1).some(row => row[0]?.trim().startsWith('3-'))
  const prices = new Map<string, number>(), names = new Map<string, string>()
  for (let index = headerIndex + 1; index < rows.length; index++) {
    const row = rows[index], sku = normalizeStockSku(row[8]), price = finitePrice(row[2] ?? '', parse)
    if (!sku || price === null) continue
    if (checkStore && row[0]?.trim() !== ENGINE_STORES[branch]) throw new Error(`แถว ${index + 1}: StoreID ไม่ตรงกับ ${branch.toUpperCase()}`)
    prices.set(sku, price)
    if (row[1]?.trim()) names.set(sku, row[1].trim())
  }
  if (!prices.size) throw new Error('ไม่พบราคาในคอลัมน์ C และ SKU ในคอลัมน์ I')
  return { branch, name, prices, names }
}

export function readPendingPromaxPrices(name: string, text: string, parse: PendingPriceParser): PendingPriceFile {
  const { rows } = parseEngineCsv(text)
  const prices = new Map<string, number>()
  for (const row of rows.slice(1)) {
    if (row[4]?.trim() !== '5') continue
    const sku = normalizeStockSku(row[1]), price = finitePrice(row[6] ?? '', parse)
    if (sku && price !== null) prices.set(sku, price)
  }
  if (!prices.size) throw new Error('ไม่พบราคา R05.103: B=SKU, E=5, G=ราคา')
  return { name, prices, names: new Map() }
}

export function readTemplateNames(name: string, text: string): Map<string, string> {
  const template = validateGrabTemplate(name, text)
  const header = template.document.rows[template.headerIndex]
  const nameColumn = nameHeaders.includes(key(header[2] ?? '')) ? 2 : 1
  return new Map(template.dataIndices.map(index => {
    const row = template.document.rows[index]
    return [normalizeStockSku(row[3]), row[nameColumn].trim()] as const
  }).filter(([, name]) => Boolean(name)))
}

export function comparePendingPrices(a: number | undefined, b: number | undefined) {
  return a === undefined || b === undefined || !Number.isFinite(a) || !Number.isFinite(b) ? 'na' : Math.round(a * 100) === Math.round(b * 100) ? 'same' : 'diff'
}
