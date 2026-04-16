'use client'

import { useState, useEffect } from 'react'

interface Product {
  [key: string]: any
}

interface SheetStat {
  name: string
  count: number
}

interface Stats {
  totalProducts: number
  sheets: SheetStat[]
}

export default function Home() {
  const [products, setProducts] = useState<Product[]>([])
  const [stats, setStats] = useState<Stats>({ totalProducts: 0, sheets: [] })
  const [search, setSearch] = useState('')
  const [selectedSheet, setSelectedSheet] = useState('all')
  const [status, setStatus] = useState('พร้อมใช้งาน')
  const [loading, setLoading] = useState(false)
  const [tooltip, setTooltip] = useState<{ x: number; y: number; url: string } | null>(null)
  const [editProduct, setEditProduct] = useState<Product | null>(null)
  const [grabFile, setGrabFile] = useState<File | null>(null)
  const [grabResults, setGrabResults] = useState<any[]>([])
  const [showGrabModal, setShowGrabModal] = useState(false)
  const [grabMismatchProducts, setGrabMismatchProducts] = useState<Product[]>([])
  const [grabSource, setGrabSource] = useState<'GRAB' | 'Promaxx'>('GRAB')
  const [isAdding, setIsAdding] = useState(false)
  const [showPriceCalcModal, setShowPriceCalcModal] = useState(false)
  const [priceCalcResults, setPriceCalcResults] = useState<any[]>([])
  const [updatingPrices, setUpdatingPrices] = useState(false)
  const [lastPriceUpdate, setLastPriceUpdate] = useState<string | null>(null)
  const [showImportModal, setShowImportModal] = useState(false)
  const [importSheets, setImportSheets] = useState<{ name: string; rows: any[] }[]>([])
  const [importingData, setImportingData] = useState(false)
  const [importLog, setImportLog] = useState('')
  const [branches, setBranches] = useState<{ id: string; name: string }[]>([])
  const [selectedBranch, setSelectedBranch] = useState<string>(() =>
    typeof window !== 'undefined' ? localStorage.getItem('selectedBranch') ?? 'all' : 'all'
  )
  const [showAddBranchModal, setShowAddBranchModal] = useState(false)
  const [newBranchName, setNewBranchName] = useState('')
  const [addingBranch, setAddingBranch] = useState(false)

  useEffect(() => {
    loadBranches()
    loadStats()
  }, [])

  useEffect(() => {
    localStorage.setItem('selectedBranch', selectedBranch)
    loadStats()
    loadProducts('all')
    setSelectedSheet('all')
  }, [selectedBranch])

  async function loadBranches() {
    const res = await fetch('/api/branches')
    const data = await res.json()
    if (data.success) setBranches(data.branches)
  }

  async function addBranch() {
    if (!newBranchName.trim()) return
    setAddingBranch(true)
    const res = await fetch('/api/branches', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: newBranchName.trim() })
    })
    const data = await res.json()
    if (data.success) {
      setBranches(prev => [...prev, data.branch])
      setNewBranchName('')
      setShowAddBranchModal(false)
      setStatus(`เพิ่มสาขา "${data.branch.name}" สำเร็จ`)
    } else {
      setStatus('เกิดข้อผิดพลาด: ' + data.error)
    }
    setAddingBranch(false)
  }

  async function loadStats() {
    const branch = typeof window !== 'undefined' ? localStorage.getItem('selectedBranch') ?? 'all' : 'all'
    const url = branch && branch !== 'all' ? `/api/stats?branch=${encodeURIComponent(branch)}` : '/api/stats'
    const res = await fetch(url)
    const data = await res.json()
    if (data.success) setStats(data)
  }

  async function loadProducts(sheet = 'all') {
    setLoading(true)
    setStatus('กำลังโหลด...')
    const branch = typeof window !== 'undefined' ? localStorage.getItem('selectedBranch') ?? 'all' : 'all'
    const branchParam = branch && branch !== 'all' ? `&branch=${encodeURIComponent(branch)}` : ''
    const url = sheet === 'all' ? `/api/products?${branchParam}` : `/api/products?sheet=${encodeURIComponent(sheet)}${branchParam}`
    const res = await fetch(url)
    const data = await res.json()
    if (data.success) {
      setProducts(data.products)
      preloadImages(data.products)
      setStatus(`พร้อมใช้งาน (${data.products.length} รายการ)`)
    }
    setLoading(false)
  }

  async function searchProducts(sku: string) {
    if (sku.length < 3) return
    setLoading(true)
    const branch = typeof window !== 'undefined' ? localStorage.getItem('selectedBranch') ?? 'all' : 'all'
    const branchParam = branch && branch !== 'all' ? `&branch=${encodeURIComponent(branch)}` : ''
    const res = await fetch(`/api/products?sku=${encodeURIComponent(sku)}${branchParam}`)
    const data = await res.json()
    if (data.success) {
      setProducts(data.products)
      preloadImages(data.products)
    }
    setLoading(false)
  }

  async function checkPrices() {
    setLoading(true)
    setStatus('กำลังตรวจสอบราคา...')
    const branch = typeof window !== 'undefined' ? localStorage.getItem('selectedBranch') ?? 'all' : 'all'
    const url = branch && branch !== 'all' ? `/api/prices?branch=${encodeURIComponent(branch)}` : '/api/prices'
    const res = await fetch(url)
    const data = await res.json()
    if (data.success) {
      setProducts(data.mismatched)
      setStatus(`พบ ${data.mismatchCount} รายการราคาไม่ตรง`)
    }
    setLoading(false)
  }

  async function saveProduct() {
    if (!editProduct) return
    const payload = {
      'ประเภทสินค้า': editProduct['ประเภทสินค้า'],
      '*ชื่อสินค้า (NAME)': editProduct['*ชื่อสินค้า (NAME)'],
      '*เลขที่ใบอนุญาตโฆษณา': editProduct['*เลขที่ใบอนุญาตโฆษณา'],
      '*ราคาสินค้า': editProduct['*ราคาสินค้า'],
      'หมวดหมู่สินค้า (CATEGORIES)': editProduct['หมวดหมู่สินค้า (CATEGORIES)'],
      'รหัสสินค้า (SKU NUMBER)': editProduct['รหัสสินค้า (SKU NUMBER)'],
      '*รูปภาพสินค้า': editProduct['*รูปภาพสินค้า']
    }
    const res = await fetch('/api/products', {
      method: isAdding ? 'POST' : 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(isAdding ? payload : { sku: editProduct['_originalSku'], ...payload })
    })
    const data = await res.json()
    if (data.success) {
      if (isAdding) {
        setProducts(prev => [...prev, data.product])
        loadStats()
        await downloadProductXlsx(data.product)
      } else {
        setProducts(prev => prev.map(p =>
          p['รหัสสินค้า (SKU NUMBER)'] === editProduct['_originalSku'] ? { ...editProduct } : p
        ))
      }
      setEditProduct(null)
      setIsAdding(false)
      setStatus(isAdding ? 'เพิ่มสินค้าสำเร็จ ✅' : 'บันทึกสำเร็จ ✅')
    } else {
      setStatus('เกิดข้อผิดพลาด: ' + data.error)
    }
  }

  async function downloadProductXlsx(p: Product) {
    const XLSX = await import('xlsx')
    const rows = [{
      '*ประเภทสินค้า': p['ประเภทสินค้า'] || '',
      '*ชื่อสินค้า': p['*ชื่อสินค้า (NAME)'] || '',
      '*เลขที่ใบอนุญาตโฆษณา': p['*เลขที่ใบอนุญาตโฆษณา'] || '',
      '*ราคาสินค้า': p['*ราคาสินค้า'] || '',
      'รหัสสินค้า': p['รหัสสินค้า (SKU NUMBER)'] || '',
      '*รูปภาพสินค้า': p['*รูปภาพสินค้า'] || '',
      'หมวดหมู่รายการสินค้า': p['หมวดหมู่สินค้า (CATEGORIES)'] || ''
    }]
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'สินค้าใหม่')
    const now = new Date()
    const dd = String(now.getDate()).padStart(2, '0')
    const mm = String(now.getMonth() + 1).padStart(2, '0')
    const yyyy = now.getFullYear()
    XLSX.writeFile(wb, `GM MME ${dd}.${mm}.${yyyy}.xlsx`)
  }

  function convertDriveLink(url: string): string {
    if (!url) return ''
    const match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
      url.match(/[?&]id=([a-zA-Z0-9_-]+)/)
    if (match) return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w300`
    return url.startsWith('http') ? url : ''
  }

  function preloadImages(productList: Product[]) {
    productList.forEach(p => {
      const url = convertDriveLink(p['*รูปภาพสินค้า'] || '')
      if (url) {
        const img = new Image()
        img.src = url
      }
    })
  }

  function parseCSVRows(text: string): string[][] {
  const firstLine = text.split('\n')[0]
  const delim = firstLine.includes(';') && !firstLine.includes(',') ? ';' : ','
  return text.split(/\r?\n/).map(line => {
    const cells: string[] = []
    let cur = '', inQuote = false
    for (let i = 0; i < line.length; i++) {
      const ch = line[i]
      if (inQuote) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++ }
        else if (ch === '"') inQuote = false
        else cur += ch
      } else {
        if (ch === '"') inQuote = true
        else if (ch === delim) { cells.push(cur.trim()); cur = '' }
        else cur += ch
      }
    }
    cells.push(cur.trim())
    return cells
  })
}

function parsePriceRobust(str: string): number | null {
  if (!str) return null
  let s = str.replace(/[฿$€£¥\s]/g, '').replace(/,/g, '')
  const n = parseFloat(s)
  return isNaN(n) ? null : n
}

async function handleGrabCheck(file: File) {
  setGrabSource('GRAB')
  setStatus('กำลังตรวจสอบราคา GRAB...')
  setShowGrabModal(true)
  setGrabResults([])

  const text = await file.text()
  const clean = text.replace(/^\uFEFF/, '')
  const rows = parseCSVRows(clean)

  // Col I (index 8) = SKU, Col C (index 2) = ราคา, เริ่มจาก row 3 (index 2)
  const grabMap: Record<string, number> = {}
  for (let i = 2; i < rows.length; i++) {
    const row = rows[i]
    if (!row || row.length < 9) continue
    const sku = row[8]?.trim()
    const price = parsePriceRobust(row[2])
    if (sku && price !== null) grabMap[sku] = price
  }

  if (Object.keys(grabMap).length === 0) {
    setStatus('ไม่พบข้อมูลใน CSV')
    setGrabResults([{ error: 'ไม่พบข้อมูลใน CSV กรุณาตรวจสอบไฟล์' }])
    return
  }

  // ดึงข้อมูลจาก Supabase เฉพาะ SKU ที่มีใน CSV
  const skus = Object.keys(grabMap)
  const grabBranch = typeof window !== 'undefined' ? localStorage.getItem('selectedBranch') ?? 'all' : 'all'
  const res = await fetch('/api/products', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ skus, branch: grabBranch }) })
  const data = await res.json()

  if (!data.success) {
    setStatus('เกิดข้อผิดพลาด')
    return
  }

  if (data.products.length === 0) {
    setStatus('ไม่มีข้อมูล')
    setGrabResults([{ error: 'ไม่พบ SKU ใน Supabase กรุณาตรวจสอบข้อมูลใน CSV' }])
    return
  }

  // เปรียบเทียบราคา
  const foundSkus = new Set(data.products.map((p: Product) => p['รหัสสินค้า (SKU NUMBER)']))

  const matched_results = data.products
    .map((p: Product) => {
      const sku = p['รหัสสินค้า (SKU NUMBER)']
      const grabPrice = grabMap[sku]
      const dbPrice = parsePriceRobust(String(p['*ราคาสินค้า']))
      if (grabPrice === undefined || dbPrice === null) return null
      const matched = Math.abs(grabPrice - dbPrice) < 0.01
      return { sku, name: p['*ชื่อสินค้า (NAME)'], grabPrice, dbPrice, matched }
    })
    .filter(Boolean)

  const notFound_results = Object.keys(grabMap)
    .filter(sku => !foundSkus.has(sku))
    .map(sku => ({ sku, grabPrice: grabMap[sku], notFound: true }))

  const results = [...matched_results, ...notFound_results]

  const mismatch = matched_results.filter((r: any) => !r.matched)
  const mismatchSkus = new Set(mismatch.map((r: any) => r.sku))
  setGrabMismatchProducts(data.products.filter((p: Product) => mismatchSkus.has(p['รหัสสินค้า (SKU NUMBER)'])))
  setGrabResults(results)
  setStatus(`GRAB: ตรง ${matched_results.filter((r: any) => r.matched).length} | ต้องแก้ไข ${mismatch.length} | ไม่พบในระบบ ${notFound_results.length} รายการ`)
}

async function handlePromaxxCheck(file: File) {
  setGrabSource('Promaxx')
  setStatus('กำลังตรวจสอบราคา Promaxx...')
  setShowGrabModal(true)
  setGrabResults([])

  const text = await file.text()
  const clean = text.replace(/^\uFEFF/, '')
  const rows = parseCSVRows(clean)

  // Col B (index 1) = SKU, Col F (index 5) = filter 0, Col G (index 6) = ราคา, เริ่ม row 2 (index 1)
  const promaxxMap: Record<string, number> = {}
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i]
    if (!row || row.length < 7) continue
    const sku = row[1]?.trim()
    const colD = row[3]?.trim()
    const colF = row[5]?.trim()
    if (colD !== '1') continue
    if (colF !== '0') continue
    const price = parsePriceRobust(row[6])
    if (sku && price !== null) promaxxMap[sku] = price
  }

  if (Object.keys(promaxxMap).length === 0) {
    setStatus('ไม่พบข้อมูลใน CSV')
    setGrabResults([{ error: 'ไม่พบข้อมูลใน CSV กรุณาตรวจสอบไฟล์' }])
    return
  }

  const skus = Object.keys(promaxxMap)
  const branch = typeof window !== 'undefined' ? localStorage.getItem('selectedBranch') ?? 'all' : 'all'
  const res = await fetch('/api/products', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ skus, branch }) })
  const data = await res.json()

  if (!data.success) { setStatus('เกิดข้อผิดพลาด'); return }
  if (data.products.length === 0) {
    setGrabResults([{ error: 'ไม่พบ SKU ใน Supabase กรุณาตรวจสอบข้อมูลใน CSV' }])
    setStatus('ไม่มีข้อมูล')
    return
  }

  const foundSkus = new Set(data.products.map((p: Product) => p['รหัสสินค้า (SKU NUMBER)']))
  const matched_results = data.products.map((p: Product) => {
    const sku = p['รหัสสินค้า (SKU NUMBER)']
    const grabPrice = promaxxMap[sku]
    const dbPrice = parsePriceRobust(String(p['*ราคาสินค้า']))
    if (grabPrice === undefined || dbPrice === null) return null
    const matched = Math.abs(grabPrice - dbPrice) < 0.01
    return { sku, name: p['*ชื่อสินค้า (NAME)'], grabPrice, dbPrice, matched }
  }).filter(Boolean)

  const notFound_results = Object.keys(promaxxMap)
    .filter(sku => !foundSkus.has(sku))
    .map(sku => ({ sku, grabPrice: promaxxMap[sku], notFound: true }))

  const results = [...matched_results, ...notFound_results]
  const mismatch = matched_results.filter((r: any) => !r.matched)
  const mismatchSkus = new Set(mismatch.map((r: any) => r.sku))
  setGrabMismatchProducts(data.products.filter((p: Product) => mismatchSkus.has(p['รหัสสินค้า (SKU NUMBER)'])))
  setGrabResults(results)
  setStatus(`Promaxx: ตรง ${matched_results.filter((r: any) => r.matched).length} | ต้องแก้ไข ${mismatch.length} | ไม่พบในระบบ ${notFound_results.length} รายการ`)
}

  const IMPORT_FIELDS = [
    'ประเภทสินค้า', '*ชื่อสินค้า (NAME)', '*เลขที่ใบอนุญาตโฆษณา',
    '*ราคาสินค้า', 'รหัสสินค้า (SKU NUMBER)', '*รูปภาพสินค้า', 'หมวดหมู่สินค้า (CATEGORIES)'
  ]

  async function handleImportXlsx(file: File) {
    const XLSX = await import('xlsx')
    const buffer = await file.arrayBuffer()
    const wb = XLSX.read(buffer, { type: 'array' })
    const sheets = wb.SheetNames.map(name => {
      const ws = wb.Sheets[name]
      const raw: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
      const rows = raw.slice(1)
        .filter(row => row.some((c: any) => String(c).trim() !== ''))
        .map(row => {
          const obj: any = {}
          IMPORT_FIELDS.forEach((field, i) => {
            const val = row[i] !== undefined ? String(row[i]).trim() : ''
            obj[field] = field === '*ราคาสินค้า' ? (parseFloat(val) || null) : val
          })
          return obj
        })
      return { name, rows }
    })
    setImportSheets(sheets)
    setImportLog('')
    setShowImportModal(true)
  }

  async function confirmImport() {
    const importBranch = typeof window !== 'undefined' ? localStorage.getItem('selectedBranch') ?? 'all' : 'all'
    const allRows = importSheets.flatMap(s => s.rows).map(row =>
      importBranch && importBranch !== 'all' ? { ...row, branch: importBranch } : row
    )
    if (allRows.length === 0) return
    setImportingData(true)
    const chunkSize = 100
    let success = 0
    for (let i = 0; i < allRows.length; i += chunkSize) {
      const chunk = allRows.slice(i, i + chunkSize)
      setImportLog(`กำลัง import... ${Math.min(i + chunkSize, allRows.length)}/${allRows.length}`)
      const res = await fetch('/api/products', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows: chunk })
      })
      const data = await res.json()
      if (!data.success) {
        setImportLog('❌ เกิดข้อผิดพลาด: ' + data.error)
        setImportingData(false)
        return
      }
      success += chunk.length
    }
    setImportLog(`✅ Import สำเร็จ ${success} รายการ`)
    loadStats()
    setImportingData(false)
  }

  async function exportNotFoundXlsx() {
    const XLSX = await import('xlsx')
    const notFoundItems = grabResults.filter((r: any) => r.notFound)
    const rows = notFoundItems.map((r: any) => ({
      'รหัสสินค้า (SKU)': r.sku,
      'ราคาใน GRAB': r.grabPrice
    }))
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'ไม่พบในระบบ')
    const now = new Date()
    const dd = String(now.getDate()).padStart(2, '0')
    const mm = String(now.getMonth() + 1).padStart(2, '0')
    const yyyy = now.getFullYear()
    XLSX.writeFile(wb, `GRAB ไม่พบในระบบ ${dd}.${mm}.${yyyy}.xlsx`)
  }

  async function exportGrabXlsx() {
    const XLSX = await import('xlsx')
    const rows = grabMismatchProducts.map(p => ({
      '*ประเภทสินค้า': p['ประเภทสินค้า'] || '',
      '*ชื่อสินค้า': p['*ชื่อสินค้า (NAME)'] || '',
      '*เลขที่ใบอนุญาตโฆษณา': p['*เลขที่ใบอนุญาตโฆษณา'] || '',
      '*ราคาสินค้า': p['*ราคาสินค้า'] || '',
      'รหัสสินค้า': p['รหัสสินค้า (SKU NUMBER)'] || '',
      '*รูปภาพสินค้า': p['*รูปภาพสินค้า'] || '',
      'หมวดหมู่รายการสินค้า': p['หมวดหมู่สินค้า (CATEGORIES)'] || ''
    }))
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'ราคาไม่ตรง')
    const now = new Date()
    const dd = String(now.getDate()).padStart(2, '0')
    const mm = String(now.getMonth() + 1).padStart(2, '0')
    const yyyy = now.getFullYear()
    XLSX.writeFile(wb, `GM MME ${dd}.${mm}.${yyyy}.xlsx`)
  }

    async function handlePriceCalcUpload(file: File) {
  setStatus('กำลังคำนวณราคา...')
  setShowPriceCalcModal(true)
  setPriceCalcResults([])

  const text = await file.text()
  const clean = text.replace(/^\uFEFF/, '')
  const rows = parseCSVRows(clean)

  setStatus('(1/3) กำลังอ่านไฟล์ CSV...')

  // Col B (index 1) = SKU, Col F (index 5) = filter เฉพาะ 0, Col G (index 6) = ราคาระดับ 0
  // เริ่มจาก row 2 (index 1) ข้าม header
  const calcMap: Record<string, { sku: string; level0: number; D: number }> = {}

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i]
    if (!row || row.length < 7) continue
    const sku = row[1]?.trim()
    const colD = row[3]?.trim()
    const colF = row[5]?.trim()
    if (colD !== '1') continue  // หน่วยเล็กที่สุดขายหน้าร้าน
    if (colF !== '0') continue
    const level0 = parsePriceRobust(row[6])
    if (!sku || level0 === null) continue

    const A = level0 * 0.95
    const B = A * 1.20
    const C = B * 0.07
    const D = Math.ceil(B + C)  // ปัดขึ้นทศนิยม

    calcMap[sku] = { sku, level0, D }
  }

  console.log('[PriceCalc] calcMap size:', Object.keys(calcMap).length, 'sample:', Object.keys(calcMap).slice(0, 3))

  if (Object.keys(calcMap).length === 0) {
    setPriceCalcResults([{ error: 'ไม่พบข้อมูลใน CSV กรุณาตรวจสอบไฟล์' }])
    setStatus('ไม่พบข้อมูล')
    return
  }

  // ดึงข้อมูลจาก Supabase เฉพาะ SKU ที่มีใน CSV
  const skus = Object.keys(calcMap)
  const calcBranch = typeof window !== 'undefined' ? localStorage.getItem('selectedBranch') ?? 'all' : 'all'
  setStatus(`(2/3) กำลังดึงข้อมูลจากระบบ... (${skus.length} SKU)`)
  const res = await fetch('/api/products', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ skus, branch: calcBranch }) })
  const data = await res.json()
  console.log('[PriceCalc] API response:', JSON.stringify(data))

  if (!data.success) {
    setStatus('เกิดข้อผิดพลาดในการดึงข้อมูล')
    return
  }

  setStatus('(3/3) กำลังคำนวณราคา...')

  // จับคู่กับข้อมูลใน Supabase
  const results = data.products.map((p: Product) => {
    const sku = p['รหัสสินค้า (SKU NUMBER)']
    const calc = calcMap[sku]
    if (!calc) return null
    const oldPrice = parsePriceRobust(String(p['*ราคาสินค้า'])) ?? 0
    return {
      sku,
      name: p['*ชื่อสินค้า (NAME)'],
      level0: calc.level0,
      newPrice: calc.D,
      oldPrice,
      changed: Math.abs(calc.D - oldPrice) > 0.01,
      _product: p
    }
  }).filter(Boolean)

  setPriceCalcResults(results)
  const changedCount = results.filter((r: any) => r.changed).length
  setStatus(`คำนวณเสร็จ — เปลี่ยนแปลง ${changedCount} รายการ`)
}

async function confirmUpdatePrices() {
  const toUpdate = priceCalcResults.filter(r => r.changed)
  if (toUpdate.length === 0) return

  setUpdatingPrices(true)
  setStatus('กำลังอัพเดทราคา...')

  const res = await fetch('/api/prices', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      items: toUpdate.map(r => ({ sku: r.sku, correctPrice: r.newPrice }))
    })
  })
  const data = await res.json()

  if (data.success) {
    const now = new Date()
    const dd = String(now.getDate()).padStart(2, '0')
    const mm = String(now.getMonth() + 1).padStart(2, '0')
    const yyyy = now.getFullYear()
    const hh = String(now.getHours()).padStart(2, '0')
    const min = String(now.getMinutes()).padStart(2, '0')
    setLastPriceUpdate(`${dd}/${mm}/${yyyy} ${hh}:${min}`)
    setStatus(`✅ อัพเดทราคาสำเร็จ ${data.updated} รายการ`)

    // Export Excel หลังอัพเดทสำเร็จ
    const XLSX = await import('xlsx')
    const exportRows = toUpdate.map((r: any) => ({
      '*ประเภทสินค้า': r._product?.['ประเภทสินค้า'] || '',
      '*ชื่อสินค้า': r._product?.['*ชื่อสินค้า (NAME)'] || '',
      '*เลขที่ใบอนุญาตโฆษณา': r._product?.['*เลขที่ใบอนุญาตโฆษณา'] || '',
      '*ราคาสินค้า': r.newPrice,
      'รหัสสินค้า': r.sku,
      '*รูปภาพสินค้า': r._product?.['*รูปภาพสินค้า'] || '',
      'หมวดหมู่รายการสินค้า': r._product?.['หมวดหมู่สินค้า (CATEGORIES)'] || ''
    }))
    const ws = XLSX.utils.json_to_sheet(exportRows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'อัพเดทราคา')
    XLSX.writeFile(wb, `GM MME ${dd}.${mm}.${yyyy}.xlsx`)

    setShowPriceCalcModal(false)
  } else {
    setStatus('เกิดข้อผิดพลาดในการอัพเดท')
  }
  setUpdatingPrices(false)
}

  return (
    <div style={{ fontFamily: 'sans-serif', height: '100vh', display: 'flex', flexDirection: 'column' }}>

      {/* Header */}
      <div style={{ background: '#f5f7fa', padding: '12px 20px', borderBottom: '1px solid #ddd', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <h1 style={{ fontSize: 18, color: '#000' }}>📦 ระบบตรวจสอบราคาสินค้า</h1>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 12, color: '#555' }}>สาขา:</span>
          <select
            value={selectedBranch}
            onChange={e => setSelectedBranch(e.target.value)}
            style={{ padding: '5px 10px', border: '1px solid #999', borderRadius: 3, fontSize: 13 }}
          >
            <option value="all">ทั้งหมด</option>
            {branches.map(b => (
              <option key={b.id} value={b.name}>{b.name}</option>
            ))}
          </select>
          <button onClick={() => setShowAddBranchModal(true)} style={{ padding: '5px 10px', border: '1px solid #999', borderRadius: 3, fontSize: 12, cursor: 'pointer', background: '#fff' }}>+ สาขาใหม่</button>
        </div>
        <span style={{ fontSize: 11, color: '#000' }}>อัพเดทราคาล่าสุด: {lastPriceUpdate ?? '-'}</span>
      </div>

      {/* Toolbar */}
      <div style={{ background: '#f8f8f8', borderBottom: '1px solid #bbb', padding: '10px 20px', display: 'flex', gap: 12, alignItems: 'center' }}>
        <input
          type="text"
          placeholder="ค้นหารหัสสินค้า..."
          value={search}
          onChange={e => { setSearch(e.target.value); searchProducts(e.target.value) }}
          style={{ padding: '6px 10px', border: '1px solid #999', borderRadius: 3, fontSize: 13, width: 220 }}
        />
        <input
        type="file" accept=".csv" id="grabInput" style={{ display: 'none' }}
        onChange={e => { if (e.target.files?.[0]) handleGrabCheck(e.target.files[0]) }}
        />
        <input
        type="file" accept=".csv" id="promaxxInput" style={{ display: 'none' }}
        onChange={e => { if (e.target.files?.[0]) handlePromaxxCheck(e.target.files[0]) }}
        />
        <input
        type="file" accept=".csv" id="priceCalcInput" style={{ display: 'none' }}
        onChange={e => { if (e.target.files?.[0]) handlePriceCalcUpload(e.target.files[0]) }}
        />
        <input
        type="file" accept=".xlsx,.xls" id="importInput" style={{ display: 'none' }}
        onChange={e => { if (e.target.files?.[0]) { handleImportXlsx(e.target.files[0]); e.target.value = '' } }}
        />
        <button onClick={() => loadProducts(selectedSheet)} style={btnStyle}>🔄 รีเฟรช</button>
        <button onClick={() => { setIsAdding(true); setEditProduct({ 'ประเภทสินค้า': '', '*ชื่อสินค้า (NAME)': '', '*เลขที่ใบอนุญาตโฆษณา': '', '*ราคาสินค้า': '', 'รหัสสินค้า (SKU NUMBER)': '', '*รูปภาพสินค้า': '', 'หมวดหมู่สินค้า (CATEGORIES)': '' }) }} style={btnStyle}>➕ เพิ่มรายการสินค้า</button>
        <button onClick={() => document.getElementById('importInput')?.click()} style={btnStyle}>📂 Import Excel</button>
        <button onClick={() => document.getElementById('priceCalcInput')?.click()} style={btnStyle} title="ใช้ไฟล์ R05.105 อัพโหลดสำหรับกรณี Promaxx Update ราคา">🧮 ตรวจสอบราคา Promaxx</button>
        <button onClick={() => document.getElementById('grabInput')?.click()} style={btnStyle} title="ไฟล์เมนูที่ download จาก GrabMart สำหรับตรวจสอบชื่อสินค้าและราคา">🛵 ตรวจสอบราคา GRAB</button>
        <button onClick={() => document.getElementById('promaxxInput')?.click()} style={btnStyle} title="ไฟล์ R05.105 จาก Promaxx">🔍 ตรวจสอบราคา Promaxx</button>
      </div>

      {/* Main Layout */}
      <div style={{ display: 'grid', gridTemplateColumns: '260px 1fr', flex: 1, overflow: 'hidden' }}>

        {/* Left Panel */}
        <div style={{ background: '#f0f0f0', borderRight: '1px solid #bbb', overflowY: 'auto' }}>
          <div style={{ padding: '6px 10px', background: '#e0e0e0', borderBottom: '1px solid #ccc', fontWeight: 700, fontSize: 13 }}>📊 สถิติ</div>
          <div style={{ padding: 15 }}>
            <div style={statCard}>
              <div style={{ fontSize: 11, color: '#000' }}>สินค้าทั้งหมด</div>
              <div style={{ fontSize: 24, fontWeight: 700, color: '#c74634' }}>{stats.totalProducts.toLocaleString()}</div>
            </div>
            <div style={statCard}>
              <div style={{ fontSize: 11, color: '#000' }}>หมวดหมู่</div>
              <div style={{ fontSize: 24, fontWeight: 700, color: '#c74634' }}>{stats.sheets.length}</div>
            </div>
          </div>

          <div style={{ padding: '6px 10px', background: '#e0e0e0', borderBottom: '1px solid #ccc', fontWeight: 700, fontSize: 13 }}>📁 หมวดหมู่</div>
          {stats.sheets.map(s => (
            <div
              key={s.name}
              onClick={() => { setSelectedSheet(s.name); loadProducts(s.name) }}
              style={{
                padding: '8px 12px', borderBottom: '1px solid #ddd', fontSize: 12,
                cursor: 'pointer', display: 'flex', justifyContent: 'space-between',
                background: selectedSheet === s.name ? '#cce5ff' : 'transparent'
              }}
            >
              <span>{s.name}</span>
              <span style={{ background: '#4a8bc4', color: '#fff', padding: '2px 8px', borderRadius: 10, fontSize: 10 }}>{s.count}</span>
            </div>
          ))}
        </div>

        {/* Right Panel */}
        <div style={{ overflowY: 'auto', background: '#fff' }}>
          {loading ? (
            <div style={{ textAlign: 'center', padding: 40, color: '#000' }}>กำลังโหลด...</div>
          ) : products.length === 0 ? (
            <div style={{ textAlign: 'center', padding: 40, color: '#000' }}>เลือกหมวดหมู่หรือค้นหาสินค้า</div>
          ) : (
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr style={{ background: '#f5f5f5' }}>
                  <th style={th}>#</th>
                  <th style={th}>หมวดหมู่</th>
                  <th style={th}>รหัสสินค้า</th>
                  <th style={th}>ชื่อสินค้า</th>
                  <th style={th}>ใบอนุญาต</th>
                  <th style={th}>ราคา</th>
                  <th style={th}>ดำเนินการ</th>
                </tr>
              </thead>
              <tbody>
                {products.map((p, i) => (
                  <tr key={i}
                    onMouseEnter={e => {
                      e.currentTarget.style.background = '#e8f4fc'
                      const imgUrl = convertDriveLink(p['*รูปภาพสินค้า'] || '')
                      setTooltip({ x: e.clientX, y: e.clientY, url: imgUrl })
                    }}
                    onMouseMove={e => {
                      setTooltip(prev => prev ? { ...prev, x: e.clientX, y: e.clientY } : null)
                    }}
                    onMouseLeave={e => {
                      e.currentTarget.style.background = 'transparent'
                      setTooltip(null)
                    }}
                  >
                    <td style={td}>{i + 1}</td>
                    <td style={td}>{p['หมวดหมู่สินค้า (CATEGORIES)'] || '-'}</td>
                    <td style={td}><strong>{p['รหัสสินค้า (SKU NUMBER)'] || '-'}</strong></td>
                    <td style={td}>{p['*ชื่อสินค้า (NAME)'] || '-'}</td>
                    <td style={td}>{p['*เลขที่ใบอนุญาตโฆษณา'] || '-'}</td>
                    <td style={{ ...td, color: '#c74634', fontWeight: 600 }}>{p['*ราคาสินค้า'] || '-'}</td>
                    <td style={td}>
                      <button
                        onClick={() => setEditProduct({ ...p, _originalSku: p['รหัสสินค้า (SKU NUMBER)'] })}
                        style={{ ...btnStyle, fontSize: 11, padding: '4px 10px' }}
                      >
                        ✏️ แก้ไข
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>

      {/* Status Bar */}
      <div style={{ background: '#f0f0f0', borderTop: '1px solid #bbb', padding: '4px 15px', fontSize: 11, color: '#000' }}>
        {status}
      </div>
      
      {showGrabModal && (
  <div style={{
    position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)',
    display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000
  }}>
    <div style={{ background: '#fff', borderRadius: 6, width: 700, maxHeight: '80vh', display: 'flex', flexDirection: 'column', boxShadow: '0 10px 40px rgba(0,0,0,0.3)' }}>

      {/* Header */}
      <div style={{ background: '#4a4a4a', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <strong>{grabSource === 'Promaxx' ? '🔍 ผลตรวจสอบราคา Promaxx' : '🛵 ผลตรวจสอบราคา GRAB'}</strong>
        <button onClick={() => setShowGrabModal(false)}
          style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer' }}>×</button>
      </div>

      {/* Summary */}
      {grabResults.length > 0 && !grabResults[0]?.error && (
        <div style={{ padding: '10px 20px', background: '#f8f8f8', borderBottom: '1px solid #ddd', display: 'flex', gap: 20, fontSize: 13 }}>
          <span>ทั้งหมด: <strong>{grabResults.length}</strong></span>
          <span style={{ color: '#28a745' }}>✓ ตรง: <strong>{grabResults.filter((r: any) => r.matched).length}</strong></span>
          <span style={{ color: '#dc3545' }}>✗ ต้องแก้ไขใน GRAB: <strong>{grabResults.filter((r: any) => !r.matched && !r.notFound).length}</strong></span>
          {grabResults.filter((r: any) => r.notFound).length > 0 && (
            <span style={{ color: '#000' }}>⚠ ไม่พบในระบบ: <strong>{grabResults.filter((r: any) => r.notFound).length}</strong></span>
          )}
        </div>
      )}

      {/* Table */}
      <div style={{ overflowY: 'auto', flex: 1 }}>
        {grabResults[0]?.error ? (
          <div style={{ padding: 40, textAlign: 'center', color: '#dc3545' }}>{grabResults[0].error}</div>
        ) : grabResults.length === 0 ? (
          <div style={{ padding: 40, textAlign: 'center', color: '#000' }}>กำลังโหลด...</div>
        ) : (
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr style={{ background: '#f5f5f5', position: 'sticky', top: 0 }}>
                <th style={th}>#</th>
                <th style={th}>SKU</th>
                <th style={th}>ชื่อสินค้า</th>
                <th style={th}>ราคาใน GRAB</th>
                <th style={th}>ราคาใน DB</th>
                <th style={th}>สถานะ</th>
              </tr>
            </thead>
            <tbody>
              {grabResults.map((r: any, i: number) => (
                <tr key={i} style={{ background: r.notFound ? '#f0f0f0' : r.matched ? 'transparent' : '#fff3cd' }}>
                  <td style={td}>{i + 1}</td>
                  <td style={{ ...td, color: '#000' }}><strong>{r.sku}</strong></td>
                  <td style={{ ...td, color: '#000' }}>{r.name || '-'}</td>
                  <td style={{ ...td, fontWeight: 600, color: '#000' }}>{r.grabPrice?.toLocaleString() ?? '-'}</td>
                  <td style={{ ...td, fontWeight: 600, color: '#000' }}>{r.dbPrice?.toLocaleString() ?? '-'}</td>
                  <td style={td}>
                    {r.notFound
                      ? <span style={{ background: '#e0e0e0', color: '#000', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>⚠ ไม่พบในระบบ</span>
                      : r.matched
                        ? <span style={{ background: '#d4edda', color: '#155724', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✓ ตรง</span>
                        : <span style={{ background: '#f8d7da', color: '#721c24', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✗ แก้ไขใน GRAB</span>
                    }
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Footer */}
      <div style={{ padding: '12px 20px', borderTop: '1px solid #ddd', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div style={{ display: 'flex', gap: 8 }}>
          <button
            onClick={exportGrabXlsx}
            disabled={grabMismatchProducts.length === 0}
            style={{ ...btnStyle, background: '#d4edda', borderColor: '#28a745', color: '#155724', opacity: grabMismatchProducts.length === 0 ? 0.5 : 1 }}
          >
            📥 Export ราคาไม่ตรง ({grabMismatchProducts.length})
          </button>
          <button
            onClick={exportNotFoundXlsx}
            disabled={grabResults.filter((r: any) => r.notFound).length === 0}
            style={{ ...btnStyle, background: '#e8e8e8', borderColor: '#999', color: '#000', opacity: grabResults.filter((r: any) => r.notFound).length === 0 ? 0.5 : 1 }}
          >
            ⚠ Export ไม่พบในระบบ ({grabResults.filter((r: any) => r.notFound).length})
          </button>
        </div>
        <button onClick={() => setShowGrabModal(false)} style={btnStyle}>ปิด</button>
      </div>
    </div>
  </div>
)}

      {showPriceCalcModal && (
  <div style={{
    position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)',
    display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000
  }}>
    <div style={{ background: '#fff', borderRadius: 6, width: 750, maxHeight: '80vh', display: 'flex', flexDirection: 'column', boxShadow: '0 10px 40px rgba(0,0,0,0.3)' }}>

      {/* Header */}
      <div style={{ background: '#4a4a4a', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <strong>🧮 ผลคำนวณราคา</strong>
        <button onClick={() => setShowPriceCalcModal(false)}
          style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer' }}>×</button>
      </div>

      {/* Summary */}
      {priceCalcResults.length > 0 && !priceCalcResults[0]?.error && (
        <div style={{ padding: '10px 20px', background: '#f8f8f8', borderBottom: '1px solid #ddd', display: 'flex', gap: 20, fontSize: 13 }}>
          <span>ทั้งหมด: <strong>{priceCalcResults.length}</strong></span>
          <span style={{ color: '#28a745' }}>เปลี่ยนแปลง: <strong>{priceCalcResults.filter(r => r.changed).length}</strong></span>
          <span style={{ color: '#000' }}>ราคาเดิม: <strong>{priceCalcResults.filter(r => !r.changed).length}</strong></span>
          <div style={{ marginLeft: 'auto', fontSize: 11, color: '#000' }}>
            สูตร: ระดับ0 × 0.95 × 1.20 × 1.07
          </div>
        </div>
      )}

      {/* Table */}
      <div style={{ overflowY: 'auto', flex: 1 }}>
        {priceCalcResults[0]?.error ? (
          <div style={{ padding: 40, textAlign: 'center', color: '#dc3545' }}>{priceCalcResults[0].error}</div>
        ) : priceCalcResults.length === 0 ? (
          <div style={{ padding: 40, textAlign: 'center', color: '#000' }}>กำลังคำนวณ...</div>
        ) : (
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr style={{ background: '#f5f5f5', position: 'sticky', top: 0 }}>
                <th style={th}>#</th>
                <th style={th}>SKU</th>
                <th style={th}>ชื่อสินค้า</th>
                <th style={th}>ราคาระดับ 0</th>
                <th style={th}>ราคาเดิม</th>
                <th style={th}>ราคาใหม่ (D)</th>
                <th style={th}>สถานะ</th>
              </tr>
            </thead>
            <tbody>
              {priceCalcResults.filter((r: any) => r.changed).map((r: any, i: number) => (
                <tr key={i} style={{ background: r.changed ? '#fff3cd' : 'transparent' }}>
                  <td style={td}>{i + 1}</td>
                  <td style={td}><strong>{r.sku}</strong></td>
                  <td style={td}>{r.name}</td>
                  <td style={td}>{r.level0.toLocaleString()}</td>
                  <td style={td}>{r.oldPrice.toLocaleString()}</td>
                  <td style={{ ...td, color: '#28a745', fontWeight: 600 }}>{r.newPrice.toLocaleString()}</td>
                  <td style={td}>
                    {r.changed
                      ? <span style={{ background: '#fff3cd', color: '#856404', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>⚠ เปลี่ยนแปลง</span>
                      : <span style={{ background: '#d4edda', color: '#155724', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✓ เดิม</span>
                    }
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Footer */}
      <div style={{ padding: '12px 20px', borderTop: '1px solid #ddd', display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
        <button onClick={() => setShowPriceCalcModal(false)} style={btnStyle}>ยกเลิก</button>
        <button
          onClick={async () => {
            const toExport = priceCalcResults.filter((r: any) => r.changed)
            if (toExport.length === 0) return
            const XLSX = await import('xlsx')
            const now = new Date()
            const dd = String(now.getDate()).padStart(2, '0')
            const mm = String(now.getMonth() + 1).padStart(2, '0')
            const yyyy = now.getFullYear()
            const exportRows = toExport.map((r: any) => ({
              '*ประเภทสินค้า': r._product?.['ประเภทสินค้า'] || '',
              '*ชื่อสินค้า': r._product?.['*ชื่อสินค้า (NAME)'] || '',
              '*เลขที่ใบอนุญาตโฆษณา': r._product?.['*เลขที่ใบอนุญาตโฆษณา'] || '',
              '*ราคาสินค้า': r.newPrice,
              'รหัสสินค้า': r.sku,
              '*รูปภาพสินค้า': r._product?.['*รูปภาพสินค้า'] || '',
              'หมวดหมู่รายการสินค้า': r._product?.['หมวดหมู่สินค้า (CATEGORIES)'] || ''
            }))
            const ws = XLSX.utils.json_to_sheet(exportRows)
            const wb = XLSX.utils.book_new()
            XLSX.utils.book_append_sheet(wb, ws, 'อัพเดทราคา')
            XLSX.writeFile(wb, `GM MME ${dd}.${mm}.${yyyy}.xlsx`)
          }}
          disabled={priceCalcResults.filter((r: any) => r.changed).length === 0}
          style={{ ...btnStyle, background: '#17a2b8', color: '#fff', borderColor: '#117a8b' }}
        >
          📥 Export Excel
        </button>
        <button
          onClick={confirmUpdatePrices}
          disabled={updatingPrices || priceCalcResults.filter(r => r.changed).length === 0}
          style={{ ...btnStyle, background: '#28a745', color: '#fff', borderColor: '#1e7e34', opacity: updatingPrices ? 0.6 : 1 }}
        >
          {updatingPrices ? '⏳ กำลังอัพเดท...' : `💾 อัพเดท ${priceCalcResults.filter(r => r.changed).length} รายการ`}
        </button>
      </div>
    </div>
  </div>
)}


      {/* Import Modal */}
      {showImportModal && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
          <div style={{ background: '#fff', borderRadius: 6, width: 500, maxHeight: '80vh', display: 'flex', flexDirection: 'column', boxShadow: '0 10px 40px rgba(0,0,0,0.3)' }}>
            <div style={{ background: '#4a4a4a', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <strong>📂 Import ข้อมูลจาก Excel</strong>
              <button onClick={() => setShowImportModal(false)} style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer' }}>×</button>
            </div>
            <div style={{ padding: '10px 20px', background: '#f8f8f8', borderBottom: '1px solid #ddd', fontSize: 13 }}>
              พบ <strong>{importSheets.length}</strong> ชีท | รวม <strong>{importSheets.reduce((s, sh) => s + sh.rows.length, 0)}</strong> รายการ
            </div>
            <div style={{ overflowY: 'auto', flex: 1, padding: 16 }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                <thead>
                  <tr style={{ background: '#f5f5f5' }}>
                    <th style={th}>#</th>
                    <th style={th}>ชีท</th>
                    <th style={th}>จำนวนรายการ</th>
                  </tr>
                </thead>
                <tbody>
                  {importSheets.map((s, i) => (
                    <tr key={i}>
                      <td style={td}>{i + 1}</td>
                      <td style={td}><strong>{s.name}</strong></td>
                      <td style={td}>{s.rows.length.toLocaleString()} รายการ</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {importLog && (
                <div style={{ marginTop: 12, padding: '8px 12px', background: importLog.startsWith('✅') ? '#d4edda' : importLog.startsWith('❌') ? '#f8d7da' : '#fff3cd', borderRadius: 4, fontSize: 13 }}>
                  {importLog}
                </div>
              )}
            </div>
            <div style={{ padding: '12px 20px', borderTop: '1px solid #ddd', display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              <button onClick={() => setShowImportModal(false)} style={btnStyle}>ปิด</button>
              <button
                onClick={confirmImport}
                disabled={importingData || importSheets.reduce((s, sh) => s + sh.rows.length, 0) === 0}
                style={{ ...btnStyle, background: '#4a8bc4', color: '#fff', borderColor: '#3d7ab3', opacity: importingData ? 0.6 : 1 }}
              >
                {importingData ? '⏳ กำลัง Import...' : `📥 Import ทั้งหมด (${importSheets.reduce((s, sh) => s + sh.rows.length, 0)} รายการ)`}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Edit Modal */}
      {editProduct && (
        <div style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)',
          display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000
        }}>
          <div style={{
            background: '#fff', borderRadius: 6, width: 480,
            boxShadow: '0 10px 40px rgba(0,0,0,0.3)', overflow: 'hidden'
          }}>
            {/* Modal Header */}
            <div style={{ background: isAdding ? '#2d6a2d' : '#4a4a4a', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <strong>{isAdding ? '➕ เพิ่มรายการสินค้า' : '✏️ แก้ไขสินค้า'}</strong>
              <button onClick={() => { setEditProduct(null); setIsAdding(false) }}
                style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer' }}>×</button>
            </div>

            {/* Modal Body */}
            <div style={{ padding: 20, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
              <div style={{ gridColumn: 'span 2' }}>
                <label style={labelStyle}>ประเภทสินค้า</label>
                <input style={inputStyle}
                  value={editProduct['ประเภทสินค้า'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, 'ประเภทสินค้า': e.target.value })}
                />
              </div>
              <div>
                <label style={labelStyle}>หมวดหมู่</label>
                <input style={inputStyle}
                  value={editProduct['หมวดหมู่สินค้า (CATEGORIES)'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, 'หมวดหมู่สินค้า (CATEGORIES)': e.target.value })}
                />
              </div>
              <div>
                <label style={labelStyle}>รหัสสินค้า</label>
                <input style={inputStyle}
                  value={editProduct['รหัสสินค้า (SKU NUMBER)'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, 'รหัสสินค้า (SKU NUMBER)': e.target.value })}
                />
              </div>
              <div style={{ gridColumn: 'span 2' }}>
                <label style={labelStyle}>ชื่อสินค้า</label>
                <input style={inputStyle}
                  value={editProduct['*ชื่อสินค้า (NAME)'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, '*ชื่อสินค้า (NAME)': e.target.value })}
                />
              </div>
              <div>
                <label style={labelStyle}>ใบอนุญาต</label>
                <input style={inputStyle}
                  value={editProduct['*เลขที่ใบอนุญาตโฆษณา'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, '*เลขที่ใบอนุญาตโฆษณา': e.target.value })}
                />
              </div>
              <div>
                <label style={labelStyle}>ราคา</label>
                <input style={inputStyle} type="number"
                  value={editProduct['*ราคาสินค้า'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, '*ราคาสินค้า': e.target.value })}
                />
              </div>
              <div style={{ gridColumn: 'span 2' }}>
                <label style={labelStyle}>รูปภาพสินค้า (Google Drive Link)</label>
                <input style={inputStyle}
                  placeholder="วาง Google Drive URL ที่นี่..."
                  value={editProduct['*รูปภาพสินค้า'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, '*รูปภาพสินค้า': e.target.value })}
                />
                {editProduct['*รูปภาพสินค้า'] && (
                  <div style={{ marginTop: 8, textAlign: 'center', background: '#f8f8f8', borderRadius: 4, padding: 8, border: '1px solid #eee' }}>
                    <img
                      src={convertDriveLink(editProduct['*รูปภาพสินค้า'])}
                      alt="preview"
                      style={{ maxWidth: '100%', maxHeight: 160, borderRadius: 4, objectFit: 'contain' }}
                      onError={e => { e.currentTarget.style.display = 'none' }}
                    />
                  </div>
                )}
              </div>
            </div>

            {/* Modal Footer */}
            <div style={{ padding: '0 20px 20px', display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              <button onClick={() => { setEditProduct(null); setIsAdding(false) }} style={btnStyle}>ยกเลิก</button>
              <button onClick={saveProduct}
                style={{ ...btnStyle, background: isAdding ? '#28a745' : '#4a8bc4', color: '#fff', borderColor: isAdding ? '#1e7e34' : '#3d7ab3' }}>
                {isAdding ? '➕ เพิ่มสินค้า' : '💾 บันทึก'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Image Tooltip */}
      {tooltip && (
        <div style={{
          position: 'fixed',
          left: tooltip.x + 15,
          top: tooltip.y + 15,
          zIndex: 9999,
          background: '#fff',
          border: '2px solid #4a8bc4',
          borderRadius: 8,
          padding: 8,
          boxShadow: '0 10px 40px rgba(0,0,0,0.3)',
          pointerEvents: 'none'
        }}>
          {tooltip.url ? (
            <img
              src={tooltip.url}
              alt=""
              style={{ maxWidth: 280, maxHeight: 280, borderRadius: 4, display: 'block' }}
              onError={e => (e.currentTarget.style.display = 'none')}
            />
          ) : (
            <div style={{ padding: 40, color: '#000', fontSize: 12 }}>ไม่มีรูป</div>
          )}
        </div>
      )}

      {/* Add Branch Modal */}
      {showAddBranchModal && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
          <div style={{ background: '#fff', borderRadius: 6, width: 360, boxShadow: '0 10px 40px rgba(0,0,0,0.3)', overflow: 'hidden' }}>
            <div style={{ background: '#2d6a2d', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <strong>+ เพิ่มสาขาใหม่</strong>
              <button onClick={() => setShowAddBranchModal(false)} style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer' }}>×</button>
            </div>
            <div style={{ padding: 20 }}>
              <label style={{ fontSize: 12, fontWeight: 600, display: 'block', marginBottom: 4 }}>ชื่อสาขา</label>
              <input
                style={{ width: '100%', padding: '8px 10px', border: '1px solid #ccc', borderRadius: 4, fontSize: 13, boxSizing: 'border-box' }}
                placeholder="เช่น สาขาสีลม, สาขาลาดพร้าว"
                value={newBranchName}
                onChange={e => setNewBranchName(e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter') addBranch() }}
                autoFocus
              />
            </div>
            <div style={{ padding: '0 20px 20px', display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              <button onClick={() => { setShowAddBranchModal(false); setNewBranchName('') }} style={btnStyle}>ยกเลิก</button>
              <button
                onClick={addBranch}
                disabled={addingBranch || !newBranchName.trim()}
                style={{ ...btnStyle, background: '#28a745', color: '#fff', borderColor: '#1e7e34', opacity: addingBranch ? 0.6 : 1 }}
              >
                {addingBranch ? '⏳ กำลังเพิ่ม...' : '✅ เพิ่มสาขา'}
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  )
}

const btnStyle: React.CSSProperties = {
  padding: '7px 16px', border: '1px solid #777', borderRadius: 20,
  fontSize: 12, fontWeight: 600, cursor: 'pointer', background: '#f5f5f5'
}
const statCard: React.CSSProperties = {
  background: '#fff', border: '1px solid #ccc', borderRadius: 4,
  padding: 15, marginBottom: 10
}
const th: React.CSSProperties = {
  padding: '8px 12px', textAlign: 'left', border: '1px solid #ccc',
  fontWeight: 600, position: 'sticky', top: 0, background: '#f5f5f5'
}
const td: React.CSSProperties = {
  padding: '8px 12px', border: '1px solid #eee', color: '#000'
}
const labelStyle: React.CSSProperties = {
  fontSize: 11, color: '#000', display: 'block', marginBottom: 4
}
const inputStyle: React.CSSProperties = {
  width: '100%', padding: '6px 10px', border: '1px solid #999',
  borderRadius: 3, fontSize: 13, fontFamily: 'sans-serif', color: '#000',
  boxSizing: 'border-box'
}

