'use client'

import { useState, useEffect, useRef } from 'react'
import { supabase } from '../lib/supabase'

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
  const [grabSource, setGrabSource] = useState<'GRAB'>('GRAB')
  const [isAdding, setIsAdding] = useState(false)
  const [showPriceCalcModal, setShowPriceCalcModal] = useState(false)
  const [priceCalcResults, setPriceCalcResults] = useState<any[]>([])
  const [updatingPrices, setUpdatingPrices] = useState(false)
  const [lastPriceUpdate, setLastPriceUpdate] = useState<string | null>(null)
  const [priceCalcGrabFile, setPriceCalcGrabFile] = useState<File | null>(null)
  const [priceCalcGrabMap, setPriceCalcGrabMap] = useState<Record<string, number>>({})
  const [showImportModal, setShowImportModal] = useState(false)
  const [importSheets, setImportSheets] = useState<{ name: string; rows: any[] }[]>([])
  const [importingData, setImportingData] = useState(false)
  const [importLog, setImportLog] = useState('')
  const [importError, setImportError] = useState('')
  const [selectedSkus, setSelectedSkus] = useState<Set<string>>(new Set())
  const [masterSkuCount, setMasterSkuCount] = useState<number | null>(null)
  const [masterFileName, setMasterFileName] = useState<string | null>(null)
  const [masterMissingCount, setMasterMissingCount] = useState<number | null>(null)
  const [exportingMissing, setExportingMissing] = useState(false)
  const [showMissingBranchModal, setShowMissingBranchModal] = useState(false)
  const [missingBranchResults, setMissingBranchResults] = useState<Record<string, Product[]>>({ src: [], kkl: [], sss: [] })
  const [missingBranchLoaded, setMissingBranchLoaded] = useState<Record<string, boolean>>({ src: false, kkl: false, sss: false })
  const [missingBranchLoading, setMissingBranchLoading] = useState(false)
  const [missingBranchTab, setMissingBranchTab] = useState<'src'|'kkl'|'sss'>('src')
  const [missingBranchExportCat, setMissingBranchExportCat] = useState<string>('all')
  const [missingBranchSearch, setMissingBranchSearch] = useState<string>('')
  const [showMasterBranchModal, setShowMasterBranchModal] = useState(false)
  const [masterBranchResults, setMasterBranchResults] = useState<Array<{sku: string, name: string, src: boolean, kkl: boolean, sss: boolean}>>([])
  const [masterBranchLoading, setMasterBranchLoading] = useState(false)
  const [masterBranchTab, setMasterBranchTab] = useState<'all'|'src'|'kkl'|'sss'>('all')
  const [masterBranchSearch, setMasterBranchSearch] = useState('')
  const [masterNameMap, setMasterNameMap] = useState<Record<string, string>>({})
  const [showMmeCheckModal, setShowMmeCheckModal] = useState(false)
  const [mmeCheckResults, setMmeCheckResults] = useState<Record<string, any[]>>({ src: [], kkl: [], sss: [] })
  const [mmeCheckGrabFileSrc, setMmeCheckGrabFileSrc] = useState<File | null>(null)
  const [mmeCheckGrabFileKkl, setMmeCheckGrabFileKkl] = useState<File | null>(null)
  const [mmeCheckGrabFileSss, setMmeCheckGrabFileSss] = useState<File | null>(null)
  const [mmeCheckMmeFile, setMmeCheckMmeFile] = useState<File | null>(null)
  const [mmeChecking, setMmeChecking] = useState(false)
  const [mmeActiveTab, setMmeActiveTab] = useState<'src'|'kkl'|'sss'|'compare'>('src')
  const [editImageSrc, setEditImageSrc] = useState<string>('')
  const [editImageOverlay, setEditImageOverlay] = useState<string>('')
  const [editImageMime, setEditImageMime] = useState<string>('image/png')
  const [editImageFileName, setEditImageFileName] = useState<string>('')
  const [editImageMsg, setEditImageMsg] = useState<string>('')
  const editCanvasRef = useRef<HTMLCanvasElement | null>(null)

  useEffect(() => {
    loadStats()
    loadProducts('all')
    loadMasterCount()
  }, [])

  async function loadMasterCount() {
    const res = await fetch('/api/master')
    const data = await res.json()
    if (data.success && data.count > 0) setMasterSkuCount(data.count)
  }

  async function loadStats() {
    const res = await fetch('/api/stats')
    const data = await res.json()
    if (data.success) setStats(data)
  }

  async function loadProducts(sheet = 'all') {
    setLoading(true)
    setStatus('กำลังโหลด...')
    const url = sheet === 'all' ? '/api/products' : `/api/products?sheet=${encodeURIComponent(sheet)}`
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
    const res = await fetch(`/api/products?q=${encodeURIComponent(sku)}`)
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
    const res = await fetch('/api/prices')
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

  function drawImageOnCanvas(src: string, overlayText: string) {
    const canvas = editCanvasRef.current
    if (!canvas) return
    const img = new window.Image()
    img.onload = () => {
      const maxW = 1024
      const scale = img.width > maxW ? maxW / img.width : 1
      const w = Math.round(img.width * scale)
      const h = Math.round(img.height * scale)
      canvas.width = w
      canvas.height = h
      const ctx = canvas.getContext('2d')
      if (!ctx) return
      ctx.fillStyle = '#fff'
      ctx.fillRect(0, 0, w, h)
      ctx.drawImage(img, 0, 0, w, h)
      const text = (overlayText || '').trim()
      if (text) {
        const fontSize = Math.max(10, Math.round(h * 0.028))
        ctx.font = `bold ${fontSize}px Arial, sans-serif`
        ctx.textBaseline = 'bottom'
        ctx.textAlign = 'right'
        ctx.fillStyle = '#000'
        ctx.fillText(text, w - 16, h - 40)
      }
    }
    img.onerror = () => { setEditImageMsg('โหลดรูปไม่สำเร็จ') }
    img.src = src
  }

  function onPickEditImage(file: File | null) {
    if (!file) return
    setEditImageMsg('')
    setEditImageMime(file.type || 'image/png')
    setEditImageFileName(file.name)
    const reader = new FileReader()
    reader.onload = () => {
      const dataUrl = String(reader.result || '')
      setEditImageSrc(dataUrl)
    }
    reader.readAsDataURL(file)
  }

  function downloadEditImage() {
    const canvas = editCanvasRef.current
    if (!canvas || !editImageSrc) {
      setEditImageMsg('กรุณาเลือกรูปภาพก่อน')
      return
    }
    const isJpeg = editImageMime.includes('jpeg') || editImageMime.includes('jpg')
    const outMime = isJpeg ? 'image/jpeg' : 'image/png'
    const ext = isJpeg ? 'jpg' : 'png'
    const name = (editProduct?.['*ชื่อสินค้า (NAME)'] || editProduct?.['รหัสสินค้า (SKU NUMBER)'] || `product-${Date.now()}`).toString().trim()
    const baseName = name.replace(/[^a-zA-Z0-9ก-๙._-]/g, '_')
    const filename = `${baseName}.${ext}`
    const url = canvas.toDataURL(outMime, 0.92)
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    a.click()
    setEditImageMsg('ดาวน์โหลดสำเร็จ ✅')
  }

  useEffect(() => {
    if (editImageSrc) drawImageOnCanvas(editImageSrc, editImageOverlay)
  }, [editImageSrc, editImageOverlay])

  useEffect(() => {
    if (!editProduct) {
      setEditImageSrc('')
      setEditImageOverlay('')
      setEditImageFileName('')
      setEditImageMsg('')
      return
    }
    const license = (editProduct['*เลขที่ใบอนุญาตโฆษณา'] || '').toString().trim()
    const sku = (editProduct['รหัสสินค้า (SKU NUMBER)'] || '').toString().trim()
    const defaultText = [license, sku].filter(Boolean).join(' ')
    setEditImageOverlay(defaultText)
  }, [editProduct])

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
    if (match) return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w200`
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

function renderBranchFlag(value: any) {
  return value ? '✓' : '-'
}


async function handleGrabCheck(file: File, branch: 'src' | 'kkl' | 'sss') {
  setGrabSource('GRAB')
  setStatus(`กำลังตรวจสอบราคา GRAB ${branch.toUpperCase()}...`)
  setShowGrabModal(true)
  setGrabResults([])

  const text = await file.text()
  const clean = text.replace(/^\uFEFF/, '')
  const rows = parseCSVRows(clean)

  const dataRows = rows.slice(2).filter(r => r && r.length >= 9)
  const hasSkuCol = dataRows.some(r => r[8]?.trim())
  if (rows.length < 3 || !hasSkuCol) {
    setStatus('ไฟล์ไม่ถูกต้อง')
    setGrabResults([{ error: 'ไฟล์ไม่ใช่ Grab_menu CSV ที่ถูกต้อง — ต้องมีอย่างน้อย 3 แถว และคอลัมน์ที่ 9 (SKU) ต้องมีข้อมูล' }])
    return
  }

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

  // ซิงก์สถานะว่าสาขาที่เลือกมี SKU ไหนบ้างตาม Grab_menu
  const skus = Object.keys(grabMap)
  const syncRes = await fetch('/api/products', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ availabilityBranch: branch, skus })
  })
  const syncData = await syncRes.json()

  if (!syncData.success) {
    setStatus('เกิดข้อผิดพลาดในการอัปเดตสถานะสาขา')
    setGrabResults([{ error: `อัปเดตสถานะสาขาไม่สำเร็จ: ${syncData.error ?? 'unknown error'}` }])
    return
  }

  // ดึงข้อมูล Master มาเทียบราคาและสถานะการมีสินค้า
  const res = await fetch('/api/products', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ skus }) })
  const data = await res.json()

  if (!data.success) {
    setStatus('เกิดข้อผิดพลาด')
    return
  }

  if (data.products.length === 0) {
    setStatus('ไม่มีข้อมูล')
    setGrabResults([{ error: 'อัพโหลดไฟล์ผิด หรือ คอลัมน์ไฟล์ไม่ถูกต้อง' }])
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
  await loadProducts(selectedSheet)
  await loadStats()
  setStatus(`Grab ${branch.toUpperCase()}: ตรง ${matched_results.filter((r: any) => r.matched).length} | ต้องแก้ไข ${mismatch.length} | ไม่พบข้อมูลใน Supabase ${notFound_results.length} | เพิ่มสถานะ ${syncData.inserted || 0} | ลบสถานะ ${syncData.deleted || 0}`)
}

  const IMPORT_FIELDS = [
    'ประเภทสินค้า', '*ชื่อสินค้า (NAME)', '*เลขที่ใบอนุญาตโฆษณา',
    '*ราคาสินค้า', 'รหัสสินค้า (SKU NUMBER)', '*รูปภาพสินค้า', 'หมวดหมู่สินค้า (CATEGORIES)'
  ]

  const IMPORT_HEADER_GROUPS: Array<{ target: string; aliases: string[] }> = [
    { target: 'ประเภทสินค้า', aliases: ['*ประเภทสินค้า', 'ประเภทสินค้า'] },
    { target: '*ชื่อสินค้า (NAME)', aliases: ['*ชื่อสินค้า', '*ชื่อสินค้า (NAME)'] },
    { target: '*เลขที่ใบอนุญาตโฆษณา', aliases: ['*เลขที่ใบอนุญาตโฆษณา'] },
    { target: '*ราคาสินค้า', aliases: ['*ราคาสินค้า'] },
    { target: 'รหัสสินค้า (SKU NUMBER)', aliases: ['รหัสสินค้า', 'รหัสสินค้า (SKU NUMBER)'] },
    { target: '*รูปภาพสินค้า', aliases: ['*รูปภาพสินค้า'] },
    { target: 'หมวดหมู่สินค้า (CATEGORIES)', aliases: ['หมวดหมู่รายการสินค้า', 'หมวดหมู่สินค้า (CATEGORIES)'] }
  ]

  function normalizeImportHeader(value: any) {
    return String(value ?? '').replace(/\s+/g, ' ').trim()
  }

  function getImportColumnMap(headerRow: any[]) {
    const normalizedHeader = (headerRow || []).map(normalizeImportHeader)
    const columnMap: Record<string, number> = {}

    IMPORT_HEADER_GROUPS.forEach(({ target, aliases }) => {
      const index = normalizedHeader.findIndex(col => aliases.includes(col))
      if (index >= 0) columnMap[target] = index
    })

    return columnMap
  }

  function findImportHeader(rawRows: any[][]) {
    const searchRows = rawRows.slice(0, 10)

    for (let rowIndex = 0; rowIndex < searchRows.length; rowIndex++) {
      const columnMap = getImportColumnMap(searchRows[rowIndex] || [])
      const hasRequiredHeaders = IMPORT_FIELDS.every(field => columnMap[field] !== undefined)
      if (hasRequiredHeaders) {
        return { headerRowIndex: rowIndex, columnMap }
      }
    }

    return { headerRowIndex: -1, columnMap: {} as Record<string, number> }
  }

  async function handleImportXlsx(file: File) {
    const XLSX = await import('xlsx')
    const buffer = await file.arrayBuffer()
    const wb = XLSX.read(buffer, { type: 'array' })
    const invalidSheets: string[] = []
    const sheets = wb.SheetNames.map(name => {
      const ws = wb.Sheets[name]
      const raw: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
      const { headerRowIndex, columnMap } = findImportHeader(raw)
      const hasInvalidHeader = headerRowIndex < 0
      if (hasInvalidHeader) invalidSheets.push(name)
      const rows = raw.slice(headerRowIndex + 1)
        .filter(row => row.some((c: any) => String(c).trim() !== ''))
        .map(row => {
          const obj: any = {}
          IMPORT_FIELDS.forEach(field => {
            const columnIndex = columnMap[field]
            const val = columnIndex !== undefined && row[columnIndex] !== undefined ? String(row[columnIndex]).trim() : ''
            obj[field] = field === '*ราคาสินค้า' ? (parseFloat(val) || null) : val
          })
          return obj
        })
      return { name, rows }
    })

    const hasRows = sheets.some(sheet => sheet.rows.length > 0)
    const errorMessage = invalidSheets.length > 0
      ? `ไฟล์นี้คอลัมน์ไม่ตรงกับรูปแบบที่ระบบรองรับในชีท: ${invalidSheets.join(', ')}`
      : !hasRows
        ? 'ไม่พบข้อมูลที่นำเข้าได้ในไฟล์นี้'
        : ''

    setImportSheets(sheets)
    setImportLog('')
    setImportError(errorMessage)
    setStatus(errorMessage || `พร้อม Import ${sheets.reduce((sum, sheet) => sum + sheet.rows.length, 0)} รายการ`)
    setShowImportModal(true)
  }

  async function confirmImport() {
    if (importError) {
      setImportLog(`❌ ${importError}`)
      return
    }

    const allRows = importSheets.flatMap(s => s.rows)
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
    await loadStats()
    await loadProducts('all')
    setSelectedSheet('all')
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
    XLSX.utils.book_append_sheet(wb, ws, 'ไม่พบข้อมูลใน Supabase')
    const now = new Date()
    const dd = String(now.getDate()).padStart(2, '0')
    const mm = String(now.getMonth() + 1).padStart(2, '0')
    const yyyy = now.getFullYear()
    XLSX.writeFile(wb, `GRAB ไม่พบข้อมูลใน Supabase ${dd}.${mm}.${yyyy}.xlsx`)
  }

  async function exportGrabXlsx() {
    const XLSX = await import('xlsx')
    const rows = toExportRows(grabMismatchProducts)
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'ราคาไม่ตรง')
    const now = new Date()
    const dd = String(now.getDate()).padStart(2, '0')
    const mm = String(now.getMonth() + 1).padStart(2, '0')
    const yyyy = now.getFullYear()
    XLSX.writeFile(wb, `GM MME ${dd}.${mm}.${yyyy}.xlsx`)
  }

  function toExportRows(list: Product[]) {
    return list.map(p => ({
      '*ประเภทสินค้า': p['ประเภทสินค้า'] || '',
      '*ชื่อสินค้า': p['*ชื่อสินค้า (NAME)'] || '',
      '*เลขที่ใบอนุญาตโฆษณา': p['*เลขที่ใบอนุญาตโฆษณา'] || '',
      '*ราคาสินค้า': p['*ราคาสินค้า'] || '',
      'รหัสสินค้า': p['รหัสสินค้า (SKU NUMBER)'] || '',
      '*รูปภาพสินค้า': p['*รูปภาพสินค้า'] || '',
      'หมวดหมู่รายการสินค้า': p['หมวดหมู่สินค้า (CATEGORIES)'] || ''
    }))
  }

  async function exportSelectedCategoryXlsx() {
    if (selectedSheet === 'all') {
      setStatus('กรุณาเลือกหมวดหมู่ก่อน Export')
      return
    }

    setStatus(`กำลังเตรียมไฟล์ Export หมวดหมู่ ${selectedSheet}...`)
    const url = `/api/products?sheet=${encodeURIComponent(selectedSheet)}`
    const res = await fetch(url)
    const data = await res.json()

    if (!data.success) {
      setStatus('เกิดข้อผิดพลาดในการ Export')
      return
    }

    const allProducts = data.products || []
    if (allProducts.length === 0) {
      setStatus('ไม่มีข้อมูลสินค้าให้ Export')
      return
    }

    const XLSX = await import('xlsx')
    const rows = toExportRows(allProducts)
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, selectedSheet)
    const now = new Date()
    const dd = String(now.getDate()).padStart(2, '0')
    const mm = String(now.getMonth() + 1).padStart(2, '0')
    const yyyy = now.getFullYear()
    XLSX.writeFile(wb, `GM MME ${selectedSheet} ${dd}.${mm}.${yyyy}.xlsx`)
    setStatus(`Export หมวดหมู่ ${selectedSheet} สำเร็จ (${allProducts.length} รายการ)`)
  }

  async function exportSelectedSkusXlsx() {
    if (selectedSkus.size === 0) return
    setStatus('กำลัง Export รายการที่เลือก...')
    const skuArray = Array.from(selectedSkus)
    const res = await fetch('/api/products', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ skus: skuArray }) })
    const data = await res.json()
    const selected = data.success ? data.products : products.filter(p => selectedSkus.has(p['รหัสสินค้า (SKU NUMBER)']))
    const XLSX = await import('xlsx')
    const rows = toExportRows(selected)
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'รายการที่เลือก')
    const now = new Date()
    const dd = String(now.getDate()).padStart(2, '0')
    const mm = String(now.getMonth() + 1).padStart(2, '0')
    const yyyy = now.getFullYear()
    XLSX.writeFile(wb, `GM MME Selected ${dd}.${mm}.${yyyy}.xlsx`)
    setStatus(`Export ${selected.length} รายการที่เลือกสำเร็จ ✅`)
  }

  async function loadMissingBranch(branch: 'src'|'kkl'|'sss') {
    if (missingBranchLoaded[branch]) return
    setMissingBranchLoading(true)
    setStatus(`กำลังโหลดสินค้าที่ขาดใน GRAB ${branch.toUpperCase()}...`)
    const res = await fetch(`/api/products?missingBranch=${branch}`)
    const data = await res.json()
    if (data.success) {
      setMissingBranchResults(prev => ({ ...prev, [branch]: data.products }))
      setMissingBranchLoaded(prev => ({ ...prev, [branch]: true }))
      setStatus(`ขาดใน GRAB ${branch.toUpperCase()}: ${data.products.length} รายการ`)
    } else {
      setStatus('เกิดข้อผิดพลาด: ' + data.error)
    }
    setMissingBranchLoading(false)
  }

  async function handleMmeVsGrabCheck() {
    if (!mmeCheckMmeFile) return
    if (!mmeCheckGrabFileSrc && !mmeCheckGrabFileKkl && !mmeCheckGrabFileSss) return
    setMmeChecking(true)
    setStatus('กำลังตรวจสอบ GM MME vs GRAB...')

    async function parseGrabMap(file: File): Promise<Record<string, number>> {
      const text = await file.text()
      const rows = parseCSVRows(text.replace(/^﻿/, ''))
      const map: Record<string, number> = {}
      for (let i = 2; i < rows.length; i++) {
        const row = rows[i]
        if (!row || row.length < 9) continue
        const sku = row[8]?.trim()
        const price = parsePriceRobust(row[2])
        if (sku && price !== null) map[sku] = price
      }
      return map
    }

    const XLSX = await import('xlsx')
    const mmeBuffer = await mmeCheckMmeFile.arrayBuffer()
    const wbMme = XLSX.read(mmeBuffer, { type: 'array' })
    const wsMme = wbMme.Sheets[wbMme.SheetNames[0]]
    const mmeRows = XLSX.utils.sheet_to_json(wsMme, { header: 1 }) as any[][]

    function compareWithMap(grabMap: Record<string, number>): any[] {
      const results: any[] = []
      for (let i = 1; i < mmeRows.length; i++) {
        const row = mmeRows[i]
        if (!row || row.length < 5) continue
        const type     = String(row[0] || '').trim()
        const name     = String(row[1] || '').trim()
        const license  = String(row[2] || '').trim()
        const mmePrice = parsePriceRobust(String(row[3] || ''))
        const sku      = String(row[4] || '').trim()
        const image    = String(row[5] || '').trim()
        const category = String(row[6] || '').trim()
        if (!sku) continue
        const grabPrice = grabMap[sku]
        if (grabPrice === undefined) {
          results.push({ sku, name, mmePrice, grabPrice: null, status: 'notFound', type, license, image, category })
        } else {
          const matched = mmePrice !== null && Math.abs(mmePrice - grabPrice) < 0.01
          results.push({ sku, name, mmePrice, grabPrice, status: matched ? 'updated' : 'notUpdated', type, license, image, category })
        }
      }
      return results
    }

    const newResults: Record<string, any[]> = { src: [], kkl: [], sss: [] }
    const [mapSrc, mapKkl, mapSss] = await Promise.all([
      mmeCheckGrabFileSrc ? parseGrabMap(mmeCheckGrabFileSrc) : Promise.resolve({}),
      mmeCheckGrabFileKkl ? parseGrabMap(mmeCheckGrabFileKkl) : Promise.resolve({}),
      mmeCheckGrabFileSss ? parseGrabMap(mmeCheckGrabFileSss) : Promise.resolve({}),
    ])
    if (mmeCheckGrabFileSrc) newResults.src = compareWithMap(mapSrc)
    if (mmeCheckGrabFileKkl) newResults.kkl = compareWithMap(mapKkl)
    if (mmeCheckGrabFileSss) newResults.sss = compareWithMap(mapSss)

    setMmeCheckResults(newResults)
    setMmeChecking(false)

    const allResults = [...newResults.src, ...newResults.kkl, ...newResults.sss]
    const updated = allResults.filter(r => r.status === 'updated').length
    const notUpdated = allResults.filter(r => r.status === 'notUpdated').length
    const notFound = allResults.filter(r => r.status === 'notFound').length
    setStatus(`GM MME: แก้ไขแล้ว ${updated} | ยังไม่แก้ไข ${notUpdated} | ไม่พบใน GRAB ${notFound}`)

    const firstTab = mmeCheckGrabFileSrc ? 'src' : mmeCheckGrabFileKkl ? 'kkl' : 'sss'
    setMmeActiveTab(firstTab)
  }

  async function handleMasterUpload(file: File) {
    setStatus('กำลังอ่าน Product Master...')
    setMasterFileName(file.name)
    const XLSX = await import('xlsx')
    const buffer = await file.arrayBuffer()
    const wb = XLSX.read(buffer, { type: 'array' })
    const ws = wb.Sheets[wb.SheetNames[0]]
    const raw: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
    const nameMap: Record<string, string> = {}
    const skus = raw.slice(1)
      .map(row => {
        const sku = String(row[0] ?? '').trim()
        const name = String(row[1] ?? '').trim()
        if (sku && name) nameMap[sku] = name
        return sku
      })
      .filter(Boolean)
    setMasterNameMap(nameMap)

    setStatus(`กำลังบันทึก ${skus.length.toLocaleString()} SKU ลง Supabase...`)
    const res = await fetch('/api/master', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ skus })
    })
    const data = await res.json()
    if (!data.success) {
      setStatus('เกิดข้อผิดพลาด: ' + data.error)
      return
    }
    setMasterSkuCount(data.count)
    setStatus(`บันทึก Product Master สำเร็จ: ${data.count.toLocaleString()} SKU — กำลังเปรียบเทียบ...`)
    await refreshMissingCount()
    setStatus(`Product Master: ${data.count.toLocaleString()} SKU`)
  }

  async function refreshMissingCount() {
    const res = await fetch('/api/master?compare=true')
    const data = await res.json()
    if (data.success) {
      setMasterSkuCount(data.masterCount)
      setMasterMissingCount(data.missingCount)
    }
  }

  async function exportMissingSkus() {
    setExportingMissing(true)
    setStatus('กำลังดึงรายการสินค้าที่ขาด...')
    const res = await fetch('/api/master?compare=true')
    const data = await res.json()
    if (!data.success) {
      setStatus('เกิดข้อผิดพลาด: ' + data.error)
      setExportingMissing(false)
      return
    }
    if (data.missingSkus.length === 0) {
      setStatus('ไม่มีรายการสินค้าที่ขาดใน GrabMart ✅')
      setExportingMissing(false)
      return
    }
    const XLSX = await import('xlsx')
    const rows = data.missingSkus.map((sku: string) => ({ 'รหัสสินค้า (SKU NUMBER)': sku }))
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'ขาดใน GrabMart')
    const now = new Date()
    const dd = String(now.getDate()).padStart(2, '0')
    const mm = String(now.getMonth() + 1).padStart(2, '0')
    const yyyy = now.getFullYear()
    XLSX.writeFile(wb, `Missing GrabMart ${dd}.${mm}.${yyyy}.xlsx`)
    setMasterMissingCount(data.missingCount)
    setStatus(`Export สำเร็จ: ${data.missingCount.toLocaleString()} รายการที่ขาดใน GrabMart`)
    setExportingMissing(false)
  }

  async function handleMasterBranchCompare() {
    setMasterBranchLoading(true)
    setShowMasterBranchModal(true)
    setMasterBranchResults([])
    setMasterBranchTab('all')
    setMasterBranchSearch('')
    setStatus('กำลังเปรียบเทียบ Product Master กับทุกสาขา...')
    const res = await fetch('/api/master?compareBranch=true')
    const data = await res.json()
    if (!data.success) {
      setStatus('เกิดข้อผิดพลาด: ' + data.error)
      setMasterBranchLoading(false)
      return
    }
    const results = (data.items as any[]).map(item => ({
      sku: String(item.sku),
      name: masterNameMap[String(item.sku)] || '',
      src: !!item.src,
      kkl: !!item.kkl,
      sss: !!item.sss
    }))
    setMasterBranchResults(results)
    setMasterBranchLoading(false)
    const missingAny = results.filter(r => !r.src || !r.kkl || !r.sss).length
    setStatus(`เปรียบเทียบสำเร็จ: ${results.length} SKU — ขาดอย่างน้อย 1 สาขา ${missingAny} รายการ`)
  }

  async function deleteSelectedCategory() {
    if (selectedSheet === 'all') {
      setStatus('กรุณาเลือกหมวดหมู่ก่อนลบ')
      return
    }

    const ok = window.confirm(
      `ยืนยันลบหมวดหมู่ "${selectedSheet}" ออกจากทุกสาขา?\nการลบนี้ไม่สามารถย้อนกลับได้`
    )
    if (!ok) return

    setStatus(`กำลังลบหมวดหมู่ ${selectedSheet}...`)
    const res = await fetch('/api/products', {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sheet: selectedSheet, branch: 'all' })
    })
    const data = await res.json()

    if (!data.success) {
      setStatus('เกิดข้อผิดพลาด: ' + data.error)
      return
    }

    setSelectedSheet('all')
    await loadStats()
    await loadProducts('all')
    setStatus(`ลบหมวดหมู่ ${selectedSheet} สำเร็จ (${data.deleted || 0} รายการ)`)
  }

  async function handlePriceCalcGrabUpload(file: File) {
    setPriceCalcGrabFile(file)
    const text = await file.text()
    const rows = parseCSVRows(text.replace(/^﻿/, ''))
    const map: Record<string, number> = {}
    for (let i = 2; i < rows.length; i++) {
      const row = rows[i]
      if (!row || row.length < 9) continue
      const sku = row[8]?.trim()
      const price = parsePriceRobust(row[2])
      if (sku && price !== null) map[sku] = price
    }
    setPriceCalcGrabMap(map)
  }

    async function handlePriceCalcUpload(file: File) {
  setStatus('กำลังคำนวณราคา...')
  setShowPriceCalcModal(true)
  setPriceCalcResults([])

  const text = await file.text()
  const clean = text.replace(/^\uFEFF/, '')
  const rows = parseCSVRows(clean)

  setStatus('(1/3) กำลังอ่านไฟล์ CSV...')

  // Col A (index 0) = SKU, Col E (index 4) = ราคาระดับ 0 (header ชื่อ '0')
  // เริ่มจาก row 2 (index 1) ข้าม header
  const calcMap: Record<string, { sku: string; level0: number; D: number }> = {}

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i]
    if (!row || row.length < 5) continue
    const sku = row[0]?.trim()
    const level0 = parsePriceRobust(row[4])
    if (!sku || level0 === null) continue

    const A = level0 * 0.95
    const B = A * 1.20
    const C = B * 0.07
    const D = Math.ceil(B + C)  // ปัดขึ้นทศนิยม

    calcMap[sku] = { sku, level0, D }
  }

  console.log('[PriceCalc] calcMap size:', Object.keys(calcMap).length, 'sample:', Object.keys(calcMap).slice(0, 3))

  if (Object.keys(calcMap).length === 0) {
    setPriceCalcResults([{ error: 'ไฟล์ไม่ถูกต้อง หรือ คอลัมน์ผิด' }])
    setStatus('ไม่พบข้อมูล')
    return
  }

  // ดึงข้อมูลจาก Supabase เฉพาะ SKU ที่มีใน CSV
  const skus = Object.keys(calcMap)
  setStatus(`(2/3) กำลังดึงข้อมูลจากระบบ... (${skus.length} SKU)`)
  const res = await fetch('/api/products', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ skus }) })
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

function handleCsvToUtf8(file: File) {
  const reader = new FileReader()
  reader.onload = (e) => {
    const buffer = e.target?.result as ArrayBuffer
    if (!buffer) return
    // ลอง decode ด้วย windows-874 (TIS-620) ก่อน
    let text: string
    try {
      text = new TextDecoder('windows-874').decode(buffer)
    } catch {
      text = new TextDecoder('utf-8').decode(buffer)
    }
    // เพิ่ม BOM สำหรับ UTF-8
    const utf8Bom = '\uFEFF'
    const blob = new Blob([utf8Bom + text], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    const baseName = file.name.replace(/\.csv$/i, '')
    a.href = url
    a.download = `${baseName}_utf8.csv`
    a.click()
    URL.revokeObjectURL(url)
    setStatus(`แปลงไฟล์ "${file.name}" เป็น UTF-8 สำเร็จ ✅`)
  }
  reader.readAsArrayBuffer(file)
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
    const exportRows = toExportRows(
      toUpdate.map((r: any) => ({
        ...r._product,
        '*ราคาสินค้า': r.newPrice,
        'รหัสสินค้า (SKU NUMBER)': r.sku
      }))
    )
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
        <h1 style={{ fontSize: 18, color: '#000' }}>Grab Master รายการสินค้าทั้งหมด by mailforspiritwish</h1>
        <span style={{ fontSize: 11, color: '#000' }}>อัพเดทราคาล่าสุด: {lastPriceUpdate ?? '-'}</span>
      </div>

      {/* Toolbar */}
      <div style={{ background: '#f8f8f8', borderBottom: '1px solid #bbb', padding: '10px 20px', display: 'flex', gap: 12, alignItems: 'center' }}>
        <input
          type="text"
          placeholder="ค้นหารหัสสินค้า / ชื่อสินค้า..."
          value={search}
          onChange={e => { setSearch(e.target.value); searchProducts(e.target.value) }}
          style={{ padding: '6px 10px', border: '1px solid #999', borderRadius: 3, fontSize: 13, width: 220 }}
        />
        <input type="file" accept=".csv" id="grabInputSrc" style={{ display: 'none' }}
          onChange={e => { if (e.target.files?.[0]) { handleGrabCheck(e.target.files[0], 'src'); e.target.value = '' } }} />
        <input type="file" accept=".csv" id="grabInputKkl" style={{ display: 'none' }}
          onChange={e => { if (e.target.files?.[0]) { handleGrabCheck(e.target.files[0], 'kkl'); e.target.value = '' } }} />
        <input type="file" accept=".csv" id="grabInputSss" style={{ display: 'none' }}
          onChange={e => { if (e.target.files?.[0]) { handleGrabCheck(e.target.files[0], 'sss'); e.target.value = '' } }} />
        <input
        type="file" accept=".csv" id="priceCalcInput" style={{ display: 'none' }}
        onChange={e => { if (e.target.files?.[0]) handlePriceCalcUpload(e.target.files[0]) }}
        />
        <input type="file" accept=".csv" id="priceCalcGrabInput" style={{ display: 'none' }}
          onChange={e => { if (e.target.files?.[0]) { handlePriceCalcGrabUpload(e.target.files[0]); e.target.value = '' } }} />
        <input
        type="file" accept=".xlsx,.xls" id="importInput" style={{ display: 'none' }}
        onChange={e => { if (e.target.files?.[0]) { handleImportXlsx(e.target.files[0]); e.target.value = '' } }}
        />
        <input
        type="file" accept=".csv" id="csvConvertInput" style={{ display: 'none' }}
        onChange={e => { if (e.target.files?.[0]) { handleCsvToUtf8(e.target.files[0]); e.target.value = '' } }}
        />
        <button onClick={() => loadProducts(selectedSheet)} style={btnStyle}>🔄 รีเฟรช</button>
        {selectedSkus.size > 0 && (
          <button
            onClick={exportSelectedSkusXlsx}
            style={{ ...btnStyle, background: '#d4edda', borderColor: '#28a745', color: '#155724' }}
          >
            📥 Export ที่เลือก ({selectedSkus.size})
          </button>
        )}
        <button onClick={() => { setIsAdding(true); setEditProduct({ 'ประเภทสินค้า': '', '*ชื่อสินค้า (NAME)': '', '*เลขที่ใบอนุญาตโฆษณา': '', '*ราคาสินค้า': '', 'รหัสสินค้า (SKU NUMBER)': '', '*รูปภาพสินค้า': '', 'หมวดหมู่สินค้า (CATEGORIES)': '' }) }} style={btnStyle}>➕ เพิ่มรายการสินค้า</button>
        <button onClick={() => document.getElementById('importInput')?.click()} style={btnStyle}>📂 Import Excel</button>
        <button onClick={() => document.getElementById('csvConvertInput')?.click()} style={btnStyle} title="แปลงไฟล์ CSV จาก POS (TIS-620) เป็น CSV UTF-8">🔄 แปลง CSV เป็น UTF-8</button>
        <button onClick={() => document.getElementById('priceCalcInput')?.click()} style={btnStyle} title="ใช้ไฟล์ Price อัพโหลดสำหรับกรณี Promaxx Update ราคา">🧮 ตรวจสอบราคา Promaxx</button>
        <button onClick={() => document.getElementById('grabInputSrc')?.click()} style={btnStyle} title="อัพโหลด Grab_menu สาขา SRC">🛵 GRAB SRC</button>
        <button onClick={() => document.getElementById('grabInputKkl')?.click()} style={btnStyle} title="อัพโหลด Grab_menu สาขา KKL">🛵 GRAB KKL</button>
        <button onClick={() => document.getElementById('grabInputSss')?.click()} style={btnStyle} title="อัพโหลด Grab_menu สาขา SSS">🛵 GRAB SSS</button>
        <input type="file" accept=".csv" id="mmeCheckGrabInputSrc" style={{ display: 'none' }}
          onChange={e => { if (e.target.files?.[0]) { setMmeCheckGrabFileSrc(e.target.files[0]); e.target.value = '' } }} />
        <input type="file" accept=".csv" id="mmeCheckGrabInputKkl" style={{ display: 'none' }}
          onChange={e => { if (e.target.files?.[0]) { setMmeCheckGrabFileKkl(e.target.files[0]); e.target.value = '' } }} />
        <input type="file" accept=".csv" id="mmeCheckGrabInputSss" style={{ display: 'none' }}
          onChange={e => { if (e.target.files?.[0]) { setMmeCheckGrabFileSss(e.target.files[0]); e.target.value = '' } }} />
        <input type="file" accept=".xlsx,.xls" id="mmeCheckMmeInput" style={{ display: 'none' }}
          onChange={e => { if (e.target.files?.[0]) { setMmeCheckMmeFile(e.target.files[0]); e.target.value = '' } }} />
        <button
          onClick={() => { setShowMmeCheckModal(true); setMmeCheckResults({ src: [], kkl: [], sss: [] }) }}
          style={{ ...btnStyle, background: '#e8f4ff', borderColor: '#4a8bc4' }}
          title="ตรวจสอบว่าสินค้าที่ยื่นแก้ไขใน GM MME ถูกอัพเดทราคาใน GRAB แล้วหรือยัง"
        >📋 ตรวจ GM MME vs GRAB</button>
        <button
          onClick={() => {
            setShowMissingBranchModal(true)
            setMissingBranchResults({ src: [], kkl: [], sss: [] })
            setMissingBranchLoaded({ src: false, kkl: false, sss: false })
            setMissingBranchTab('src')
            setTimeout(() => loadMissingBranch('src'), 0)
          }}
          style={{ ...btnStyle, background: '#fff3e0', borderColor: '#e65100' }}
          title="ตรวจสอบสินค้าที่ยังไม่มีใน GRAB แต่ละสาขา"
        >🔍 สินค้าขาด GRAB</button>

      {/* Main Layout */}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '260px 1fr', flex: 1, overflow: 'hidden' }}>

        {/* Left Panel */}
        <div style={{ background: '#f0f0f0', borderRight: '1px solid #bbb', overflowY: 'auto' }}>
          <div style={{ padding: '6px 10px', background: '#e0e0e0', borderBottom: '1px solid #ccc', fontWeight: 700, fontSize: 13 }}>📊 สถิติ</div>
          <div style={{ padding: 15 }}>
            <div style={statCard}>
              <div style={{ fontSize: 11, color: '#555', marginBottom: 6 }}>📋 Product Master</div>
              <input
                type="file"
                id="masterInput"
                accept=".xlsx,.xls,.csv"
                style={{ display: 'none' }}
                onChange={e => { if (e.target.files?.[0]) { handleMasterUpload(e.target.files[0]); e.target.value = '' } }}
              />
              <button
                onClick={() => document.getElementById('masterInput')?.click()}
                style={{ ...btnStyle, width: '100%', fontSize: 11, padding: '5px 8px', borderRadius: 4 }}
              >
                📂 {masterFileName ? masterFileName : 'อัพโหลด Product Master'}
              </button>
              {masterSkuCount !== null && (
                <div style={{ marginTop: 10 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, marginBottom: 4 }}>
                    <span style={{ color: '#555' }}>Product Master</span>
                    <strong style={{ color: '#c74634' }}>{masterSkuCount.toLocaleString()}</strong>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, marginBottom: 6 }}>
                    <span style={{ color: '#555' }}>GrabMart</span>
                    <strong style={{ color: '#4a8bc4' }}>{stats.totalProducts.toLocaleString()}</strong>
                  </div>
                  <div style={{ background: '#eee', borderRadius: 4, height: 6, overflow: 'hidden' }}>
                    <div style={{
                      background: stats.totalProducts >= masterSkuCount ? '#28a745' : '#c74634',
                      height: '100%',
                      width: `${Math.min(100, Math.round(stats.totalProducts / masterSkuCount * 100))}%`,
                      transition: 'width 0.3s'
                    }} />
                  </div>
                  <div style={{ fontSize: 10, color: '#555', marginTop: 4, textAlign: 'right' }}>
                    {stats.totalProducts.toLocaleString()} / {masterSkuCount.toLocaleString()} SKU
                    ({masterSkuCount > 0 ? Math.round(stats.totalProducts / masterSkuCount * 100) : 0}%)
                  </div>
                  <button
                    onClick={exportMissingSkus}
                    disabled={exportingMissing}
                    style={{
                      ...btnStyle,
                      width: '100%',
                      marginTop: 8,
                      fontSize: 11,
                      padding: '5px 8px',
                      borderRadius: 4,
                      background: '#fff3cd',
                      borderColor: '#f0ad4e',
                      color: '#856404',
                      opacity: exportingMissing ? 0.6 : 1
                    }}
                  >
                    {exportingMissing
                      ? '⏳ กำลังดึงข้อมูล...'
                      : masterMissingCount !== null
                        ? `📥 Export ที่ขาดใน GrabMart (${masterMissingCount.toLocaleString()})`
                        : '📥 Export ที่ขาดใน GrabMart'}
                  </button>
                  <button
                    onClick={handleMasterBranchCompare}
                    disabled={masterBranchLoading}
                    style={{
                      ...btnStyle,
                      width: '100%',
                      marginTop: 6,
                      fontSize: 11,
                      padding: '5px 8px',
                      borderRadius: 4,
                      background: '#e8f4fd',
                      borderColor: '#4a8bc4',
                      color: '#1a5276',
                      opacity: masterBranchLoading ? 0.6 : 1
                    }}
                  >
                    {masterBranchLoading ? '⏳ กำลังเปรียบเทียบ...' : '🔍 เปรียบเทียบ 3 สาขา'}
                  </button>
                </div>
              )}
            </div>
            <div style={statCard}>
              <div style={{ fontSize: 11, color: '#000' }}>หมวดหมู่</div>
              <div style={{ fontSize: 24, fontWeight: 700, color: '#c74634' }}>{stats.sheets.length}</div>
            </div>
          </div>

          <div style={{ padding: '6px 10px', background: '#e0e0e0', borderBottom: '1px solid #ccc', fontWeight: 700, fontSize: 13 }}>📁 หมวดหมู่</div>
          <div style={{ padding: '8px 10px', borderBottom: '1px solid #ddd' }}>
            <button
              onClick={exportSelectedCategoryXlsx}
              disabled={selectedSheet === 'all'}
              style={{
                ...btnStyle,
                width: '100%',
                padding: '7px 10px',
                borderRadius: 6,
                fontSize: 11,
                opacity: selectedSheet === 'all' ? 0.5 : 1
              }}
            >
              📥 Export หมวดหมู่ที่เลือก
            </button>
            <button
              onClick={deleteSelectedCategory}
              disabled={selectedSheet === 'all'}
              style={{
                ...btnStyle,
                width: '100%',
                marginTop: 8,
                padding: '7px 10px',
                borderRadius: 6,
                fontSize: 11,
                background: '#f8d7da',
                borderColor: '#dc3545',
                color: '#721c24',
                opacity: selectedSheet === 'all' ? 0.5 : 1
              }}
              title="ลบหมวดหมู่นี้ออกจากทุกสาขา"
            >
              🗑️ ลบหมวดหมู่นี้ทุกสาขา
            </button>
          </div>
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
                  <th style={{ ...th, width: 32, textAlign: 'center' }}>
                    <input
                      type="checkbox"
                      title="เลือกทั้งหมด"
                      checked={products.length > 0 && products.every(p => selectedSkus.has(p['รหัสสินค้า (SKU NUMBER)']))}
                      onChange={e => {
                        setSelectedSkus((prev: Set<string>) => {
                          const next = new Set(prev)
                          if (e.target.checked) {
                            products.forEach((p: Product) => next.add(p['รหัสสินค้า (SKU NUMBER)']))
                          } else {
                            products.forEach((p: Product) => next.delete(p['รหัสสินค้า (SKU NUMBER)']))
                          }
                          return next
                        })
                      }}
                    />
                  </th>
                  <th style={th}>#</th>
                  <th style={th}>หมวดหมู่</th>
                  <th style={th}>รหัสสินค้า</th>
                  <th style={th}>ชื่อสินค้า</th>
                  <th style={th}>ใบอนุญาต</th>
                  <th style={th}>ราคา</th>
                  <th style={th}>SRC</th>
                  <th style={th}>KKL</th>
                  <th style={th}>SSS</th>
                  <th style={th}>ดำเนินการ</th>
                </tr>
              </thead>
              <tbody>
                {products.map((p, i) => {
                  const sku = p['รหัสสินค้า (SKU NUMBER)']
                  const isChecked = selectedSkus.has(sku)
                  return (
                  <tr key={i}
                    style={{ background: isChecked ? '#e8f8e8' : 'transparent' }}
                    onMouseEnter={e => {
                      if (!isChecked) e.currentTarget.style.background = '#e8f4fc'
                      const imgUrl = convertDriveLink(p['*รูปภาพสินค้า'] || '')
                      if (imgUrl) setTooltip({ x: e.clientX, y: e.clientY, url: imgUrl })
                    }}
                    onMouseMove={e => {
                      setTooltip(prev => prev ? { ...prev, x: e.clientX, y: e.clientY } : null)
                    }}
                    onMouseLeave={e => {
                      e.currentTarget.style.background = isChecked ? '#e8f8e8' : 'transparent'
                      setTooltip(null)
                    }}
                  >
                    <td style={{ ...td, textAlign: 'center' }}>
                      <input
                        type="checkbox"
                        checked={isChecked}
                        onChange={e => {
                          setSelectedSkus(prev => {
                            const next = new Set(prev)
                            e.target.checked ? next.add(sku) : next.delete(sku)
                            return next
                          })
                        }}
                      />
                    </td>
                    <td style={td}>{i + 1}</td>
                    <td style={td}>{p['หมวดหมู่สินค้า (CATEGORIES)'] || '-'}</td>
                    <td style={td}><strong>{p['รหัสสินค้า (SKU NUMBER)'] || '-'}</strong></td>
                    <td style={td}>{p['*ชื่อสินค้า (NAME)'] || '-'}</td>
                    <td style={td}>{p['*เลขที่ใบอนุญาตโฆษณา'] || '-'}</td>
                    <td style={{ ...td, color: '#c74634', fontWeight: 600 }}>{p['*ราคาสินค้า'] || '-'}</td>
                    <td style={{ ...td, textAlign: 'center', fontWeight: 700 }}>{renderBranchFlag(p.src)}</td>
                    <td style={{ ...td, textAlign: 'center', fontWeight: 700 }}>{renderBranchFlag(p.kkl)}</td>
                    <td style={{ ...td, textAlign: 'center', fontWeight: 700 }}>{renderBranchFlag(p.sss)}</td>
                    <td style={td}>
                      <button
                        onClick={() => setEditProduct({ ...p, _originalSku: p['รหัสสินค้า (SKU NUMBER)'] })}
                        style={{ ...btnStyle, fontSize: 11, padding: '4px 10px' }}
                      >
                        ✏️ แก้ไข
                      </button>
                    </td>
                  </tr>
                  )
                })}
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
        <strong>🛵 ผลตรวจสอบราคา GRAB</strong>
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
            <span style={{ color: '#000' }}>⚠ ไม่พบข้อมูลใน Supabase: <strong>{grabResults.filter((r: any) => r.notFound).length}</strong></span>
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
                      ? <span style={{ background: '#e0e0e0', color: '#000', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>⚠ ไม่พบข้อมูลใน Supabase</span>
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
            ⚠ Export ไม่พบข้อมูลใน Supabase ({grabResults.filter((r: any) => r.notFound).length})
          </button>
        </div>
        <button onClick={() => setShowGrabModal(false)} style={btnStyle}>ปิด</button>
      </div>
    </div>
  </div>
)}

      {showPriceCalcModal && (() => {
        const hasGrabMap = Object.keys(priceCalcGrabMap).length > 0
        const enriched = priceCalcResults.filter(r => !r.error).map((r: any) => {
          const grabCurrent = hasGrabMap ? priceCalcGrabMap[r.sku] : undefined
          const effectiveChanged = hasGrabMap
            ? (grabCurrent !== undefined ? Math.abs(r.newPrice - grabCurrent) > 0.01 : r.changed)
            : r.changed
          return { ...r, grabCurrent, effectiveChanged }
        })
        const displayRows = enriched.filter(r => r.effectiveChanged)
        return (
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
            <div style={{ background: '#fff', borderRadius: 6, width: 800, maxHeight: '85vh', display: 'flex', flexDirection: 'column', boxShadow: '0 10px 40px rgba(0,0,0,0.3)' }}>

              {/* Header */}
              <div style={{ background: '#4a4a4a', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderRadius: '6px 6px 0 0' }}>
                <strong>🧮 ผลคำนวณราคา</strong>
                <button onClick={() => setShowPriceCalcModal(false)} style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer' }}>×</button>
              </div>

              {/* Grab_menu upload bar */}
              <div style={{ padding: '10px 20px', background: '#f0f7ff', borderBottom: '1px solid #c8dff8', display: 'flex', gap: 10, alignItems: 'center' }}>
                <span style={{ fontSize: 12, color: '#444', fontWeight: 600, whiteSpace: 'nowrap' }}>Grab_menu (.csv):</span>
                <button onClick={() => document.getElementById('priceCalcGrabInput')?.click()}
                  style={{ ...btnStyle, fontSize: 11, padding: '5px 12px', background: priceCalcGrabFile ? '#d4edda' : '#fff', borderColor: priceCalcGrabFile ? '#28a745' : '#999', maxWidth: 280, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                  title={priceCalcGrabFile ? priceCalcGrabFile.name : 'เลือกไฟล์ Grab_menu เพื่อเปรียบเทียบราคาจริง'}
                >
                  {priceCalcGrabFile ? `✅ ${priceCalcGrabFile.name}` : '📂 เลือกไฟล์ Grab_menu'}
                </button>
                {!priceCalcGrabFile && (
                  <span style={{ fontSize: 11, color: '#888' }}>← อัพโหลดเพื่อแสดงราคา Grab จริง (ปัจจุบันใช้ราคาจาก DB)</span>
                )}
                {hasGrabMap && (
                  <span style={{ fontSize: 11, color: '#28a745', marginLeft: 4 }}>✅ โหลด {Object.keys(priceCalcGrabMap).length.toLocaleString()} SKU จาก Grab_menu</span>
                )}
              </div>

              {/* Summary */}
              {priceCalcResults.length > 0 && !priceCalcResults[0]?.error && (
                <div style={{ padding: '10px 20px', background: '#f8f8f8', borderBottom: '1px solid #ddd', display: 'flex', gap: 20, fontSize: 13 }}>
                  <span>ทั้งหมด: <strong>{enriched.length}</strong></span>
                  <span style={{ color: '#28a745' }}>เปลี่ยนแปลง: <strong>{displayRows.length}</strong></span>
                  <span style={{ color: '#000' }}>ราคาเดิม: <strong>{enriched.length - displayRows.length}</strong></span>
                  <div style={{ marginLeft: 'auto', fontSize: 11, color: '#666' }}>
                    สูตร Grab = ระดับ 0 × 0.95 × 1.20 × 1.07
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
                        <th style={th}>ราคาระดับ 0 ใหม่</th>
                        <th style={{ ...th, color: hasGrabMap ? '#000' : '#999' }}>
                          ราคาขาย Grab ปัจจุบัน{!hasGrabMap && <span style={{ fontWeight: 400, fontSize: 10 }}> (DB)</span>}
                        </th>
                        <th style={th}>ราคาขาย Grab ใหม่</th>
                        <th style={th}>สถานะ</th>
                      </tr>
                    </thead>
                    <tbody>
                      {displayRows.map((r: any, i: number) => (
                        <tr key={i} style={{ background: '#fff3cd' }}>
                          <td style={td}>{i + 1}</td>
                          <td style={td}><strong>{r.sku}</strong></td>
                          <td style={td}>{r.name}</td>
                          <td style={td}>{r.level0.toLocaleString()}</td>
                          <td style={td}>
                            {hasGrabMap
                              ? r.grabCurrent !== undefined
                                ? <strong>{r.grabCurrent.toLocaleString()}</strong>
                                : <span style={{ color: '#999', fontSize: 11 }}>ไม่พบใน Grab</span>
                              : r.oldPrice.toLocaleString()
                            }
                          </td>
                          <td style={{ ...td, color: '#28a745', fontWeight: 600 }}>{r.newPrice.toLocaleString()}</td>
                          <td style={td}>
                            <span style={{ background: '#fff3cd', color: '#856404', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>⚠ เปลี่ยนแปลง</span>
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
                    const toExport = displayRows
                    if (toExport.length === 0) return
                    const XLSX = await import('xlsx')
                    const now = new Date()
                    const dd = String(now.getDate()).padStart(2, '0')
                    const mm = String(now.getMonth() + 1).padStart(2, '0')
                    const yyyy = now.getFullYear()
                    const exportRows = toExportRows(
                      toExport.map((r: any) => ({
                        ...r._product,
                        '*ราคาสินค้า': r.newPrice,
                        'รหัสสินค้า (SKU NUMBER)': r.sku
                      }))
                    )
                    const ws = XLSX.utils.json_to_sheet(exportRows)
                    const wb = XLSX.utils.book_new()
                    XLSX.utils.book_append_sheet(wb, ws, 'อัพเดทราคา')
                    XLSX.writeFile(wb, `GM MME ${dd}.${mm}.${yyyy}.xlsx`)
                  }}
                  disabled={displayRows.length === 0}
                  style={{ ...btnStyle, background: '#17a2b8', color: '#fff', borderColor: '#117a8b' }}
                >
                  📥 Export Excel
                </button>
                <button
                  onClick={confirmUpdatePrices}
                  disabled={updatingPrices || displayRows.length === 0}
                  style={{ ...btnStyle, background: '#28a745', color: '#fff', borderColor: '#1e7e34', opacity: updatingPrices ? 0.6 : 1 }}
                >
                  {updatingPrices ? '⏳ กำลังอัพเดท...' : `💾 อัพเดท ${displayRows.length} รายการ`}
                </button>
              </div>
            </div>
          </div>
        )
      })()}


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
              {importError && (
                <div style={{ marginTop: 12, padding: '8px 12px', background: '#f8d7da', borderRadius: 4, fontSize: 13, color: '#721c24' }}>
                  ❌ {importError}
                </div>
              )}
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
                disabled={importingData || !!importError || importSheets.reduce((s, sh) => s + sh.rows.length, 0) === 0}
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
            background: '#fff', borderRadius: 6, width: 560, maxHeight: '92vh', overflowY: 'auto',
            boxShadow: '0 10px 40px rgba(0,0,0,0.3)'
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
                <div style={{ marginTop: 4, padding: 14, borderRadius: 10, border: '2px dashed #1a73e8', background: '#f4f8ff' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                    <label style={{ ...labelStyle, margin: 0, fontSize: 15, color: '#0b57d0', fontWeight: 700 }}>
                      รูปภาพสินค้า + ข้อความมุมขวาล่าง
                    </label>
                    <span style={{ fontSize: 11, fontWeight: 700, color: '#0b57d0', background: '#dce9ff', borderRadius: 999, padding: '4px 8px' }}>
                      ขั้นตอนสำคัญ
                    </span>
                  </div>
                  <div style={{ marginTop: 6, fontSize: 12, lineHeight: 1.5, color: '#355070' }}>
                    1. เลือกรูปสินค้า 2. ตรวจ Preview 3. กดอัพโหลด Drive 4. กดบันทึก
                  </div>
                  <div style={{ marginTop: 10, display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                    <input
                      id="editProductImageInput"
                      type="file"
                      accept="image/png,image/jpeg,image/webp"
                      onChange={e => onPickEditImage(e.target.files?.[0] || null)}
                      style={{ display: 'none' }}
                    />
                    <label
                      htmlFor="editProductImageInput"
                      style={{ ...btnStyle, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', minWidth: 170, background: '#0b57d0', color: '#fff', borderColor: '#0842a0', cursor: 'pointer' }}
                    >
                      📂 เลือกรูปภาพสินค้า
                    </label>
                    <div style={{ flex: '1 1 220px', minHeight: 38, display: 'flex', alignItems: 'center', padding: '8px 10px', borderRadius: 6, border: '1px solid #c7d7f7', background: '#fff', fontSize: 12, color: editImageFileName ? '#1f2328' : '#6b7280' }}>
                      {editImageFileName || 'ยังไม่ได้เลือกรูป (รองรับ PNG / JPEG / WebP)'}
                    </div>
                  </div>
                  <div style={{ marginTop: 10, fontSize: 12, color: '#4b5563' }}>
                    ข้อความมุมขวาล่างจะใส่ SKU ให้อัตโนมัติ และแก้ไขเองได้
                  </div>
                  <div style={{ marginTop: 8, display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: '#355070' }}>ข้อความบนรูป:</div>
                  <input
                    style={{ ...inputStyle, flex: '1 1 160px', marginTop: 0 }}
                    placeholder="ข้อความบนรูป (มุมขวาล่าง)"
                    value={editImageOverlay}
                    onChange={e => setEditImageOverlay(e.target.value)}
                  />
                  <button
                    type="button"
                    onClick={downloadEditImage}
                    disabled={!editImageSrc}
                    style={{ ...btnStyle, background: '#1a73e8', color: '#fff', borderColor: '#1558b3' }}
                  >
                    ⬇️ Download รูป
                  </button>
                  </div>
                  <div style={{ marginTop: 10, textAlign: 'center', background: '#fff', borderRadius: 8, padding: 10, border: '1px solid #dbe4f0' }}>
                    {editImageSrc ? (
                      <canvas
                        ref={editCanvasRef}
                        style={{ maxWidth: '100%', maxHeight: 220, borderRadius: 4, background: '#fff' }}
                      />
                    ) : (
                      <div style={{ padding: '20px 12px', fontSize: 12, color: '#6b7280' }}>
                        Preview รูปจะขึ้นที่นี่หลังจากเลือกรูปภาพ
                      </div>
                    )}
                  </div>
                </div>
                {editImageMsg && (
                  <div style={{ marginTop: 6, fontSize: 12, color: editImageMsg.includes('สำเร็จ') ? '#1a7f37' : '#a52828' }}>
                    {editImageMsg}
                  </div>
                )}
                <label style={{ ...labelStyle, marginTop: 10 }}>รูปภาพสินค้า (Google Drive Link)</label>
                <input style={inputStyle}
                  placeholder="ลิงก์จะถูกใส่อัตโนมัติหลังอัพโหลด หรือวาง Google Drive URL ที่นี่..."
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

      {/* GM MME vs GRAB Modal */}
      {showMmeCheckModal && (() => {
        const branches: { key: 'src'|'kkl'|'sss'; label: string; file: File|null; inputId: string }[] = [
          { key: 'src', label: 'SRC', file: mmeCheckGrabFileSrc, inputId: 'mmeCheckGrabInputSrc' },
          { key: 'kkl', label: 'KKL', file: mmeCheckGrabFileKkl, inputId: 'mmeCheckGrabInputKkl' },
          { key: 'sss', label: 'SSS', file: mmeCheckGrabFileSss, inputId: 'mmeCheckGrabInputSss' },
        ]
        const hasAnyGrab = !!(mmeCheckGrabFileSrc || mmeCheckGrabFileKkl || mmeCheckGrabFileSss)
        const canCheck = hasAnyGrab && !!mmeCheckMmeFile && !mmeChecking
        const tabResults = mmeActiveTab !== 'compare' ? (mmeCheckResults[mmeActiveTab] ?? []) : []
        const hasResults = Object.values(mmeCheckResults).some(r => r.length > 0)
        const checkedBranches = branches.filter(b => mmeCheckResults[b.key].length > 0)
        const tabHasFile = branches.find(b => b.key === mmeActiveTab)?.file

        // build compare rows: union of all SKUs
        const skuMap: Record<string, { name: string; mmePrice: number|null; src?: any; kkl?: any; sss?: any }> = {}
        checkedBranches.forEach(b => {
          mmeCheckResults[b.key].forEach((r: any) => {
            if (!skuMap[r.sku]) skuMap[r.sku] = { name: r.name, mmePrice: r.mmePrice }
            skuMap[r.sku][b.key] = r
          })
        })
        const compareRows = Object.entries(skuMap).map(([sku, v]) => ({ sku, ...v }))
        const compareNotUpdatedCount = compareRows.filter(r =>
          checkedBranches.some(b => r[b.key as 'src'|'kkl'|'sss']?.status === 'notUpdated')
        ).length

        const statusBadge = (r: any) => {
          if (!r) return <span style={{ color: '#bbb', fontSize: 11 }}>—</span>
          if (r.status === 'updated') return <span style={{ background: '#d4edda', color: '#155724', padding: '2px 6px', borderRadius: 10, fontSize: 10 }}>✅</span>
          if (r.status === 'notFound') return <span style={{ background: '#e0e0e0', color: '#555', padding: '2px 6px', borderRadius: 10, fontSize: 10 }}>⚠</span>
          return <span style={{ background: '#f8d7da', color: '#721c24', padding: '2px 6px', borderRadius: 10, fontSize: 10 }}>❌</span>
        }

        return (
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
            <div style={{ background: '#fff', borderRadius: 6, width: mmeActiveTab === 'compare' ? 920 : 800, maxHeight: '88vh', display: 'flex', flexDirection: 'column', boxShadow: '0 10px 40px rgba(0,0,0,0.3)', transition: 'width 0.2s' }}>

              {/* Header */}
              <div style={{ background: '#2d5fa6', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderRadius: '6px 6px 0 0' }}>
                <strong>📋 ตรวจสอบ GM MME vs GRAB Menu</strong>
                <button onClick={() => setShowMmeCheckModal(false)} style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer' }}>×</button>
              </div>

              {/* Upload section */}
              <div style={{ padding: '12px 20px', background: '#f0f7ff', borderBottom: '1px solid #c8dff8' }}>
                <div style={{ display: 'flex', gap: 10, marginBottom: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                  <span style={{ fontSize: 12, color: '#444', whiteSpace: 'nowrap', fontWeight: 600 }}>Grab_menu (.csv):</span>
                  {branches.map(b => (
                    <button key={b.key} onClick={() => document.getElementById(b.inputId)?.click()}
                      style={{ ...btnStyle, fontSize: 11, padding: '5px 10px', background: b.file ? '#d4edda' : '#fff', borderColor: b.file ? '#28a745' : '#999', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                      title={b.file ? b.file.name : `เลือกไฟล์ Grab_menu สาขา ${b.label}`}
                    >
                      {b.file ? `✅ ${b.label}` : `📂 ${b.label}`}
                    </button>
                  ))}
                </div>
                <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
                  <span style={{ fontSize: 12, color: '#444', whiteSpace: 'nowrap', fontWeight: 600 }}>GM MME (.xlsx):</span>
                  <button onClick={() => document.getElementById('mmeCheckMmeInput')?.click()}
                    style={{ ...btnStyle, fontSize: 11, padding: '5px 10px', background: mmeCheckMmeFile ? '#d4edda' : '#fff', borderColor: mmeCheckMmeFile ? '#28a745' : '#999', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                    title={mmeCheckMmeFile ? mmeCheckMmeFile.name : 'เลือกไฟล์ GM MME'}
                  >
                    {mmeCheckMmeFile ? `✅ ${mmeCheckMmeFile.name}` : '📂 เลือกไฟล์'}
                  </button>
                  <button onClick={handleMmeVsGrabCheck} disabled={!canCheck}
                    style={{ ...btnStyle, background: canCheck ? '#2d5fa6' : '#ccc', color: canCheck ? '#fff' : '#999', borderColor: '#2d5fa6', marginLeft: 'auto', opacity: mmeChecking ? 0.6 : 1 }}
                  >
                    {mmeChecking ? '⏳ กำลังตรวจสอบ...' : '🔍 ตรวจสอบ'}
                  </button>
                </div>
              </div>

              {/* Tabs */}
              {hasResults && (
                <div style={{ display: 'flex', borderBottom: '2px solid #ddd', background: '#fafafa' }}>
                  {checkedBranches.map(b => {
                    const notUpd = mmeCheckResults[b.key].filter((r: any) => r.status === 'notUpdated').length
                    const isActive = mmeActiveTab === b.key
                    return (
                      <button key={b.key} onClick={() => setMmeActiveTab(b.key)}
                        style={{ padding: '8px 18px', border: 'none', borderBottom: isActive ? '2px solid #2d5fa6' : '2px solid transparent', background: 'none', cursor: 'pointer', fontWeight: isActive ? 700 : 400, color: isActive ? '#2d5fa6' : '#555', fontSize: 13, marginBottom: -2 }}
                      >
                        🛵 {b.label}
                        {notUpd > 0 && <span style={{ marginLeft: 5, background: '#dc3545', color: '#fff', borderRadius: 10, fontSize: 10, padding: '1px 6px' }}>{notUpd}</span>}
                      </button>
                    )
                  })}
                  {checkedBranches.length >= 2 && (
                    <button onClick={() => setMmeActiveTab('compare')}
                      style={{ padding: '8px 18px', border: 'none', borderBottom: mmeActiveTab === 'compare' ? '2px solid #6f42c1' : '2px solid transparent', background: 'none', cursor: 'pointer', fontWeight: mmeActiveTab === 'compare' ? 700 : 400, color: mmeActiveTab === 'compare' ? '#6f42c1' : '#555', fontSize: 13, marginBottom: -2, marginLeft: 'auto' }}
                    >
                      📊 เปรียบเทียบ
                      {compareNotUpdatedCount > 0 && <span style={{ marginLeft: 5, background: '#dc3545', color: '#fff', borderRadius: 10, fontSize: 10, padding: '1px 6px' }}>{compareNotUpdatedCount}</span>}
                    </button>
                  )}
                </div>
              )}

              {/* Summary bar */}
              {mmeActiveTab !== 'compare' && tabResults.length > 0 && (
                <div style={{ padding: '8px 20px', background: '#f8f8f8', borderBottom: '1px solid #ddd', display: 'flex', gap: 20, fontSize: 13 }}>
                  <span>ทั้งหมด: <strong>{tabResults.length}</strong></span>
                  <span style={{ color: '#28a745' }}>✅ แก้ไขแล้ว: <strong>{tabResults.filter((r: any) => r.status === 'updated').length}</strong></span>
                  <span style={{ color: '#dc3545' }}>❌ ยังไม่แก้ไข: <strong>{tabResults.filter((r: any) => r.status === 'notUpdated').length}</strong></span>
                  {tabResults.filter((r: any) => r.status === 'notFound').length > 0 && (
                    <span style={{ color: '#856404' }}>⚠ ไม่พบใน GRAB: <strong>{tabResults.filter((r: any) => r.status === 'notFound').length}</strong></span>
                  )}
                </div>
              )}
              {mmeActiveTab === 'compare' && compareRows.length > 0 && (
                <div style={{ padding: '8px 20px', background: '#f3eeff', borderBottom: '1px solid #d6c8f5', display: 'flex', gap: 20, fontSize: 13 }}>
                  <span>SKU ทั้งหมด: <strong>{compareRows.length}</strong></span>
                  <span style={{ color: '#dc3545' }}>❌ มีสาขาที่ยังไม่แก้ไข: <strong>{compareNotUpdatedCount}</strong></span>
                  <span style={{ color: '#28a745' }}>✅ แก้ไขครบทุกสาขา: <strong>{compareRows.filter(r => checkedBranches.every(b => r[b.key as 'src'|'kkl'|'sss']?.status === 'updated')).length}</strong></span>
                </div>
              )}

              {/* Table area */}
              <div style={{ overflowY: 'auto', flex: 1 }}>
                {!hasResults ? (
                  <div style={{ padding: 40, textAlign: 'center', color: '#888', fontSize: 13 }}>
                    เลือกไฟล์ Grab_menu อย่างน้อย 1 สาขา และ GM MME (.xlsx) แล้วกด &quot;ตรวจสอบ&quot;
                  </div>
                ) : mmeActiveTab === 'compare' ? (
                  /* Compare table */
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: '#f0ebff', position: 'sticky', top: 0 }}>
                        <th style={th}>#</th>
                        <th style={th}>SKU</th>
                        <th style={th}>ชื่อสินค้า</th>
                        <th style={{ ...th, textAlign: 'center' }}>ราคา GM MME</th>
                        {checkedBranches.map(b => (
                          <th key={b.key} style={{ ...th, textAlign: 'center', minWidth: 110 }}>🛵 {b.label}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {compareRows.map((r, i) => {
                        const anyNotUpdated = checkedBranches.some(b => r[b.key as 'src'|'kkl'|'sss']?.status === 'notUpdated')
                        const allUpdated = checkedBranches.every(b => r[b.key as 'src'|'kkl'|'sss']?.status === 'updated')
                        return (
                          <tr key={r.sku} style={{ background: anyNotUpdated ? '#fff0f0' : allUpdated ? '#f6fff6' : 'transparent' }}>
                            <td style={td}>{i + 1}</td>
                            <td style={{ ...td, color: '#000' }}><strong>{r.sku}</strong></td>
                            <td style={{ ...td, color: '#000' }}>{r.name || '-'}</td>
                            <td style={{ ...td, textAlign: 'center', fontWeight: 600 }}>{r.mmePrice?.toLocaleString() ?? '-'}</td>
                            {checkedBranches.map(b => {
                              const br = r[b.key as 'src'|'kkl'|'sss']
                              return (
                                <td key={b.key} style={{ ...td, textAlign: 'center' }}>
                                  {statusBadge(br)}
                                  {br?.grabPrice != null && (
                                    <div style={{ fontSize: 10, color: '#666', marginTop: 2 }}>{br.grabPrice.toLocaleString()}</div>
                                  )}
                                </td>
                              )
                            })}
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                ) : tabResults.length === 0 ? (
                  <div style={{ padding: 40, textAlign: 'center', color: '#888', fontSize: 13 }}>
                    {tabHasFile ? 'ไม่มีข้อมูล' : 'ไม่ได้อัพโหลดไฟล์สาขานี้'}
                  </div>
                ) : (
                  /* Single branch table */
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: '#f5f5f5', position: 'sticky', top: 0 }}>
                        <th style={th}>#</th>
                        <th style={th}>SKU</th>
                        <th style={th}>ชื่อสินค้า</th>
                        <th style={th}>ราคาใน GM MME</th>
                        <th style={th}>ราคาใน GRAB ปัจจุบัน</th>
                        <th style={th}>สถานะ</th>
                      </tr>
                    </thead>
                    <tbody>
                      {tabResults.map((r: any, i: number) => (
                        <tr key={i} style={{ background: r.status === 'updated' ? 'transparent' : r.status === 'notFound' ? '#f5f5f5' : '#fff0f0' }}>
                          <td style={td}>{i + 1}</td>
                          <td style={{ ...td, color: '#000' }}><strong>{r.sku}</strong></td>
                          <td style={{ ...td, color: '#000' }}>{r.name || '-'}</td>
                          <td style={{ ...td, fontWeight: 600 }}>{r.mmePrice?.toLocaleString() ?? '-'}</td>
                          <td style={{ ...td, fontWeight: 600 }}>{r.grabPrice?.toLocaleString() ?? '-'}</td>
                          <td style={td}>
                            {r.status === 'updated'
                              ? <span style={{ background: '#d4edda', color: '#155724', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✅ แก้ไขแล้ว</span>
                              : r.status === 'notFound'
                                ? <span style={{ background: '#e0e0e0', color: '#555', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>⚠ ไม่พบใน GRAB</span>
                                : <span style={{ background: '#f8d7da', color: '#721c24', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>❌ ยังไม่แก้ไข</span>
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
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                  {checkedBranches.map(b => {
                    const notUpdated = mmeCheckResults[b.key].filter((r: any) => r.status === 'notUpdated')
                    return (
                      <button key={b.key}
                        disabled={notUpdated.length === 0}
                        onClick={async () => {
                          if (notUpdated.length === 0) return
                          const XLSX = await import('xlsx')
                          const rows = notUpdated.map((r: any) => ({
                            '*ประเภทสินค้า': r.type || '',
                            '*ชื่อสินค้า': r.name || '',
                            '*เลขที่ใบอนุญาตโฆษณา': r.license || '',
                            '*ราคาสินค้า': r.mmePrice ?? '',
                            'รหัสสินค้า': r.sku || '',
                            '*รูปภาพสินค้า': r.image || '',
                            'หมวดหมู่รายการสินค้า': r.category || '',
                          }))
                          const ws = XLSX.utils.json_to_sheet(rows)
                          notUpdated.forEach((r: any, i: number) => {
                            const url = r.image || ''
                            if (!url) return
                            const cellRef = `F${i + 2}`
                            if (ws[cellRef]) ws[cellRef].l = { Target: url }
                          })
                          const wb2 = XLSX.utils.book_new()
                          XLSX.utils.book_append_sheet(wb2, ws, 'ยังไม่แก้ไข')
                          const now = new Date()
                          const dd = String(now.getDate()).padStart(2, '0')
                          const mm2 = String(now.getMonth() + 1).padStart(2, '0')
                          const yyyy = now.getFullYear()
                          XLSX.writeFile(wb2, `GM MME ยังไม่แก้ไข ${b.label} ${dd}.${mm2}.${yyyy}.xlsx`)
                        }}
                        style={{ ...btnStyle, background: '#f8d7da', borderColor: '#dc3545', color: '#721c24', opacity: notUpdated.length === 0 ? 0.5 : 1 }}
                      >
                        📥 Export {b.label} ({notUpdated.length})
                      </button>
                    )
                  })}
                </div>
                <button onClick={() => setShowMmeCheckModal(false)} style={btnStyle}>ปิด</button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* Missing Branch Modal */}
      {showMissingBranchModal && (() => {
        const branches: { key: 'src'|'kkl'|'sss'; label: string }[] = [
          { key: 'src', label: 'SRC' },
          { key: 'kkl', label: 'KKL' },
          { key: 'sss', label: 'SSS' },
        ]
        const tabData = missingBranchResults[missingBranchTab] ?? []
        const isLoaded = missingBranchLoaded[missingBranchTab]
        return (
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
            <div style={{ background: '#fff', borderRadius: 6, width: 820, maxHeight: '88vh', display: 'flex', flexDirection: 'column', boxShadow: '0 10px 40px rgba(0,0,0,0.3)' }}>

              {/* Header */}
              <div style={{ background: '#e65100', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderRadius: '6px 6px 0 0' }}>
                <strong>🔍 สินค้าที่ขาดใน GRAB</strong>
                <button onClick={() => setShowMissingBranchModal(false)} style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer' }}>×</button>
              </div>

              {/* Tabs */}
              <div style={{ display: 'flex', borderBottom: '2px solid #ddd', background: '#fafafa' }}>
                {branches.map(b => {
                  const isActive = missingBranchTab === b.key
                  const count = missingBranchLoaded[b.key] ? missingBranchResults[b.key].length : null
                  return (
                    <button key={b.key}
                      onClick={() => { setMissingBranchTab(b.key); setMissingBranchExportCat('all'); setMissingBranchSearch(''); loadMissingBranch(b.key) }}
                      style={{ padding: '8px 24px', border: 'none', borderBottom: isActive ? '2px solid #e65100' : '2px solid transparent', background: 'none', cursor: 'pointer', fontWeight: isActive ? 700 : 400, color: isActive ? '#e65100' : '#555', fontSize: 13, marginBottom: -2 }}
                    >
                      🛵 {b.label}
                      {count !== null && (
                        <span style={{ marginLeft: 6, background: count > 0 ? '#e65100' : '#28a745', color: '#fff', borderRadius: 10, fontSize: 10, padding: '1px 6px' }}>{count}</span>
                      )}
                    </button>
                  )
                })}
                <div style={{ marginLeft: 'auto', padding: '8px 16px', fontSize: 11, color: '#888', alignSelf: 'center' }}>
                  ข้อมูลอ้างอิงจาก branch field ใน DB (อัพเดทเมื่ออัพโหลด GRAB CSV)
                </div>
              </div>

              {/* Summary + Search */}
              {isLoaded && (
                <div style={{ padding: '8px 20px', background: '#fff8f2', borderBottom: '1px solid #ffe0cc', display: 'flex', gap: 16, fontSize: 13, alignItems: 'center', flexWrap: 'wrap' }}>
                  <span>ทั้งหมด: <strong style={{ color: tabData.length > 0 ? '#e65100' : '#28a745' }}>{tabData.length} รายการ</strong></span>
                  <input
                    type="text"
                    placeholder="ค้นหาชื่อ / SKU..."
                    value={missingBranchSearch}
                    onChange={e => setMissingBranchSearch(e.target.value)}
                    style={{ padding: '3px 10px', fontSize: 12, borderRadius: 4, border: '1px solid #ccc', width: 200 }}
                  />
                  {missingBranchSearch && (
                    <span style={{ color: '#888' }}>พบ {tabData.filter((p: any) => {
                      const q = missingBranchSearch.toLowerCase()
                      return (p['*ชื่อสินค้า (NAME)'] || '').toLowerCase().includes(q) || (p['รหัสสินค้า (SKU NUMBER)'] || '').toLowerCase().includes(q)
                    }).length} รายการ</span>
                  )}
                </div>
              )}

              {/* Table */}
              <div style={{ overflowY: 'auto', flex: 1 }}>
                {!isLoaded || missingBranchLoading ? (
                  <div style={{ padding: 40, textAlign: 'center', color: '#888', fontSize: 13 }}>⏳ กำลังโหลด...</div>
                ) : tabData.length === 0 ? (
                  <div style={{ padding: 40, textAlign: 'center', color: '#28a745', fontSize: 14 }}>
                    ✅ ไม่มีสินค้าที่ขาดใน GRAB {missingBranchTab.toUpperCase()}
                  </div>
                ) : (
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: '#f5f5f5', position: 'sticky', top: 0 }}>
                        <th style={th}>#</th>
                        <th style={th}>SKU</th>
                        <th style={th}>ชื่อสินค้า</th>
                        <th style={th}>ราคา</th>
                        <th style={th}>หมวดหมู่</th>
                        <th style={{ ...th, textAlign: 'center' }}>SRC</th>
                        <th style={{ ...th, textAlign: 'center' }}>KKL</th>
                        <th style={{ ...th, textAlign: 'center' }}>SSS</th>
                      </tr>
                    </thead>
                    <tbody>
                      {tabData.filter((p: any) => {
                        if (!missingBranchSearch) return true
                        const q = missingBranchSearch.toLowerCase()
                        return (p['*ชื่อสินค้า (NAME)'] || '').toLowerCase().includes(q) || (p['รหัสสินค้า (SKU NUMBER)'] || '').toLowerCase().includes(q)
                      }).map((p: any, i: number) => (
                        <tr key={p['รหัสสินค้า (SKU NUMBER)']} style={{ background: '#fff8f2' }}>
                          <td style={td}>{i + 1}</td>
                          <td style={{ ...td, color: '#000' }}><strong>{p['รหัสสินค้า (SKU NUMBER)']}</strong></td>
                          <td style={{ ...td, color: '#000' }}>{p['*ชื่อสินค้า (NAME)'] || '-'}</td>
                          <td style={{ ...td, fontWeight: 600 }}>{p['*ราคาสินค้า'] || '-'}</td>
                          <td style={td}>{p['หมวดหมู่สินค้า (CATEGORIES)'] || '-'}</td>
                          {(['src','kkl','sss'] as const).map(b => (
                            <td key={b} style={{ ...td, textAlign: 'center' }}>
                              {p[b]
                                ? <span style={{ background: '#d4edda', color: '#155724', padding: '2px 6px', borderRadius: 10, fontSize: 10 }}>✅</span>
                                : <span style={{ background: '#f8d7da', color: '#721c24', padding: '2px 6px', borderRadius: 10, fontSize: 10 }}>❌</span>
                              }
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>

              {/* Footer */}
              <div style={{ padding: '12px 20px', borderTop: '1px solid #ddd', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                  {(() => {
                    const cats = isLoaded ? [...new Set(tabData.map((p: any) => p['หมวดหมู่สินค้า (CATEGORIES)'] || '').filter(Boolean))].sort((a, b) => a.localeCompare(b, 'th')) : []
                    const exportData = missingBranchExportCat === 'all' ? tabData : tabData.filter((p: any) => (p['หมวดหมู่สินค้า (CATEGORIES)'] || '') === missingBranchExportCat)
                    return (<>
                      <select
                        value={missingBranchExportCat}
                        onChange={e => setMissingBranchExportCat(e.target.value)}
                        disabled={!isLoaded || tabData.length === 0}
                        style={{ padding: '5px 8px', fontSize: 12, borderRadius: 4, border: '1px solid #ccc', cursor: 'pointer' }}
                      >
                        <option value="all">ทุกหมวดหมู่ ({tabData.length})</option>
                        {cats.map((c: string) => {
                          const cnt = tabData.filter((p: any) => (p['หมวดหมู่สินค้า (CATEGORIES)'] || '') === c).length
                          return <option key={c} value={c}>{c} ({cnt})</option>
                        })}
                      </select>
                      <button
                        disabled={!isLoaded || exportData.length === 0}
                        onClick={async () => {
                          if (exportData.length === 0) return
                          const XLSX = await import('xlsx')
                          const sorted = [...exportData].sort((a: any, b: any) =>
                            (a['หมวดหมู่สินค้า (CATEGORIES)'] || '').localeCompare(b['หมวดหมู่สินค้า (CATEGORIES)'] || '', 'th')
                          )
                          const rows = sorted.map((p: any) => ({
                            '*ประเภทสินค้า': p['ประเภทสินค้า'] || '',
                            '*ชื่อสินค้า': p['*ชื่อสินค้า (NAME)'] || '',
                            '*เลขที่ใบอนุญาตโฆษณา': p['*เลขที่ใบอนุญาตโฆษณา'] || '',
                            '*ราคาสินค้า': p['*ราคาสินค้า'] || '',
                            'รหัสสินค้า': p['รหัสสินค้า (SKU NUMBER)'] || '',
                            '*รูปภาพสินค้า': p['*รูปภาพสินค้า'] || '',
                            'หมวดหมู่รายการสินค้า': p['หมวดหมู่สินค้า (CATEGORIES)'] || '',
                          }))
                          const ws = XLSX.utils.json_to_sheet(rows)
                          sorted.forEach((p: any, i: number) => {
                            const url = p['*รูปภาพสินค้า'] || ''
                            if (!url) return
                            const cellRef = `F${i + 2}`
                            if (ws[cellRef]) ws[cellRef].l = { Target: url }
                          })
                          const wb = XLSX.utils.book_new()
                          const sheetName = missingBranchExportCat === 'all' ? `ขาด ${missingBranchTab.toUpperCase()}` : missingBranchExportCat.slice(0, 31)
                          XLSX.utils.book_append_sheet(wb, ws, sheetName)
                          const now = new Date()
                          const dd = String(now.getDate()).padStart(2, '0')
                          const mm2 = String(now.getMonth() + 1).padStart(2, '0')
                          const yyyy = now.getFullYear()
                          const catSuffix = missingBranchExportCat === 'all' ? '' : ` ${missingBranchExportCat}`
                          XLSX.writeFile(wb, `สินค้าขาด GRAB ${missingBranchTab.toUpperCase()}${catSuffix} ${dd}.${mm2}.${yyyy}.xlsx`)
                        }}
                        style={{ ...btnStyle, background: '#fff3e0', borderColor: '#e65100', color: '#e65100', opacity: (!isLoaded || exportData.length === 0) ? 0.5 : 1 }}
                      >
                        📥 Export ({exportData.length})
                      </button>
                    </>)
                  })()}
                  <button
                    disabled={missingBranchLoading}
                    onClick={() => {
                      setMissingBranchLoaded({ src: false, kkl: false, sss: false })
                      setMissingBranchResults({ src: [], kkl: [], sss: [] })
                      setTimeout(() => loadMissingBranch(missingBranchTab), 0)
                    }}
                    style={{ ...btnStyle, opacity: missingBranchLoading ? 0.5 : 1 }}
                  >
                    🔄 รีโหลด
                  </button>
                </div>
                <button onClick={() => setShowMissingBranchModal(false)} style={btnStyle}>ปิด</button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* Master Branch Compare Modal */}
      {showMasterBranchModal && (() => {
        const branchTabs: { key: 'all'|'src'|'kkl'|'sss'; label: string }[] = [
          { key: 'all', label: 'ทั้งหมด' },
          { key: 'src', label: 'ขาด SRC' },
          { key: 'kkl', label: 'ขาด KKL' },
          { key: 'sss', label: 'ขาด SSS' },
        ]
        const filtered = masterBranchResults.filter(r => {
          if (masterBranchTab === 'src') return !r.src
          if (masterBranchTab === 'kkl') return !r.kkl
          if (masterBranchTab === 'sss') return !r.sss
          return true
        }).filter(r => {
          if (!masterBranchSearch) return true
          const q = masterBranchSearch.toLowerCase()
          return r.sku.toLowerCase().includes(q) || r.name.toLowerCase().includes(q)
        })
        const missingCounts = {
          src: masterBranchResults.filter(r => !r.src).length,
          kkl: masterBranchResults.filter(r => !r.kkl).length,
          sss: masterBranchResults.filter(r => !r.sss).length,
        }
        const branchBadge = (has: boolean) => (
          <span style={{
            display: 'inline-block', padding: '2px 8px', borderRadius: 10, fontSize: 11, fontWeight: 600,
            background: has ? '#d4edda' : '#f8d7da',
            color: has ? '#155724' : '#721c24'
          }}>{has ? '✓' : '✗'}</span>
        )
        return (
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
            <div style={{ background: '#fff', borderRadius: 8, width: '90vw', maxWidth: 900, height: '85vh', display: 'flex', flexDirection: 'column', boxShadow: '0 20px 60px rgba(0,0,0,0.3)' }}>
              <div style={{ padding: '12px 20px', borderBottom: '1px solid #ddd', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div>
                  <strong style={{ fontSize: 15 }}>🔍 Product Master vs สาขา</strong>
                  <span style={{ marginLeft: 12, fontSize: 12, color: '#555' }}>
                    {masterBranchResults.length} SKU
                    {missingCounts.src + missingCounts.kkl + missingCounts.sss > 0 && (
                      <> — ขาด SRC: <strong style={{ color: '#c74634' }}>{missingCounts.src}</strong>{' '}
                      KKL: <strong style={{ color: '#c74634' }}>{missingCounts.kkl}</strong>{' '}
                      SSS: <strong style={{ color: '#c74634' }}>{missingCounts.sss}</strong></>
                    )}
                  </span>
                </div>
                <button onClick={() => setShowMasterBranchModal(false)} style={{ ...btnStyle, padding: '4px 12px' }}>✕ ปิด</button>
              </div>
              <div style={{ display: 'flex', borderBottom: '2px solid #ddd', background: '#fafafa' }}>
                {branchTabs.map(t => {
                  const isActive = masterBranchTab === t.key
                  const cnt = t.key === 'all' ? masterBranchResults.length : missingCounts[t.key]
                  return (
                    <button key={t.key}
                      onClick={() => setMasterBranchTab(t.key)}
                      style={{
                        padding: '8px 16px', border: 'none', cursor: 'pointer', fontSize: 12, fontWeight: isActive ? 700 : 400,
                        background: isActive ? '#fff' : 'transparent',
                        borderBottom: isActive ? '2px solid #4a8bc4' : '2px solid transparent',
                        color: isActive ? '#4a8bc4' : (t.key !== 'all' && cnt > 0 ? '#c74634' : '#555')
                      }}
                    >{t.label} ({cnt})</button>
                  )
                })}
              </div>
              <div style={{ padding: '8px 16px', display: 'flex', gap: 8, alignItems: 'center', borderBottom: '1px solid #eee' }}>
                <input
                  type="text"
                  placeholder="ค้นหา SKU / ชื่อสินค้า..."
                  value={masterBranchSearch}
                  onChange={e => setMasterBranchSearch(e.target.value)}
                  style={{ padding: '4px 10px', fontSize: 12, borderRadius: 4, border: '1px solid #ccc', width: 220 }}
                />
                {masterBranchSearch && <span style={{ color: '#888', fontSize: 12 }}>พบ {filtered.length} รายการ</span>}
                <button
                  onClick={async () => {
                    if (filtered.length === 0) return
                    const XLSX = await import('xlsx')
                    const rows = filtered.map((r, i) => ({
                      '#': i + 1,
                      'SKU': r.sku,
                      'ชื่อสินค้า': r.name,
                      'SRC': r.src ? '✓' : '✗',
                      'KKL': r.kkl ? '✓' : '✗',
                      'SSS': r.sss ? '✓' : '✗',
                    }))
                    const ws = XLSX.utils.json_to_sheet(rows)
                    const wb = XLSX.utils.book_new()
                    const tabLabel = masterBranchTab === 'all' ? 'ทั้งหมด' : `ขาด ${masterBranchTab.toUpperCase()}`
                    XLSX.utils.book_append_sheet(wb, ws, tabLabel.slice(0, 31))
                    const now = new Date()
                    const dd = String(now.getDate()).padStart(2, '0')
                    const mm = String(now.getMonth() + 1).padStart(2, '0')
                    const yyyy = now.getFullYear()
                    XLSX.writeFile(wb, `Master vs สาขา ${tabLabel} ${dd}.${mm}.${yyyy}.xlsx`)
                  }}
                  disabled={filtered.length === 0}
                  style={{ ...btnStyle, fontSize: 11, padding: '4px 10px', background: '#e8f5e9', borderColor: '#388e3c', color: '#1b5e20', opacity: filtered.length === 0 ? 0.5 : 1 }}
                >📥 Export</button>
              </div>
              <div style={{ overflowY: 'auto', flex: 1 }}>
                {masterBranchLoading ? (
                  <div style={{ padding: 40, textAlign: 'center', color: '#888', fontSize: 13 }}>⏳ กำลังโหลด...</div>
                ) : filtered.length === 0 ? (
                  <div style={{ padding: 40, textAlign: 'center', color: '#28a745', fontSize: 14 }}>
                    {masterBranchResults.length === 0 ? 'ไม่มีข้อมูล — กรุณาอัพโหลด Product Master ก่อน' : '✅ ไม่มีรายการที่ขาดในสาขานี้'}
                  </div>
                ) : (
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr>
                        <th style={{ ...th, width: 40 }}>#</th>
                        <th style={th}>SKU</th>
                        <th style={th}>ชื่อสินค้า</th>
                        <th style={{ ...th, width: 70, textAlign: 'center' }}>SRC</th>
                        <th style={{ ...th, width: 70, textAlign: 'center' }}>KKL</th>
                        <th style={{ ...th, width: 70, textAlign: 'center' }}>SSS</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filtered.map((r, i) => (
                        <tr key={r.sku} style={{ background: (!r.src || !r.kkl || !r.sss) ? '#fff8f8' : 'transparent' }}>
                          <td style={{ ...td, color: '#888', textAlign: 'center' }}>{i + 1}</td>
                          <td style={{ ...td, fontFamily: 'monospace', fontWeight: 600 }}>{r.sku}</td>
                          <td style={td}>{r.name || <span style={{ color: '#aaa' }}>—</span>}</td>
                          <td style={{ ...td, textAlign: 'center' }}>{branchBadge(r.src)}</td>
                          <td style={{ ...td, textAlign: 'center' }}>{branchBadge(r.kkl)}</td>
                          <td style={{ ...td, textAlign: 'center' }}>{branchBadge(r.sss)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          </div>
        )
      })()}

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

