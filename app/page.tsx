'use client'

import { useState, useEffect, useEffectEvent, useRef } from 'react'
import Link from 'next/link'
import NextImage from 'next/image'

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

type BranchKey = 'src' | 'kkl' | 'sss'
type PriceState = 'mismatch' | 'matched' | 'unpriced' | 'absent'

function WorkspaceIcon({ name, className = '' }: { name: string; className?: string }) {
  const paths: Record<string, React.ReactNode> = {
    grid: <><rect x="3" y="3" width="7" height="7" rx="1.5" /><rect x="14" y="3" width="7" height="7" rx="1.5" /><rect x="3" y="14" width="7" height="7" rx="1.5" /><rect x="14" y="14" width="7" height="7" rx="1.5" /></>,
    box: <><path d="m3 7 9-4 9 4v10l-9 4-9-4Z" /><path d="m3 7 9 4 9-4M12 11v10M7 5l10 5" /></>,
    tag: <><path d="M3 3h8l10 10-8 8L3 11Z" /><circle cx="7.5" cy="7.5" r="1" /></>,
    search: <><circle cx="10.5" cy="10.5" r="6.5" /><path d="m16 16 5 5" /></>,
    branch: <><path d="M6 5v10a4 4 0 0 0 4 4h6M6 8h8a4 4 0 0 0 4-4" /><circle cx="6" cy="4" r="2" /><circle cx="18" cy="3" r="2" /><circle cx="18" cy="19" r="2" /></>,
    upload: <><path d="M12 16V3m-5 5 5-5 5 5M4 16v5h16v-5" /></>,
    download: <><path d="M12 3v13m-5-5 5 5 5-5M4 17v4h16v-4" /></>,
    check: <path d="m5 12 4 4 10-10" />,
    clock: <><circle cx="12" cy="12" r="9" /><path d="M12 7v6l4 2" /></>,
    arrow: <path d="M4 12h15m-6-6 6 6-6 6" />,
    chevron: <path d="m9 5 7 7-7 7" />,
    filter: <path d="M3 5h18l-7 8v6l-4 2v-8Z" />,
    settings: <><path d="M4 7h16M4 17h16" /><circle cx="9" cy="7" r="3" /><circle cx="15" cy="17" r="3" /></>,
    refresh: <><path d="M20 7v5h-5M4 17v-5h5" /><path d="M6 6a8 8 0 0 1 13 3M5 15a8 8 0 0 0 13 3" /></>,
    plus: <path d="M12 4v16M4 12h16" />,
    close: <path d="m6 6 12 12M18 6 6 18" />,
  }
  return <svg className={`pc-icon ${className}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths[name] ?? paths.box}</svg>
}

export default function Home() {
  const [products, setProducts] = useState<Product[]>([])
  const [stats, setStats] = useState<Stats>({ totalProducts: 0, sheets: [] })
  const [search, setSearch] = useState('')
  const [selectedSheet, setSelectedSheet] = useState('all')
  const [status, setStatus] = useState('พร้อมใช้งาน')
  const [loading, setLoading] = useState(true)
  const [tooltip, setTooltip] = useState<{ x: number; y: number; url: string } | null>(null)
  const [editProduct, setEditProduct] = useState<Product | null>(null)
  const [grabFile, setGrabFile] = useState<File | null>(null)
  const [grabResults, setGrabResults] = useState<any[]>([])
  const [showGrabModal, setShowGrabModal] = useState(false)
  const [grabMismatchProducts, setGrabMismatchProducts] = useState<Product[]>([])
  const [grabSource, setGrabSource] = useState<'GRAB'>('GRAB')
  const [grabSearch, setGrabSearch] = useState('')
  const [grabShowOnlyMismatch, setGrabShowOnlyMismatch] = useState(true)
  const [grabEdits, setGrabEdits] = useState<Record<string, number>>({})
  const [editingGrabSku, setEditingGrabSku] = useState<string | null>(null)
  const [editingGrabValue, setEditingGrabValue] = useState('')
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
  const [showPromaxxModal, setShowPromaxxModal] = useState(false)
  const [promaxxResults, setPromaxxResults] = useState<any[]>([])
  const [promaxxChecking, setPromaxxChecking] = useState(false)
  const [promaxxProgress, setPromaxxProgress] = useState({ current: 0, total: 0, phase: '' })
  const [editImageSrc, setEditImageSrc] = useState<string>('')
  const [editImageOverlay, setEditImageOverlay] = useState<string>('')
  const [editImageMime, setEditImageMime] = useState<string>('image/png')
  const [editImageFileName, setEditImageFileName] = useState<string>('')
  const [editImageMsg, setEditImageMsg] = useState<string>('')
  const [uploadingDrive, setUploadingDrive] = useState(false)
  const [driveLinkOwner, setDriveLinkOwner] = useState<{ type: 'anin' | 'external' | 'error' | 'loading'; email: string } | null>(null)
  const [saveBtnAnim, setSaveBtnAnim] = useState(false)
  const [cleaningImage, setCleaningImage] = useState(false)
  const editCanvasRef = useRef<HTMLCanvasElement | null>(null)
  const searchReqId = useRef(0)
  const searchInputRef = useRef<HTMLInputElement>(null)
  const inspectorRef = useRef<HTMLElement>(null)
  const [activeSku, setActiveSku] = useState<string | null>(null)
  const [inspectionBranch, setInspectionBranch] = useState<BranchKey>('src')
  const [branchFilter, setBranchFilter] = useState<'all' | BranchKey>('all')
  const [priceFilter, setPriceFilter] = useState<'all' | PriceState>('all')
  const [showAllPrices, setShowAllPrices] = useState(false)
  const [toolsOpen, setToolsOpen] = useState(false)
  const [filtersOpen, setFiltersOpen] = useState(false)
  const [pageNumber, setPageNumber] = useState(1)
  const [loadError, setLoadError] = useState('')
  const [priceDrafts, setPriceDrafts] = useState<Record<string, string>>({})
  const [priceReview, setPriceReview] = useState<{ sku: string; price: number; previous: number | null } | null>(null)
  const [savingPrice, setSavingPrice] = useState(false)
  const [inspectorError, setInspectorError] = useState('')

  const onSearchCriteriaChange = useEffectEvent(() => {
    if (search.trim().length >= 3) searchProducts(search.trim())
    else loadProducts(selectedSheet)
  })

  useEffect(() => {
    loadStats()
    loadMasterCount()
  }, [])

  useEffect(() => {
    // Invalidate an older request immediately, including during the debounce window.
    searchReqId.current += 1
    const timeout = setTimeout(() => {
      onSearchCriteriaChange()
    }, search.trim().length >= 3 ? 250 : 0)
    return () => { clearTimeout(timeout); searchReqId.current += 1 }
  }, [search, selectedSheet])

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key === 'k') {
        event.preventDefault()
        searchInputRef.current?.focus()
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [])

  async function loadMasterCount() {
    try {
      const res = await fetch('/api/master')
      if (!res.ok) return
      const data = await res.json()
      if (data.success) setMasterSkuCount(data.count)
    } catch { /* The product request displays the connection error. */ }
  }

  async function loadStats() {
    try {
      const res = await fetch('/api/stats')
      if (!res.ok) return
      const data = await res.json()
      if (data.success) setStats(data)
    } catch { /* Keep the last successfully loaded summary. */ }
  }

  async function loadProducts(sheet = 'all') {
    const url = sheet === 'all' ? '/api/products' : `/api/products?sheet=${encodeURIComponent(sheet)}`
    await fetchProductList(url)
  }

  async function searchProducts(sku: string) {
    await fetchProductList(`/api/products?q=${encodeURIComponent(sku)}`)
  }

  async function fetchProductList(url: string) {
    const reqId = ++searchReqId.current
    setLoading(true)
    setLoadError('')
    setPriceReview(null)
    setStatus('กำลังโหลดสินค้า...')
    try {
      const res = await fetch(url)
      if (!res.ok) throw new Error('โหลดข้อมูลไม่สำเร็จ')
      const data = await res.json()
      if (!data.success || !Array.isArray(data.products)) throw new Error('โหลดข้อมูลไม่สำเร็จ')
      if (reqId !== searchReqId.current) return
      setProducts(data.products)
      preloadImages(data.products)
      setStatus(`โหลดสินค้าแล้ว ${data.products.length.toLocaleString('th-TH')} รายการ`)
    } catch {
      if (reqId !== searchReqId.current) return
      setLoadError('ไม่สามารถโหลดสินค้าได้ กรุณาตรวจสอบการเชื่อมต่อแล้วลองอีกครั้ง')
      setStatus('โหลดข้อมูลสินค้าไม่สำเร็จ')
    } finally {
      if (reqId === searchReqId.current) setLoading(false)
    }
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
      const savedSku = String(isAdding ? data.product['รหัสสินค้า (SKU NUMBER)'] : editProduct['รหัสสินค้า (SKU NUMBER)'])
      setActiveSku(savedSku)
      setPriceDrafts(previous => { const next = { ...previous }; delete next[savedSku]; delete next[String(editProduct['_originalSku'])]; return next })
      resetPriceReview()
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
      // Download and Drive upload both re-encode from this canvas, so this cap is what decides
      // the exported file's resolution — CSS keeps the on-screen preview small regardless.
      const maxW = 3072
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
      const text = (overlayText || '').trim() ? `เลขที่ใบอนุญาต ${(overlayText || '').trim()}` : ''
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

  async function cleanGreenBorderAndText() {
    const canvas = editCanvasRef.current
    if (!canvas) return
    setCleaningImage(true)
    setEditImageMsg('กำลังลบกรอบ...')
    try {
      let src = editImageSrc
      if (!src) {
        const link = editProduct?.['*รูปภาพสินค้า'] || ''
        const match = link.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || link.match(/[?&]id=([a-zA-Z0-9_-]+)/)
        if (!match) { setEditImageMsg('❌ ไม่พบลิงก์รูปภาพ'); setCleaningImage(false); return }
        src = `https://drive.google.com/thumbnail?id=${match[1]}&sz=w1024`
      }
      const img = new window.Image()
      img.crossOrigin = 'anonymous'
      await new Promise<void>((resolve, reject) => { img.onload = () => resolve(); img.onerror = () => reject(new Error('โหลดรูปไม่สำเร็จ')); img.src = src })
      const w = img.width, h = img.height
      const tmpCanvas = document.createElement('canvas')
      tmpCanvas.width = w; tmpCanvas.height = h
      const ctx = tmpCanvas.getContext('2d')!
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, h)
      ctx.drawImage(img, 0, 0, w, h)
      const imageData = ctx.getImageData(0, 0, w, h)
      const px = imageData.data

      // Treat the picture as islands of non-white pixels joined 8-ways. The decorative frame
      // touches the image border, the promo ribbon touches (or fuses into) it, and an old logo
      // badge sits in a corner — only the product is a big island floating over the middle. So
      // keeping just that island drops every decoration without ever guessing at colours, which
      // is what previous hue-based attempts kept getting wrong on green/yellow packaging.
      const INK = 12
      const total = w * h
      const ink = new Uint8Array(total)
      for (let p = 0; p < total; p++) {
        const i = p * 4
        if (255 - px[i] > INK || 255 - px[i + 1] > INK || 255 - px[i + 2] > INK) ink[p] = 1
      }

      type Island = { label: number; area: number; minX: number; minY: number; maxX: number; maxY: number }
      const labels = new Int32Array(total)
      const stack = new Int32Array(total)
      const islands: Island[] = []
      let nextLabel = 0
      for (let seed = 0; seed < total; seed++) {
        if (!ink[seed] || labels[seed]) continue
        nextLabel++
        let sp = 0
        stack[sp++] = seed
        labels[seed] = nextLabel
        let area = 0, minX = w, minY = h, maxX = 0, maxY = 0
        while (sp > 0) {
          const p = stack[--sp]
          const x = p % w, y = (p / w) | 0
          area++
          if (x < minX) minX = x
          if (x > maxX) maxX = x
          if (y < minY) minY = y
          if (y > maxY) maxY = y
          for (let dy = -1; dy <= 1; dy++) {
            const ny = y + dy
            if (ny < 0 || ny >= h) continue
            for (let dx = -1; dx <= 1; dx++) {
              const nx = x + dx
              if (nx < 0 || nx >= w) continue
              const np = ny * w + nx
              if (ink[np] && !labels[np]) { labels[np] = nextLabel; stack[sp++] = np }
            }
          }
        }
        islands.push({ label: nextLabel, area, minX, minY, maxX, maxY })
      }

      const edge = 2
      const cx0 = w * 0.2, cx1 = w * 0.8, cy0 = h * 0.2, cy1 = h * 0.8
      const touchesBorder = (s: Island) => s.minX <= edge || s.minY <= edge || s.maxX >= w - 1 - edge || s.maxY >= h - 1 - edge
      const overlapsCenter = (s: Island) => s.maxX >= cx0 && s.minX <= cx1 && s.maxY >= cy0 && s.minY <= cy1
      let product: Island | null = null
      for (const s of islands) {
        if (touchesBorder(s) || !overlapsCenter(s)) continue
        if (!product || s.area > product.area) product = s
      }
      // Bail out rather than hand back a mangled photo — a no-op beats a clipped product.
      if (!product || product.area / total > 0.9 ||
        product.maxX - product.minX < w * 0.05 || product.maxY - product.minY < h * 0.05) {
        setEditImageMsg('❌ ไม่พบสินค้าในรูป')
        setCleaningImage(false)
        return
      }

      // Islands sitting inside the product's own bounds belong to it (a detached cap, a sticker).
      const keep = new Uint8Array(nextLabel + 1)
      keep[product.label] = 1
      const padX = (product.maxX - product.minX) * 0.08, padY = (product.maxY - product.minY) * 0.08
      let bx0 = product.minX, by0 = product.minY, bx1 = product.maxX, by1 = product.maxY
      for (const s of islands) {
        if (keep[s.label] || touchesBorder(s)) continue
        if (s.minX >= product.minX - padX && s.maxX <= product.maxX + padX &&
          s.minY >= product.minY - padY && s.maxY <= product.maxY + padY) {
          keep[s.label] = 1
          if (s.minX < bx0) bx0 = s.minX
          if (s.minY < by0) by0 = s.minY
          if (s.maxX > bx1) bx1 = s.maxX
          if (s.maxY > by1) by1 = s.maxY
        }
      }

      const outCanvas = document.createElement('canvas')
      outCanvas.width = w; outCanvas.height = h
      const outCtx = outCanvas.getContext('2d')!
      outCtx.fillStyle = '#fff'
      outCtx.fillRect(0, 0, w, h)
      // Copy the product's own bounding box across, whitening only pixels that belong to some
      // other island. Unlabelled pixels are copied as-is so light areas inside the product survive.
      const regionW = bx1 - bx0 + 1, regionH = by1 - by0 + 1
      const region = ctx.getImageData(bx0, by0, regionW, regionH)
      const rpx = region.data
      for (let y = 0; y < regionH; y++) {
        for (let x = 0; x < regionW; x++) {
          const lbl = labels[(by0 + y) * w + (bx0 + x)]
          if (lbl && !keep[lbl]) {
            const ri = (y * regionW + x) * 4
            rpx[ri] = 255; rpx[ri + 1] = 255; rpx[ri + 2] = 255
          }
        }
      }
      const prodTmp = document.createElement('canvas')
      prodTmp.width = regionW; prodTmp.height = regionH
      prodTmp.getContext('2d')!.putImageData(region, 0, 0)
      outCtx.drawImage(prodTmp, bx0, by0)

      const logoImg = new window.Image()
      logoImg.crossOrigin = 'anonymous'
      await new Promise<void>((resolve) => { logoImg.onload = () => resolve(); logoImg.onerror = () => resolve(); logoImg.src = '/LOGOBIGYA.jpg' })
      if (logoImg.complete && logoImg.naturalWidth > 0) {
        // LOGOBIGYA.jpg carries a wide white margin around the badge, so scaling the whole file
        // would land the badge much smaller than asked for — trim to the badge itself first.
        const lw = logoImg.naturalWidth, lh = logoImg.naturalHeight
        const lCanvas = document.createElement('canvas')
        lCanvas.width = lw; lCanvas.height = lh
        const lCtx = lCanvas.getContext('2d')!
        lCtx.fillStyle = '#fff'; lCtx.fillRect(0, 0, lw, lh)
        lCtx.drawImage(logoImg, 0, 0)
        const lpx = lCtx.getImageData(0, 0, lw, lh).data
        let lx0 = lw, ly0 = lh, lx1 = -1, ly1 = -1
        for (let y = 0; y < lh; y++) {
          for (let x = 0; x < lw; x++) {
            const i = (y * lw + x) * 4
            if (255 - lpx[i] > INK || 255 - lpx[i + 1] > INK || 255 - lpx[i + 2] > INK) {
              if (x < lx0) lx0 = x
              if (x > lx1) lx1 = x
              if (y < ly0) ly0 = y
              if (y > ly1) ly1 = y
            }
          }
        }
        if (lx1 > lx0 && ly1 > ly0) {
          const srcW = lx1 - lx0 + 1, srcH = ly1 - ly0 + 1
          const logoW = Math.round(w * 0.18)
          const logoH = Math.round(srcH * (logoW / srcW))
          const margin = Math.round(w * 0.06)
          outCtx.drawImage(lCanvas, lx0, ly0, srcW, srcH, w - logoW - margin, Math.round(h * 0.06), logoW, logoH)
        }
      }

      setEditImageSrc(outCanvas.toDataURL('image/png'))
      setEditImageOverlay('')
      setEditImageMsg(`ลบกรอบสำเร็จ ✅ (${w}x${h})`)
    } catch (e: any) { setEditImageMsg('❌ ' + (e?.message || 'ลบกรอบไม่สำเร็จ')) }
    setCleaningImage(false)
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
    const baseName = name.replace(/[^a-zA-Z0-9ก-๙ ._-]/g, '_')
    const filename = `${baseName}.${ext}`
    const url = canvas.toDataURL(outMime, 0.92)
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    a.click()
    setEditImageMsg('ดาวน์โหลดสำเร็จ ✅')
  }

  function uploadEditImageToDrive() {
    const canvas = editCanvasRef.current
    if (!canvas || !editImageSrc) {
      setEditImageMsg('กรุณาเลือกรูปภาพก่อน')
      return
    }
    const isJpeg = editImageMime.includes('jpeg') || editImageMime.includes('jpg')
    const outMime = isJpeg ? 'image/jpeg' : 'image/png'
    const ext = isJpeg ? 'jpg' : 'png'
    const name = (editProduct?.['*ชื่อสินค้า (NAME)'] || editProduct?.['รหัสสินค้า (SKU NUMBER)'] || `product-${Date.now()}`).toString().trim()
    const baseName = name.replace(/[^a-zA-Z0-9ก-๙ ._-]/g, '_')
    setUploadingDrive(true)
    setEditImageMsg('กำลังอัพโหลดไป Google Drive...')
    canvas.toBlob(async blob => {
      if (!blob) {
        setUploadingDrive(false)
        setEditImageMsg('❌ สร้างไฟล์จากรูปไม่สำเร็จ')
        return
      }
      try {
        const fd = new FormData()
        fd.append('file', blob, `${baseName}.${ext}`)
        fd.append('filename', baseName)
        const res = await fetch('/api/upload-drive', { method: 'POST', body: fd })
        const data = await res.json()
        if (data.success && data.link) {
          setEditProduct((prev: Product | null) => prev ? { ...prev, '*รูปภาพสินค้า': data.link } : prev)
          setEditImageMsg('อัพโหลด Drive สำเร็จ ✅ — กด 💾 บันทึก เพื่อเก็บลิงก์ลง DB')
        } else {
          setEditImageMsg('❌ ' + (data.error || 'อัพโหลดไม่สำเร็จ'))
        }
      } catch (e) {
        setEditImageMsg('❌ ' + String(e))
      }
      setUploadingDrive(false)
    }, outMime, 0.92)
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
    setEditImageOverlay(license)
  }, [editProduct])

  useEffect(() => {
    if (!editProduct) {
      setDriveLinkOwner(null)
      return
    }
    const link = editProduct['*รูปภาพสินค้า'] || ''
    const id = extractDriveFileId(link)
    if (!id) {
      setDriveLinkOwner(null)
      return
    }
    setDriveLinkOwner({ type: 'loading', email: '' })
    let cancelled = false
    const t = setTimeout(() => {
      fetch(`/api/upload-drive?fileId=${encodeURIComponent(id)}`)
        .then(r => r.json())
        .then(d => {
          if (!cancelled) setDriveLinkOwner(d.success ? { type: d.type, email: d.owner || '' } : { type: 'error', email: '' })
        })
        .catch(() => { if (!cancelled) setDriveLinkOwner({ type: 'error', email: '' }) })
    }, 400)
    return () => { cancelled = true; clearTimeout(t) }
  }, [editProduct?.['*รูปภาพสินค้า']])

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

  function extractDriveFileId(url: string): string | null {
    if (!url) return null
    const match = url.match(/\/file\/d\/([a-zA-Z0-9_-]{20,})/) ||
      url.match(/[?&]id=([a-zA-Z0-9_-]{20,})/)
    if (match) return match[1]
    return /^[a-zA-Z0-9_-]{25,}$/.test(url.trim()) ? url.trim() : null
  }

  function driveDownloadUrl(url: string): string | null {
    const id = extractDriveFileId(url)
    return id ? `https://drive.usercontent.google.com/download?id=${id}&export=download&confirm=t` : null
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

  const skus = Object.keys(grabMap)

  // Step 1: sync branch availability
  setStatus(`(1/3) กำลัง sync สาขา ${branch.toUpperCase()} (${skus.length} SKU)...`)
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

  // Step 2: ดึงข้อมูลจาก DB มาเทียบราคาก่อน (แสดงผลใน modal ได้เลย)
  setStatus(`(2/3) กำลังดึงข้อมูลจาก DB...`)
  const res = await fetch('/api/products', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ skus })
  })
  const data = await res.json()

  if (!data.success) {
    setStatus('เกิดข้อผิดพลาดดึงข้อมูล')
    setGrabResults([{ error: `ดึงข้อมูลไม่สำเร็จ: ${data.error ?? 'unknown'}` }])
    return
  }

  if (data.products.length === 0) {
    setStatus('ไม่มีข้อมูล')
    setGrabResults([{ error: 'อัพโหลดไฟล์ผิด หรือ คอลัมน์ไฟล์ไม่ถูกต้อง' }])
    return
  }

  // เปรียบเทียบราคา — แสดงผลทันที
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

  // แสดงผลใน modal ทันที
  setGrabResults(results)
  setStatus(`Grab ${branch.toUpperCase()}: ตรง ${matched_results.filter((r: any) => r.matched).length} | ต้องแก้ไข ${mismatch.length} | ไม่พบใน DB ${notFound_results.length} | sync +${syncData.inserted || 0} -${syncData.deleted || 0} | กำลังบันทึกราคา...`)

  // Step 3: บันทึกราคาสาขาลง DB (background — ไม่บล็อก UI)
  const PRICE_CHUNK = 200
  let priceError = ''
  for (let i = 0; i < skus.length; i += PRICE_CHUNK) {
    const chunkSkus = skus.slice(i, i + PRICE_CHUNK)
    const chunkMap: Record<string, number> = {}
    chunkSkus.forEach(s => { chunkMap[s] = grabMap[s] })
    try {
      const priceRes = await fetch('/api/products', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ updateBranchPrice: branch, priceMap: chunkMap })
      })
      const priceData = await priceRes.json()
      if (!priceData.success) { priceError = priceData.error ?? 'unknown'; break }
    } catch (e) {
      priceError = String(e)
      break
    }
  }

  await loadProducts(selectedSheet)
  await loadStats()

  if (priceError) {
    setStatus(`Grab ${branch.toUpperCase()}: ตรง ${matched_results.filter((r: any) => r.matched).length} | ต้องแก้ไข ${mismatch.length} | ⚠ บันทึกราคาไม่สำเร็จ: ${priceError}`)
  } else {
    setStatus(`Grab ${branch.toUpperCase()}: ตรง ${matched_results.filter((r: any) => r.matched).length} | ต้องแก้ไข ${mismatch.length} | ไม่พบใน DB ${notFound_results.length} | บันทึกราคาสำเร็จ ✅`)
  }
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

  async function handlePromaxxVsDbCheck(file: File) {
    setShowPromaxxModal(true)
    setPromaxxResults([])
    setPromaxxChecking(true)
    setStatus('กำลังอ่านไฟล์ R05.103...')

    const text = await file.text()
    const clean = text.replace(/^\uFEFF/, '')
    const rows = parseCSVRows(clean)

    // Col B (index 1) = SKU, Col E (index 4) = filter (เลข 5), Col G (index 6) = ราคา
    const promaxxMap: Record<string, number> = {}
    let skippedCount = 0
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i]
      if (!row || row.length < 7) continue
      const sku = row[1]?.trim()
      const filter = row[4]?.trim()
      const price = parsePriceRobust(row[6])
      if (!sku) continue
      if (filter !== '5') { skippedCount++; continue }
      if (price === null) { console.warn(`[Promaxx] SKU ${sku} row ${i+1}: filter=5 แต่ parse ราคา Col G ไม่ได้:`, row[6]); continue }
      promaxxMap[sku] = price
    }
    console.log(`[Promaxx] parsed ${Object.keys(promaxxMap).length} SKUs (skipped ${skippedCount} non-filter-5)`)

    if (Object.keys(promaxxMap).length === 0) {
      setPromaxxResults([{ error: 'ไม่พบข้อมูลในไฟล์ — ตรวจสอบว่าเป็นไฟล์ R05.103 และคอลัมน์ E มีค่าเป็น 5' }])
      setPromaxxChecking(false)
      setStatus('ไม่พบข้อมูล')
      return
    }

    const skus = Object.keys(promaxxMap)
    setStatus(`(1/2) กำลังดึงข้อมูลจาก DB... (${skus.length} SKU)`)
    setPromaxxProgress({ current: 0, total: skus.length, phase: 'fetch' })

    const chunks: string[][] = []
    for (let i = 0; i < skus.length; i += 300) chunks.push(skus.slice(i, i + 300))
    const allProducts: Product[] = []
    for (let chunkIdx = 0; chunkIdx < chunks.length; chunkIdx++) {
      const chunk = chunks[chunkIdx]
      setPromaxxProgress({ current: chunkIdx * 300, total: skus.length, phase: 'fetch' })
      const res = await fetch('/api/products', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ skus: chunk })
      })
      const data = await res.json()
      if (data.success) allProducts.push(...data.products)
    }
    console.log(`[Promaxx] fetched ${allProducts.length} products from DB`)

    const foundSkus = new Set(allProducts.map((p: Product) => p['รหัสสินค้า (SKU NUMBER)']))
    const matched = allProducts.map((p: Product) => {
      const sku = p['รหัสสินค้า (SKU NUMBER)']
      const promaxxPrice = promaxxMap[sku]
      const dbPrice = parsePriceRobust(String(p['*ราคาสินค้า']))
      if (promaxxPrice === undefined || dbPrice === null) return null
      const same = Math.abs(promaxxPrice - dbPrice) < 0.01
      return { sku, name: p['*ชื่อสินค้า (NAME)'], promaxxPrice, dbPrice, matched: same }
    }).filter(Boolean)

    const notFound = skus
      .filter(sku => !foundSkus.has(sku))
      .map(sku => ({ sku, promaxxPrice: promaxxMap[sku], notFound: true }))

    // แสดงผลใน modal ทันที
    setPromaxxResults([...matched, ...notFound])
    setPromaxxChecking(false)
    const mismatch = matched.filter((r: any) => !r.matched).length
    setStatus(`Promaxx: ตรง ${matched.filter((r: any) => r.matched).length} | ไม่ตรง ${mismatch} | ไม่พบใน DB ${notFound.length} | กำลังบันทึกราคา...`)

    // Step 2: บันทึก promaxx_price ลง DB (background, chunk 200)
    const CHUNK = 200
    let priceError = ''
    let savedCount = 0
    const totalChunks = Math.ceil(skus.length / CHUNK)
    setPromaxxProgress({ current: 0, total: skus.length, phase: 'save' })
    
    for (let i = 0; i < skus.length; i += CHUNK) {
      const chunkIdx = Math.floor(i / CHUNK)
      setPromaxxProgress({ current: i, total: skus.length, phase: 'save' })
      setStatus(`Promaxx: กำลังบันทึกราคา... (${chunkIdx + 1}/${totalChunks})`)
      
      const chunkSkus = skus.slice(i, i + CHUNK)
      const chunkMap: Record<string, number> = {}
      chunkSkus.forEach(s => { chunkMap[s] = promaxxMap[s] })
      try {
        const res = await fetch('/api/products', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ updatePromaxxPrice: true, priceMap: chunkMap })
        })
        const data = await res.json()
        if (!data.success) { priceError = data.error ?? 'unknown'; console.error(`[Promaxx] chunk ${i}-${i+CHUNK} failed:`, data.error); break }
        savedCount += data.updated ?? chunkSkus.length
        console.log(`[Promaxx] saved chunk ${i}-${Math.min(i+CHUNK, skus.length)} (${data.updated} SKUs)`)
      } catch (e) { priceError = String(e); console.error(`[Promaxx] chunk ${i} error:`, e); break }
    }

    setPromaxxProgress({ current: 0, total: 0, phase: '' })
    await loadProducts(selectedSheet)
    console.log(`[Promaxx] total saved: ${savedCount}/${skus.length}`)
    if (priceError) {
      setStatus(`Promaxx: ตรง ${matched.filter((r: any) => r.matched).length} | ไม่ตรง ${mismatch} | ⚠ บันทึกราคาไม่สำเร็จ: ${priceError}`)
    } else {
      setStatus(`Promaxx: ตรง ${matched.filter((r: any) => r.matched).length} | ไม่ตรง ${mismatch} | ไม่พบใน DB ${notFound.length} | บันทึกราคาสำเร็จ ${savedCount}/${skus.length} ✅`)
    }
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

  function hasProductBranch(product: Product, branch: BranchKey) {
    return !!product[branch] || String(product.branch ?? '').split(',').some(value => value.trim().toLowerCase() === branch)
  }

  function productPrice(value: unknown) {
    const price = parsePriceRobust(String(value ?? ''))
    return price !== null && Number.isFinite(price) ? price : null
  }

  function formatMoney(value: unknown) {
    const price = productPrice(value)
    return price === null ? '—' : new Intl.NumberFormat('th-TH', { style: 'currency', currency: 'THB', minimumFractionDigits: 2 }).format(price)
  }

  function pricesDiffer(left: unknown, right: unknown) {
    const first = productPrice(left)
    const second = productPrice(right)
    return first !== null && second !== null && Math.round(first * 100) !== Math.round(second * 100)
  }

  function priceState(product: Product): PriceState {
    if (!hasProductBranch(product, inspectionBranch)) return 'absent'
    const dbPrice = productPrice(product['*ราคาสินค้า'])
    const branchPrice = productPrice(product[`${inspectionBranch}_price`])
    if (dbPrice === null || branchPrice === null) return 'unpriced'
    return pricesDiffer(branchPrice, dbPrice) ? 'mismatch' : 'matched'
  }

  const priceLabels: Record<PriceState, string> = { mismatch: 'ต้องตรวจ', matched: 'ราคาตรง', unpriced: 'ยังไม่มีราคา', absent: 'ไม่มีในสาขา' }
  const branchNames: Record<BranchKey, string> = { src: 'SRC', kkl: 'KKL', sss: 'SSS' }
  const branches = Object.keys(branchNames) as BranchKey[]
  const normalizedSearch = search.trim().toLowerCase()
  const filteredProducts = products.filter(product =>
    (selectedSheet === 'all' || product['หมวดหมู่สินค้า (CATEGORIES)'] === selectedSheet) &&
    (branchFilter === 'all' || hasProductBranch(product, branchFilter)) &&
    (priceFilter === 'all' || priceState(product) === priceFilter) &&
    (!normalizedSearch || String(product['รหัสสินค้า (SKU NUMBER)'] ?? '').toLowerCase().includes(normalizedSearch) || String(product['*ชื่อสินค้า (NAME)'] ?? '').toLowerCase().includes(normalizedSearch))
  )
  const pageSize = 12
  const pageCount = Math.max(1, Math.ceil(filteredProducts.length / pageSize))
  const currentPage = Math.min(pageNumber, pageCount)
  const pageProducts = filteredProducts.slice((currentPage - 1) * pageSize, currentPage * pageSize)
  const inspectedProduct = filteredProducts.find(product => product['รหัสสินค้า (SKU NUMBER)'] === activeSku) ?? pageProducts[0] ?? null
  const inspectedSku = inspectedProduct?.['รหัสสินค้า (SKU NUMBER)'] as string | undefined
  const originalPrice = productPrice(inspectedProduct?.['*ราคาสินค้า'])
  const draftPrice = inspectedSku ? (Object.hasOwn(priceDrafts, inspectedSku) ? priceDrafts[inspectedSku] : (originalPrice === null ? '' : originalPrice.toFixed(2))) : ''
  const proposedPrice = productPrice(draftPrice)
  const hasPriceChange = draftPrice.trim() !== '' && proposedPrice !== null && proposedPrice !== originalPrice
  const comparisonPrice = inspectedProduct && hasProductBranch(inspectedProduct, inspectionBranch) ? productPrice(inspectedProduct[`${inspectionBranch}_price`]) : null
  const priceDifference = comparisonPrice !== null && originalPrice !== null ? comparisonPrice - originalPrice : null
  const issueCount = products.filter(product => priceState(product) === 'mismatch').length
  const activeReview = priceReview && priceReview.sku === inspectedSku && priceReview.price === proposedPrice ? priceReview : null

  function resetPriceReview() {
    setPriceReview(null)
    setInspectorError('')
  }

  function selectInspectorProduct(product: Product) {
    if (savingPrice) return
    setTooltip(null)
    setActiveSku(product['รหัสสินค้า (SKU NUMBER)'])
    resetPriceReview()
    if (window.matchMedia('(max-width: 820px)').matches) inspectorRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' })
  }

  function reviewInspectorPrice() {
    if (!inspectedSku) return
    if (proposedPrice === null || proposedPrice < 0 || Math.abs(proposedPrice * 100 - Math.round(proposedPrice * 100)) > 0.00001) {
      setInspectorError('กรุณาระบุราคาตั้งแต่ 0 บาท และทศนิยมไม่เกิน 2 ตำแหน่ง')
      return
    }
    setInspectorError('')
    setPriceReview({ sku: inspectedSku, price: proposedPrice, previous: originalPrice })
  }

  async function saveInspectorPrice() {
    if (!activeReview || savingPrice) return
    const review = activeReview
    setSavingPrice(true)
    setInspectorError('')
    try {
      const response = await fetch('/api/products', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sku: review.sku, '*ราคาสินค้า': review.price }),
      })
      if (!response.ok) throw new Error('บันทึกไม่สำเร็จ')
      const data = await response.json()
      if (!data.success) throw new Error('บันทึกไม่สำเร็จ')
      setProducts(previous => previous.map(product => product['รหัสสินค้า (SKU NUMBER)'] === review.sku ? { ...product, '*ราคาสินค้า': review.price } : product))
      setPriceDrafts(previous => { const next = { ...previous }; delete next[review.sku]; return next })
      setPriceReview(null)
      setLastPriceUpdate(new Intl.DateTimeFormat('th-TH', { dateStyle: 'short', timeStyle: 'short', timeZone: 'Asia/Bangkok' }).format(new Date()))
      setStatus(`บันทึกราคา ${review.sku} สำเร็จ · ${formatMoney(review.price)}`)
    } catch {
      setInspectorError('บันทึกราคาไม่สำเร็จ ข้อมูลที่กรอกยังอยู่ กรุณาลองอีกครั้ง')
      setStatus('บันทึกราคาไม่สำเร็จ')
    } finally {
      setSavingPrice(false)
    }
  }

  function openMissingBranches() {
    setShowMissingBranchModal(true)
    setMissingBranchResults({ src: [], kkl: [], sss: [] })
    setMissingBranchLoaded({ src: false, kkl: false, sss: false })
    setMissingBranchTab('src')
    setTimeout(() => loadMissingBranch('src'), 0)
  }

  function openAddProduct() {
    setIsAdding(true)
    setEditProduct({ 'ประเภทสินค้า': '', '*ชื่อสินค้า (NAME)': '', '*เลขที่ใบอนุญาตโฆษณา': '', '*ราคาสินค้า': '', 'รหัสสินค้า (SKU NUMBER)': '', '*รูปภาพสินค้า': '', 'หมวดหมู่สินค้า (CATEGORIES)': '' })
  }

  return (
    <div className="pc-app">
      <a className="pc-skip-link" href="#product-workspace">ข้ามไปยังรายการสินค้า</a>
      <aside className="pc-rail" aria-label="เมนูหลัก">
        <Link href="/" className="pc-brand" aria-label="Price Checker หน้าหลัก"><WorkspaceIcon name="tag" /></Link>
        <nav>
          <button className={!toolsOpen && priceFilter === 'all' ? 'is-active' : ''} title="รายการสินค้า" aria-label="รายการสินค้า" onClick={() => { setToolsOpen(false); setPriceFilter('all'); setPageNumber(1) }}><WorkspaceIcon name="box" /></button>
          <button className={priceFilter === 'mismatch' ? 'is-active' : ''} title="ตรวจสอบราคา" aria-label="ตรวจสอบราคา" onClick={() => { setPriceFilter('mismatch'); setPageNumber(1); resetPriceReview() }}><WorkspaceIcon name="tag" /></button>
          <button title="เปรียบเทียบ 3 สาขา" aria-label="เปรียบเทียบ 3 สาขา" onClick={handleMasterBranchCompare}><WorkspaceIcon name="branch" /></button>
          <button title="นำเข้าสินค้า" aria-label="นำเข้าสินค้า" onClick={() => document.getElementById('importInput')?.click()}><WorkspaceIcon name="upload" /></button>
          <button className={toolsOpen ? 'is-active' : ''} title="เครื่องมือทั้งหมด" aria-label="เครื่องมือทั้งหมด" aria-expanded={toolsOpen} aria-controls="workflow-tools" onClick={() => setToolsOpen(value => !value)}><WorkspaceIcon name="grid" /></button>
        </nav>
        <div className="pc-rail-bottom"><span className="pc-rail-divider" /><span className="pc-workspace-avatar" title="ANIN Workspace">AN</span></div>
      </aside>

      <main className="pc-workspace" id="product-workspace">
        <header className="pc-page-header">
          <div><div className="pc-eyebrow">PRICE CHECKER / PRODUCT OPERATIONS</div><h1>ตรวจและปรับราคาสินค้า<span className="pc-heading-dot">.</span></h1><p>ดูผลต่าง ตรวจค่าที่เสนอ แล้วยืนยันก่อนเปลี่ยนราคา</p></div>
          <div className="pc-header-actions">
            <span className="pc-workspace-label">ANIN <i /> WORKSPACE</span>
            <button className="pc-button pc-button-subtle" onClick={openAddProduct}><WorkspaceIcon name="plus" />เพิ่มสินค้า</button>
            <button className="pc-button" aria-expanded={toolsOpen} aria-controls="workflow-tools" onClick={() => setToolsOpen(value => !value)}><WorkspaceIcon name="settings" />เครื่องมือ</button>
          </div>
        </header>

        <div className="pc-context-strip"><span><i className="pc-context-dot" />รายการที่โหลด <b>{products.length.toLocaleString('th-TH')}</b></span><span className="pc-context-issues"><WorkspaceIcon name="clock" />ต้องตรวจ {inspectionBranch.toUpperCase()} <b>{issueCount.toLocaleString('th-TH')}</b></span><span className="pc-last-update">{lastPriceUpdate ? `บันทึกราคาล่าสุด ${lastPriceUpdate}` : 'จัดการข้อมูลสินค้าและราคาของทั้ง 3 สาขา'}</span></div>

        <div hidden>
          <input type="file" accept=".csv" id="grabInputSrc" onChange={e => { if (e.target.files?.[0]) { handleGrabCheck(e.target.files[0], 'src'); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="grabInputKkl" onChange={e => { if (e.target.files?.[0]) { handleGrabCheck(e.target.files[0], 'kkl'); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="grabInputSss" onChange={e => { if (e.target.files?.[0]) { handleGrabCheck(e.target.files[0], 'sss'); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="priceCalcInput" onChange={e => { if (e.target.files?.[0]) { handlePriceCalcUpload(e.target.files[0]); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="priceCalcGrabInput" onChange={e => { if (e.target.files?.[0]) { handlePriceCalcGrabUpload(e.target.files[0]); e.target.value = '' } }} />
          <input type="file" accept=".xlsx,.xls" id="importInput" onChange={e => { if (e.target.files?.[0]) { handleImportXlsx(e.target.files[0]); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="csvConvertInput" onChange={e => { if (e.target.files?.[0]) { handleCsvToUtf8(e.target.files[0]); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="promaxxCheckInput" onChange={e => { if (e.target.files?.[0]) { handlePromaxxVsDbCheck(e.target.files[0]); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="mmeCheckGrabInputSrc" onChange={e => { if (e.target.files?.[0]) { setMmeCheckGrabFileSrc(e.target.files[0]); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="mmeCheckGrabInputKkl" onChange={e => { if (e.target.files?.[0]) { setMmeCheckGrabFileKkl(e.target.files[0]); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="mmeCheckGrabInputSss" onChange={e => { if (e.target.files?.[0]) { setMmeCheckGrabFileSss(e.target.files[0]); e.target.value = '' } }} />
          <input type="file" accept=".xlsx,.xls" id="mmeCheckMmeInput" onChange={e => { if (e.target.files?.[0]) { setMmeCheckMmeFile(e.target.files[0]); e.target.value = '' } }} />
          <input type="file" accept=".xlsx,.xls,.csv" id="masterInput" onChange={e => { if (e.target.files?.[0]) { handleMasterUpload(e.target.files[0]); e.target.value = '' } }} />
        </div>

        {toolsOpen && <section className="pc-workflow-tools" id="workflow-tools" aria-label="เครื่องมือจัดการสินค้า">
          <div className="pc-tools-heading"><div><span className="pc-eyebrow">YOUR WORKFLOW</span><h2>เครื่องมือจัดการสินค้า</h2></div><button className="pc-icon-button" aria-label="ปิดเครื่องมือ" onClick={() => setToolsOpen(false)}><WorkspaceIcon name="close" /></button></div>
          <div className="pc-tool-groups">
            <section><h3><WorkspaceIcon name="upload" />นำเข้าและเตรียมไฟล์</h3><button onClick={openAddProduct}>เพิ่มสินค้ารายการเดียว <WorkspaceIcon name="plus" /></button><button onClick={() => document.getElementById('importInput')?.click()}>นำเข้า Menu Bulk <small>Excel</small></button><button onClick={() => document.getElementById('masterInput')?.click()}>นำเข้า Product Master <small>Excel / CSV</small></button><button onClick={() => document.getElementById('csvConvertInput')?.click()}>แปลง CSV เป็น UTF-8 <small>POS</small></button></section>
            <section><h3><WorkspaceIcon name="tag" />ตรวจสอบราคา</h3><button onClick={() => document.getElementById('priceCalcInput')?.click()}>คำนวณและอัปเดตราคา DB <small>R05.105</small></button><button onClick={() => document.getElementById('promaxxCheckInput')?.click()}>เทียบราคา Promaxx <small>R05.103</small></button><button onClick={() => { setShowMmeCheckModal(true); setMmeCheckResults({ src: [], kkl: [], sss: [] }) }}>ตรวจ GM MME กับ GRAB <WorkspaceIcon name="chevron" /></button></section>
            <section><h3><WorkspaceIcon name="branch" />GRAB และสาขา</h3><div className="pc-grab-imports">{branches.map(branch => <button key={branch} onClick={() => document.getElementById(`grabInput${branch[0].toUpperCase()}${branch.slice(1)}`)?.click()} title={`นำเข้า GRAB CSV ${branch.toUpperCase()}`}>GRAB {branch.toUpperCase()}</button>)}</div><button onClick={openMissingBranches}>ตรวจสินค้าที่ขาดใน GRAB <WorkspaceIcon name="chevron" /></button><button onClick={handleMasterBranchCompare}>เปรียบเทียบ Product Master 3 สาขา <WorkspaceIcon name="chevron" /></button></section>
            <section><h3><WorkspaceIcon name="download" />ส่งออกและข้อมูลกลาง</h3><button disabled={exportingMissing} onClick={exportMissingSkus}>{exportingMissing ? 'กำลังเตรียมไฟล์...' : 'ส่งออก SKU ที่ขาดจาก Master'}<WorkspaceIcon name="download" /></button><button disabled={selectedSheet === 'all'} onClick={exportSelectedCategoryXlsx}>ส่งออกหมวดหมู่ที่เลือก <WorkspaceIcon name="download" /></button><p className="pc-master-summary">Product Master <b>{masterSkuCount === null ? 'ยังไม่โหลด' : masterSkuCount.toLocaleString('th-TH')}</b>{masterFileName && <small>{masterFileName}</small>}</p><p className="pc-master-summary">SKU ที่ยังขาด <b>{masterMissingCount === null ? 'ยังไม่ตรวจ' : masterMissingCount.toLocaleString('th-TH')}</b></p></section>
          </div>
          <div className="pc-tools-bottom"><span>เลือกหมวดหมู่เพื่อส่งออกหรือจัดการสินค้าทั้งหมวด</span><button className="pc-danger-link" disabled={selectedSheet === 'all'} onClick={deleteSelectedCategory}>ลบหมวดหมู่ที่เลือกทุกสาขา</button></div>
        </section>}

        <div className="pc-inspection-grid">
          <section className="pc-list-column" aria-label="รายการสินค้า">
            <div className="pc-search-toolbar"><label className="pc-search"><WorkspaceIcon name="search" /><input ref={searchInputRef} type="search" aria-label="ค้นหาสินค้า" placeholder="ค้นหาชื่อสินค้า หรือรหัส SKU..." value={search} disabled={savingPrice} onChange={event => { setSearch(event.target.value); setPageNumber(1); resetPriceReview() }} /><kbd>⌘ K</kbd></label><button className={`pc-button ${filtersOpen ? 'is-selected' : ''}`} aria-expanded={filtersOpen} aria-controls="product-filters" onClick={() => setFiltersOpen(value => !value)}><WorkspaceIcon name="filter" />ตัวกรอง{(branchFilter !== 'all' || selectedSheet !== 'all') && <span className="pc-count">{Number(branchFilter !== 'all') + Number(selectedSheet !== 'all')}</span>}</button><button className="pc-icon-button pc-refresh" title="รีเฟรชสินค้า" aria-label="รีเฟรชสินค้า" disabled={loading || savingPrice} onClick={() => search.trim().length >= 3 ? searchProducts(search.trim()) : loadProducts(selectedSheet)}><WorkspaceIcon name="refresh" /></button></div>

            {filtersOpen && <div className="pc-filters" id="product-filters"><label>หมวดหมู่<select aria-label="กรองหมวดหมู่" value={selectedSheet} disabled={savingPrice} onChange={event => { setSelectedSheet(event.target.value); setPageNumber(1); resetPriceReview() }}><option value="all">ทุกหมวดหมู่ ({stats.sheets.length})</option>{stats.sheets.map(sheet => <option key={sheet.name} value={sheet.name}>{sheet.name} ({sheet.count})</option>)}</select></label><label>สินค้าที่มีในสาขา<select aria-label="กรองสาขา" value={branchFilter} disabled={savingPrice} onChange={event => { setBranchFilter(event.target.value as 'all' | BranchKey); setPageNumber(1); resetPriceReview() }}><option value="all">ทุกสาขา</option>{branches.map(branch => <option key={branch} value={branch}>{branchNames[branch]}</option>)}</select></label><button className="pc-text-button" onClick={() => { setSelectedSheet('all'); setBranchFilter('all'); setPriceFilter('all'); setSearch(''); setPageNumber(1); resetPriceReview() }}>ล้างตัวกรอง</button></div>}

            <div className="pc-table-card">
              <div className="pc-table-heading"><div><h2>{selectedSheet === 'all' ? 'สินค้าทั้งหมด' : selectedSheet}<span className="pc-count">{filteredProducts.length.toLocaleString('th-TH')}</span></h2><p>เทียบราคา GRAB <b>{branchNames[inspectionBranch]}</b> กับราคา DB · เลือกสินค้าเพื่อดูรายละเอียด</p></div><div className="pc-price-tabs" aria-label="กรองสถานะราคา"><button className={priceFilter === 'all' ? 'is-active' : ''} aria-pressed={priceFilter === 'all'} onClick={() => { setPriceFilter('all'); setPageNumber(1); resetPriceReview() }}>ทั้งหมด</button><button className={priceFilter === 'mismatch' ? 'is-active' : ''} aria-pressed={priceFilter === 'mismatch'} onClick={() => { setPriceFilter('mismatch'); setPageNumber(1); resetPriceReview() }}>ต้องตรวจ <b>{issueCount}</b></button></div></div>
              <div className="pc-table-options"><div className="pc-branch-switch" aria-label="เลือกสาขาเปรียบเทียบ">{branches.map(branch => <button key={branch} className={inspectionBranch === branch ? 'is-active' : ''} aria-pressed={inspectionBranch === branch} disabled={savingPrice} onClick={() => { setInspectionBranch(branch); setPageNumber(1); resetPriceReview() }}>{branchNames[branch]}</button>)}</div><button className="pc-text-button" aria-pressed={showAllPrices} onClick={() => setShowAllPrices(value => !value)}><WorkspaceIcon name="grid" />{showAllPrices ? 'มุมมองกระชับ' : 'ดูราคาทุกระบบ'}</button></div>

              {loadError && <div className="pc-load-error" role="alert"><WorkspaceIcon name="clock" /><div><strong>โหลดข้อมูลสินค้าไม่สำเร็จ</strong><p>{loadError}</p></div><button className="pc-button" onClick={() => search.trim().length >= 3 ? searchProducts(search.trim()) : loadProducts(selectedSheet)}>ลองอีกครั้ง</button></div>}
              {loading ? <div className="pc-loading" role="status" aria-label="กำลังโหลดสินค้า"><span className="pc-loading-label"><i />กำลังโหลดข้อมูลสินค้า...</span>{Array.from({ length: 6 }, (_, index) => <div className="pc-skeleton-row" key={index}><i /><span /><span /><b /></div>)}</div> : !loadError && filteredProducts.length === 0 ? <div className="pc-empty"><WorkspaceIcon name="search" /><h3>{search || priceFilter !== 'all' || selectedSheet !== 'all' || branchFilter !== 'all' ? 'ไม่พบสินค้าที่ตรงกับเงื่อนไข' : 'ยังไม่มีรายการสินค้า'}</h3><p>{search || priceFilter !== 'all' || selectedSheet !== 'all' || branchFilter !== 'all' ? 'ลองเปลี่ยนคำค้นหา สาขา หรือสถานะราคา' : 'เริ่มต้นด้วยการนำเข้าไฟล์ Excel หรือเพิ่มสินค้าใหม่'}</p>{search || priceFilter !== 'all' || selectedSheet !== 'all' || branchFilter !== 'all' ? <button className="pc-button" onClick={() => { setSearch(''); setPriceFilter('all'); setSelectedSheet('all'); setBranchFilter('all'); setPageNumber(1) }}>แสดงสินค้าทั้งหมด</button> : <button className="pc-button pc-button-primary" onClick={() => document.getElementById('importInput')?.click()}><WorkspaceIcon name="upload" />นำเข้าสินค้า</button>}</div> : !loadError && <div className="pc-table-scroll"><table className={`pc-products-table ${showAllPrices ? 'pc-expanded-prices' : ''}`}><thead><tr><th className="pc-checkbox-cell"><input type="checkbox" aria-label="เลือกสินค้าทั้งหมดในหน้านี้" checked={pageProducts.length > 0 && pageProducts.every(product => selectedSkus.has(product['รหัสสินค้า (SKU NUMBER)']))} ref={input => { if (input) input.indeterminate = pageProducts.some(product => selectedSkus.has(product['รหัสสินค้า (SKU NUMBER)'])) && !pageProducts.every(product => selectedSkus.has(product['รหัสสินค้า (SKU NUMBER)'])) }} onChange={event => setSelectedSkus(previous => { const next = new Set(previous); pageProducts.forEach(product => event.target.checked ? next.add(product['รหัสสินค้า (SKU NUMBER)']) : next.delete(product['รหัสสินค้า (SKU NUMBER)'])); return next })} /></th><th className="pc-product-column">สินค้า / SKU</th><th>ราคา DB</th>{showAllPrices ? <><th>Promaxx</th>{branches.map(branch => <th key={branch}>GRAB {branchNames[branch]}</th>)}</> : <th>GRAB {branchNames[inspectionBranch]}</th>}<th>สถานะ</th><th><span className="pc-sr-only">รายละเอียด</span></th></tr></thead><tbody>{pageProducts.map(product => {
                const sku = String(product['รหัสสินค้า (SKU NUMBER)'])
                const state = priceState(product)
                const image = convertDriveLink(product['*รูปภาพสินค้า'] || '')
                return <tr key={sku} data-sku={sku} className={`${sku === inspectedSku ? 'pc-row-active' : ''} ${selectedSkus.has(sku) ? 'pc-row-checked' : ''}`} onClick={event => { if (!(event.target as HTMLElement).closest('input,button,a')) selectInspectorProduct(product) }} onMouseEnter={event => { if (image && window.matchMedia('(hover: hover) and (pointer: fine) and (min-width: 821px)').matches) setTooltip({ x: event.clientX, y: event.clientY, url: image }) }} onMouseMove={event => setTooltip(previous => previous ? { ...previous, x: event.clientX, y: event.clientY } : null)} onMouseLeave={() => setTooltip(null)}>
                  <td className="pc-checkbox-cell"><input type="checkbox" aria-label={`เลือก ${product['*ชื่อสินค้า (NAME)'] || sku}`} checked={selectedSkus.has(sku)} onChange={event => setSelectedSkus(previous => { const next = new Set(previous); event.target.checked ? next.add(sku) : next.delete(sku); return next })} /></td>
                  <td><button className="pc-product-link" aria-label={`ตรวจสินค้า ${product['*ชื่อสินค้า (NAME)'] || sku}`} disabled={savingPrice} onClick={() => selectInspectorProduct(product)}><span className="pc-product-thumb"><WorkspaceIcon name="box" />{image && <NextImage src={image} alt="" fill sizes="33px" unoptimized loading="lazy" onError={event => { event.currentTarget.hidden = true }} />}</span><span><strong>{product['*ชื่อสินค้า (NAME)'] || 'ยังไม่ระบุชื่อสินค้า'}</strong><small>{sku}</small><span className="pc-row-branches">{branches.filter(branch => hasProductBranch(product, branch)).map(branch => <i key={branch}>{branchNames[branch]}</i>)}</span></span></button></td>
                  <td className="pc-table-price">{formatMoney(product['*ราคาสินค้า'])}</td>
                  {showAllPrices ? <><td className={pricesDiffer(product.promaxx_price, product['*ราคาสินค้า']) ? 'pc-price-mismatch' : ''}>{formatMoney(product.promaxx_price)}</td>{branches.map(branch => <td key={branch} className={hasProductBranch(product, branch) && pricesDiffer(product[`${branch}_price`], product['*ราคาสินค้า']) ? 'pc-price-mismatch' : ''}>{hasProductBranch(product, branch) ? formatMoney(product[`${branch}_price`]) : '—'}</td>)}</> : <td className={`pc-table-price ${state === 'mismatch' ? 'pc-price-mismatch' : ''}`}>{hasProductBranch(product, inspectionBranch) ? formatMoney(product[`${inspectionBranch}_price`]) : '—'}</td>}
                  <td><span className={`pc-status pc-status-${state}`}><WorkspaceIcon name={state === 'matched' ? 'check' : state === 'mismatch' ? 'clock' : 'tag'} />{priceLabels[state]}</span></td><td><button className="pc-row-open" aria-label={`ดูรายละเอียด ${product['*ชื่อสินค้า (NAME)'] || sku}`} disabled={savingPrice} onClick={() => selectInspectorProduct(product)}><WorkspaceIcon name="chevron" /></button></td>
                </tr>
              })}</tbody></table></div>}

              <div className="pc-table-footer"><span>{selectedSkus.size ? <button className="pc-selected-export" onClick={exportSelectedSkusXlsx}><WorkspaceIcon name="download" />ส่งออกที่เลือก ({selectedSkus.size.toLocaleString('th-TH')})</button> : `${filteredProducts.length ? (currentPage - 1) * pageSize + 1 : 0}–${Math.min(currentPage * pageSize, filteredProducts.length)} จาก ${filteredProducts.length.toLocaleString('th-TH')} รายการที่โหลด`}</span><div className="pc-pagination"><button aria-label="หน้าก่อนหน้า" disabled={currentPage === 1 || loading} onClick={() => setPageNumber(currentPage - 1)}>‹</button><span><b>{currentPage}</b> / {pageCount}</span><button aria-label="หน้าถัดไป" disabled={currentPage === pageCount || loading} onClick={() => setPageNumber(currentPage + 1)}>›</button></div></div>
            </div>
            <div className="pc-list-footnote"><WorkspaceIcon name="tag" /><span>ราคาไม่ตรงจะแสดงทั้งสีและสถานะ เพื่อช่วยตรวจรายการได้ชัดเจน</span><button className="pc-text-button" onClick={openMissingBranches}>ดูสินค้าที่ขาดใน GRAB <WorkspaceIcon name="arrow" /></button></div>
          </section>

          <aside className="pc-inspector" ref={inspectorRef} aria-label="รายละเอียดและปรับราคาสินค้า" aria-busy={loading}>
            <div className="pc-inspector-heading"><span className="pc-eyebrow">PRODUCT DETAILS</span><WorkspaceIcon name="tag" /></div>
            {inspectedProduct && !loadError ? <>
              <div className="pc-product-image"><span className="pc-image-placeholder"><WorkspaceIcon name="box" /><span>รูปภาพสินค้า</span></span>{convertDriveLink(inspectedProduct['*รูปภาพสินค้า'] || '') && <NextImage key={inspectedSku} src={convertDriveLink(inspectedProduct['*รูปภาพสินค้า'])} alt={inspectedProduct['*ชื่อสินค้า (NAME)'] || 'รูปสินค้า'} fill sizes="(max-width: 820px) 80vw, 350px" unoptimized onError={event => { event.currentTarget.hidden = true }} />}</div>
              <div className="pc-inspector-product"><span className="pc-product-category">{inspectedProduct['หมวดหมู่สินค้า (CATEGORIES)'] || 'ยังไม่ระบุหมวดหมู่'}</span><h2>{inspectedProduct['*ชื่อสินค้า (NAME)'] || 'ยังไม่ระบุชื่อสินค้า'}</h2><p className="pc-inspector-sku">{inspectedSku}</p><span className={`pc-status pc-status-${priceState(inspectedProduct)}`}><WorkspaceIcon name={priceState(inspectedProduct) === 'matched' ? 'check' : 'clock'} />{priceLabels[priceState(inspectedProduct)]}</span></div>
              <div className="pc-comparison-prices"><div><small>ราคาใน DB</small><strong>{formatMoney(inspectedProduct['*ราคาสินค้า'])}</strong></div><div><small>ราคา GRAB {branchNames[inspectionBranch]}</small><strong>{formatMoney(comparisonPrice)}</strong></div></div>
              <div className={`pc-price-difference ${priceDifference !== null && !pricesDiffer(comparisonPrice, originalPrice) ? 'is-matched' : ''}`}><WorkspaceIcon name="tag" /><span>{priceDifference === null ? 'ยังเปรียบเทียบราคาไม่ได้' : 'ผลต่างจากราคา DB'}</span><strong>{priceDifference === null ? '—' : `${priceDifference > 0 ? '+' : priceDifference < 0 ? '−' : ''} ${formatMoney(Math.abs(priceDifference))}`}</strong></div>
              <div className="pc-inspector-field"><label htmlFor="proposed-price">ราคาที่เสนอ <span>THB</span></label><input id="proposed-price" type="text" inputMode="decimal" autoComplete="off" aria-describedby="price-help inspector-error" value={draftPrice} disabled={savingPrice || loading} onChange={event => { if (!inspectedSku) return; setPriceDrafts(previous => ({ ...previous, [inspectedSku]: event.target.value })); resetPriceReview() }} /><div id="price-help" className="pc-field-help"><span>{hasPriceChange ? 'มีราคาที่ยังไม่ได้บันทึก' : 'ปรับราคา DB แล้วตรวจทานก่อนบันทึก'}</span>{hasPriceChange && <button disabled={savingPrice} onClick={() => { setPriceDrafts(previous => { const next = { ...previous }; if (inspectedSku) delete next[inspectedSku]; return next }); resetPriceReview() }}>คืนค่าเดิม</button>}</div></div>
              <div className="pc-inspector-branches"><span>สาขาที่จำหน่าย</span><div>{branches.map(branch => <span className={hasProductBranch(inspectedProduct, branch) ? 'is-available' : ''} key={branch}><WorkspaceIcon name={hasProductBranch(inspectedProduct, branch) ? 'check' : 'close'} />{branchNames[branch]}</span>)}</div></div>
              <div className="pc-reference-price"><span>ราคา Promaxx</span><strong>{productPrice(inspectedProduct.promaxx_price) === null ? 'ยังไม่มีราคา' : formatMoney(inspectedProduct.promaxx_price)}</strong></div>
              <div className="pc-reference-price"><span>เลขที่ใบอนุญาต</span><strong>{inspectedProduct['*เลขที่ใบอนุญาตโฆษณา'] || 'ยังไม่ระบุ'}</strong></div>
              {inspectedProduct.correctPrice !== undefined && <div className="pc-reference-price"><span>ราคาอ้างอิง</span><strong>{formatMoney(inspectedProduct.correctPrice)}</strong></div>}
              {inspectorError && <p id="inspector-error" className="pc-inspector-error" role="alert">{inspectorError}</p>}
              {activeReview ? <section className="pc-review" aria-label="ตรวจทานราคาก่อนบันทึก"><h3><WorkspaceIcon name="check" />ตรวจทานก่อนบันทึก</h3><p>ปรับราคา DB ของสินค้านี้</p><div><span>{formatMoney(activeReview.previous)}</span><WorkspaceIcon name="arrow" /><strong>{formatMoney(activeReview.price)}</strong></div><button className="pc-button pc-button-primary" disabled={savingPrice} onClick={saveInspectorPrice}>{savingPrice ? 'กำลังบันทึก...' : 'ยืนยันบันทึกราคา'}<WorkspaceIcon name="check" /></button><button className="pc-text-button" disabled={savingPrice} onClick={resetPriceReview}>กลับไปแก้ไข</button></section> : <button className="pc-button pc-button-primary pc-review-button" disabled={!hasPriceChange || savingPrice || loading} onClick={reviewInspectorPrice}>ตรวจทานการปรับราคา<WorkspaceIcon name="arrow" /></button>}
              <button className="pc-edit-details" disabled={savingPrice} onClick={() => { setIsAdding(false); setEditProduct({ ...inspectedProduct, _originalSku: inspectedSku }) }}>แก้ไขรายละเอียดและรูปสินค้า<WorkspaceIcon name="chevron" /></button>
            </> : <div className="pc-inspector-empty"><WorkspaceIcon name="box" /><h2>เลือกสินค้าที่ต้องการตรวจ</h2><p>เลือกสินค้าจากตารางเพื่อดูราคา สาขา และปรับข้อมูลในพื้นที่นี้</p><div className="pc-empty-guide"><span><b>01</b>เลือกสินค้าจากรายการ</span><span><b>02</b>ตรวจราคาและผลต่าง</span><span><b>03</b>ตรวจทานก่อนบันทึก</span></div></div>}
          </aside>
        </div>
        <footer className="pc-statusbar"><span role="status" aria-live="polite"><i className={loadError || inspectorError ? 'has-error' : ''} />{status}</span><span>PRICE CHECKER <b>/</b> SPLIT INSPECTOR</span></footer>
      </main>

      {showGrabModal && (() => {
        const closeGrabModal = () => {
          setShowGrabModal(false)
          setEditingGrabSku(null)
          setEditingGrabValue('')
          setGrabSearch('')
          setGrabShowOnlyMismatch(true)
          setGrabEdits({})
        }
        const search = grabSearch.trim().toLowerCase()
        const displayGrabResults = grabResults
          .map((r: any) => {
            const editedPrice = grabEdits[r.sku]
            const effectiveGrabPrice = editedPrice !== undefined ? editedPrice : r.grabPrice
            const effectiveMatched = !r.notFound && r.dbPrice !== null && r.dbPrice !== undefined && Math.abs(effectiveGrabPrice - r.dbPrice) < 0.01
            const isEdited = editedPrice !== undefined && editedPrice !== r.grabPrice
            return { ...r, effectiveGrabPrice, effectiveMatched, isEdited }
          })
          .filter((r: any) => {
            if (search && !r.sku.toLowerCase().includes(search) && !(r.name || '').toLowerCase().includes(search)) return false
            if (grabShowOnlyMismatch && !r.notFound && r.effectiveMatched) return false
            return true
          })
        const totalCount = grabResults.length
        const matchedCount = displayGrabResults.filter((r: any) => r.effectiveMatched).length
        const mismatchCount = displayGrabResults.filter((r: any) => !r.effectiveMatched && !r.notFound).length
        const notFoundCount = displayGrabResults.filter((r: any) => r.notFound).length

        return (
  <div style={{
    position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)',
    display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000
  }}>
    <div style={{ background: '#fff', borderRadius: 6, width: 700, maxHeight: '80vh', display: 'flex', flexDirection: 'column', boxShadow: '0 10px 40px rgba(0,0,0,0.3)' }}>

      {/* Header */}
      <div style={{ background: '#4a4a4a', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <strong>🛵 ผลตรวจสอบราคา GRAB</strong>
        <button onClick={closeGrabModal}
          style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer' }}>×</button>
      </div>

      {/* Summary */}
      {grabResults.length > 0 && !grabResults[0]?.error && (
        <div style={{ padding: '10px 20px', background: '#f8f8f8', borderBottom: '1px solid #ddd', display: 'flex', gap: 20, fontSize: 13, flexWrap: 'wrap' }}>
          <span>แสดง: <strong>{displayGrabResults.length}</strong> / {totalCount}</span>
          <span style={{ color: '#28a745' }}>✓ ตรง: <strong>{matchedCount}</strong></span>
          <span style={{ color: '#dc3545' }}>✗ ต้องแก้ไขใน GRAB: <strong>{mismatchCount}</strong></span>
          {notFoundCount > 0 && (
            <span style={{ color: '#000' }}>⚠ ไม่พบข้อมูลใน Supabase: <strong>{notFoundCount}</strong></span>
          )}
        </div>
      )}

      {/* Toolbar */}
      {grabResults.length > 0 && !grabResults[0]?.error && (
        <div style={{ padding: '8px 20px', background: '#fff', borderBottom: '1px solid #eee', display: 'flex', gap: 12, alignItems: 'center' }}>
          <input
            type="text"
            placeholder="ค้นหา SKU / ชื่อสินค้า..."
            value={grabSearch}
            onChange={(e) => setGrabSearch(e.target.value)}
            style={{ padding: '6px 10px', border: '1px solid #ccc', borderRadius: 4, fontSize: 12, width: 220 }}
          />
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, cursor: 'pointer' }}>
            <input
              type="checkbox"
              checked={grabShowOnlyMismatch}
              onChange={(e) => setGrabShowOnlyMismatch(e.target.checked)}
            />
            เฉพาะราคาไม่ตรง
          </label>
          <span style={{ fontSize: 11, color: '#666', marginLeft: 'auto' }}>
            💡 คลิกที่ราคาใน GRAB เพื่อแก้ไข (ไม่บันทึกลง DB)
          </span>
        </div>
      )}

      {/* Table */}
      <div style={{ overflowY: 'auto', flex: 1 }}>
        {grabResults[0]?.error ? (
          <div style={{ padding: 40, textAlign: 'center', color: '#dc3545' }}>{grabResults[0].error}</div>
        ) : grabResults.length === 0 ? (
          <div style={{ padding: 40, textAlign: 'center', color: '#000' }}>กำลังโหลด...</div>
        ) : displayGrabResults.length === 0 ? (
          <div style={{ padding: 40, textAlign: 'center', color: '#000' }}>ไม่พบรายการที่ตรงกับเงื่อนไข</div>
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
              {displayGrabResults.map((r: any, i: number) => {
                const isEditing = editingGrabSku === r.sku
                const commitEdit = () => {
                  const parsed = parsePriceRobust(editingGrabValue)
                  if (parsed !== null) {
                    setGrabEdits(prev => ({ ...prev, [r.sku]: parsed }))
                  }
                  setEditingGrabSku(null)
                  setEditingGrabValue('')
                }
                const cancelEdit = () => {
                  setEditingGrabSku(null)
                  setEditingGrabValue('')
                }
                return (
                  <tr key={r.sku || i} style={{ background: r.notFound ? '#f0f0f0' : r.effectiveMatched ? 'transparent' : '#fff3cd' }}>
                    <td style={td}>{i + 1}</td>
                    <td style={{ ...td, color: '#000' }}><strong>{r.sku}</strong></td>
                    <td style={{ ...td, color: '#000' }}>{r.name || '-'}</td>
                    <td style={{ ...td, fontWeight: 600, color: '#000' }}>
                      {r.notFound ? (
                        <span>-</span>
                      ) : isEditing ? (
                        <input
                          type="number"
                          autoFocus
                          value={editingGrabValue}
                          onChange={(e) => setEditingGrabValue(e.target.value)}
                          onBlur={commitEdit}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') commitEdit()
                            else if (e.key === 'Escape') cancelEdit()
                          }}
                          style={{ width: 110, padding: '3px 6px', border: '1px solid #007bff', borderRadius: 3, fontSize: 12 }}
                        />
                      ) : (
                        <span
                          onClick={() => { setEditingGrabSku(r.sku); setEditingGrabValue(String(r.effectiveGrabPrice)) }}
                          title="คลิกเพื่อแก้ไข"
                          style={{ cursor: 'pointer', padding: '2px 6px', borderRadius: 3, color: r.isEdited ? '#007bff' : '#000' }}
                          onMouseEnter={(e) => (e.currentTarget.style.background = '#e7f1ff')}
                          onMouseLeave={(e) => (e.currentTarget.style.background = 'transparent')}
                        >
                          {r.effectiveGrabPrice?.toLocaleString() ?? '-'}
                          {r.isEdited && <span style={{ marginLeft: 4, fontSize: 10 }}>✎</span>}
                        </span>
                      )}
                    </td>
                    <td style={{ ...td, fontWeight: 600, color: '#000' }}>{r.dbPrice?.toLocaleString() ?? '-'}</td>
                    <td style={td}>
                      {r.notFound
                        ? <span style={{ background: '#e0e0e0', color: '#000', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>⚠ ไม่พบข้อมูลใน Supabase</span>
                        : r.effectiveMatched
                          ? <span style={{ background: '#d4edda', color: '#155724', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✓ ตรง</span>
                          : <span style={{ background: '#f8d7da', color: '#721c24', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✗ แก้ไขใน GRAB</span>
                      }
                    </td>
                  </tr>
                )
              })}
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
        <button onClick={closeGrabModal} style={btnStyle}>ปิด</button>
      </div>
    </div>
  </div>
        )
      })()}

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
          <div className="edit-modal-anim" role="dialog" aria-modal="true" aria-label={isAdding ? 'เพิ่มรายการสินค้า' : 'แก้ไขรายละเอียดสินค้า'} style={{
            background: '#fff', borderRadius: 6, width: 560, maxHeight: '92vh', overflowY: 'auto',
            boxShadow: '0 10px 40px rgba(0,0,0,0.3)'
          }}>
            {/* Modal Header */}
            <div style={{ background: isAdding ? '#2d6a2d' : '#4a4a4a', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <strong>{isAdding ? '➕ เพิ่มรายการสินค้า' : '✏️ แก้ไขสินค้า'}</strong>
              <button aria-label="ปิดการแก้ไขสินค้า" onClick={() => { setEditProduct(null); setIsAdding(false) }}
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
                    ใส่เลขที่ใบอนุญาต (เลขอย่างเดียว) — ระบบจะพิมพ์คำว่า <b style={{ color: '#355070' }}>"เลขที่ใบอนุญาต"</b> นำหน้าให้อัตโนมัติ เช่น เลขที่ใบอนุญาต 65-2-3-2-0003417
                  </div>
                  <div style={{ marginTop: 8, display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: '#355070' }}>ข้อความบนรูป:</div>
                  <input
                    style={{ ...inputStyle, flex: '1 1 160px', marginTop: 0 }}
                    placeholder="ใส่เลขที่ใบอนุญาต (เลขอย่างเดียว)"
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
                  <button
                    type="button"
                    onClick={cleanGreenBorderAndText}
                    disabled={cleaningImage || (!editImageSrc && !editProduct?.['*รูปภาพสินค้า'])}
                    style={{ ...btnStyle, background: '#e65100', color: '#fff', borderColor: '#bf360c', opacity: cleaningImage ? 0.6 : 1 }}
                  >
                    {cleaningImage ? '⏳ กำลังลบกรอบ...' : '🧹 ลบกรอบ+ข้อความ'}
                  </button>
                  <button
                    type="button"
                    onClick={uploadEditImageToDrive}
                    disabled={!editImageSrc || uploadingDrive}
                    style={{ ...btnStyle, background: '#188038', color: '#fff', borderColor: '#0d652d', opacity: uploadingDrive ? 0.6 : 1 }}
                  >
                    {uploadingDrive ? '⏳ กำลังอัพโหลด...' : '☁️ อัพโหลด Drive'}
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
                <label style={{ ...labelStyle, marginTop: 10, display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                  รูปภาพสินค้า (Google Drive Link)
                  {driveLinkOwner?.type === 'anin' && (
                    <span title={driveLinkOwner.email} style={{ background: '#d4edda', color: '#155724', padding: '2px 8px', borderRadius: 10, fontSize: 10, fontWeight: 600 }}>🏢 Anin (องค์กร)</span>
                  )}
                  {driveLinkOwner?.type === 'external' && (
                    <span title={driveLinkOwner.email} style={{ background: '#fff3cd', color: '#856404', padding: '2px 8px', borderRadius: 10, fontSize: 10, fontWeight: 600 }}>🌐 External (เมล์นอก)</span>
                  )}
                </label>
                <input style={inputStyle}
                  placeholder="ลิงก์จะถูกใส่อัตโนมัติหลังอัพโหลด หรือวาง Google Drive URL ที่นี่..."
                  value={editProduct['*รูปภาพสินค้า'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, '*รูปภาพสินค้า': e.target.value })}
                />
                {driveDownloadUrl(editProduct['*รูปภาพสินค้า'] || '') && (
                  <div style={{ marginTop: 6 }}>
                    <a
                      href={driveDownloadUrl(editProduct['*รูปภาพสินค้า'])!}
                      target="_blank"
                      rel="noreferrer"
                      style={{ ...btnStyle, background: '#8ab4f8', color: '#174ea6', borderColor: '#aecbfa', textDecoration: 'none', display: 'inline-block', padding: '4px 12px', fontSize: 12, borderRadius: 4 }}
                    >
                      ⬇️ Download รูปจาก Drive
                    </a>
                  </div>
                )}
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
                  value={editProduct['*ราคาสินค้า'] ?? ''}
                  onChange={e => setEditProduct({ ...editProduct, '*ราคาสินค้า': e.target.value })}
                />
              </div>
            </div>

            {/* Modal Footer */}
            <div style={{ padding: '0 20px 20px', display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              <button onClick={() => { setEditProduct(null); setIsAdding(false) }} style={btnStyle}>ยกเลิก</button>
              <button onClick={() => { setSaveBtnAnim(true); setTimeout(() => setSaveBtnAnim(false), 200); saveProduct(); }}
                className={saveBtnAnim ? 'btn-press-anim' : ''}
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
            <div style={{ background: '#fff', borderRadius: 6, width: 960, maxHeight: '88vh', display: 'flex', flexDirection: 'column', boxShadow: '0 10px 40px rgba(0,0,0,0.3)' }}>

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
                        <th style={{ ...th, textAlign: 'center', minWidth: 70 }}>SRC ราคา</th>
                        <th style={{ ...th, textAlign: 'center', minWidth: 70 }}>KKL ราคา</th>
                        <th style={{ ...th, textAlign: 'center', minWidth: 70 }}>SSS ราคา</th>
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
                          {(['src','kkl','sss'] as const).map(b => {
                            const val = p[`${b}_price`]
                            const dbPrice = parsePriceRobust(String(p['*ราคาสินค้า']))
                            const mismatch = val !== null && val !== undefined && dbPrice !== null && Math.abs(Number(val) - dbPrice) >= 0.01
                            return (
                              <td key={`${b}-price`} style={{ ...td, textAlign: 'center', background: mismatch ? '#fff3cd' : 'transparent' }}>
                                {val !== null && val !== undefined ? (
                                  <span style={{ fontWeight: 600, color: mismatch ? '#856404' : '#155724', fontSize: 11 }}>
                                    {Number(val).toLocaleString()}
                                    {mismatch && <span style={{ display: 'block', fontSize: 9, color: '#856404' }}>≠ DB</span>}
                                  </span>
                                ) : (
                                  <span style={{ color: '#ccc', fontSize: 11 }}>-</span>
                                )}
                              </td>
                            )
                          })}
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


      {/* Promaxx vs DB Modal */}
      {showPromaxxModal && (() => {
        const mismatchList = promaxxResults.filter((r: any) => !r.notFound && !r.matched)
        const matchedList = promaxxResults.filter((r: any) => !r.notFound && r.matched)
        const notFoundList = promaxxResults.filter((r: any) => r.notFound)
        return (
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
            <div style={{ background: '#fff', borderRadius: 6, width: 740, maxHeight: '85vh', display: 'flex', flexDirection: 'column', boxShadow: '0 10px 40px rgba(0,0,0,0.3)' }}>
              <div style={{ background: '#6d28d9', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderRadius: '6px 6px 0 0' }}>
                <strong>🔍 ผลตรวจสอบ Promaxx (R05.103) vs DB</strong>
                <button onClick={() => setShowPromaxxModal(false)} style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer' }}>×</button>
              </div>
              {/* Progress bar */}
              {promaxxProgress.total > 0 && (
                <div style={{ padding: '10px 20px', background: '#f0e6ff', borderBottom: '1px solid #d1b3ff' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11, color: '#5b21b6', marginBottom: 4 }}>
                    <span>{promaxxProgress.phase === 'fetch' ? '🔄 กำลังดึงข้อมูล...' : '💾 กำลังบันทึกราคา...'}</span>
                    <span>{promaxxProgress.current.toLocaleString()} / {promaxxProgress.total.toLocaleString()}</span>
                  </div>
                  <div style={{ background: '#e9d5ff', borderRadius: 999, height: 6, overflow: 'hidden' }}>
                    <div style={{ background: '#7c3aed', height: '100%', width: `${Math.min(100, (promaxxProgress.current / promaxxProgress.total) * 100)}%`, transition: 'width 0.3s' }} />
                  </div>
                </div>
              )}
              {promaxxResults.length > 0 && !promaxxResults[0]?.error && (
                <div style={{ padding: '10px 20px', background: '#f8f8f8', borderBottom: '1px solid #ddd', display: 'flex', gap: 20, fontSize: 13 }}>
                  <span>ทั้งหมด: <strong>{promaxxResults.length}</strong></span>
                  <span style={{ color: '#28a745' }}>✓ ตรง: <strong>{matchedList.length}</strong></span>
                  <span style={{ color: '#dc3545' }}>✗ ราคาไม่ตรง: <strong>{mismatchList.length}</strong></span>
                  {notFoundList.length > 0 && <span style={{ color: '#856404' }}>⚠ ไม่พบใน DB: <strong>{notFoundList.length}</strong></span>}
                </div>
              )}
              <div style={{ overflowY: 'auto', flex: 1 }}>
                {promaxxChecking ? (
                  <div style={{ padding: 40, textAlign: 'center', color: '#888' }}>⏳ กำลังโหลด...</div>
                ) : promaxxResults[0]?.error ? (
                  <div style={{ padding: 40, textAlign: 'center', color: '#dc3545' }}>{promaxxResults[0].error}</div>
                ) : promaxxResults.length === 0 ? (
                  <div style={{ padding: 40, textAlign: 'center', color: '#888' }}>ไม่มีข้อมูล</div>
                ) : (
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: '#f5f5f5', position: 'sticky', top: 0 }}>
                        <th style={th}>#</th>
                        <th style={th}>SKU</th>
                        <th style={th}>ชื่อสินค้า</th>
                        <th style={th}>ราคาใน Promaxx</th>
                        <th style={th}>ราคาใน DB</th>
                        <th style={th}>สถานะ</th>
                      </tr>
                    </thead>
                    <tbody>
                      {promaxxResults.map((r: any, i: number) => (
                        <tr key={i} style={{ background: r.notFound ? '#f5f5f5' : r.matched ? 'transparent' : '#fff3cd' }}>
                          <td style={td}>{i + 1}</td>
                          <td style={td}><strong>{r.sku}</strong></td>
                          <td style={td}>{r.name || '-'}</td>
                          <td style={td}>{r.notFound ? '-' : r.promaxxPrice?.toLocaleString()}</td>
                          <td style={td}>{r.notFound ? '-' : r.dbPrice?.toLocaleString()}</td>
                          <td style={td}>
                            {r.notFound
                              ? <span style={{ background: '#e0e0e0', color: '#555', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>⚠ ไม่พบใน DB</span>
                              : r.matched
                                ? <span style={{ background: '#d4edda', color: '#155724', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✓ ตรง</span>
                                : <span style={{ background: '#fff3cd', color: '#856404', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✗ ราคาไม่ตรง</span>
                            }
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
              <div style={{ padding: '12px 20px', borderTop: '1px solid #ddd', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <button
                  disabled={mismatchList.length === 0}
                  onClick={async () => {
                    if (mismatchList.length === 0) return
                    const XLSX = await import('xlsx')
                    const rows = mismatchList.map((r: any) => ({
                      'รหัสสินค้า (SKU NUMBER)': r.sku,
                      'ชื่อสินค้า': r.name || '',
                      'ราคาใน Promaxx': r.promaxxPrice,
                      'ราคาใน DB': r.dbPrice,
                    }))
                    const ws = XLSX.utils.json_to_sheet(rows)
                    const wb = XLSX.utils.book_new()
                    XLSX.utils.book_append_sheet(wb, ws, 'ราคาไม่ตรง')
                    const now = new Date()
                    const dd = String(now.getDate()).padStart(2, '0')
                    const mm = String(now.getMonth() + 1).padStart(2, '0')
                    const yyyy = now.getFullYear()
                    XLSX.writeFile(wb, `Promaxx vs DB ราคาไม่ตรง ${dd}.${mm}.${yyyy}.xlsx`)
                  }}
                  style={{ ...btnStyle, background: '#f3e8ff', borderColor: '#8b5cf6', color: '#5b21b6', opacity: mismatchList.length === 0 ? 0.5 : 1 }}
                >
                  📥 Export ราคาไม่ตรง ({mismatchList.length})
                </button>
                <button onClick={() => setShowPromaxxModal(false)} style={btnStyle}>ปิด</button>
              </div>
            </div>
          </div>
        )
      })()}

    </div>
  )
}

const btnStyle: React.CSSProperties = {
  padding: '7px 16px', border: '1px solid #777', borderRadius: 20,
  fontSize: 12, fontWeight: 600, cursor: 'pointer', background: '#f5f5f5'
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

