'use client'

import { useState, useEffect, useEffectEvent, useRef } from 'react'
import Link from 'next/link'
import NextImage from 'next/image'
import { supabase } from '../lib/supabase'
import { ENGINE_BRANCHES, createResultsZip, decodeCsv, engineDate, normalizeStockSku, parseEngineCsv, pendingSkuSets, processGrabTemplate, templateName, validateGrabTemplateUpload, type EngineBranch, type StockReport } from '../lib/process-engine'
import { loadPendingSnapshot, savePendingMutation } from '../lib/process-engine-source'
import { loadAgentLastSeen, loadMssqlStock, requestLiveStock, waitForStockRequest } from '../lib/stock-source'
import { agentOnline, requestStatusText } from '../lib/stock-live'
import { saveGrabTemplate } from '../lib/grab-template-source'
import { GRAB_CHANGE_LABELS, grabChangesCsv, grabCsvWithBom, type SavedGrabReview } from '../lib/grab-template-review'
import { buildPendingPriceChecks, comparePendingPrices, mutatePendingEntries, parsePendingInput, pendingBranchCoverage, pendingManualBranches, pendingPricePair, readTemplateNames, samePendingStockMembership, type PendingMutation, type PendingPriceCheck, type PendingSkuEntry, type PendingSnapshot } from '../lib/pending-sku'

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
const PRODUCT_PAGE_SIZE = 12
type PriceState = 'mismatch' | 'matched' | 'unpriced' | 'absent'
type EngineOutput = { name: string; content: string; branch: EngineBranch; total: number; updated: number; matched: number }
type GrabFileInfo = { name: string; updated_at: string }
type EngineEntryState = { id: number; status: 'idle' | 'loading' | 'error'; pending: boolean; prices: boolean; templates: boolean; error: string }
type PendingPromaxxRow = { sku: string; product: Product | null; grab: number | null; promaxx: number | null; ready: boolean; reason: string }
type PendingPromaxxReview = { branch: EngineBranch; rows: PendingPromaxxRow[]; loading: boolean; error: string }
type WorkspaceTool = { name: string; detail: string; action: () => void; disabled?: boolean; busy?: boolean }

function WorkspaceIcon({ name, className = '' }: { name: string; className?: string }) {
  const paths: Record<string, React.ReactNode> = {
    'nav-products': <><path d="m3 7 9-4 9 4-9 4-9-4Zm0 0v10l9 4 9-4V7M12 11v10" /><path d="M8 7c2-3 6-3 8 0" /></>,
    'nav-price-check': <><path d="m9 19-7-7V3h9l6 6" /><circle cx="6.5" cy="7" r="1" /><circle cx="15" cy="15" r="4.5" /><path d="m18.3 18.3 3.7 3.7m-9-7 1.4 1.4 2.4-2.8" /></>,
    'nav-branches': <><path d="M12 10v2M5 14v-2h14v2" />{['translate(8 1)', 'translate(1 14)', 'translate(15 14)'].map(transform => <g key={transform} transform={transform}><path d="m0 3 1-3h6l1 3H0Zm1 0v6h6V3M3 9V6h2v3M3 3v1m2-1v1" /></g>)}</>,
    'nav-import': <><path d="M12 2v12m-4-4 4 4 4-4M7 9H5a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-8a2 2 0 0 0-2-2h-2" /></>,
    'nav-engine': <><path d="M9.5 2h5l.5 2.5 2 1.2 2.5-.8L22 9.1l-2 1.7v2.4l2 1.7-2.5 4.2-2.5-.8-2 1.2-.5 2.5h-5L9 19.5l-2-1.2-2.5.8L2 14.9l2-1.7v-2.4L2 9.1l2.5-4.2 2.5.8L9 4.5Z" /><path d="m13 7-4 6h3l-1 4 4-6h-3l1-4Z" /></>,
    'nav-tools': <><g transform="translate(24 0) scale(-1 1)"><path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94L14.7 6.3Z" /></g><path d="m14 8 5-6 3 3-6 5-2-2Zm0 2-2 2M9 14l-7 6 2 2 7-6" /></>,
    grid: <><rect x="3" y="3" width="7" height="7" rx="1.5" /><rect x="14" y="3" width="7" height="7" rx="1.5" /><rect x="3" y="14" width="7" height="7" rx="1.5" /><rect x="14" y="14" width="7" height="7" rx="1.5" /></>,
    box: <><path d="m3 7 9-4 9 4v10l-9 4-9-4Z" /><path d="m3 7 9 4 9-4M12 11v10M7 5l10 5" /></>,
    tag: <><path d="M3 3h8l10 10-8 8L3 11Z" /><circle cx="7.5" cy="7.5" r="1" /></>,
    search: <><circle cx="10.5" cy="10.5" r="6.5" /><path d="m16 16 5 5" /></>,
    branch: <><path d="M6 5v10a4 4 0 0 0 4 4h6M6 8h8a4 4 0 0 0 4-4" /><circle cx="6" cy="4" r="2" /><circle cx="18" cy="3" r="2" /><circle cx="18" cy="19" r="2" /></>,
    upload: <><path d="M12 16V3m-5 5 5-5 5 5M4 16v5h16v-5" /></>,
    download: <><path d="M12 3v13m-5-5 5 5 5-5M4 17v4h16v-4" /></>,
    check: <path d="m5 12 4 4 10-10" />,
    alert: <><path d="m12 3 10 18H2L12 3Z" /><path d="M12 9v5m0 3v.1" /></>,
    file: <><path d="M14 2H4v20h16V8l-6-6Zm0 0v6h6M8 13h8M8 17h6" /></>,
    clock: <><circle cx="12" cy="12" r="9" /><path d="M12 7v6l4 2" /></>,
    arrow: <path d="M4 12h15m-6-6 6 6-6 6" />,
    chevron: <path d="m9 5 7 7-7 7" />,
    filter: <path d="M3 5h18l-7 8v6l-4 2v-8Z" />,
    settings: <><path d="M4 7h16M4 17h16" /><circle cx="9" cy="7" r="3" /><circle cx="15" cy="17" r="3" /></>,
    refresh: <><path d="M20 7v5h-5M4 17v-5h5" /><path d="M6 6a8 8 0 0 1 13 3M5 15a8 8 0 0 0 13 3" /></>,
    database: <><ellipse cx="12" cy="5" rx="8" ry="3" /><path d="M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3" /></>,
    plus: <path d="M12 4v16M4 12h16" />,
    close: <path d="m6 6 12 12M18 6 6 18" />,
    engine: <><path d="m13 2-9 12h7l-1 8 10-12h-7Z" /></>,
  }
  return <svg className={`pc-icon ${className}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths[name] ?? paths.box}</svg>
}

function StudioModal({ label, width = 900, onClose, children }: { label: string; width?: number; onClose: () => void; children: React.ReactNode }) {
  const dialogRef = useRef<HTMLDialogElement>(null)
  useEffect(() => {
    const dialog = dialogRef.current
    if (!dialog) return
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const previousOverflow = document.body.style.overflow
    dialog.showModal()
    document.body.style.overflow = 'hidden'
    return () => {
      dialog.close()
      document.body.style.overflow = previousOverflow
      requestAnimationFrame(() => { if (opener?.isConnected) opener.focus() })
    }
  }, [])
  useEffect(() => {
    // Disabling a busy action can move browser focus back to the body.
    // Keep keyboard navigation inside the active popup while its data loads.
    const dialog = dialogRef.current
    if (dialog?.open && document.activeElement === document.body) dialog.focus()
  })
  return <dialog ref={dialogRef} className="pc-modal" tabIndex={-1} aria-label={label} style={{ width: `min(${width}px, calc(100vw - 32px))` }} onCancel={event => { event.preventDefault(); onClose() }}>{children}</dialog>
}

type PopupNotice = { message: string; confirm: boolean; resolve: (answer: boolean) => void }

export default function Home() {
  const [popupNotice, setPopupNotice] = useState<PopupNotice | null>(null)
  const popupQueue = useRef<PopupNotice[]>([])
  function notifyPopup(message: string, confirm = false) {
    return new Promise<boolean>(resolve => {
      const notice = { message, confirm, resolve }
      popupQueue.current.push(notice)
      if (popupQueue.current.length === 1) setPopupNotice(notice)
    })
  }
  function dismissPopup(answer: boolean) {
    const notice = popupQueue.current.shift()
    setPopupNotice(popupQueue.current[0] ?? null)
    notice?.resolve(answer)
  }
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
  const [savingNewPromaxx, setSavingNewPromaxx] = useState(false)
  const [newPromaxxSaveError, setNewPromaxxSaveError] = useState('')
  const priceCalcBusyRef = useRef(false)
  const [updatingPrices, setUpdatingPrices] = useState(false)
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
  const productRequestAbort = useRef<AbortController | null>(null)
  const searchInputRef = useRef<HTMLInputElement>(null)
  const inspectorRef = useRef<HTMLElement>(null)
  const [activeSku, setActiveSku] = useState<string | null>(null)
  const [inspectionBranch, setInspectionBranch] = useState<BranchKey>('src')
  const [branchFilter, setBranchFilter] = useState<'all' | BranchKey>('all')
  const [priceFilter, setPriceFilter] = useState<'all' | PriceState>('all')
  const [catalogOpen, setCatalogOpen] = useState(true)
  const [toolsOpen, setToolsOpen] = useState(false)
  const [toolSearch, setToolSearch] = useState('')
  const [filtersOpen, setFiltersOpen] = useState(false)
  const [pageNumber, setPageNumber] = useState(1)
  const [loadError, setLoadError] = useState('')
  const [priceDrafts, setPriceDrafts] = useState<Record<string, string>>({})
  const [priceReview, setPriceReview] = useState<{ sku: string; price: number; previous: number | null } | null>(null)
  const [savingPrice, setSavingPrice] = useState(false)
  const [inspectorError, setInspectorError] = useState('')
  const [engineOpen, setEngineOpen] = useState(false)
  const [engineEntry, setEngineEntry] = useState<EngineEntryState>({ id: 0, status: 'idle', pending: false, prices: false, templates: false, error: '' })
  const engineEntryRequest = useRef(0)
  const engineNavRef = useRef<HTMLButtonElement>(null)
  const [engineReports, setEngineReports] = useState<Partial<Record<EngineBranch, StockReport>>>({})
  const [engineSubtract, setEngineSubtract] = useState<Record<EngineBranch, string>>({ src: '5', sss: '2', kkl: '2' })
  const [engineBusy, setEngineBusy] = useState('')
  const engineBusyRef = useRef(false)
  const [engineLog, setEngineLog] = useState<string[]>([])
  const [engineProgress, setEngineProgress] = useState({ current: 0, total: 0, label: 'รอเริ่มงาน' })
  const [engineChecks, setEngineChecks] = useState<Partial<Record<EngineBranch, boolean>> | null>(null)
  const [engineSummary, setEngineSummary] = useState<{ processed: number; updated: number; skipped: number } | null>(null)
  const [engineOutputs, setEngineOutputs] = useState<EngineOutput[]>([])
  const [engineWarnings, setEngineWarnings] = useState<string[]>([])
  const [enginePendingCounts, setEnginePendingCounts] = useState<Record<EngineBranch, number> | null>(null)
  const [grabFiles, setGrabFiles] = useState<GrabFileInfo[]>([])
  const [grabFilesLoading, setGrabFilesLoading] = useState(false)
  const grabFilesRequest = useRef(0)
  const [grabFilesError, setGrabFilesError] = useState('')
  const [grabUploadProblems, setGrabUploadProblems] = useState<string[]>([])
  const [grabUploadNotice, setGrabUploadNotice] = useState('')
  const [grabUploadResults, setGrabUploadResults] = useState<SavedGrabReview[]>([])
  const [grabChangeOpen, setGrabChangeOpen] = useState(false)
  const [grabChangeBranch, setGrabChangeBranch] = useState<EngineBranch>('src')
  const [grabExpandedRow, setGrabExpandedRow] = useState<string | null>(null)
  const [grabChangeLimit, setGrabChangeLimit] = useState(50)
  const [grabExportError, setGrabExportError] = useState('')
  const grabChangeDialogRef = useRef<HTMLDialogElement>(null)
  const grabChangeOpenerRef = useRef<HTMLElement | null>(null)
  const grabChangedFiles = grabUploadResults.filter(review => review.changed)
  const grabActiveReview = grabChangedFiles.find(review => review.branch === grabChangeBranch) ?? grabChangedFiles[0]
  const grabChangedRows = grabActiveReview?.rows.filter(row => row.type !== 'unchanged') ?? []
  const [grabManagerOpen, setGrabManagerOpen] = useState(false)
  const [grabAdminUser, setGrabAdminUser] = useState<string | null>(null)
  const [grabAdminEmail, setGrabAdminEmail] = useState('')
  const [grabAdminPassword, setGrabAdminPassword] = useState('')
  const [grabAdminError, setGrabAdminError] = useState('')
  const engineLogRef = useRef<HTMLDivElement>(null)
  const [pendingSnapshot, setPendingSnapshot] = useState<PendingSnapshot | null>(null)
  const pendingSnapshotRef = useRef<PendingSnapshot | null>(null)
  const pendingReadyRef = useRef(false)
  const pendingLoadRequest = useRef(0)
  const [pendingLoading, setPendingLoading] = useState(false)
  const [pendingError, setPendingError] = useState('')
  const [pendingSaveError, setPendingSaveError] = useState('')
  const [pendingActionNotice, setPendingActionNotice] = useState('')
  const [pendingExportNotice, setPendingExportNotice] = useState<{ error: boolean; message: string } | null>(null)
  const [pendingPromaxxReview, setPendingPromaxxReview] = useState<PendingPromaxxReview | null>(null)
  const pendingPromaxxRequest = useRef(0)
  const [pendingProduct, setPendingProduct] = useState<{ sku: string; name: string } | null>(null)
  const [pendingInput, setPendingInput] = useState('')
  const [pendingSelectedBranches, setPendingSelectedBranches] = useState<Record<EngineBranch, boolean>>({ src: true, sss: true, kkl: true })
  const [pendingTab, setPendingTab] = useState<EngineBranch>('src')
  const [pendingAuthOpen, setPendingAuthOpen] = useState(false)
  const [pendingDbPrices, setPendingDbPrices] = useState({ prices: new Map<string, number>(), loaded: false, loading: false, error: '' })
  const pendingDbRequest = useRef(0)
  const pendingNameCache = useRef(new Map<string, { updatedAt: string; names: Map<string, string> }>())
  const [pendingPriceRevision, setPendingPriceRevision] = useState(0)
  const pendingAutoRequest = useRef('')
  const [pendingAutoPrices, setPendingAutoPrices] = useState(new Map<string, PendingPriceCheck>())
  const [pendingAutoStatus, setPendingAutoStatus] = useState({ loaded: false, saved: false, checkedAt: '', error: '' })

  useEffect(() => {
    let mounted = true
    void supabase.auth.getSession().then(({ data }) => {
      if (mounted) setGrabAdminUser(data.session ? data.session.user.email ?? 'บัญชีที่ล็อกอิน' : null)
    })
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      if (mounted) setGrabAdminUser(session ? session.user.email ?? 'บัญชีที่ล็อกอิน' : null)
    })
    return () => { mounted = false; subscription.unsubscribe() }
  }, [])

  useEffect(() => {
    if (engineLogRef.current) engineLogRef.current.scrollTop = engineLogRef.current.scrollHeight
  }, [engineLog, engineOpen])

  useEffect(() => {
    const dialog = grabChangeDialogRef.current
    if (!dialog) return
    if (!grabChangeOpen) { if (dialog.open) dialog.close(); return }
    if (!dialog.open) dialog.showModal()
    const previousOverflow = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    return () => { document.body.style.overflow = previousOverflow }
  }, [grabChangeOpen])

  useEffect(() => {
    if (!engineOpen && !toolsOpen) return
    const entryId = engineOpen ? engineEntry.id : null
    void refreshGrabFiles().then(files => {
      if (entryId !== null) markEngineEntry(entryId, files !== null ? { templates: true } : { status: 'error', error: 'โหลดรายการ Template ไม่สำเร็จ' })
    })
  }, [engineOpen, toolsOpen, engineEntry.id])

  useEffect(() => {
    if ((catalogOpen && !engineOpen && !toolsOpen) || engineBusy || engineBusyRef.current || !pendingPriceRevision) return
    const request = `${pendingPriceRevision}:${grabAdminUser ?? 'anonymous'}`
    if (pendingAutoRequest.current === request) return
    pendingAutoRequest.current = request
    const entryId = engineOpen ? engineEntry.id : null
    void refreshPendingCard(entryId === null ? undefined : () => markEngineEntry(entryId, { pending: true })).then(ready => {
      if (entryId !== null) markEngineEntry(entryId, ready ? { pending: true, prices: true } : { status: 'error', error: 'โหลดรายการ PENDING หรือราคาสินค้าไม่สำเร็จ' })
    })
  }, [pendingPriceRevision, grabAdminUser, engineBusy, catalogOpen, engineOpen, toolsOpen, engineEntry.id])

  const onSearchCriteriaChange = useEffectEvent(() => {
    if (search.trim().length >= 3) searchProducts(search.trim())
    else loadProducts(selectedSheet)
  })

  useEffect(() => {
    loadStats()
  }, [])

  useEffect(() => {
    if (toolsOpen) void loadMasterCount()
  }, [toolsOpen])

  useEffect(() => {
    // Invalidate an older request immediately, including during the debounce window.
    searchReqId.current += 1
    productRequestAbort.current?.abort()
    const timeout = setTimeout(() => {
      onSearchCriteriaChange()
    }, search.trim().length >= 3 ? 250 : 0)
    return () => { clearTimeout(timeout); searchReqId.current += 1; productRequestAbort.current?.abort() }
  }, [search, selectedSheet])

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key === 'k') {
        event.preventDefault()
        setEngineOpen(false); setToolsOpen(false)
        requestAnimationFrame(() => searchInputRef.current?.focus())
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [])

  function engineMessage(error: unknown) {
    return error instanceof Error ? error.message : 'เกิดข้อผิดพลาด กรุณาลองอีกครั้ง'
  }

  function logEngine(message: string) {
    const time = new Intl.DateTimeFormat('en-GB', { timeZone: 'Asia/Bangkok', hour: '2-digit', minute: '2-digit', second: '2-digit' }).format(new Date())
    setEngineLog(previous => [...previous, `[${time}] ${message}`])
  }

  function invalidateEngineResults() {
    setEngineOutputs([])
    setEngineSummary(null)
    setEngineWarnings([])
  }

  function acceptPendingSnapshot(snapshot: PendingSnapshot) {
    if (pendingSnapshotRef.current && !samePendingStockMembership(pendingSnapshotRef.current.entries, snapshot.entries)) invalidateEngineResults()
    pendingSnapshotRef.current = snapshot
    pendingReadyRef.current = true
    setPendingSnapshot(snapshot)
    setPendingLoading(false)
    setPendingError('')
    const sets = pendingSkuSets(snapshot.entries)
    setEnginePendingCounts({ src: sets.src.size, sss: sets.sss.size, kkl: sets.kkl.size })
  }

  function beginEngineEntry() {
    const id = ++engineEntryRequest.current
    setEngineEntry({ id, status: 'loading', pending: false, prices: false, templates: false, error: '' })
    setPendingPriceRevision(value => value + 1)
    setEngineOpen(true)
    setToolsOpen(false)
  }

  function openProcessEngine() {
    if (!engineOpen) beginEngineEntry()
  }

  function markEngineEntry(id: number, update: Partial<Omit<EngineEntryState, 'id'>>) {
    setEngineEntry(previous => {
      if (previous.id !== id || previous.status !== 'loading') return previous
      const next = { ...previous, ...update }
      return next.pending && next.prices && next.templates ? { ...next, status: 'idle' } : next
    })
  }

  async function refreshPendingCard(onPrepared?: () => void) {
    if (engineBusyRef.current) return false
    engineBusyRef.current = true; setEngineBusy('pending-auto')
    const request = ++pendingLoadRequest.current
    pendingReadyRef.current = false
    setPendingLoading(true); setPendingError(''); setPendingSaveError(''); setPendingActionNotice('')
    try {
      const snapshot = await resolvePendingPrices()
      if (request !== pendingLoadRequest.current) return false
      onPrepared?.()
      return await refreshPendingDatabase(snapshot.entries)
    } catch (error) {
      if (request !== pendingLoadRequest.current) return false
      invalidateEngineResults()
      setPendingError(engineMessage(error)); setPendingLoading(false)
      setPendingAutoStatus(previous => ({ ...previous, error: engineMessage(error) }))
      logEngine(`PENDING LOAD ERROR · ${engineMessage(error)}`)
      return false
    } finally { engineBusyRef.current = false; setEngineBusy('') }
  }

  async function resolvePendingPrices() {
    const [response, stored, session] = await Promise.all([
      fetch('/api/products?pendingPriceCheck=true', { cache: 'no-store' }),
      loadPendingSnapshot(),
      supabase.auth.getSession(),
    ])
    if (!response.ok) throw new Error('ตรวจราคา GRAB / Promaxx ไม่สำเร็จ กรุณาโหลดใหม่')
    const data = await response.json()
    if (!data.success || !Array.isArray(data.products)) throw new Error('อ่านราคา GRAB / Promaxx ไม่สำเร็จ')
    if (session.error) throw new Error('ตรวจสอบ session ไม่สำเร็จ กรุณาเข้าสู่ระบบใหม่')
    const checks = buildPendingPriceChecks(data.products, parsePriceRobust)
    const mutation: PendingMutation = { kind: 'syncPrices', checks }
    const preview = mutatePendingEntries(stored.entries, mutation)
    const snapshot = session.data.session
      ? await savePendingMutation(mutation)
      : { ...stored, entries: preview.entries }
    acceptPendingSnapshot(snapshot)
    setPendingAutoPrices(checks)
    setPendingAutoStatus({ loaded: true, saved: Boolean(session.data.session) || !preview.changed, checkedAt: new Intl.DateTimeFormat('th-TH', { timeZone: 'Asia/Bangkok', hour: '2-digit', minute: '2-digit' }).format(new Date()), error: '' })
    if (preview.changed) logEngine(`PENDING AUTO · GRAB ≠ Promaxx · ${session.data.session ? 'บันทึกแล้ว' : 'ใช้ผลตรวจในหน่วยความจำจนกว่าจะเข้าสู่ระบบ'}`)
    return snapshot
  }

  async function refreshPendingDatabase(entries = pendingSnapshotRef.current?.entries ?? []) {
    const request = ++pendingDbRequest.current
    setPendingDbPrices(previous => ({ ...previous, loading: true, error: '' }))
    try {
      const skus = entries.filter(entry => entry.branches.length).map(entry => entry.sku), chunks: string[][] = []
      for (let index = 0; index < skus.length; index += 300) chunks.push(skus.slice(index, index + 300))
      const batches = await Promise.all(chunks.map(async chunk => {
        const response = await fetch('/api/products', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ skus: chunk }) })
        if (!response.ok) throw new Error('โหลดราคา DATABASE ไม่สำเร็จ')
        const data = await response.json()
        if (!data.success || !Array.isArray(data.products)) throw new Error('โหลดราคา DATABASE ไม่สำเร็จ')
        return data.products as Product[]
      }))
      const prices = new Map<string, number>()
      for (const product of batches.flat()) {
        const sku = normalizeStockSku(product['รหัสสินค้า (SKU NUMBER)'])
        const price = parsePriceRobust(String(product['*ราคาสินค้า'] ?? ''))
        if (sku && price !== null && Number.isFinite(price)) prices.set(sku, price)
      }
      if (request !== pendingDbRequest.current) return false
      setPendingDbPrices({ prices, loaded: true, loading: false, error: '' })
      return true
    } catch (error) {
      if (request === pendingDbRequest.current) setPendingDbPrices(previous => ({ ...previous, loading: false, error: `${engineMessage(error)}${previous.loaded ? ' · แสดงราคาที่โหลดสำเร็จครั้งก่อน' : ''}` }))
      return false
    }
  }

  async function findPendingNames(entries: PendingSkuEntry[], preferred: EngineBranch[]) {
    const remaining = new Set(entries.filter(entry => entry.branches.length && !entry.name.trim()).map(entry => entry.sku))
    const names = new Map<string, string>()
    if (!remaining.size) return names
    const files = await refreshGrabFiles()
    if (!files) throw new Error('เพิ่มรายการแล้ว แต่ยังค้นหาชื่อไม่ได้เพราะอ่านรายการ Template ไม่สำเร็จ')
    const branches = [...new Set([...preferred, ...ENGINE_BRANCHES])]
    for (const branch of branches) {
      if (!remaining.size) break
      const file = files.find(item => item.name === templateName(branch))
      if (!file) continue
      try {
        let cached = pendingNameCache.current.get(file.name)
        if (!cached || cached.updatedAt !== file.updated_at) {
          const { data, error } = await supabase.from('grab_files').select('content').eq('name', file.name).maybeSingle()
          if (error || !data || typeof data.content !== 'string') throw new Error(error?.message ?? 'ไม่พบ Template')
          cached = { updatedAt: file.updated_at, names: readTemplateNames(file.name, data.content) }
          pendingNameCache.current.set(file.name, cached)
        }
        for (const sku of remaining) {
          const name = cached.names.get(sku)
          if (name) { names.set(sku, name); remaining.delete(sku) }
        }
      } catch (error) { logEngine(`PENDING NAME WARNING · ${file.name}: ${engineMessage(error)}`) }
    }
    if (remaining.size) logEngine(`PENDING · ไม่พบชื่อใน Template ${remaining.size} SKU`)
    return names
  }

  async function editPending(mutation: PendingMutation) {
    if (engineBusyRef.current || !grabAdminUser || !pendingReadyRef.current) return
    engineBusyRef.current = true; setEngineBusy('pending')
    ++pendingLoadRequest.current; ++pendingDbRequest.current
    setPendingDbPrices(previous => ({ ...previous, loading: false }))
    setPendingLoading(false); setPendingSaveError(''); setPendingActionNotice(''); setPendingExportNotice(null)
    if (mutation.kind !== 'enrichNames') invalidateEngineResults()
    let saved = false
    try {
      const { data } = await supabase.auth.getSession()
      if (!data.session) throw new Error('session หมดอายุ กรุณาเข้าสู่ระบบก่อนบันทึก PENDING')
      const snapshot = await savePendingMutation(mutation)
      acceptPendingSnapshot(snapshot); saved = true
      logEngine(`PENDING SAVED · ${mutation.kind} · ${snapshot.entries.filter(entry => entry.branches.length).length} SKUs ที่ยังรออนุมัติ`)
      if (mutation.kind === 'resolveBranch') setPendingActionNotice(`นำ ${mutation.sku} ออกจาก PENDING เฉพาะ ${mutation.branch.toUpperCase()} แล้ว · จะเพิ่มกลับเมื่อคู่ราคาเปลี่ยนและยังไม่ตรงกัน`)
      if (mutation.kind === 'add') setPendingInput('')
      if (mutation.kind !== 'enrichNames') void refreshPendingDatabase(snapshot.entries)
      if (mutation.kind === 'add') {
        const wanted = new Set(mutation.skus)
        const names = await findPendingNames(snapshot.entries.filter(entry => wanted.has(entry.sku)), mutation.branches)
        if (names.size) acceptPendingSnapshot(await savePendingMutation({ kind: 'enrichNames', names }))
      }
    } catch (error) {
      pendingReadyRef.current = false
      const message = `${saved ? 'รายการบันทึกแล้ว แต่เติมชื่อไม่สำเร็จ: ' : ''}${engineMessage(error)}`
      setPendingSaveError(`${message} · โหลดรายการใหม่ก่อนแก้ไขต่อ`)
      logEngine(`PENDING SAVE ERROR · ${message}`)
    } finally { engineBusyRef.current = false; setEngineBusy('') }
  }

  function addPendingSkus() {
    const branches = ENGINE_BRANCHES.filter(branch => pendingSelectedBranches[branch])
    if (!branches.length) { setPendingSaveError('เลือกสาขาอย่างน้อย 1 สาขา'); return }
    const skus = parsePendingInput(pendingInput)
    if (!skus.length) { setPendingSaveError('กรุณากรอก SKU'); return }
    void editPending({ kind: 'add', skus, branches })
  }

  function selectPendingBranch(branch: EngineBranch) {
    setPendingTab(branch)
    setPendingActionNotice('')
    setPendingExportNotice(null)
  }

  async function loadPendingExportProducts(branch: EngineBranch) {
    const branchName = branch.toUpperCase() as 'SRC' | 'SSS' | 'KKL'
    const snapshot = await loadPendingSnapshot()
    const skus = snapshot.entries.filter(entry => entry.branches.includes(branchName)).map(entry => entry.sku)
    const productsBySku = new Map<string, Product>()
    for (let offset = 0; offset < skus.length; offset += 1200) {
      const chunks = Array.from({ length: Math.min(4, Math.ceil((skus.length - offset) / 300)) }, (_, index) => skus.slice(offset + index * 300, offset + (index + 1) * 300))
      const batches = await Promise.all(chunks.map(async chunk => {
        const response = await fetch('/api/products', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ skus: chunk }) })
        if (!response.ok) throw new Error('โหลดข้อมูลสินค้าสำหรับ Export ไม่สำเร็จ กรุณาลองใหม่')
        const data = await response.json()
        if (!data.success || !Array.isArray(data.products)) throw new Error('อ่านข้อมูลสินค้าสำหรับ Export ไม่สำเร็จ กรุณาลองใหม่')
        return data.products as Product[]
      }))
      for (const product of batches.flat()) productsBySku.set(normalizeStockSku(product['รหัสสินค้า (SKU NUMBER)']), product)
    }
    return { skus, productsBySku }
  }

  function pendingPromaxxRows(branch: EngineBranch, skus: string[], productsBySku: Map<string, Product>): PendingPromaxxRow[] {
    return skus.map(sku => {
      const product = productsBySku.get(sku) ?? null
      const grab = product ? productPrice(product[`${branch}_price`]) : null
      const promaxx = product ? productPrice(product.promaxx_price) : null
      const reason = !product ? 'ไม่พบข้อมูลสินค้า'
        : !hasProductBranch(product, branch) ? 'ไม่มีสินค้าในสาขานี้'
        : promaxx === null || promaxx < 0 ? 'ไม่มีราคา PROMAXX ที่ใช้ได้'
        : grab === null || grab < 0 ? 'ไม่มีราคา GRAB ที่ใช้ได้'
        : !pricesDiffer(grab, promaxx) ? 'ราคาตรงแล้ว' : ''
      return { sku, product, grab, promaxx, ready: !reason, reason }
    })
  }

  function closePendingPromaxxReview() {
    pendingPromaxxRequest.current++
    setPendingPromaxxReview(null)
  }

  async function reviewPendingPromaxxExport(branch: EngineBranch) {
    if (engineBusyRef.current) return
    const request = ++pendingPromaxxRequest.current
    engineBusyRef.current = true; setEngineBusy('pending-price-review'); setPendingExportNotice(null)
    setPendingPromaxxReview({ branch, rows: [], loading: true, error: '' })
    try {
      const { skus, productsBySku } = await loadPendingExportProducts(branch)
      if (request !== pendingPromaxxRequest.current) return
      setPendingPromaxxReview({ branch, rows: pendingPromaxxRows(branch, skus, productsBySku), loading: false, error: '' })
    } catch (error) {
      if (request === pendingPromaxxRequest.current) setPendingPromaxxReview({ branch, rows: [], loading: false, error: engineMessage(error) })
    } finally { engineBusyRef.current = false; setEngineBusy('') }
  }

  async function exportPendingGmMmeXlsx(branch: EngineBranch, review?: PendingPromaxxReview) {
    if (engineBusyRef.current) return
    const request = review ? pendingPromaxxRequest.current : null
    engineBusyRef.current = true; setEngineBusy('pending-export'); setPendingExportNotice(null)
    if (review) setPendingPromaxxReview({ ...review, loading: true, error: '' })
    const branchName = branch.toUpperCase() as 'SRC' | 'SSS' | 'KKL'
    try {
      const { skus, productsBySku } = await loadPendingExportProducts(branch)
      if (review) {
        if (request !== pendingPromaxxRequest.current) return
        const rows = pendingPromaxxRows(branch, skus, productsBySku)
        if (JSON.stringify(rows) !== JSON.stringify(review.rows)) {
          setPendingPromaxxReview({ branch, rows, loading: false, error: 'ข้อมูล PENDING หรือราคาเปลี่ยนแล้ว กรุณาตรวจทานรายการล่าสุดก่อน Export อีกครั้ง' })
          return
        }
      }
      if (!skus.length) {
        setPendingExportNotice({ error: false, message: `ไม่มี PENDING ที่บันทึกไว้สำหรับ ${branchName}` })
        return
      }
      const missing = skus.filter(sku => !productsBySku.has(sku))
      const selected = review ? review.rows.filter(row => row.ready && row.product).map(row => ({ ...row.product!, '*ราคาสินค้า': row.promaxx! })) : skus.flatMap(sku => {
        const product = productsBySku.get(sku)
        return product ? [{ ...product, '*ราคาสินค้า': productPrice(product['*ราคาสินค้า']) ?? '' }] : []
      })
      if (!selected.length) throw new Error(review ? 'ไม่มีรายการที่พร้อมปรับราคา GRAB ตาม PROMAXX' : `ไม่พบข้อมูลสินค้าใน Supabase สำหรับ PENDING ${branchName} (${skus.length} SKU)`)

      const XLSX = await import('xlsx')
      if (review && request !== pendingPromaxxRequest.current) return
      const rows = toExportRows(selected)
      const ws = XLSX.utils.json_to_sheet(rows)
      ws['!cols'] = [16.5, 24.5, 24.5, 24.5, 24.5, 24.5, 33.64].map(wch => ({ wch }))
      ws['!rows'] = [{ hpt: 21 }, ...rows.map(() => ({ hpt: 141 }))]
      rows.forEach((row, index) => {
        ws[`D${index + 2}`].z = '#,##0.00'
        const imageUrl = String(row['*รูปภาพสินค้า']).trim()
        if (/^https?:\/\//i.test(imageUrl)) ws[`F${index + 2}`].l = { Target: imageUrl, Tooltip: 'เปิดรูปภาพสินค้า' }
      })
      const wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, ws, 'TEMPLATE')
      const priceSource = review ? 'PROMAXX' : 'DB Cloud'
      XLSX.writeFile(wb, `GM MME PENDING ${branchName} ${priceSource} ${engineDate()}.xlsx`)
      const message = `Export PENDING ${branchName} · ราคา ${priceSource} สำเร็จ ${selected.length.toLocaleString('th-TH')}/${skus.length.toLocaleString('th-TH')} รายการ${missing.length ? ` · ไม่พบข้อมูลสินค้า ${missing.length} SKU: ${missing.slice(0, 10).join(', ')}${missing.length > 10 ? ' …' : ''}` : ''}${review ? ' · นำไฟล์ไปปรับราคาบน Grab แล้วนำเข้าราคา GRAB ล่าสุดเพื่อยืนยันผล' : ''}`
      setPendingExportNotice({ error: missing.length > 0, message })
      setStatus(message)
      if (review) closePendingPromaxxReview()
    } catch (error) {
      if (review) {
        if (request === pendingPromaxxRequest.current) setPendingPromaxxReview({ ...review, loading: false, error: `Export ไม่สำเร็จ: ${engineMessage(error)}` })
      } else setPendingExportNotice({ error: true, message: `Export PENDING ${branchName} ไม่สำเร็จ: ${engineMessage(error)}` })
    } finally { engineBusyRef.current = false; setEngineBusy('') }
  }

  function pendingPriceText(loaded: boolean, price: number | undefined) {
    return !loaded ? '·' : price === undefined ? '—' : price.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  }

  async function refreshGrabFiles() {
    const request = ++grabFilesRequest.current
    setGrabFilesLoading(true)
    setGrabFilesError('')
    try {
      const { data, error } = await supabase.from('grab_files').select('name,updated_at').order('name')
      if (error) throw new Error(error.message)
      const files = (data ?? []) as GrabFileInfo[]
      if (request !== grabFilesRequest.current) return null
      setGrabFiles(files)
      return files
    } catch (error) {
      if (request === grabFilesRequest.current) setGrabFilesError(`อ่านรายการ Template ไม่สำเร็จ: ${engineMessage(error)}`)
      return null
    } finally { if (request === grabFilesRequest.current) setGrabFilesLoading(false) }
  }

  // Reads the stock already stored in Supabase (the agent's last run). No busy handling here:
  // callers own engineBusyRef.
  async function loadStockFromDb() {
    invalidateEngineResults()
    setEngineChecks(null)
    const next: Partial<Record<EngineBranch, StockReport>> = {}, problems: string[] = [], warnings: string[] = []
    for (const branch of ENGINE_BRANCHES) {
      try {
        const { report, warning } = await loadMssqlStock(branch)
        next[branch] = report
        logEngine(`MSSQL ${branch.toUpperCase()} · ${report.stock.size} SKUs · sync ${report.syncedAt}`)
        if (warning) { warnings.push(warning); logEngine(`⚠ ${warning}`) }
      } catch (error) {
        const problem = `${branch.toUpperCase()}: ${engineMessage(error)}`
        problems.push(problem); logEngine(`BLOCKED ${problem}`)
        // A failed refresh must not leave an older report for that branch selected.
        setEngineReports(previous => { const copy = { ...previous }; delete copy[branch]; return copy })
      }
    }
    setEngineReports(previous => ({ ...previous, ...next }))
    const messages = [...problems, ...warnings]
    if (messages.length) window.alert(`ดึงจากฐานข้อมูล\n\n${messages.join('\n')}`)
  }

  async function handleEngineLoadFromDb() {
    if (engineBusyRef.current) return
    engineBusyRef.current = true
    setEngineBusy('mssql')
    try { await loadStockFromDb() } finally { engineBusyRef.current = false; setEngineBusy('') }
  }

  // Live stock: ask the office agent (tools/stock-sync --serve) to query MSSQL right now.
  // Signed-in users only; the database function enforces it too.
  async function handleEngineLoadLive() {
    if (engineBusyRef.current) return
    if (!grabAdminUser) {
      setPendingAuthOpen(true); setGrabAdminError('')
      window.alert('ต้องเข้าสู่ระบบก่อนดึงจากฐานข้อมูล\n\nใช้บัญชีเดียวกับ GRAB CSV MANAGER ฟอร์มเข้าสู่ระบบอยู่ที่การ์ด PENDING ด้านล่าง')
      setTimeout(() => document.getElementById('pending-login')?.scrollIntoView({ behavior: 'smooth', block: 'center' }), 100)
      return
    }
    engineBusyRef.current = true
    setEngineBusy('mssql')
    try {
      const lastSeen = await loadAgentLastSeen()
      if (!agentOnline(lastSeen)) {
        logEngine('⚠ agent ไม่ตอบสนอง ดึงจากฐานข้อมูลไม่ได้')
        if (window.confirm('เครื่อง agent (BigyaMainPC) ไม่ตอบสนอง จึงดึงข้อมูลสดไม่ได้\n\nต้องการใช้ข้อมูลที่ดึงไว้ล่าสุดแทนหรือไม่? (เวลาของข้อมูลแสดงในการ์ดแต่ละสาขา)')) await loadStockFromDb()
        return
      }
      logEngine('LIVE · สั่งดึงจากฐานข้อมูล...')
      const id = await requestLiveStock()
      const row = await waitForStockRequest(id, update => logEngine(`LIVE · ${requestStatusText(update)}`))
      logEngine(`LIVE · ${requestStatusText(row)}`)
      await loadStockFromDb()
    } catch (error) {
      const message = engineMessage(error)
      logEngine(`BLOCKED ดึงจากฐานข้อมูลไม่สำเร็จ: ${message}`)
      window.alert(`ดึงจากฐานข้อมูล ไม่สำเร็จ\n\n${message}\n\nรายงานที่โหลดไว้เดิมยังไม่ถูกเปลี่ยน`)
    } finally { engineBusyRef.current = false; setEngineBusy('') }
  }

  async function checkEnginePairs() {
    if (engineBusyRef.current) return
    engineBusyRef.current = true; setEngineBusy('check'); setEngineChecks(null)
    try {
      const files = await refreshGrabFiles()
      if (!files) { logEngine('CHECK ไม่สำเร็จ: อ่านรายการ Template จาก Supabase ไม่ได้'); return }
      const checks: Partial<Record<EngineBranch, boolean>> = {}
      for (const branch of ENGINE_BRANCHES) {
        const report = engineReports[branch]
        if (!report) continue
        checks[branch] = files.some(file => file.name === templateName(branch))
        logEngine(`${checks[branch] ? '✔ ถูกต้อง' : '✕ ไม่พบ'} ${report.name} → ${templateName(branch)}`)
      }
      setEngineChecks(checks)
      logEngine(`พร้อม ${Object.values(checks).filter(Boolean).length} คู่จากทั้งหมด ${Object.keys(checks).length} คู่`)
    } finally { engineBusyRef.current = false; setEngineBusy('') }
  }

  async function runProcessEngine() {
    if (engineBusyRef.current) return
    const branches = ENGINE_BRANCHES.filter(branch => engineReports[branch])
    if (!branches.length) return
    const subtracts = {} as Record<EngineBranch, number>
    for (const branch of branches) {
      const value = engineSubtract[branch].trim()
      if (!value || !Number.isFinite(Number(value)) || Number(value) < 0) {
        await notifyPopup(`SUBTRACT ${branch.toUpperCase()} ต้องเป็นตัวเลขตั้งแต่ 0 ขึ้นไป`)
        return
      }
      subtracts[branch] = Number(value)
    }
    engineBusyRef.current = true; setEngineBusy('run')
    ++pendingLoadRequest.current
    pendingReadyRef.current = false
    setPendingLoading(true); setPendingError('')
    invalidateEngineResults(); setEnginePendingCounts(null)
    setEngineProgress({ current: 0, total: branches.length, label: 'กำลังโหลด PENDING SKU UPDATE' })
    const outputs: EngineOutput[] = [], warnings: string[] = []
    const nameSuggestions = new Map<string, string>()
    let skipped = 0, updated = 0
    try {
      const snapshot = await resolvePendingPrices()
      void refreshPendingDatabase(snapshot.entries)
      const pending = pendingSkuSets(snapshot.entries)
      logEngine(`PENDING LOADED · ${ENGINE_BRANCHES.map(branch => `${branch.toUpperCase()}=${pending[branch].size}`).join(' / ')}`)
      const date = engineDate()
      for (let index = 0; index < branches.length; index++) {
        const branch = branches[index], name = templateName(branch), report = engineReports[branch]!
        setEngineProgress({ current: index, total: branches.length, label: `กำลังประมวลผล ${branch.toUpperCase()} (${index + 1}/${branches.length})` })
        logEngine(`RUN ${branch.toUpperCase()} · ${report.name} → ${name} · SUBTRACT=${subtracts[branch]}`)
        try {
          const { data, error } = await supabase.from('grab_files').select('content').eq('name', name).maybeSingle()
          if (error) throw new Error(`โหลด Template ไม่สำเร็จ: ${error.message}`)
          if (!data) { skipped++; logEngine(`SKIPPED ${branch.toUpperCase()} · ไม่พบ ${name}`); continue }
          if (typeof data.content !== 'string') throw new Error('เนื้อหา Template ไม่ใช่ข้อความ CSV')
          const result = processGrabTemplate(name, data.content, report, subtracts[branch], pending[branch])
          if (grabAdminUser) {
            const templateNames = readTemplateNames(name, data.content)
            for (const entry of snapshot.entries) {
              const productName = templateNames.get(entry.sku)
              if (!entry.name.trim() && pending[branch].has(entry.sku) && productName && !nameSuggestions.has(entry.sku)) nameSuggestions.set(entry.sku, productName)
            }
          }
          const outputName = name.replace(/\.CSV$/, `_${date}.csv`)
          outputs.push({ name: outputName, content: result.content, branch, total: result.total, updated: result.updated, matched: result.matched })
          updated += result.updated
          logEngine(`✔ ${branch.toUpperCase()} · ${result.total} แถว · MATCHED ${result.matched} · UPDATED ${result.updated} · IN STOCK ${result.inStock} · PENDING ${result.pendingCount}`)
          if (!result.matched) {
            const warning = `${branch.toUpperCase()}: ไม่มี SKU ตรงกับ R01.102 เลย คอลัมน์ CurrentStock/MaxStock จะว่างทั้งไฟล์ กรุณาตรวจว่าเลือกไฟล์ถูกสาขาก่อนดาวน์โหลด`
            warnings.push(warning); logEngine(`⚠ ${warning}`)
          }
          logEngine(`READY ${outputName}`)
        } catch (error) {
          skipped++
          const warning = `${branch.toUpperCase()}: ${engineMessage(error)}`
          warnings.push(warning); logEngine(`BLOCKED / SKIPPED ${warning}`)
        } finally {
          setEngineProgress({ current: index + 1, total: branches.length, label: `ประมวลผลแล้ว ${index + 1}/${branches.length} สาขา` })
        }
        await new Promise(resolve => setTimeout(resolve, 0))
      }
      if (grabAdminUser && nameSuggestions.size) {
        try { acceptPendingSnapshot(await savePendingMutation({ kind: 'enrichNames', names: nameSuggestions })) }
        catch (error) {
          pendingReadyRef.current = false
          setPendingSaveError(`สร้างผลสต็อกแล้ว แต่บันทึกชื่อ PENDING ไม่สำเร็จ: ${engineMessage(error)} · โหลดรายการใหม่ก่อนแก้ไขต่อ`)
          logEngine(`PENDING NAME WARNING · ${engineMessage(error)}`)
        }
      }
      if (pendingSnapshotRef.current && !samePendingStockMembership(snapshot.entries, pendingSnapshotRef.current.entries)) throw new Error('รายการ PENDING เปลี่ยนระหว่าง RUN กรุณารันใหม่เพื่อใช้รายการล่าสุด')
      setEngineOutputs(outputs)
      setEngineSummary({ processed: outputs.length, updated, skipped })
      setEngineWarnings(warnings)
      logEngine(`PROCESSED ${outputs.length} สาขา / UPDATED ${updated} แถว / SKIPPED ${skipped} สาขา`)
      if (warnings.length) await notifyPopup(`พบปัญหาระหว่างประมวลผล กรุณาตรวจ SYSTEM LOG ก่อนดาวน์โหลด\n\n${warnings.join('\n\n')}`)
    } catch (error) {
      const warning = engineMessage(error)
      pendingReadyRef.current = false; setPendingLoading(false); setPendingError(warning)
      setEngineWarnings([warning]); setEngineSummary({ processed: 0, updated: 0, skipped: branches.length })
      setEngineProgress({ current: 0, total: branches.length, label: 'หยุด: ตรวจสอบ PENDING SKU UPDATE' })
      logEngine(`RUN STOPPED · ${warning}`)
      await notifyPopup(`RUN หยุดทำงาน\n\n${warning}\n\nไม่มีการสร้างไฟล์ผลลัพธ์`)
    } finally { engineBusyRef.current = false; setEngineBusy('') }
  }

  async function uploadGrabTemplates(files: File[]) {
    if (!files.length || engineBusyRef.current) return
    engineBusyRef.current = true; setEngineBusy('templates')
    setGrabUploadProblems([])
    setGrabUploadNotice(''); setGrabExportError(''); setGrabChangeOpen(false)
    const problems: string[] = []
    const reviews: SavedGrabReview[] = []
    try {
      for (const file of files) {
        try {
          const { text, encoding } = decodeCsv(await file.arrayBuffer())
          const template = validateGrabTemplateUpload(text)
          if (reviews.some(item => item.name === template.name)) throw new Error(`มีไฟล์สาขา ${template.branch.toUpperCase()} ในชุดนี้แล้ว กรุณาเลือกหนึ่งไฟล์ต่อสาขา`)
          const review = await saveGrabTemplate(file.name, text)
          reviews.push(review)
          pendingNameCache.current.delete(review.name)
          logEngine(`SAVED ${file.name} → ${review.name} · ${review.total} แถว · ${encoding} · แก้ไข ${review.counts.updated} / เพิ่ม ${review.counts.added} / นำออก ${review.counts.removed}`)
        } catch (error) {
          const problem = `${file.name}: ${engineMessage(error)}`
          problems.push(problem); logEngine(`BLOCKED ${problem}`)
        }
      }
      setGrabUploadProblems(problems)
      if (reviews.length) {
        invalidateEngineResults(); setEngineChecks(null)
        setGrabUploadResults(reviews)
        const changed = reviews.filter(review => review.changed).length
        const added = reviews.filter(review => review.isNew).length
        const unchanged = reviews.length - changed - added
        setGrabUploadNotice(`บันทึกสำเร็จ ${reviews.length} ไฟล์ · ข้อมูลเปลี่ยน ${changed} ไฟล์ · ข้อมูลไม่เปลี่ยน ${unchanged} ไฟล์${added ? ` · เพิ่มไฟล์ใหม่ ${added} ไฟล์` : ''}`)
      } else setGrabUploadResults([])
      await refreshGrabFiles()
      logEngine(`UPLOAD TEMPLATE · SAVED ${reviews.length} / BLOCKED ${problems.length}`)
      if (reviews.some(review => review.changed)) openGrabChangeNotice(reviews)
      else if (problems.length) await notifyPopup(`บันทึกสำเร็จ ${reviews.length} ไฟล์ / BLOCKED ${problems.length} ไฟล์\n\n${problems.join('\n\n')}`)
    } finally { engineBusyRef.current = false; setEngineBusy('') }
  }

  function openGrabChangeNotice(reviews = grabUploadResults) {
    const review = reviews.find(item => item.changed)
    if (!review) return
    setGrabChangeBranch(review.branch)
    setGrabExpandedRow(review.rows.find(row => row.type !== 'unchanged')?.key ?? null)
    setGrabChangeLimit(50)
    setGrabChangeOpen(true)
  }

  function downloadGrabCsv(name: string, text: string) {
    const url = URL.createObjectURL(new Blob([grabCsvWithBom(text)], { type: 'text/csv;charset=utf-8' }))
    const link = document.createElement('a')
    link.href = url; link.download = name
    document.body.appendChild(link); link.click(); link.remove()
    setTimeout(() => URL.revokeObjectURL(url), 1000)
  }

  async function exportStoredGrabTemplate(name: string) {
    if (engineBusyRef.current) return
    engineBusyRef.current = true; setEngineBusy('template-export'); setGrabExportError('')
    try {
      const { data, error } = await supabase.from('grab_files').select('content').eq('name', name).maybeSingle()
      if (error) throw new Error(error.message)
      if (!data || typeof data.content !== 'string') throw new Error('ไม่พบเนื้อหา CSV ของไฟล์นี้')
      downloadGrabCsv(name, data.content)
      logEngine(`EXPORT ${name}`)
    } catch (error) { setGrabExportError(`Export ${name} ไม่สำเร็จ: ${engineMessage(error)}`) }
    finally { engineBusyRef.current = false; setEngineBusy('') }
  }

  function downloadEngineOutputs() {
    if (!engineOutputs.length || engineBusyRef.current) return
    const zip = createResultsZip(engineOutputs)
    const url = URL.createObjectURL(new Blob([zip], { type: 'application/zip' }))
    const link = document.createElement('a')
    link.href = url; link.download = `GRAB_ALL_${engineDate()}.zip`
    document.body.appendChild(link); link.click(); link.remove()
    setTimeout(() => URL.revokeObjectURL(url), 1000)
    logEngine(`DOWNLOAD ALL · ${engineOutputs.length} ไฟล์ CSV ใน ZIP`)
  }

  async function signInGrabManager(event: React.FormEvent) {
    event.preventDefault()
    if (engineBusyRef.current) return
    engineBusyRef.current = true; setEngineBusy('auth'); setGrabAdminError('')
    try {
      const { error } = await supabase.auth.signInWithPassword({ email: grabAdminEmail.trim(), password: grabAdminPassword })
      if (error) throw new Error('เข้าสู่ระบบไม่สำเร็จ กรุณาตรวจอีเมลและรหัสผ่านของบัญชีที่มีสิทธิ์บันทึก')
      logEngine('เข้าสู่ระบบสำหรับจัดการ Template และ PENDING แล้ว')
    } catch (error) { setGrabAdminError(engineMessage(error)) }
    finally { setGrabAdminPassword(''); engineBusyRef.current = false; setEngineBusy('') }
  }

  async function signOutGrabManager() {
    if (engineBusyRef.current) return
    engineBusyRef.current = true; setEngineBusy('auth'); setGrabAdminError('')
    try {
      const { error } = await supabase.auth.signOut({ scope: 'local' })
      if (error) throw new Error(error.message)
      setPendingAuthOpen(false); setGrabAdminPassword('')
      logEngine('ออกจากระบบจัดการ Template และ PENDING แล้ว')
    } catch (error) { setGrabAdminError(engineMessage(error)) }
    finally { engineBusyRef.current = false; setEngineBusy('') }
  }

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
    productRequestAbort.current?.abort()
    const controller = new AbortController()
    productRequestAbort.current = controller
    setLoading(true)
    setLoadError('')
    setPriceReview(null)
    setStatus('กำลังโหลดสินค้า...')
    try {
      const res = await fetch(url, { signal: controller.signal })
      if (!res.ok) throw new Error('โหลดข้อมูลไม่สำเร็จ')
      const data = await res.json()
      if (!data.success || !Array.isArray(data.products)) throw new Error('โหลดข้อมูลไม่สำเร็จ')
      if (reqId !== searchReqId.current) return
      setProducts(data.products)
      setPendingPriceRevision(value => value + 1)
      preloadImages(data.products)
      setStatus(`โหลดสินค้าแล้ว ${data.products.length.toLocaleString('th-TH')} รายการ`)
    } catch {
      if (controller.signal.aborted || reqId !== searchReqId.current) return
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
    productList.slice(0, PRODUCT_PAGE_SIZE).forEach(p => {
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
    XLSX.writeFile(wb, `GM MME DB Cloud ${dd}.${mm}.${yyyy}.xlsx`)
  }

  function toExportRows(list: Product[]) {
    return list.map(p => ({
      '*ประเภทสินค้า': p['ประเภทสินค้า'] || '',
      '*ชื่อสินค้า': p['*ชื่อสินค้า (NAME)'] || '',
      '*เลขที่ใบอนุญาตโฆษณา': p['*เลขที่ใบอนุญาตโฆษณา'] || '',
      '*ราคาสินค้า': p['*ราคาสินค้า'] ?? '',
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

    const ok = await notifyPopup(
      `ยืนยันลบหมวดหมู่ "${selectedSheet}" ออกจากทุกสาขา?\nการลบนี้ไม่สามารถย้อนกลับได้`, true
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
    if (priceCalcBusyRef.current) return
    priceCalcBusyRef.current = true
    setSavingNewPromaxx(true)
    setNewPromaxxSaveError('')
    setStatus('กำลังคำนวณราคา...')
    setShowPriceCalcModal(true)
    setPriceCalcResults([])
    let persisting = false

    try {
      const { text } = decodeCsv(await file.arrayBuffer())
      const { rows } = parseEngineCsv(text)
      const headerIndex = rows.slice(0, 10).findIndex(row => {
        const cells = row.map(cell => cell.trim().toUpperCase())
        return cells.includes('CF_ITEMID') || (cells[0] === 'SKU' && cells[4] === '0')
      })
      if (headerIndex < 0) throw new Error('ไม่พบหัวตาราง R05.105: ต้องมี CF_ITEMID หรือหัวตาราง SKU / ระดับ 0')
      const header = rows[headerIndex].map(cell => cell.trim().toUpperCase())
      const isRawReport = header.includes('CF_ITEMID')
      const skuColumn = isRawReport ? header.indexOf('CF_ITEMID') : 0
      const priceColumn = isRawReport ? header.indexOf('CF_ITEMPRICELEVEL_PRICE') : 4
      const multipleColumn = header.indexOf('CF_BASEMULTIPLE')
      const levelColumn = header.indexOf('CF_LEVELNO')
      if (isRawReport && [priceColumn, multipleColumn, levelColumn].some(column => column < 0)) {
        throw new Error('หัวตาราง R05.105 ต้องมี CF_ITEMID, CF_BASEMULTIPLE, CF_LEVELNO และ CF_ITEMPRICELEVEL_PRICE')
      }
      // Raw Promaxx: B=SKU, G=price, D=base multiple 1, F=level 0.
      // Retain the explicit compact SKU/.../0 format at A/E.
      const calcMap = new Map<string, { level0: number; newPrice: number }>()
      for (const row of rows.slice(headerIndex + 1)) {
        if (isRawReport && (parsePriceRobust(row[multipleColumn] ?? '') !== 1 || parsePriceRobust(row[levelColumn] ?? '') !== 0)) continue
        const sku = row[skuColumn]?.trim(), level0 = parsePriceRobust(row[priceColumn] ?? '')
        if (!sku || level0 === null || !Number.isFinite(level0) || level0 < 0) continue
        const base = level0 * 0.95 * 1.20
        const newPrice = Math.ceil(base + base * 0.07)
        if (Number.isFinite(newPrice)) calcMap.set(sku, { level0, newPrice })
      }
      if (!calcMap.size) throw new Error('ไม่พบราคา R05.105 ที่ถูกต้องสำหรับราคาระดับ 0')

      setStatus(`กำลังดึงข้อมูลสินค้า... (${calcMap.size} SKU)`)
      const response = await fetch('/api/products', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ skus: [...calcMap.keys()] }),
      })
      const data = await response.json()
      if (!response.ok || !data.success || !Array.isArray(data.products)) throw new Error('ดึงข้อมูลสินค้าเพื่อคำนวณไม่สำเร็จ')
      const results = (data.products as Product[]).flatMap(product => {
        const sku = String(product['รหัสสินค้า (SKU NUMBER)'])
        const calc = calcMap.get(sku)
        if (!calc) return []
        const oldPrice = parsePriceRobust(String(product['*ราคาสินค้า'])) ?? 0
        return [{ sku, name: product['*ชื่อสินค้า (NAME)'], ...calc, oldPrice, changed: Math.abs(calc.newPrice - oldPrice) > 0.01, _product: product }]
      })
      if (!results.length) throw new Error('ไม่พบ SKU จากไฟล์ R05.105 ในฐานข้อมูล')
      setPriceCalcResults(results)
      setStatus(`กำลังบันทึก New Promaxx... (${results.length} SKU)`)
      persisting = true

      // Persist every matched calculation, including results equal to the current DB price.
      for (let offset = 0; offset < results.length; offset += 300) {
        const chunk = results.slice(offset, offset + 300)
        const saved = await fetch('/api/products', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ updateNewPromaxxPrice: true, priceMap: Object.fromEntries(chunk.map(result => [result.sku, result.newPrice])) }),
        })
        const result = await saved.json()
        if (!saved.ok || !result.success || result.updated !== chunk.length) throw new Error(result.error || 'บันทึก New Promaxx ไม่ครบ กรุณาคำนวณใหม่เพื่อลองอีกครั้ง')
        const prices = new Map(chunk.map(result => [result.sku, result.newPrice]))
        setProducts(previous => previous.map(product => {
          const price = prices.get(String(product['รหัสสินค้า (SKU NUMBER)']))
          return price === undefined ? product : { ...product, new_promaxx_price: price }
        }))
      }
      setStatus(`คำนวณและบันทึก New Promaxx แล้ว ${results.length} รายการ`)
    } catch (error) {
      const message = error instanceof Error ? error.message : 'คำนวณหรือบันทึก New Promaxx ไม่สำเร็จ'
      setNewPromaxxSaveError(`${message}${persisting ? ' · บันทึกอาจยังไม่ครบ โหลดรายการสินค้าใหม่เพื่อตรวจสอบก่อนลองอีกครั้ง' : ''}`)
      setStatus('คำนวณหรือบันทึก New Promaxx ไม่สำเร็จ')
    } finally {
      priceCalcBusyRef.current = false
      setSavingNewPromaxx(false)
    }
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
    const referencePrice = productPrice(product.promaxx_price)
    const branchPrice = productPrice(product[`${inspectionBranch}_price`])
    if (referencePrice === null || branchPrice === null) return 'unpriced'
    return pricesDiffer(branchPrice, referencePrice) ? 'mismatch' : 'matched'
  }

  const priceLabels: Record<PriceState, string> = { mismatch: 'ต้องตรวจ', matched: 'ราคาตรง', unpriced: 'ยังไม่มีราคา', absent: 'ไม่มีในสาขา' }
  const branchNames: Record<BranchKey, string> = { src: 'SRC', kkl: 'KKL', sss: 'SSS' }
  const branches = Object.keys(branchNames) as BranchKey[]
  const normalizedSearch = search.trim().toLowerCase()
  const filteredProducts = products.filter(product =>
    (selectedSheet === 'all' || product['หมวดหมู่สินค้า (CATEGORIES)'] === selectedSheet) &&
    (catalogOpen || branchFilter === 'all' || hasProductBranch(product, branchFilter)) &&
    (catalogOpen || priceFilter === 'all' || priceState(product) === priceFilter) &&
    (!normalizedSearch || String(product['รหัสสินค้า (SKU NUMBER)'] ?? '').toLowerCase().includes(normalizedSearch) || String(product['*ชื่อสินค้า (NAME)'] ?? '').toLowerCase().includes(normalizedSearch))
  )
  const pageSize = PRODUCT_PAGE_SIZE
  const pageCount = Math.max(1, Math.ceil(filteredProducts.length / pageSize))
  const currentPage = Math.min(pageNumber, pageCount)
  const pageProducts = filteredProducts.slice((currentPage - 1) * pageSize, currentPage * pageSize)
  const inspectedProduct = filteredProducts.find(product => product['รหัสสินค้า (SKU NUMBER)'] === activeSku) ?? pageProducts[0] ?? null
  const inspectedSku = inspectedProduct?.['รหัสสินค้า (SKU NUMBER)'] as string | undefined
  const originalPrice = productPrice(inspectedProduct?.['*ราคาสินค้า'])
  const referencePrice = productPrice(inspectedProduct?.promaxx_price)
  const draftPrice = inspectedSku ? (Object.hasOwn(priceDrafts, inspectedSku) ? priceDrafts[inspectedSku] : (originalPrice === null ? '' : originalPrice.toFixed(2))) : ''
  const proposedPrice = productPrice(draftPrice)
  const hasPriceChange = draftPrice.trim() !== '' && proposedPrice !== null && proposedPrice !== originalPrice
  const comparisonPrice = inspectedProduct && hasProductBranch(inspectedProduct, inspectionBranch) ? productPrice(inspectedProduct[`${inspectionBranch}_price`]) : null
  const priceDifference = comparisonPrice !== null && referencePrice !== null ? comparisonPrice - referencePrice : null
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
      setStatus(`บันทึกราคา DB / Cloud ${review.sku} สำเร็จ · ${formatMoney(review.price)}`)
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

  const toolGroups: { key: string; icon: string; title: string; items: WorkspaceTool[] }[] = [
    { key: 'import', icon: 'upload', title: 'นำเข้าและเตรียมไฟล์', items: [
      { name: 'เพิ่มสินค้ารายการเดียว', detail: 'แบบฟอร์ม', action: openAddProduct },
      { name: 'นำเข้า Menu Bulk', detail: 'Excel', action: () => document.getElementById('importInput')?.click() },
      { name: 'นำเข้า Product Master', detail: 'Excel / CSV', action: () => document.getElementById('masterInput')?.click() },
      { name: 'แปลง CSV เป็น UTF-8', detail: 'POS', action: () => document.getElementById('csvConvertInput')?.click() },
    ] },
    { key: 'price', icon: 'tag', title: 'ตรวจสอบราคา', items: [
      { name: 'คำนวณและอัปเดตราคา DB', detail: 'R05.105', action: () => document.getElementById('priceCalcInput')?.click(), disabled: savingNewPromaxx, busy: savingNewPromaxx },
      { name: 'เทียบราคา Promaxx', detail: 'R05.103', action: () => document.getElementById('promaxxCheckInput')?.click() },
      { name: 'ตรวจ GM MME กับ GRAB', detail: '3 สาขา', action: () => { setShowMmeCheckModal(true); setMmeCheckResults({ src: [], kkl: [], sss: [] }) } },
    ] },
    { key: 'grab', icon: 'branch', title: 'GRAB และสาขา', items: [
      ...branches.map(branch => ({ name: 'นำเข้า GRAB ' + branchNames[branch], detail: 'CSV', action: () => document.getElementById('grabInput' + branch[0].toUpperCase() + branch.slice(1))?.click() })),
      { name: 'ตรวจสินค้าที่ขาดใน GRAB', detail: 'ตรวจสอบ', action: openMissingBranches },
      { name: 'เปรียบเทียบ Product Master 3 สาขา', detail: 'Master', action: handleMasterBranchCompare },
    ] },
    { key: 'export', icon: 'download', title: 'ส่งออกและข้อมูลกลาง', items: [
      { name: 'ส่งออก SKU ที่ขาดจาก Master', detail: 'Excel', action: exportMissingSkus, disabled: exportingMissing, busy: exportingMissing },
      { name: 'ส่งออกหมวดหมู่ที่เลือก', detail: 'Excel', action: exportSelectedCategoryXlsx, disabled: selectedSheet === 'all' },
    ] },
  ]
  const normalizeToolText = (text: string) => text.toLowerCase().replace(/[\s._/-]+/g, '')
  const toolSearchTerms = toolSearch.trim().split(/\s+/).filter(Boolean).map(normalizeToolText)
  const matchesTool = (text: string) => toolSearchTerms.every(term => normalizeToolText(text).includes(term))
  const filteredToolGroups = toolGroups.map(group => ({ ...group, items: group.items.filter(tool => matchesTool(group.title + ' ' + tool.name + ' ' + tool.detail)) })).filter(group => group.items.length)
  const showGrabManager = matchesTool('GRAB CSV MANAGER Template แม่แบบ เทมเพลต อัปโหลดไฟล์ สาขา SRC KKL SSS')
  const matchingToolCount = filteredToolGroups.reduce((total, group) => total + group.items.length, 0) + (showGrabManager ? 1 : 0)

  const productPendingSets = pendingSnapshot ? pendingSkuSets(pendingSnapshot.entries) : null
  const productPendingError = pendingError || pendingSaveError
  const productPendingReady = Boolean(productPendingSets && !pendingLoading && !productPendingError)
  const pendingProductCoverage = pendingProduct && productPendingSets ? pendingBranchCoverage(pendingProduct.sku, productPendingSets) : null

  const pendingCanEdit = Boolean(grabAdminUser && pendingReadyRef.current && !pendingLoading && !engineBusy)
  const pendingEntries = pendingSnapshot?.entries.filter(entry => entry.branches.length) ?? []
  const pendingRows = pendingEntries.filter(entry => entry.branches.includes(pendingTab.toUpperCase() as 'SRC' | 'SSS' | 'KKL'))
  const pendingCounts = ENGINE_BRANCHES.map(branch => pendingEntries.filter(entry => entry.branches.includes(branch.toUpperCase() as 'SRC' | 'SSS' | 'KKL')).length)
  const pendingMismatchCount = pendingRows.filter(entry => {
    const check = pendingAutoPrices.get(entry.sku)
    return comparePendingPrices(check?.prices[pendingTab], check?.promaxx) === 'diff'
  }).length

  return (
    <div className="pc-app">
      <a className="pc-skip-link" href="#product-workspace">ข้ามไปยังพื้นที่ทำงาน</a>
      <aside className="pc-rail" aria-label="เมนูหลัก">
        <Link href="/" className="pc-brand" aria-label="Price Checker หน้าหลัก"><WorkspaceIcon name="tag" /></Link>
        <nav>
          <button className={!engineOpen && !toolsOpen && catalogOpen ? 'is-active' : ''} title="รายการสินค้า" aria-label="รายการสินค้า" onClick={() => { setEngineOpen(false); setToolsOpen(false); setCatalogOpen(true); setPriceFilter('all'); setPageNumber(1) }}><WorkspaceIcon name="nav-products" /><span className="pc-nav-label">รายการสินค้า</span></button>
          <button className={!engineOpen && !toolsOpen && !catalogOpen ? 'is-active' : ''} title="ตรวจสอบราคา" aria-label="ตรวจสอบราคา" onClick={() => { setEngineOpen(false); setToolsOpen(false); setCatalogOpen(false); setPriceFilter('mismatch'); setPageNumber(1); resetPriceReview() }}><WorkspaceIcon name="nav-price-check" /><span className="pc-nav-label">ตรวจสอบราคา</span></button>
          <button title="เปรียบเทียบ 3 สาขา" aria-label="เปรียบเทียบ 3 สาขา" onClick={handleMasterBranchCompare}><WorkspaceIcon name="nav-branches" /><span className="pc-nav-label">เปรียบเทียบ 3 สาขา</span></button>
          <button title="นำเข้าสินค้า" aria-label="นำเข้าสินค้า" onClick={() => document.getElementById('importInput')?.click()}><WorkspaceIcon name="nav-import" /><span className="pc-nav-label">นำเข้าสินค้า</span></button>
          <button className={engineOpen ? 'is-active' : ''} ref={engineNavRef} title="PROCESS ENGINE" aria-label="PROCESS ENGINE" aria-pressed={engineOpen} onClick={openProcessEngine}><WorkspaceIcon name="nav-engine" /><span className="pc-nav-label">PROCESS ENGINE</span></button>
          <button className={!engineOpen && toolsOpen ? 'is-active' : ''} title="เครื่องมือทั้งหมด" aria-label="เครื่องมือทั้งหมด" aria-expanded={!engineOpen && toolsOpen} aria-controls="workflow-tools" onClick={() => { setEngineOpen(false); setToolsOpen(value => !value) }}><WorkspaceIcon name="nav-tools" /><span className="pc-nav-label">เครื่องมือทั้งหมด</span></button>
        </nav>
        <div className="pc-rail-bottom"><span className="pc-rail-divider" /><span className="pc-workspace-avatar" title="ANIN Workspace">AN</span></div>
      </aside>

      <main className="pc-workspace" id="product-workspace">
        <header className="pc-page-header">
          <div><div className="pc-eyebrow">PRICE CHECKER / {engineOpen ? 'STOCK OPERATIONS' : toolsOpen ? 'TOOLS' : 'PRODUCT OPERATIONS'}</div><h1>{engineOpen ? 'PROCESS ENGINE' : toolsOpen ? 'เครื่องมือทั้งหมด' : catalogOpen ? 'รายการสินค้า' : 'ตรวจและปรับราคาสินค้า'}<span className="pc-heading-dot">.</span></h1>{(engineOpen || toolsOpen) && <p>{engineOpen ? 'เติมสต็อก R01.102 ลงใน Template ที่ Grab อนุมัติ แยกตามสาขา' : 'เลือกหมวดที่ต้องการ แล้วเริ่มงานได้ทันที'}</p>}</div>
          <div className="pc-header-actions">
            <span className="pc-workspace-label">ANIN <i /> WORKSPACE</span>
            {engineOpen ? <button className="pc-button" onClick={() => { setEngineOpen(false); setToolsOpen(false); setCatalogOpen(true); setPriceFilter('all'); setPageNumber(1) }}><WorkspaceIcon name="box" />รายการสินค้า</button> : <><button className="pc-button pc-button-subtle" onClick={openAddProduct}><WorkspaceIcon name="plus" />เพิ่มสินค้า</button><button className="pc-button" onClick={openProcessEngine}><WorkspaceIcon name="engine" />PROCESS ENGINE</button><button className="pc-button" aria-expanded={toolsOpen} aria-controls="workflow-tools" onClick={() => setToolsOpen(value => !value)}><WorkspaceIcon name="settings" />เครื่องมือ</button></>}
          </div>
        </header>




        {engineOpen && <div className="pe-entry-shell">
          {engineEntry.status !== 'idle' && <div className="pe-entry-overlay" data-state={engineEntry.status}>
            <div className="pe-entry-panel-wrap">
              <div className={`pe-entry-card${engineEntry.status === 'error' ? ' is-error' : ''}`} role={engineEntry.status === 'error' ? 'alert' : 'status'} aria-live={engineEntry.status === 'error' ? 'assertive' : 'polite'} aria-atomic="true">
                {engineEntry.status === 'loading' ? <>
                  <div className="pe-entry-orbit" aria-hidden="true"><i className="pe-entry-ring outer" /><i className="pe-entry-ring" /><div className="pe-entry-core"><WorkspaceIcon name="database" /></div>{ENGINE_BRANCHES.map((branch, index) => <div className={`pe-entry-track track-${index}`} key={branch}><span className="pe-entry-node">{branch.toUpperCase()}</span></div>)}</div>
                  <h2>กำลังเชื่อมข้อมูลให้พร้อม</h2>
                  <p className="pe-entry-caption">{!engineEntry.pending ? 'กำลังโหลด PENDING และตรวจราคา GRAB / Promaxx' : !engineEntry.prices ? 'กำลังอ่านราคา DB ของสินค้า' : 'กำลังอ่านรายการ Template'}</p>
                  <p className="pe-entry-count">เตรียมข้อมูลพร้อมแล้ว <strong>{Number(engineEntry.pending) + Number(engineEntry.prices) + Number(engineEntry.templates)} / 3</strong></p>
                  <div className="pe-entry-ticks" aria-hidden="true"><i className={engineEntry.pending ? 'is-done' : ''} /><i className={engineEntry.prices ? 'is-done' : ''} /><i className={engineEntry.templates ? 'is-done' : ''} /></div>
                </> : <>
                  <span className="pe-entry-error-icon"><WorkspaceIcon name="alert" /></span>
                  <h2>โหลดข้อมูลไม่สำเร็จ</h2><p className="pe-entry-caption">{engineEntry.error} กรุณาลองอีกครั้ง</p>
                  <button className="pc-button pc-button-primary" onClick={() => { engineNavRef.current?.focus(); beginEngineEntry() }}><WorkspaceIcon name="refresh" />ลองอีกครั้ง</button>
                </>}
              </div>
            </div>
          </div>}
          <section className="pe-workspace" aria-label="PROCESS ENGINE workspace" aria-busy={Boolean(engineBusy) || engineEntry.status === 'loading'} inert={engineEntry.status !== 'idle'}>
          <div className="pe-notice"><WorkspaceIcon name="clock" /><p>R01.102 และผลลัพธ์เก็บในหน่วยความจำของหน้านี้เท่านั้น <strong>ดาวน์โหลดก่อนรีเฟรชหรือปิดหน้า</strong> · Template เก็บไว้ใน Supabase และใช้ซ้ำได้ทุกวัน</p></div>
          <section className="pe-card" aria-labelledby="pe-reports-heading">
            <div className="pe-card-heading"><div><span className="pc-eyebrow">01 / PREPARE</span><h2 id="pe-reports-heading">ดึงจำนวนสินค้าในระบบ</h2><p>เข้าสู่ระบบก่อนดึงข้อมูล</p></div><button className={`pc-button pe-mssql-button${engineBusy === 'mssql' ? ' is-loading' : ''}`} disabled={Boolean(engineBusy)} aria-busy={engineBusy === 'mssql'} onClick={() => void handleEngineLoadLive()}>
                <span className="pe-mssql-button-icon" aria-hidden="true">{engineBusy === 'mssql' ? <span className="pe-mssql-orbit"><WorkspaceIcon name="database" /></span> : <WorkspaceIcon name="refresh" />}</span>
                <span>{engineBusy === 'mssql' ? 'กำลังดึงข้อมูล...' : 'ดึงจากฐานข้อมูล'}</span>
              </button></div>
            <div className="pe-branch-grid">{ENGINE_BRANCHES.map(branch => {
              const report = engineReports[branch]
              return <div className={`pe-branch-report ${report ? 'is-ready' : ''}`} key={branch}><div><strong>{branch.toUpperCase()}</strong><span className={`pc-status ${report ? 'pc-status-matched' : 'pc-status-unpriced'}`}>{report ? (report.source === 'mssql' ? 'MSSQL' : 'อัปโหลดแล้ว') : 'รอข้อมูล'}</span></div>{report ? <><b>{report.name}</b><p>{report.stock.size.toLocaleString('th-TH')} SKUs · {report.source === 'mssql' && report.syncedAt ? `ข้อมูล ณ ${new Intl.DateTimeFormat('th-TH', { timeZone: 'Asia/Bangkok', day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }).format(new Date(report.syncedAt))}` : report.encoding}</p><button className="pc-text-button" disabled={Boolean(engineBusy)} aria-label={`นำรายงาน ${branch.toUpperCase()} ออก`} onClick={() => { setEngineReports(previous => { const next = { ...previous }; delete next[branch]; return next }); setEngineChecks(null); invalidateEngineResults() }}><WorkspaceIcon name="close" />นำไฟล์ออก</button></> : <p>ดึงจากฐานข้อมูล</p>}</div>
            })}</div>
          </section>

          <section className="pe-card pn-card" aria-labelledby="pending-heading" aria-busy={pendingLoading || engineBusy === 'pending' || engineBusy === 'pending-auto' || engineBusy === 'pending-export'}>
            <div className="pe-card-heading">
              <div><span className="pc-eyebrow">02 / PENDING APPROVAL</span><h2 id="pending-heading">PENDING SKU UPDATE <span className="pc-count">{pendingSnapshot ? pendingEntries.length : '—'} SKUs</span></h2><p>ระงับขายแยกสาขา ทั้งรายการราคาไม่ตรงที่เพิ่มอัตโนมัติ และรายการที่เพิ่มเองเพื่อรออนุมัติ</p></div>
              <div className="pn-heading-actions">
                <button className="pc-button pc-button-primary" disabled={Boolean(engineBusy) || pendingLoading} aria-busy={engineBusy === 'pending-price-review'} onClick={() => void reviewPendingPromaxxExport(pendingTab)}><WorkspaceIcon name="download" />Export ปรับราคา GRAB ตาม PROMAXX · {pendingTab.toUpperCase()}</button>
                <button className="pc-button" disabled={Boolean(engineBusy) || pendingLoading} aria-busy={engineBusy === 'pending-export'} onClick={() => void exportPendingGmMmeXlsx(pendingTab)}><WorkspaceIcon name="download" />{engineBusy === 'pending-export' ? 'กำลัง Export...' : `Export GM MME · DB / Cloud · ${pendingTab.toUpperCase()}`}</button>
                <button className="pc-button" disabled={Boolean(engineBusy) || pendingLoading} onClick={() => void refreshPendingCard()}><WorkspaceIcon name="refresh" />โหลดรายการใหม่</button>{!grabAdminUser && <button className="pc-button" disabled={Boolean(engineBusy)} aria-expanded={pendingAuthOpen} aria-controls="pending-login" onClick={() => { setPendingAuthOpen(value => !value); setGrabAdminError('') }}>เข้าสู่ระบบเพื่อจัดการ PENDING</button>}
              </div>
            </div>
            {grabAdminUser && <div className="pe-admin-session"><span>เข้าสู่ระบบแล้ว: {grabAdminUser} · จัดการ PENDING ได้</span><button className="pc-text-button" disabled={Boolean(engineBusy)} onClick={() => void signOutGrabManager()}>ออกจากระบบ</button></div>}
            {pendingAuthOpen && !grabAdminUser && <form id="pending-login" className="pe-admin-login" onSubmit={signInGrabManager}>
              <h3>เข้าสู่ระบบเพื่อจัดการ PENDING</h3><p>ใช้บัญชีเดียวกับ GRAB CSV MANAGER ผู้ใช้ที่ยังไม่ล็อกอินดูรายการและเทียบราคาได้</p>
              <div><label htmlFor="pending-admin-email">อีเมล<input id="pending-admin-email" type="email" required autoComplete="username" disabled={Boolean(engineBusy)} value={grabAdminEmail} onChange={event => setGrabAdminEmail(event.target.value)} /></label><label htmlFor="pending-admin-password">รหัสผ่านบัญชี<input id="pending-admin-password" type="password" required autoComplete="current-password" disabled={Boolean(engineBusy)} value={grabAdminPassword} onChange={event => setGrabAdminPassword(event.target.value)} /></label><button className="pc-button pc-button-primary" disabled={Boolean(engineBusy)} type="submit">เข้าสู่ระบบ</button></div>
              {grabAdminError && <p className="pe-error" role="alert">{grabAdminError}</p>}
            </form>}
            {pendingError && <p className="pe-error" role="alert">{pendingError} · ปิดการแก้ไขเพื่อป้องกันเขียนทับข้อมูลเดิม กรุณาโหลดรายการใหม่</p>}
            {pendingSaveError && <p className="pe-error" role="alert">{pendingSaveError}</p>}
            <p className="pc-pending-note">Export เฉพาะสาขาที่เปิดอยู่ · เลือก PROMAXX เพื่อปรับราคาขาย GRAB หรือ DB / Cloud เพื่อส่งออกราคาบน Cloud</p>
            {pendingActionNotice && <p className="pc-pending-note" role="status">{pendingActionNotice}</p>}
            {pendingExportNotice && <p className={pendingExportNotice.error ? 'pe-error' : 'pc-pending-note'} role={pendingExportNotice.error ? 'alert' : 'status'}>{pendingExportNotice.message}</p>}
            <p className="pn-source-status">{engineBusy === 'pending-auto' ? 'กำลังตรวจราคา GRAB / Promaxx...' : pendingAutoStatus.loaded ? `ตรวจราคาล่าสุด ${pendingAutoStatus.checkedAt}` : 'ยังไม่ได้ตรวจราคา'}</p>
            {!pendingAutoStatus.saved && pendingAutoStatus.loaded && <p className="pn-mismatch-note" role="status">ผลตรวจราคาอัตโนมัติบางรายการยังไม่บันทึกลงฐานข้อมูล · เข้าสู่ระบบเพื่อบันทึกและจัดการ PENDING</p>}
            <div className="pn-add-area">
              <h3>เพิ่มสินค้าเข้า PENDING</h3><p className="pn-source-status">เลือกสาขาที่รออนุมัติ แล้วกรอกหรือวาง SKU ที่ต้องการเพิ่ม</p>
              <fieldset disabled={!pendingCanEdit}><legend>สาขาที่รออนุมัติ</legend><div className="pn-branch-checkboxes">{ENGINE_BRANCHES.map(branch => <label key={branch}><input type="checkbox" aria-label={`เลือกสาขา Pending ${branch.toUpperCase()}`} checked={pendingSelectedBranches[branch]} onChange={event => setPendingSelectedBranches(previous => ({ ...previous, [branch]: event.target.checked }))} />{branch.toUpperCase()}</label>)}</div></fieldset>
              <label className="pn-input-label" htmlFor="pending-sku-input">SKU ที่ต้องการเพิ่ม<span>ขึ้นบรรทัดใหม่ หรือคั่นด้วย , ; | และ Tab · Enter เพิ่ม · Shift+Enter ขึ้นบรรทัดใหม่</span></label>
              <div className="pn-add-input"><textarea id="pending-sku-input" rows={3} disabled={!pendingCanEdit} placeholder={grabAdminUser ? 'กรอกหรือวาง SKU หลายรายการ...' : 'เข้าสู่ระบบเพื่อเพิ่มรายการ PENDING'} value={pendingInput} onChange={event => setPendingInput(event.target.value)} onKeyDown={event => { if (event.key === 'Enter' && !event.shiftKey && !event.nativeEvent.isComposing) { event.preventDefault(); addPendingSkus() } }} /><button className="pc-button pc-button-primary" disabled={!pendingCanEdit || !pendingInput.trim()} onClick={addPendingSkus}>{engineBusy === 'pending' ? 'กำลังบันทึก...' : '+ ADD'}</button></div>
            </div>
            <p className="pn-price-note">ราคา GRAB ใช้ข้อมูลที่นำเข้าจากแต่ละสาขา · PROMAXX เป็นราคาขายอ้างอิง · DB / Cloud เป็นราคาสำหรับตรวจสอบบน Cloud</p>
            {pendingDbPrices.error && <p className="pe-error" role="alert">{pendingDbPrices.error}</p>}
            <div className="pn-tabs" role="tablist" aria-label="สาขาของ PENDING">{ENGINE_BRANCHES.map((branch, index) => <button key={branch} id={`pending-tab-${branch}`} role="tab" type="button" aria-selected={pendingTab === branch} aria-controls="pending-table-panel" tabIndex={pendingTab === branch ? 0 : -1} className={pendingTab === branch ? 'is-active' : ''} onClick={() => selectPendingBranch(branch)} onKeyDown={event => { if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) { event.preventDefault(); const next = event.key === 'Home' ? ENGINE_BRANCHES[0] : event.key === 'End' ? ENGINE_BRANCHES[2] : ENGINE_BRANCHES[(index + (event.key === 'ArrowRight' ? 1 : 2)) % 3]; selectPendingBranch(next); document.getElementById(`pending-tab-${next}`)?.focus() } }}>{branch.toUpperCase()} <span>({pendingCounts[index]})</span></button>)}</div>
            <div id="pending-table-panel" role="tabpanel" aria-labelledby={`pending-tab-${pendingTab}`}>
              <h3 className="pn-manual-heading">PENDING · {pendingTab.toUpperCase()} <span className="pc-count">{pendingRows.length} SKUs</span></h3>
              <p className="pn-legend">สาเหตุแสดงที่มาของรายการ · SKU เดียวรวมเป็นหนึ่งแถวต่อสาขา · นำออกจาก PENDING เพื่อยกเลิกพักขายเฉพาะสาขานี้</p>
              {pendingMismatchCount > 0 && <p className="pn-mismatch-note" role="status">≠ {pendingTab.toUpperCase()}: ราคา GRAB ไม่ตรงกับ PROMAXX {pendingMismatchCount} SKU</p>}
              <div className="pn-table-scroll"><table className="pn-table" aria-label={`รายการ PENDING ${pendingTab.toUpperCase()}`}><thead><tr><th>SKU</th><th>ชื่อสินค้า</th><th>ราคา DB / Cloud</th><th>GRAB {pendingTab.toUpperCase()}</th><th>PROMAXX</th><th>สาเหตุ</th><th>จัดการ</th></tr></thead><tbody>
                {pendingRows.map(entry => {
                  const branch = pendingTab.toUpperCase() as 'SRC' | 'SSS' | 'KKL'
                  const check = pendingAutoPrices.get(entry.sku)
                  const automatic = entry.autoPriceBranches?.includes(branch)
                  const manual = pendingManualBranches(entry).includes(branch)
                  const reason = automatic ? `${check?.states[pendingTab] === 'diff' ? 'ราคาไม่ตรง' : 'รอตรวจราคา'}${manual ? ' + เพิ่มเอง' : ''}` : 'เพิ่มเอง'
                  const grab = check?.prices[pendingTab], database = pendingDbPrices.loaded ? pendingDbPrices.prices.get(entry.sku) : undefined, promax = check?.promaxx
                  const promaxDiff = comparePendingPrices(grab, promax) === 'diff'
                  return <tr key={entry.sku} data-pending-sku={entry.sku} className={promaxDiff ? 'pn-row-mismatch' : ''}><td><code>{entry.sku}</code></td><td>{entry.name || check?.name || <span className="pn-no-name">ยังไม่พบชื่อสินค้า</span>}</td><td>{pendingPriceText(pendingDbPrices.loaded, database)}</td><td>{pendingPriceText(pendingAutoStatus.loaded, grab)}{promaxDiff && <span className="pn-difference" aria-label="ราคา GRAB ไม่ตรงกับ PROMAXX">≠</span>}</td><td className={promaxDiff ? 'pn-price-diff' : ''}>{pendingPriceText(pendingAutoStatus.loaded, promax)}</td><td><span className={`pn-reason${automatic ? ' is-automatic' : ''}`}>{reason}</span></td><td><button className="pn-delete" disabled={!pendingCanEdit} aria-label={`นำ ${entry.sku} ออกจาก PENDING เฉพาะ ${pendingTab.toUpperCase()}`} onClick={() => void editPending({ kind: 'resolveBranch', sku: entry.sku, branch: pendingTab, prices: pendingPricePair(check, pendingTab) })}>นำออกจาก PENDING</button></td></tr>
                })}
                {!pendingRows.length && <tr><td colSpan={7}>{pendingLoading ? 'กำลังโหลด PENDING...' : pendingError ? 'ยังอ่านรายการไม่ได้ กรุณาโหลดใหม่' : `ไม่มีรายการ PENDING สำหรับ ${pendingTab.toUpperCase()}`}</td></tr>}
              </tbody></table></div>
            </div>
            <p className="pn-footer-note">จำนวนรวมนับ SKU ที่ยังอยู่ใน PENDING ไม่ซ้ำทุกสาขา รวม WRD/ONN เดิม · รายการที่นำออกจะไม่กลับมาจากคู่ราคาเดิม แม้รีเฟรชหรือ RUN · หากคู่ราคาเปลี่ยนและยังไม่ตรงกัน ระบบจะเพิ่มกลับให้ตรวจใหม่</p>
          </section>

          <section className="pe-card" aria-labelledby="pe-run-heading">
            <div className="pe-card-heading"><div><span className="pc-eyebrow">03 / PROCESS</span><h2 id="pe-run-heading">ตรวจคู่ไฟล์และเติมสต็อก</h2><p>ตรวจโครงสร้างและ StoreID ทุกครั้งก่อนสร้างผลลัพธ์ · เปลี่ยนเฉพาะคอลัมน์ E / F / P</p></div><button className="pc-button" disabled={Boolean(engineBusy) || !Object.keys(engineReports).length} onClick={() => void checkEnginePairs()}>✔ CHECK</button></div>
            {grabFilesError && <div className="pe-error" role="alert">{grabFilesError}<p>ตรวจตาราง grab_files และสิทธิ์อ่านใน Supabase แล้วกด CHECK อีกครั้ง</p></div>}
            {engineChecks && <div className="pe-pairs" role="status">{ENGINE_BRANCHES.filter(branch => engineReports[branch]).map(branch => <div key={branch} className={engineChecks[branch] ? 'pe-pair-ready' : 'pe-pair-missing'}><span>{engineChecks[branch] ? '✔ ถูกต้อง' : '✕ ไม่พบ'}</span><code>{engineReports[branch]!.name} → {templateName(branch)}</code></div>)}<strong>พร้อม {Object.values(engineChecks).filter(Boolean).length} คู่จากทั้งหมด {Object.keys(engineChecks).length} คู่</strong></div>}
            <fieldset className="pe-subtract" disabled={Boolean(engineBusy)}><legend>SUBTRACT PER BRANCH</legend><p>จำนวนที่กันไว้ขายหน้าร้าน หักออกจากสต็อกของแต่ละ SKU ก่อนส่งไป Grab</p><div>{ENGINE_BRANCHES.map(branch => <label key={branch} htmlFor={`subtract-${branch}`}>{branch.toUpperCase()}<input id={`subtract-${branch}`} type="number" min="0" step="any" inputMode="decimal" value={engineSubtract[branch]} onChange={event => { setEngineSubtract(previous => ({ ...previous, [branch]: event.target.value })); invalidateEngineResults() }} /></label>)}</div></fieldset>
            <div className="pe-pending"><strong>PENDING</strong><span>{enginePendingCounts ? ENGINE_BRANCHES.map(branch => `${branch.toUpperCase()} ${enginePendingCounts[branch]}`).join(' · ') : 'อ่านรายการล่าสุดจาก Supabase ทุกครั้งที่ RUN'}</span><p>รายการที่อยู่ใน PENDING ของสาขานั้นจะถูกระงับขาย หากอ่านรายการไม่ได้ ระบบจะหยุด RUN</p></div>
            <div className="pe-actions"><button className="pc-button pc-button-primary" disabled={Boolean(engineBusy) || !Object.keys(engineReports).length} onClick={() => void runProcessEngine()}>{engineBusy === 'run' ? 'กำลังประมวลผล...' : '▶▶ RUN'}</button>{engineOutputs.length > 0 && !engineBusy && <button className="pc-button pe-download" onClick={downloadEngineOutputs}><WorkspaceIcon name="download" />DOWNLOAD ALL ({engineOutputs.length})</button>}<span>{engineOutputs.length ? 'ไฟล์ CSV ทั้งหมดจะรวมใน ZIP เดียว' : 'DOWNLOAD ALL จะปรากฏเมื่อประมวลผลสำเร็จอย่างน้อย 1 สาขา'}</span></div>
            {engineSummary && <div className="pe-summary" role="status"><div><span>PROCESSED</span><strong>{engineSummary.processed}</strong><small>สาขาสำเร็จ</small></div><div><span>UPDATED</span><strong>{engineSummary.updated.toLocaleString('th-TH')}</strong><small>แถวที่เปลี่ยนค่า</small></div><div><span>SKIPPED</span><strong>{engineSummary.skipped}</strong><small>สาขาที่ข้าม</small></div></div>}
            {engineWarnings.length > 0 && <div className="pe-error" role="alert"><strong>กรุณาตรวจ SYSTEM LOG ก่อนดาวน์โหลด</strong>{engineWarnings.map((warning, index) => <p key={index}>{warning}</p>)}</div>}
            {engineOutputs.length > 0 && <ul className="pe-output-list">{engineOutputs.map(output => <li key={output.name}><WorkspaceIcon name="check" /><code>{output.name}</code><span>{output.total.toLocaleString('th-TH')} แถว · SKU ตรง {output.matched.toLocaleString('th-TH')}</span></li>)}</ul>}
          </section>

          <section className="pe-card pe-log-card" aria-labelledby="pe-log-heading"><div className="pe-card-heading"><div><span className="pc-eyebrow">SYSTEM LOG</span><h2 id="pe-log-heading">สถานะการทำงาน</h2></div><button className="pc-button" disabled={Boolean(engineBusy)} onClick={() => { setEngineLog([]); setEngineProgress({ current: 0, total: 0, label: 'รอเริ่มงาน' }) }}>✕ CLEAR</button></div><div className="pe-progress"><label htmlFor="engine-progress" role="status">{engineProgress.label}<span>{engineProgress.current}/{engineProgress.total}</span></label><progress id="engine-progress" value={engineProgress.current} max={engineProgress.total || 1} /></div><div className="pe-log" ref={engineLogRef} role="log" aria-live="polite" aria-relevant="additions"><pre>{engineLog.length ? engineLog.join('\n') : 'พร้อมทำงาน · อัปโหลด R01.102 แล้วกด CHECK'}</pre></div></section>
        </section></div>}


        <dialog className="gc-drawer" ref={grabChangeDialogRef} aria-labelledby="grab-change-title" aria-describedby="grab-change-description" onClose={() => { setGrabChangeOpen(false); grabChangeOpenerRef.current?.focus() }} onCancel={() => setGrabChangeOpen(false)}>
          {grabActiveReview && <div className="gc-shell">
            <header className="gc-header">
              <span className="gc-alert-icon"><WorkspaceIcon name="alert" /></span>
              <div><span className="gc-eyebrow">UPLOAD COMPLETE / DATA CHANGED</span><h2 id="grab-change-title">แจ้งการเปลี่ยนแปลง</h2><p id="grab-change-description">{grabChangedFiles.length} ไฟล์มีข้อมูลต่างจากไฟล์เดิมที่เก็บไว้</p></div>
              <button className="gc-close" aria-label="ปิดแจ้งการเปลี่ยนแปลง" onClick={() => setGrabChangeOpen(false)}><WorkspaceIcon name="close" /></button>
            </header>
            <div className="gc-success"><WorkspaceIcon name="check" />บันทึกไฟล์ที่ผ่านการตรวจสอบเรียบร้อยแล้ว</div>
            <div className="gc-body">
              {grabUploadProblems.length > 0 && <details className="gc-upload-errors"><summary>มี {grabUploadProblems.length} ไฟล์บันทึกไม่สำเร็จ</summary><ul>{grabUploadProblems.map((problem, index) => <li key={index}><pre>{problem}</pre></li>)}</ul></details>}
              {grabChangedFiles.length > 1 && <div className="gc-file-tabs" role="group" aria-label="เลือกไฟล์ที่ข้อมูลเปลี่ยน">{grabChangedFiles.map(review => <button key={review.name} aria-pressed={review.name === grabActiveReview.name} onClick={() => { setGrabChangeBranch(review.branch); setGrabExpandedRow(review.rows.find(row => row.type !== 'unchanged')?.key ?? null); setGrabChangeLimit(50) }}><strong>{review.branch.toUpperCase()}</strong><small>{review.counts.updated + review.counts.added + review.counts.removed} แถวเปลี่ยน</small></button>)}</div>}
              <div className="gc-file-tag"><WorkspaceIcon name="file" /><strong>{grabActiveReview.name}</strong><span>{grabActiveReview.total.toLocaleString('th-TH')} แถว</span></div>
              <p className="gc-source">จาก {grabActiveReview.sourceName}<br />{new Intl.DateTimeFormat('th-TH', { timeZone: 'Asia/Bangkok', dateStyle: 'medium', timeStyle: 'short' }).format(new Date(grabActiveReview.updatedAt))}</p>
              <div className="gc-total"><strong>{grabChangedRows.length.toLocaleString('th-TH')}</strong><span>แถวมีข้อมูลเปลี่ยนแปลง</span></div>
              <div className="gc-stats">{(['updated', 'added', 'removed'] as const).map(type => <div key={type} className={type}><strong>{grabActiveReview.counts[type].toLocaleString('th-TH')}</strong><span>{GRAB_CHANGE_LABELS[type]}</span></div>)}</div>
              <div className="gc-list-heading"><h3>รายการที่เปลี่ยน</h3><span>ค่าเดิม → ค่าใหม่</span></div>
              <div className="gc-rows">{grabChangedRows.slice(0, grabChangeLimit).map((row, index) => <article key={row.key} className="gc-row">
                <button className="gc-row-toggle" aria-expanded={grabExpandedRow === row.key} aria-controls={`grab-change-row-${grabActiveReview.branch}-${index}`} onClick={() => setGrabExpandedRow(current => current === row.key ? null : row.key)}>
                  <span className={`gc-type ${row.type}`}>{GRAB_CHANGE_LABELS[row.type]}</span><span className="gc-product"><strong>SKU {row.sku}</strong><small>{row.name}</small></span><span className="gc-cell-count">{row.changes.length} คอลัมน์</span><WorkspaceIcon name="chevron" className={grabExpandedRow === row.key ? 'is-expanded' : ''} />
                </button>
                {grabExpandedRow === row.key && <div className="gc-row-content" id={`grab-change-row-${grabActiveReview.branch}-${index}`}><p className="gc-row-reference">แถวเดิม {row.before?.row ?? '—'} → แถวใหม่ {row.after?.row ?? '—'}</p><div className="gc-cells">{row.changes.map(change => <div key={change.column}><label>{String.fromCharCode(65 + change.column)} / {change.name}</label><p><span className="gc-before" title="ค่าเดิม">{change.before === undefined ? 'ไม่มีแถวนี้' : change.before === '' ? '(ว่าง)' : change.before}</span><WorkspaceIcon name="arrow" /><strong className="gc-after" title="ค่าใหม่">{change.after === undefined ? 'ไม่มีแถวนี้' : change.after === '' ? '(ว่าง)' : change.after}</strong></p></div>)}</div></div>}
              </article>)}</div>
              {grabChangedRows.length > grabChangeLimit && <button className="gc-more" onClick={() => setGrabChangeLimit(value => value + 50)}>แสดงเพิ่มอีก {Math.min(50, grabChangedRows.length - grabChangeLimit)} แถว · แสดงแล้ว {grabChangeLimit.toLocaleString('th-TH')}/{grabChangedRows.length.toLocaleString('th-TH')}</button>}
              <p className="gc-unchanged">อีก {grabActiveReview.counts.unchanged.toLocaleString('th-TH')} แถวข้อมูลตรงกับไฟล์เดิม · Export รายงานได้ครบทุกแถวที่เปลี่ยน</p>
            </div>
            <footer className="gc-footer"><div><button className="gc-button" onClick={() => downloadGrabCsv(grabActiveReview.name, grabActiveReview.candidate)}><WorkspaceIcon name="download" />Export CSV</button><button className="gc-button" onClick={() => downloadGrabCsv(grabActiveReview.name.replace(/\.CSV$/, '_CHANGES.csv'), grabChangesCsv(grabActiveReview))}><WorkspaceIcon name="file" />Export รายงาน</button></div><button className="gc-button gc-primary" autoFocus onClick={() => setGrabChangeOpen(false)}>รับทราบ</button></footer>
          </div>}
        </dialog>

        <div hidden>
          <input type="file" accept=".csv" id="grabInputSrc" onChange={e => { if (e.target.files?.[0]) { handleGrabCheck(e.target.files[0], 'src'); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="grabInputKkl" onChange={e => { if (e.target.files?.[0]) { handleGrabCheck(e.target.files[0], 'kkl'); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="grabInputSss" onChange={e => { if (e.target.files?.[0]) { handleGrabCheck(e.target.files[0], 'sss'); e.target.value = '' } }} />
          <input type="file" accept=".csv" id="priceCalcInput" disabled={savingNewPromaxx} onChange={e => { if (e.target.files?.[0]) { handlePriceCalcUpload(e.target.files[0]); e.target.value = '' } }} />
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

        {!engineOpen && toolsOpen && <section className="pc-workflow-tools" id="workflow-tools" aria-label="เครื่องมือจัดการสินค้า">
          <div className="pc-tools-heading">
            <h2 className="pc-sr-only">เครื่องมือจัดการสินค้า</h2>
            <label className="pc-tools-search"><WorkspaceIcon name="search" /><input type="search" aria-label="ค้นหาเครื่องมือ" placeholder="ค้นหาเครื่องมือ เช่น นำเข้า, ราคา, GRAB" value={toolSearch} onChange={event => setToolSearch(event.target.value)} onKeyDown={event => { if (event.key === 'Escape') { event.preventDefault(); setToolSearch('') } }} /></label>
            <div className="pc-tools-heading-actions"><span className="pc-tools-count" role="status">{matchingToolCount} รายการ</span><button className="pc-icon-button" aria-label="ปิดเครื่องมือ" onClick={() => setToolsOpen(false)}><WorkspaceIcon name="close" /></button></div>
          </div>
          <div className="pc-tool-groups">
            {filteredToolGroups.map(group => <section key={group.key} className="pc-tool-card" aria-labelledby={'tool-group-' + group.key}>
              <h3 id={'tool-group-' + group.key}><WorkspaceIcon name={group.icon} />{group.title}<span className="pc-tool-card-count">{group.items.length}</span></h3>
              {group.key === 'grab' && <p className="pc-tool-description">นำเข้า = GRAB Menu (C = ราคา, I = SKU)</p>}
              {group.items.map(tool => <button key={tool.name} className="pc-tool-action" disabled={tool.disabled} onClick={tool.action}><span>{tool.busy ? 'กำลังเตรียมไฟล์...' : tool.name}</span><small>{tool.detail}</small><WorkspaceIcon name="arrow" /></button>)}
              {group.key === 'export' && <div className="pc-tools-master"><p className="pc-master-summary">Product Master <b>{masterSkuCount === null ? 'ยังไม่โหลด' : masterSkuCount.toLocaleString('th-TH')}</b>{masterFileName && <small>{masterFileName}</small>}</p><p className="pc-master-summary">SKU ที่ยังขาด <b>{masterMissingCount === null ? 'ยังไม่ตรวจ' : masterMissingCount.toLocaleString('th-TH')}</b></p></div>}
            </section>)}
          </div>
          {matchingToolCount === 0 && <div className="pc-tools-empty" role="status"><WorkspaceIcon name="search" /><h3>ไม่พบเครื่องมือที่ตรงกับคำค้นหา</h3><p>ลองค้นหาด้วยชื่อเครื่องมือ รหัสรายงาน หรือสาขา</p><button className="pc-button" onClick={() => setToolSearch('')}>แสดงเครื่องมือทั้งหมด</button></div>}
        <section className="pc-tool-card pe-manager pc-tools-manager" aria-label="GRAB CSV MANAGER" hidden={!showGrabManager}>
          <h3 className="pe-manager-heading"><button className="pe-manager-toggle" aria-expanded={grabManagerOpen} aria-controls="pe-manager-content" onClick={() => setGrabManagerOpen(value => !value)}><WorkspaceIcon name="file" /><span>GRAB CSV MANAGER</span><small className="pc-tool-card-count">{grabFiles.length} ไฟล์</small><WorkspaceIcon name="chevron" className={grabManagerOpen ? 'is-expanded' : ''} /></button></h3>
          <p className="pe-manager-description">Template ที่ Grab อนุมัติแล้ว</p>
          {grabManagerOpen && <div id="pe-manager-content" className="pe-manager-content"><p>อัปโหลดเฉพาะเมื่อ Grab อนุมัติการเปลี่ยนรายการสินค้า เลือกไฟล์ CSV ชื่อใดก็ได้ ไม่จำกัดชื่อหรือนามสกุล ระบบตรวจเนื้อหาไฟล์และแยกสาขาจาก StoreID ในคอลัมน์ A และบันทึกเป็น GRAB_SRC.CSV, GRAB_SSS.CSV หรือ GRAB_KKL.CSV</p><p>รองรับ UTF-8 / TIS-620 · ชื่อและลำดับคอลัมน์ A–P ต้องตรง Template สต็อกทุกคอลัมน์ · ทุกแถวต้องมี 16 คอลัมน์และ StoreID ของสาขาเดียวกัน</p><div className="pe-actions"><button className="pc-button pc-button-primary" disabled={Boolean(engineBusy)} onClick={event => { grabChangeOpenerRef.current = event.currentTarget; document.getElementById('engineTemplateInput')?.click() }}><WorkspaceIcon name="upload" />UPLOAD CSV</button><button className="pc-button" disabled={Boolean(engineBusy) || grabFilesLoading} onClick={() => void refreshGrabFiles()}><WorkspaceIcon name="refresh" />รีเฟรชรายการ</button><button className="pc-button" disabled={Boolean(engineBusy) || !grabChangedFiles.length} onClick={event => { grabChangeOpenerRef.current = event.currentTarget; openGrabChangeNotice() }}><WorkspaceIcon name="clock" />ดูผลอัปโหลดล่าสุด</button><input hidden id="engineTemplateInput" type="file" multiple disabled={Boolean(engineBusy)} onChange={event => { const files = Array.from(event.target.files ?? []); event.target.value = ''; void uploadGrabTemplates(files) }} /></div>{grabUploadNotice && <p className="gc-upload-notice" role="status">{grabUploadNotice}</p>}{grabExportError && <p className="pe-error" role="alert">{grabExportError}</p>}{grabUploadProblems.length > 0 && <div className="pe-error pe-upload-warning" role="alert"><strong>ตรวจพบปัญหาในไฟล์ CSV {grabUploadProblems.length} ไฟล์</strong><p>ไฟล์ที่ไม่ผ่านจะไม่ถูกบันทึก กรุณาตรวจคอลัมน์ตามรายละเอียดด้านล่างแล้วอัปโหลดใหม่</p><ul>{grabUploadProblems.map((problem, index) => <li key={index}><pre>{problem}</pre></li>)}</ul></div>}{grabAdminUser ? <div className="pe-admin-session"><span>เข้าสู่ระบบแล้ว: {grabAdminUser}</span><button className="pc-text-button" disabled={Boolean(engineBusy)} onClick={() => void signOutGrabManager()}>ออกจากระบบ</button></div> : <form className="pe-admin-login" onSubmit={signInGrabManager}><h3>เข้าสู่ระบบเพื่อบันทึก Template</h3><p>ใช้บัญชีที่มีสิทธิ์เขียนข้อมูล หากตั้ง policy เป็น authenticated ต้องเข้าสู่ระบบก่อนอัปโหลด</p><div><label htmlFor="pe-admin-email">อีเมล<input id="pe-admin-email" type="email" required autoComplete="username" disabled={Boolean(engineBusy)} value={grabAdminEmail} onChange={event => setGrabAdminEmail(event.target.value)} /></label><label htmlFor="pe-admin-password">รหัสผ่านบัญชี<input id="pe-admin-password" type="password" required autoComplete="current-password" disabled={Boolean(engineBusy)} value={grabAdminPassword} onChange={event => setGrabAdminPassword(event.target.value)} /></label><button className="pc-button" disabled={Boolean(engineBusy)} type="submit">เข้าสู่ระบบ</button></div></form>}{grabAdminError && <p className="pe-error" role="alert">{grabAdminError}</p>}{grabFilesError && <p className="pe-error" role="alert">{grabFilesError}</p>}<div className="pe-table-scroll"><table className="pe-files-table"><thead><tr><th>FILE NAME</th><th>UPLOADED AT (เวลาไทย)</th><th>EXPORT</th></tr></thead><tbody>{grabFilesLoading ? <tr><td colSpan={3}>กำลังโหลดรายการ...</td></tr> : grabFiles.length ? grabFiles.map(file => <tr key={file.name}><td><code>{file.name}</code></td><td>{Number.isNaN(Date.parse(file.updated_at)) ? '—' : new Intl.DateTimeFormat('th-TH', { timeZone: 'Asia/Bangkok', dateStyle: 'medium', timeStyle: 'short' }).format(new Date(file.updated_at))}</td><td><button className="pc-button" disabled={Boolean(engineBusy)} aria-label={`Export ${file.name}`} onClick={() => void exportStoredGrabTemplate(file.name)}><WorkspaceIcon name="download" />Export CSV</button></td></tr>) : <tr><td colSpan={3}>{grabFilesError ? 'ยังอ่านรายการไม่ได้' : 'ยังไม่มี Template ในระบบ'}</td></tr>}</tbody></table></div></div>}
        </section>
          <div className="pc-tools-bottom"><div className="pc-tools-category"><label htmlFor="tools-category">หมวดหมู่สำหรับส่งออกหรือจัดการ</label><select id="tools-category" value={selectedSheet} disabled={loading || savingPrice} onChange={event => { setSelectedSheet(event.target.value); setPageNumber(1) }}><option value="all">เลือกหมวดหมู่</option>{stats.sheets.map(sheet => <option key={sheet.name} value={sheet.name}>{sheet.name} ({sheet.count})</option>)}</select></div><button className="pc-danger-link" disabled={selectedSheet === 'all'} onClick={deleteSelectedCategory}>ลบหมวดหมู่ที่เลือกทุกสาขา</button></div>
        </section>}

        {catalogOpen && <section className="pc-catalog" aria-label="รายการสินค้าแบบการ์ด" hidden={engineOpen || toolsOpen}>
          <h2 className="pc-sr-only">Image Catalogue · รายการสินค้า</h2>
          <div className="pc-catalog-toolbar">
            <label className="pc-search pc-catalog-search"><WorkspaceIcon name="search" /><input ref={searchInputRef} type="search" aria-label="ค้นหาสินค้า" placeholder="ค้นหาชื่อสินค้า หรือรหัส SKU..." value={search} disabled={savingPrice} onChange={event => { setSearch(event.target.value); setPageNumber(1) }} /><kbd>⌘ K</kbd></label>
            <select className="pc-catalog-category" aria-label="หมวดหมู่สินค้า" value={selectedSheet} disabled={savingPrice} onChange={event => { setSelectedSheet(event.target.value); setPageNumber(1) }}><option value="all">ทุกหมวดหมู่</option>{stats.sheets.map(sheet => <option key={sheet.name} value={sheet.name}>{sheet.name} ({sheet.count})</option>)}</select>
            <button className="pc-icon-button" title="รีเฟรชสินค้า" aria-label="รีเฟรชสินค้า" disabled={loading || savingPrice} onClick={() => search.trim().length >= 3 ? searchProducts(search.trim()) : loadProducts(selectedSheet)}><WorkspaceIcon name="refresh" /></button>
          </div>
          <div className="pc-catalog-summary">
            <label><input type="checkbox" aria-label="เลือกสินค้าทั้งหมดในหน้านี้" disabled={loading || !pageProducts.length} checked={pageProducts.length > 0 && pageProducts.every(product => selectedSkus.has(product['รหัสสินค้า (SKU NUMBER)']))} ref={input => { if (input) input.indeterminate = pageProducts.some(product => selectedSkus.has(product['รหัสสินค้า (SKU NUMBER)'])) && !pageProducts.every(product => selectedSkus.has(product['รหัสสินค้า (SKU NUMBER)'])) }} onChange={event => setSelectedSkus(previous => { const next = new Set(previous); pageProducts.forEach(product => event.target.checked ? next.add(product['รหัสสินค้า (SKU NUMBER)']) : next.delete(product['รหัสสินค้า (SKU NUMBER)'])); return next })} />เลือกสินค้าในหน้านี้</label>
            <span>{filteredProducts.length.toLocaleString('th-TH')} รายการ</span>
          </div>
          {loadError && <div className="pc-load-error" role="alert"><WorkspaceIcon name="refresh" /><div><strong>โหลดสินค้าไม่สำเร็จ</strong><p>{loadError}</p></div><button className="pc-button" onClick={() => search.trim().length >= 3 ? searchProducts(search.trim()) : loadProducts(selectedSheet)}>ลองใหม่</button></div>}
          {loading ? <div className="pc-catalog-grid" role="status" aria-label="กำลังโหลดสินค้า">{Array.from({ length: 8 }, (_, index) => <div className="pc-catalog-skeleton" aria-hidden="true" key={index}><i /><span /><span /><b /></div>)}</div> : !loadError && !filteredProducts.length ? <div className="pc-empty"><WorkspaceIcon name="search" /><h3>{search || selectedSheet !== 'all' ? 'ไม่พบสินค้าที่ตรงกับเงื่อนไข' : 'ยังไม่มีรายการสินค้า'}</h3><p>{search || selectedSheet !== 'all' ? 'ลองเปลี่ยนคำค้นหา หรือหมวดหมู่สินค้า' : 'เริ่มต้นด้วยการนำเข้าไฟล์ Excel หรือเพิ่มสินค้าใหม่'}</p>{(search || selectedSheet !== 'all') && <button className="pc-button" onClick={() => { setSearch(''); setSelectedSheet('all'); setPageNumber(1) }}>แสดงสินค้าทั้งหมด</button>}</div> : !loadError && <div className="pc-catalog-grid" role="list" aria-label="สินค้า">
            {pageProducts.map(product => {
              const sku = String(product['รหัสสินค้า (SKU NUMBER)'])
              const name = String(product['*ชื่อสินค้า (NAME)'] || 'ยังไม่ระบุชื่อสินค้า')
              const category = String(product['หมวดหมู่สินค้า (CATEGORIES)'] || 'ยังไม่ระบุหมวดหมู่')
              const imageSource = product['*รูปภาพสินค้า'] || ''
              const image = convertDriveLink(imageSource)
              return <article className={`pc-catalog-card${selectedSkus.has(sku) ? ' is-selected' : ''}`} key={sku} role="listitem" data-sku={sku} aria-label={`ข้อมูลสินค้า ${name}`}>
                <div className="pc-catalog-media">
                  <span className="pc-catalog-placeholder" aria-hidden="true"><WorkspaceIcon name="box" /><span>รูปสินค้า</span></span>
                  {image && <a className="pc-catalog-image-link" href={imageSource} target="_blank" rel="noopener noreferrer" aria-label={`เปิดรูปสินค้า ${name} ในแท็บใหม่`} title="เปิดรูปภาพในแท็บใหม่"><NextImage key={image} className="pc-catalog-photo" src={image} alt={name} fill sizes="(max-width: 619px) 90vw, (max-width: 900px) 44vw, (max-width: 1200px) 27vw, 22vw" unoptimized onError={event => { event.currentTarget.hidden = true }} /></a>}
                  <label className="pc-catalog-selection"><input type="checkbox" aria-label={`เลือก ${name}`} checked={selectedSkus.has(sku)} onChange={event => setSelectedSkus(previous => { const next = new Set(previous); if (event.target.checked) next.add(sku); else next.delete(sku); return next })} /></label>
                </div>
                <div className="pc-catalog-body"><span className="pc-catalog-category-tag" title={category}>{category}</span><h3>{name}</h3><p className="pc-catalog-sku"><span>SKU</span> {sku}</p><div className="pc-catalog-price"><span>ราคา DB</span><strong>{formatMoney(product['*ราคาสินค้า'])}</strong></div></div>
                <footer className="pc-catalog-card-footer"><button className="pc-catalog-edit" aria-label={`แก้ไขสินค้า ${name}`} disabled={savingPrice} onClick={() => { setIsAdding(false); setEditProduct({ ...product, _originalSku: sku }) }}>แก้ไขสินค้า<WorkspaceIcon name="chevron" /></button></footer>
              </article>
            })}
          </div>}
          <div className="pc-catalog-footer"><span>{selectedSkus.size ? <button className="pc-selected-export" onClick={exportSelectedSkusXlsx}><WorkspaceIcon name="download" />ส่งออกที่เลือก ({selectedSkus.size.toLocaleString('th-TH')})</button> : `${filteredProducts.length ? (currentPage - 1) * pageSize + 1 : 0}–${Math.min(currentPage * pageSize, filteredProducts.length)} จาก ${filteredProducts.length.toLocaleString('th-TH')} รายการที่โหลด`}</span><div className="pc-pagination"><button aria-label="หน้าก่อนหน้า" disabled={currentPage === 1 || loading} onClick={() => setPageNumber(currentPage - 1)}>‹</button><span><b>{currentPage}</b> / {pageCount}</span><button aria-label="หน้าถัดไป" disabled={currentPage === pageCount || loading} onClick={() => setPageNumber(currentPage + 1)}>›</button></div></div>
        </section>}

        {!catalogOpen && <div className="pc-inspection-grid" hidden={engineOpen || toolsOpen}>
          <section className="pc-list-column" aria-label="รายการสินค้า">
            <div className="pc-search-toolbar"><label className="pc-search"><WorkspaceIcon name="search" /><input ref={searchInputRef} type="search" aria-label="ค้นหาสินค้า" placeholder="ค้นหาชื่อสินค้า หรือรหัส SKU..." value={search} disabled={savingPrice} onChange={event => { setSearch(event.target.value); setPageNumber(1); resetPriceReview() }} /><kbd>⌘ K</kbd></label><button className={`pc-button ${filtersOpen ? 'is-selected' : ''}`} aria-expanded={filtersOpen} aria-controls="product-filters" onClick={() => setFiltersOpen(value => !value)}><WorkspaceIcon name="filter" />ตัวกรอง{(branchFilter !== 'all' || selectedSheet !== 'all') && <span className="pc-count">{Number(branchFilter !== 'all') + Number(selectedSheet !== 'all')}</span>}</button><button className="pc-icon-button pc-refresh" title="รีเฟรชสินค้า" aria-label="รีเฟรชสินค้า" disabled={loading || savingPrice} onClick={() => search.trim().length >= 3 ? searchProducts(search.trim()) : loadProducts(selectedSheet)}><WorkspaceIcon name="refresh" /></button></div>

            {filtersOpen && <div className="pc-filters" id="product-filters"><label>หมวดหมู่<select aria-label="กรองหมวดหมู่" value={selectedSheet} disabled={savingPrice} onChange={event => { setSelectedSheet(event.target.value); setPageNumber(1); resetPriceReview() }}><option value="all">ทุกหมวดหมู่ ({stats.sheets.length})</option>{stats.sheets.map(sheet => <option key={sheet.name} value={sheet.name}>{sheet.name} ({sheet.count})</option>)}</select></label><label>สินค้าที่มีในสาขา<select aria-label="กรองสาขา" value={branchFilter} disabled={savingPrice} onChange={event => { setBranchFilter(event.target.value as 'all' | BranchKey); setPageNumber(1); resetPriceReview() }}><option value="all">ทุกสาขา</option>{branches.map(branch => <option key={branch} value={branch}>{branchNames[branch]}</option>)}</select></label><button className="pc-text-button" onClick={() => { setSelectedSheet('all'); setBranchFilter('all'); setPriceFilter('all'); setSearch(''); setPageNumber(1); resetPriceReview() }}>ล้างตัวกรอง</button></div>}

            <div className="pc-table-card">
              <div className="pc-table-heading"><div><h2>{selectedSheet === 'all' ? 'สินค้าทั้งหมด' : selectedSheet}<span className="pc-count">{filteredProducts.length.toLocaleString('th-TH')}</span></h2><p>เทียบราคา GRAB <b>{branchNames[inspectionBranch]}</b> กับ PROMAXX · ราคา DB / Cloud ใช้ประกอบการตรวจสอบ</p></div><div className="pc-price-tabs" aria-label="กรองสถานะราคา"><button className={priceFilter === 'all' ? 'is-active' : ''} aria-pressed={priceFilter === 'all'} onClick={() => { setPriceFilter('all'); setPageNumber(1); resetPriceReview() }}>ทั้งหมด</button><button className={priceFilter === 'mismatch' ? 'is-active' : ''} aria-pressed={priceFilter === 'mismatch'} onClick={() => { setPriceFilter('mismatch'); setPageNumber(1); resetPriceReview() }}>ต้องตรวจ <b>{issueCount}</b></button></div></div>
              <div className="pc-table-options"><div className="pc-branch-switch" aria-label="เลือกสาขาเปรียบเทียบ">{branches.map(branch => <button key={branch} className={inspectionBranch === branch ? 'is-active' : ''} aria-pressed={inspectionBranch === branch} disabled={savingPrice} onClick={() => { setInspectionBranch(branch); setPageNumber(1); resetPriceReview() }}>{branchNames[branch]}</button>)}</div><span className="pc-price-view"><WorkspaceIcon name="grid" />ดูราคาทุกระบบ</span></div>

              {loadError && <div className="pc-load-error" role="alert"><WorkspaceIcon name="clock" /><div><strong>โหลดข้อมูลสินค้าไม่สำเร็จ</strong><p>{loadError}</p></div><button className="pc-button" onClick={() => search.trim().length >= 3 ? searchProducts(search.trim()) : loadProducts(selectedSheet)}>ลองอีกครั้ง</button></div>}
              {loading ? <div className="pc-loading" role="status" aria-label="กำลังโหลดสินค้า"><span className="pc-loading-label"><i />กำลังโหลดข้อมูลสินค้า...</span>{Array.from({ length: 6 }, (_, index) => <div className="pc-skeleton-row" key={index}><i /><span /><span /><b /></div>)}</div> : !loadError && filteredProducts.length === 0 ? <div className="pc-empty"><WorkspaceIcon name="search" /><h3>{search || priceFilter !== 'all' || selectedSheet !== 'all' || branchFilter !== 'all' ? 'ไม่พบสินค้าที่ตรงกับเงื่อนไข' : 'ยังไม่มีรายการสินค้า'}</h3><p>{search || priceFilter !== 'all' || selectedSheet !== 'all' || branchFilter !== 'all' ? 'ลองเปลี่ยนคำค้นหา สาขา หรือสถานะราคา' : 'เริ่มต้นด้วยการนำเข้าไฟล์ Excel หรือเพิ่มสินค้าใหม่'}</p>{search || priceFilter !== 'all' || selectedSheet !== 'all' || branchFilter !== 'all' ? <button className="pc-button" onClick={() => { setSearch(''); setPriceFilter('all'); setSelectedSheet('all'); setBranchFilter('all'); setPageNumber(1) }}>แสดงสินค้าทั้งหมด</button> : <button className="pc-button pc-button-primary" onClick={() => document.getElementById('importInput')?.click()}><WorkspaceIcon name="upload" />นำเข้าสินค้า</button>}</div> : !loadError && <div className="pc-table-scroll"><table className="pc-products-table pc-expanded-prices"><thead><tr><th className="pc-checkbox-cell"><input type="checkbox" aria-label="เลือกสินค้าทั้งหมดในหน้านี้" checked={pageProducts.length > 0 && pageProducts.every(product => selectedSkus.has(product['รหัสสินค้า (SKU NUMBER)']))} ref={input => { if (input) input.indeterminate = pageProducts.some(product => selectedSkus.has(product['รหัสสินค้า (SKU NUMBER)'])) && !pageProducts.every(product => selectedSkus.has(product['รหัสสินค้า (SKU NUMBER)'])) }} onChange={event => setSelectedSkus(previous => { const next = new Set(previous); pageProducts.forEach(product => event.target.checked ? next.add(product['รหัสสินค้า (SKU NUMBER)']) : next.delete(product['รหัสสินค้า (SKU NUMBER)'])); return next })} /></th><th className="pc-product-column">สินค้า / SKU</th><th scope="col" className="pc-pending-cell" title="✓ อยู่ใน PENDING ครบ 3 สาขา · ✕ ยังไม่ครบ · คลิกดูรายละเอียด">Pending</th><th>ราคา DB / Cloud</th><th>PROMAXX</th><th>New Promaxx</th>{branches.map(branch => <th key={branch}>GRAB {branchNames[branch]}</th>)}<th>สถานะ</th><th><span className="pc-sr-only">รายละเอียด</span></th></tr></thead><tbody>{pageProducts.map(product => {
                const sku = String(product['รหัสสินค้า (SKU NUMBER)'])
                const state = priceState(product)
                const coverage = productPendingReady && productPendingSets ? pendingBranchCoverage(sku, productPendingSets) : null
                const pendingDescription = coverage ? coverage.complete ? 'อยู่ใน PENDING ครบ 3 สาขา' : `ยังไม่อยู่ใน PENDING: ${coverage.missing.map(branch => branch.toUpperCase()).join(', ')}` : productPendingError ? 'อ่าน PENDING ไม่สำเร็จ' : 'กำลังโหลด PENDING'
                const image = convertDriveLink(product['*รูปภาพสินค้า'] || '')
                return <tr key={sku} data-sku={sku} className={`${sku === inspectedSku ? 'pc-row-active' : ''} ${selectedSkus.has(sku) ? 'pc-row-checked' : ''}`} onClick={event => { if (!(event.target as HTMLElement).closest('input,button,a')) selectInspectorProduct(product) }} onMouseEnter={event => { if (image && window.matchMedia('(hover: hover) and (pointer: fine) and (min-width: 821px)').matches) setTooltip({ x: event.clientX, y: event.clientY, url: image }) }} onMouseMove={event => setTooltip(previous => previous ? { ...previous, x: event.clientX, y: event.clientY } : null)} onMouseLeave={() => setTooltip(null)}>
                  <td className="pc-checkbox-cell"><input type="checkbox" aria-label={`เลือก ${product['*ชื่อสินค้า (NAME)'] || sku}`} checked={selectedSkus.has(sku)} onChange={event => setSelectedSkus(previous => { const next = new Set(previous); event.target.checked ? next.add(sku) : next.delete(sku); return next })} /></td>
                  <td><button className="pc-product-link" aria-label={`ตรวจสินค้า ${product['*ชื่อสินค้า (NAME)'] || sku}`} disabled={savingPrice} onClick={() => selectInspectorProduct(product)}><span className="pc-product-thumb"><WorkspaceIcon name="box" />{image && <NextImage src={image} alt="" fill sizes="33px" unoptimized loading="lazy" onError={event => { event.currentTarget.hidden = true }} />}</span><span><strong>{product['*ชื่อสินค้า (NAME)'] || 'ยังไม่ระบุชื่อสินค้า'}</strong><small>{sku}</small><span className="pc-row-branches">{branches.filter(branch => hasProductBranch(product, branch)).map(branch => <i key={branch}>{branchNames[branch]}</i>)}</span></span></button></td>
                  <td className="pc-pending-cell"><button className={`pc-pending-status ${coverage ? coverage.complete ? 'is-complete' : 'is-incomplete' : 'is-unavailable'}`} aria-label={`ดู Pending ของ SKU ${sku}: ${pendingDescription}`} title={pendingDescription} aria-haspopup="dialog" aria-expanded={pendingProduct?.sku === sku} onClick={() => { setTooltip(null); setPendingProduct({ sku, name: String(product['*ชื่อสินค้า (NAME)'] || '') }) }}><WorkspaceIcon name={coverage ? coverage.complete ? 'check' : 'close' : productPendingError ? 'alert' : 'clock'} /><small>{coverage ? `${coverage.present.length}/3` : '—'}</small></button></td>
                  <td className="pc-table-price">{formatMoney(product['*ราคาสินค้า'])}</td>
                  <td>{formatMoney(product.promaxx_price)}</td><td>{formatMoney(product.new_promaxx_price)}</td>{branches.map(branch => <td key={branch} className={hasProductBranch(product, branch) && pricesDiffer(product[`${branch}_price`], product.promaxx_price) ? 'pc-price-mismatch' : ''}>{hasProductBranch(product, branch) ? formatMoney(product[`${branch}_price`]) : '—'}</td>)}
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
              <div className="pc-comparison-prices"><div><small>PROMAXX · ราคาขายอ้างอิง</small><strong>{formatMoney(referencePrice)}</strong></div><div><small>ราคา GRAB {branchNames[inspectionBranch]}</small><strong>{formatMoney(comparisonPrice)}</strong></div></div>
              <div className={`pc-price-difference ${priceDifference !== null && !pricesDiffer(comparisonPrice, referencePrice) ? 'is-matched' : ''}`}><WorkspaceIcon name="tag" /><span>{priceDifference === null ? 'ยังเปรียบเทียบราคาไม่ได้' : 'GRAB − PROMAXX'}</span><strong>{priceDifference === null ? '—' : `${priceDifference > 0 ? '+' : priceDifference < 0 ? '−' : ''} ${formatMoney(Math.abs(priceDifference))}`}</strong></div>
              <div className="pc-inspector-field"><label htmlFor="proposed-price">ราคา DB / Cloud ที่เสนอ <span>THB</span></label><input id="proposed-price" type="text" inputMode="decimal" autoComplete="off" aria-describedby="price-help inspector-error" value={draftPrice} disabled={savingPrice || loading} onChange={event => { if (!inspectedSku) return; setPriceDrafts(previous => ({ ...previous, [inspectedSku]: event.target.value })); resetPriceReview() }} /><div id="price-help" className="pc-field-help"><span>{hasPriceChange ? 'มีราคาที่ยังไม่ได้บันทึก' : 'ปรับราคาบน Cloud แล้วตรวจทานก่อนบันทึก'}</span>{hasPriceChange && <button disabled={savingPrice} onClick={() => { setPriceDrafts(previous => { const next = { ...previous }; if (inspectedSku) delete next[inspectedSku]; return next }); resetPriceReview() }}>คืนค่าเดิม</button>}</div></div>
              <div className="pc-inspector-branches"><span>สาขาที่จำหน่าย</span><div>{branches.map(branch => <span className={hasProductBranch(inspectedProduct, branch) ? 'is-available' : ''} key={branch}><WorkspaceIcon name={hasProductBranch(inspectedProduct, branch) ? 'check' : 'close'} />{branchNames[branch]}</span>)}</div></div>
              <div className="pc-reference-price"><span>ราคา DB / Cloud ปัจจุบัน</span><strong>{formatMoney(originalPrice)}</strong></div>
              <div className="pc-reference-price"><span>เลขที่ใบอนุญาต</span><strong>{inspectedProduct['*เลขที่ใบอนุญาตโฆษณา'] || 'ยังไม่ระบุ'}</strong></div>
              {inspectedProduct.correctPrice !== undefined && <div className="pc-reference-price"><span>ราคาอ้างอิง</span><strong>{formatMoney(inspectedProduct.correctPrice)}</strong></div>}
              {inspectorError && <p id="inspector-error" className="pc-inspector-error" role="alert">{inspectorError}</p>}
              {activeReview ? <section className="pc-review" aria-label="ตรวจทานราคาก่อนบันทึก"><h3><WorkspaceIcon name="check" />ตรวจทานก่อนบันทึก</h3><p>ปรับราคา DB / Cloud ของสินค้านี้</p><div><span>{formatMoney(activeReview.previous)}</span><WorkspaceIcon name="arrow" /><strong>{formatMoney(activeReview.price)}</strong></div><button className="pc-button pc-button-primary" disabled={savingPrice} onClick={saveInspectorPrice}>{savingPrice ? 'กำลังบันทึก...' : 'ยืนยันบันทึกราคา DB / Cloud'}<WorkspaceIcon name="check" /></button><button className="pc-text-button" disabled={savingPrice} onClick={resetPriceReview}>กลับไปแก้ไข</button></section> : <button className="pc-button pc-button-primary pc-review-button" disabled={!hasPriceChange || savingPrice || loading} onClick={reviewInspectorPrice}>ตรวจทานราคา DB / Cloud<WorkspaceIcon name="arrow" /></button>}
              <button className="pc-edit-details" disabled={savingPrice} onClick={() => { setIsAdding(false); setEditProduct({ ...inspectedProduct, _originalSku: inspectedSku }) }}>แก้ไขรายละเอียดและรูปสินค้า<WorkspaceIcon name="chevron" /></button>
            </> : <div className="pc-inspector-empty"><WorkspaceIcon name="box" /><h2>เลือกสินค้าที่ต้องการตรวจ</h2><p>เลือกสินค้าจากตารางเพื่อดูราคา สาขา และปรับข้อมูลในพื้นที่นี้</p><div className="pc-empty-guide"><span><b>01</b>เลือกสินค้าจากรายการ</span><span><b>02</b>ตรวจราคาและผลต่าง</span><span><b>03</b>ตรวจทานก่อนบันทึก</span></div></div>}
          </aside>
        </div>}
        <footer className="pc-statusbar"><span role="status" aria-live="polite"><i className={loadError || inspectorError ? 'has-error' : ''} />{status}</span><span>PRICE CHECKER <b>/</b> {engineOpen ? 'PROCESS ENGINE' : toolsOpen ? 'TOOLS' : catalogOpen ? 'IMAGE CATALOGUE' : 'SPLIT INSPECTOR'}</span></footer>
      </main>

      {pendingPromaxxReview && <StudioModal label="ตรวจทานราคา GRAB ตาม PROMAXX" width={1000} onClose={closePendingPromaxxReview}>
        <div className="pc-modal-panel" aria-busy={pendingPromaxxReview.loading}>
          <header className="pc-modal-header"><div><h2><WorkspaceIcon name="tag" />ตรวจทานราคา GRAB ตาม PROMAXX</h2><span>GM MME · สาขา {pendingPromaxxReview.branch.toUpperCase()} · ราคาเดิม → ราคา PROMAXX</span></div><button className="pc-modal-close" aria-label="ปิด" onClick={closePendingPromaxxReview}>×</button></header>
          <div className="pc-modal-body pn-export-body">
            <p>ไฟล์นี้ใช้ราคา PROMAXX เพื่อปรับราคาขายบน Grab หลังได้รับอนุมัติ ให้นำเข้าไฟล์ราคา GRAB ล่าสุดอีกครั้งเพื่อยืนยันผล</p>
            {pendingPromaxxReview.error && <p className="pe-error" role="alert">{pendingPromaxxReview.error}</p>}
            {pendingPromaxxReview.loading ? <p className="pn-export-loading" role="status">กำลังตรวจ PENDING และราคาล่าสุด...</p> : <>
              <div className="pn-export-summary" role="status"><strong>พร้อมปรับราคา {pendingPromaxxReview.rows.filter(row => row.ready).length} รายการ</strong><span>ไม่ส่งออก {pendingPromaxxReview.rows.filter(row => !row.ready).length} รายการ · ดูเหตุผลในตาราง</span></div>
              <p className="pn-export-scroll-hint">เลื่อนตารางด้านข้างเพื่อดูราคาและผลตรวจสอบ</p>
              <div className="pn-table-scroll"><table className="pn-table" aria-label="ตรวจทานราคาเดิมและราคาใหม่"><thead><tr><th>SKU</th><th>ชื่อสินค้า</th><th>GRAB เดิม</th><th>PROMAXX ใหม่</th><th>ผลตรวจสอบ</th></tr></thead><tbody>
                {pendingPromaxxReview.rows.map(row => <tr key={row.sku} className={row.ready ? '' : 'pn-export-skipped'}><td><code>{row.sku}</code></td><td>{row.product?.['*ชื่อสินค้า (NAME)'] || '—'}</td><td>{formatMoney(row.grab)}</td><td>{formatMoney(row.promaxx)}</td><td>{row.ready ? 'พร้อมปรับราคา' : row.reason}</td></tr>)}
                {!pendingPromaxxReview.rows.length && <tr><td colSpan={5}>{pendingPromaxxReview.error ? 'ตรวจสอบข้อความด้านบน แล้วกดโหลดข้อมูลใหม่' : 'ไม่มี PENDING ที่บันทึกไว้สำหรับสาขานี้'}</td></tr>}
              </tbody></table></div>
            </>}
          </div>
          <footer className="pc-modal-footer pn-export-footer"><button className="pc-button" onClick={closePendingPromaxxReview}>ยกเลิก</button><button className="pc-button" disabled={Boolean(engineBusy)} onClick={() => void reviewPendingPromaxxExport(pendingPromaxxReview.branch)}><WorkspaceIcon name="refresh" />โหลดข้อมูลใหม่</button><button className="pc-button pc-button-primary" disabled={Boolean(engineBusy) || pendingPromaxxReview.loading || !pendingPromaxxReview.rows.some(row => row.ready)} onClick={() => void exportPendingGmMmeXlsx(pendingPromaxxReview.branch, pendingPromaxxReview)}><WorkspaceIcon name="download" />{pendingPromaxxReview.loading ? 'กำลังตรวจข้อมูล...' : `Export PROMAXX (${pendingPromaxxReview.rows.filter(row => row.ready).length} รายการ)`}</button></footer>
        </div>
      </StudioModal>}

      {pendingProduct && <StudioModal label="สถานะ Pending แยกสาขา" width={440} onClose={() => setPendingProduct(null)}>
        <div className="pc-modal-panel">
          <header className="pc-modal-header"><h2><WorkspaceIcon name="branch" />Pending แยกสาขา</h2><button className="pc-modal-close" aria-label="ปิดสถานะ Pending" onClick={() => setPendingProduct(null)}><WorkspaceIcon name="close" /></button></header>
          <div className="pc-modal-body pc-pending-details">
            <div className="pc-pending-product"><code>{pendingProduct.sku}</code><strong>{pendingProduct.name || 'ยังไม่ระบุชื่อสินค้า'}</strong></div>
            {productPendingReady && pendingProductCoverage ? <>
              <p className={`pc-pending-summary ${pendingProductCoverage.complete ? 'is-complete' : 'is-incomplete'}`}>{pendingProductCoverage.complete ? 'อยู่ใน PENDING ครบทั้ง 3 สาขา' : `ยังไม่อยู่ใน PENDING: ${pendingProductCoverage.missing.map(branch => branch.toUpperCase()).join(', ')}`}</p>
              <ul className="pc-pending-branches" aria-label="สถานะ PENDING ของแต่ละสาขา">{ENGINE_BRANCHES.map(branch => {
                const present = pendingProductCoverage.present.includes(branch)
                return <li key={branch} className={present ? 'is-pending' : 'is-missing'}><WorkspaceIcon name={present ? 'check' : 'close'} /><strong>{branch.toUpperCase()}</strong><span>{present ? 'อยู่ใน PENDING' : 'ยังไม่อยู่ใน PENDING'}</span></li>
              })}</ul>
              <p className="pc-pending-note">รวมรายการที่เพิ่มเองและรายการอัตโนมัติ เฉพาะ SRC, KKL และ SSS</p>
              {pendingAutoStatus.loaded && !pendingAutoStatus.saved && <p className="pc-pending-note">รวมผลตรวจอัตโนมัติในหน้านี้ที่ยังไม่ได้บันทึก เข้าสู่ระบบเพื่อบันทึก PENDING</p>}
            </> : productPendingError ? <p className="pe-error" role="alert">{productPendingError}</p> : <p className="pc-pending-note" role="status">กำลังโหลด PENDING SKU UPDATE...</p>}
          </div>
          <footer className="pc-modal-footer">{productPendingError && <button className="pc-button" disabled={Boolean(engineBusy)} onClick={() => void refreshPendingCard()}><WorkspaceIcon name="refresh" />ลองโหลด PENDING อีกครั้ง</button>}<button className="pc-button" onClick={() => setPendingProduct(null)}>ปิด</button></footer>
        </div>
      </StudioModal>}

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
  <StudioModal label="ผลตรวจสอบราคา GRAB" width={900} onClose={closeGrabModal}>
    <div className="pc-modal-panel">

      {/* Header */}
      <header className="pc-modal-header">
        <h2><WorkspaceIcon name="branch" />ผลตรวจสอบราคา GRAB</h2>
        <button onClick={closeGrabModal}
          className="pc-modal-close" aria-label="ปิดหน้าต่าง"><WorkspaceIcon name="close" /></button>
      </header>

      {/* Summary */}
      {grabResults.length > 0 && !grabResults[0]?.error && (
        <div style={{ padding: '10px 20px', background: 'var(--pc-bg)', borderBottom: '1px solid var(--pc-line)', display: 'flex', gap: 20, fontSize: 13, flexWrap: 'wrap' }} className="pc-modal-toolbar">
          <span>แสดง: <strong>{displayGrabResults.length}</strong> / {totalCount}</span>
          <span style={{ color: 'var(--pc-success)' }}>✓ ตรง: <strong>{matchedCount}</strong></span>
          <span style={{ color: 'var(--pc-danger)' }}>✗ ต้องแก้ไขใน GRAB: <strong>{mismatchCount}</strong></span>
          {notFoundCount > 0 && (
            <span style={{ color: 'var(--pc-ink)' }}>⚠ ไม่พบข้อมูลใน Supabase: <strong>{notFoundCount}</strong></span>
          )}
        </div>
      )}

      {/* Toolbar */}
      {grabResults.length > 0 && !grabResults[0]?.error && (
        <div style={{ padding: '8px 20px', background: '#fff', borderBottom: '1px solid var(--pc-line)', display: 'flex', gap: 12, alignItems: 'center' }} className="pc-modal-toolbar">
          <input
            type="text"
            placeholder="ค้นหา SKU / ชื่อสินค้า..."
            value={grabSearch}
            onChange={(e) => setGrabSearch(e.target.value)}
            style={{ padding: '6px 10px', border: '1px solid var(--pc-line)', borderRadius: 4, fontSize: 12, width: 220 }}
          />
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, cursor: 'pointer' }}>
            <input
              type="checkbox"
              checked={grabShowOnlyMismatch}
              onChange={(e) => setGrabShowOnlyMismatch(e.target.checked)}
            />
            เฉพาะราคาไม่ตรง
          </label>
          <span style={{ fontSize: 11, color: 'var(--pc-muted)', marginLeft: 'auto' }}>
            💡 คลิกที่ราคาใน GRAB เพื่อแก้ไข (ไม่บันทึกลง DB)
          </span>
        </div>
      )}

      {/* Table */}
      <div style={{ overflowY: 'auto', flex: 1 }} className="pc-modal-body">
        {grabResults[0]?.error ? (
          <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-danger)' }}>{grabResults[0].error}</div>
        ) : grabResults.length === 0 ? (
          <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-ink)' }}>กำลังโหลด...</div>
        ) : displayGrabResults.length === 0 ? (
          <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-ink)' }}>ไม่พบรายการที่ตรงกับเงื่อนไข</div>
        ) : (
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr style={{ background: 'var(--pc-bg)', position: 'sticky', top: 0 }}>
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
                  <tr key={r.sku || i} style={{ background: r.notFound ? 'var(--pc-bg)' : r.effectiveMatched ? 'transparent' : 'var(--pc-soft)' }}>
                    <td style={td}>{i + 1}</td>
                    <td style={{ ...td, color: 'var(--pc-ink)' }}><strong>{r.sku}</strong></td>
                    <td style={{ ...td, color: 'var(--pc-ink)' }}>{r.name || '-'}</td>
                    <td style={{ ...td, fontWeight: 600, color: 'var(--pc-ink)' }}>
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
                          style={{ width: 110, padding: '3px 6px', border: '1px solid var(--pc-accent)', borderRadius: 3, fontSize: 12 }}
                        />
                      ) : (
                        <span
                          onClick={() => { setEditingGrabSku(r.sku); setEditingGrabValue(String(r.effectiveGrabPrice)) }}
                          title="คลิกเพื่อแก้ไข"
                          style={{ cursor: 'pointer', padding: '2px 6px', borderRadius: 3, color: r.isEdited ? 'var(--pc-accent)' : 'var(--pc-ink)' }}
                          onMouseEnter={(e) => (e.currentTarget.style.background = 'var(--pc-soft)')}
                          onMouseLeave={(e) => (e.currentTarget.style.background = 'transparent')}
                        >
                          {r.effectiveGrabPrice?.toLocaleString() ?? '-'}
                          {r.isEdited && <span style={{ marginLeft: 4, fontSize: 10 }}>✎</span>}
                        </span>
                      )}
                    </td>
                    <td style={{ ...td, fontWeight: 600, color: 'var(--pc-ink)' }}>{r.dbPrice?.toLocaleString() ?? '-'}</td>
                    <td style={td}>
                      {r.notFound
                        ? <span style={{ background: 'var(--pc-bg)', color: 'var(--pc-ink)', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>⚠ ไม่พบข้อมูลใน Supabase</span>
                        : r.effectiveMatched
                          ? <span style={{ background: 'var(--pc-success-soft)', color: 'var(--pc-success)', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✓ ตรง</span>
                          : <span style={{ background: 'var(--pc-danger-soft)', color: 'var(--pc-danger)', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✗ แก้ไขใน GRAB</span>
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
      <div className="pc-modal-footer">
        <div style={{ display: 'flex', gap: 8 }}>
          <button
            onClick={exportGrabXlsx}
            disabled={grabMismatchProducts.length === 0}
            style={{ ...btnStyle, background: 'var(--pc-soft)', borderColor: 'var(--pc-line)', color: 'var(--pc-accent-dark)', opacity: grabMismatchProducts.length === 0 ? 0.5 : 1 }}
           className="pc-popup-action"><WorkspaceIcon name="download" />
            Export ราคา DB / Cloud ({grabMismatchProducts.length})
          </button>
          <button
            onClick={exportNotFoundXlsx}
            disabled={grabResults.filter((r: any) => r.notFound).length === 0}
            style={{ ...btnStyle, background: 'var(--pc-bg)', borderColor: 'var(--pc-muted)', color: 'var(--pc-ink)', opacity: grabResults.filter((r: any) => r.notFound).length === 0 ? 0.5 : 1 }}
           className="pc-popup-action"><WorkspaceIcon name="download" />
            Export ไม่พบข้อมูลใน Supabase ({grabResults.filter((r: any) => r.notFound).length})
          </button>
        </div>
        <button onClick={closeGrabModal} style={btnStyle}>ปิด</button>
      </div>
    </div>
  </StudioModal>
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
          <StudioModal label="ผลคำนวณราคา" width={1000} onClose={() => setShowPriceCalcModal(false)}>
            <div className="pc-modal-panel">

              {/* Header */}
              <header className="pc-modal-header">
                <h2><WorkspaceIcon name="tag" />ผลคำนวณราคา</h2>
                <button onClick={() => setShowPriceCalcModal(false)} className="pc-modal-close" aria-label="ปิดหน้าต่าง"><WorkspaceIcon name="close" /></button>
              </header>

              {newPromaxxSaveError ? <p className="pe-error pc-modal-toolbar" role="alert">{newPromaxxSaveError}</p> : <p className="pc-modal-toolbar" role="status">{savingNewPromaxx ? 'กำลังคำนวณและบันทึก New Promaxx...' : 'บันทึก New Promaxx แล้ว · ราคา DB จะเปลี่ยนเมื่อกดอัปเดทเท่านั้น'}</p>}

              {/* Grab_menu upload bar */}
              <div style={{ padding: '10px 20px', background: 'var(--pc-bg)', borderBottom: '1px solid var(--pc-line)', display: 'flex', gap: 10, alignItems: 'center' }} className="pc-modal-toolbar">
                <span style={{ fontSize: 12, color: 'var(--pc-ink)', fontWeight: 600, whiteSpace: 'nowrap' }}>Grab_menu (.csv):</span>
                <button onClick={() => document.getElementById('priceCalcGrabInput')?.click()}
                  style={{ ...btnStyle, fontSize: 11, padding: '5px 12px', background: priceCalcGrabFile ? 'var(--pc-success-soft)' : '#fff', borderColor: priceCalcGrabFile ? 'var(--pc-success)' : 'var(--pc-muted)', maxWidth: 280, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                  title={priceCalcGrabFile ? priceCalcGrabFile.name : 'เลือกไฟล์ Grab_menu เพื่อเปรียบเทียบราคาจริง'}
                 className="pc-popup-action"><WorkspaceIcon name="upload" />
                  {priceCalcGrabFile ? `${priceCalcGrabFile.name}` : 'เลือกไฟล์ Grab_menu'}
                </button>
                {!priceCalcGrabFile && (
                  <span style={{ fontSize: 11, color: 'var(--pc-muted)' }}>← อัพโหลดเพื่อแสดงราคา Grab จริง (ปัจจุบันใช้ราคาจาก DB)</span>
                )}
                {hasGrabMap && (
                  <span style={{ fontSize: 11, color: 'var(--pc-success)', marginLeft: 4 }}>✅ โหลด {Object.keys(priceCalcGrabMap).length.toLocaleString()} SKU จาก Grab_menu</span>
                )}
              </div>

              {/* Summary */}
              {priceCalcResults.length > 0 && !priceCalcResults[0]?.error && (
                <div style={{ padding: '10px 20px', background: 'var(--pc-bg)', borderBottom: '1px solid var(--pc-line)', display: 'flex', gap: 20, fontSize: 13 }} className="pc-modal-toolbar">
                  <span>ทั้งหมด: <strong>{enriched.length}</strong></span>
                  <span style={{ color: 'var(--pc-success)' }}>เปลี่ยนแปลง: <strong>{displayRows.length}</strong></span>
                  <span style={{ color: 'var(--pc-ink)' }}>ราคาเดิม: <strong>{enriched.length - displayRows.length}</strong></span>
                  <div style={{ marginLeft: 'auto', fontSize: 11, color: 'var(--pc-muted)' }}>
                    สูตร Grab = ระดับ 0 × 0.95 × 1.20 × 1.07
                  </div>
                </div>
              )}

              {/* Table */}
              <div style={{ overflowY: 'auto', flex: 1 }} className="pc-modal-body">
                {priceCalcResults[0]?.error ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-danger)' }}>{priceCalcResults[0].error}</div>
                ) : priceCalcResults.length === 0 ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-ink)' }}>{newPromaxxSaveError ? 'ยังไม่มีผลคำนวณ' : 'กำลังคำนวณ...'}</div>
                ) : (
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: 'var(--pc-bg)', position: 'sticky', top: 0 }}>
                        <th style={th}>#</th>
                        <th style={th}>SKU</th>
                        <th style={th}>ชื่อสินค้า</th>
                        <th style={th}>ราคาระดับ 0 ใหม่</th>
                        <th style={{ ...th, color: hasGrabMap ? 'var(--pc-ink)' : 'var(--pc-muted)' }}>
                          ราคาขาย Grab ปัจจุบัน{!hasGrabMap && <span style={{ fontWeight: 400, fontSize: 10 }}> (DB)</span>}
                        </th>
                        <th style={th}>ราคาขาย Grab ใหม่</th>
                        <th style={th}>สถานะ</th>
                      </tr>
                    </thead>
                    <tbody>
                      {displayRows.map((r: any, i: number) => (
                        <tr key={i} style={{ background: 'var(--pc-soft)' }}>
                          <td style={td}>{i + 1}</td>
                          <td style={td}><strong>{r.sku}</strong></td>
                          <td style={td}>{r.name}</td>
                          <td style={td}>{r.level0.toLocaleString()}</td>
                          <td style={td}>
                            {hasGrabMap
                              ? r.grabCurrent !== undefined
                                ? <strong>{r.grabCurrent.toLocaleString()}</strong>
                                : <span style={{ color: 'var(--pc-muted)', fontSize: 11 }}>ไม่พบใน Grab</span>
                              : r.oldPrice.toLocaleString()
                            }
                          </td>
                          <td style={{ ...td, color: 'var(--pc-success)', fontWeight: 600 }}>{r.newPrice.toLocaleString()}</td>
                          <td style={td}>
                            <span style={{ background: 'var(--pc-soft)', color: 'var(--pc-accent-dark)', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>⚠ เปลี่ยนแปลง</span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>

              {/* Footer */}
              <div className="pc-modal-footer">
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
                  style={{ ...btnStyle, background: 'var(--pc-accent)', color: '#fff', borderColor: 'var(--pc-accent)' }}
                 className="pc-popup-action"><WorkspaceIcon name="download" />
                  Export Excel
                </button>
                <button
                  onClick={confirmUpdatePrices}
                  disabled={updatingPrices || savingNewPromaxx || displayRows.length === 0}
                  style={{ ...btnStyle, background: 'var(--pc-accent)', color: '#fff', borderColor: 'var(--pc-accent)', opacity: updatingPrices ? 0.6 : 1 }}
                 className="pc-popup-action"><WorkspaceIcon name="check" />
                  {updatingPrices ? 'กำลังอัพเดท...' : `อัพเดท ${displayRows.length} รายการ`}
                </button>
              </div>
            </div>
          </StudioModal>
        )
      })()}


      {/* Import Modal */}
      {showImportModal && (
        <StudioModal label="Import ข้อมูลจาก Excel" width={600} onClose={() => setShowImportModal(false)}>
          <div className="pc-modal-panel">
            <header className="pc-modal-header">
              <h2><WorkspaceIcon name="upload" />Import ข้อมูลจาก Excel</h2>
              <button onClick={() => setShowImportModal(false)} className="pc-modal-close" aria-label="ปิดหน้าต่าง"><WorkspaceIcon name="close" /></button>
            </header>
            <div style={{ padding: '10px 20px', background: 'var(--pc-bg)', borderBottom: '1px solid var(--pc-line)', fontSize: 13 }} className="pc-modal-toolbar">
              พบ <strong>{importSheets.length}</strong> ชีท | รวม <strong>{importSheets.reduce((s, sh) => s + sh.rows.length, 0)}</strong> รายการ
            </div>
            <div style={{ overflowY: 'auto', flex: 1, padding: 16 }} className="pc-modal-body">
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                <thead>
                  <tr style={{ background: 'var(--pc-bg)' }}>
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
                <div style={{ marginTop: 12, padding: '8px 12px', background: 'var(--pc-danger-soft)', borderRadius: 4, fontSize: 13, color: 'var(--pc-danger)' }}>
                  ❌ {importError}
                </div>
              )}
              {importLog && (
                <div style={{ marginTop: 12, padding: '8px 12px', background: importLog.startsWith('✅') ? 'var(--pc-success-soft)' : importLog.startsWith('❌') ? 'var(--pc-danger-soft)' : 'var(--pc-soft)', borderRadius: 4, fontSize: 13 }}>
                  {importLog}
                </div>
              )}
            </div>
            <div className="pc-modal-footer">
              <button onClick={() => setShowImportModal(false)} style={btnStyle}>ปิด</button>
              <button
                onClick={confirmImport}
                disabled={importingData || !!importError || importSheets.reduce((s, sh) => s + sh.rows.length, 0) === 0}
                style={{ ...btnStyle, background: 'var(--pc-accent)', color: '#fff', borderColor: 'var(--pc-accent)', opacity: importingData ? 0.6 : 1 }}
               className="pc-popup-action"><WorkspaceIcon name="upload" />
                {importingData ? 'กำลัง Import...' : `Import ทั้งหมด (${importSheets.reduce((s, sh) => s + sh.rows.length, 0)} รายการ)`}
              </button>
            </div>
          </div>
        </StudioModal>
      )}

      {/* Edit Modal */}
      {editProduct && (
        <StudioModal label="รายละเอียดสินค้า" width={720} onClose={() => { setEditProduct(null); setIsAdding(false) }}>
          <div className="pc-modal-panel">
            {/* Modal Header */}
            <header className="pc-modal-header">
              <h2><WorkspaceIcon name="box" />{isAdding ? 'เพิ่มรายการสินค้า' : 'แก้ไขสินค้า'}</h2>
              <button aria-label="ปิดการแก้ไขสินค้า" onClick={() => { setEditProduct(null); setIsAdding(false) }}
                className="pc-modal-close"><WorkspaceIcon name="close" /></button>
            </header>

            {/* Modal Body */}
            <div style={{ padding: 20, display: 'grid', gap: 12 }} className="pc-modal-body pc-modal-form">
              <div style={{ gridColumn: '1 / -1' }}>
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
              <div style={{ gridColumn: '1 / -1' }}>
                <div style={{ marginTop: 4, padding: 14, borderRadius: 10, border: '2px dashed var(--pc-accent)', background: 'var(--pc-bg)' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                    <label style={{ ...labelStyle, margin: 0, fontSize: 15, color: 'var(--pc-accent)', fontWeight: 700 }}>
                      รูปภาพสินค้า + ข้อความมุมขวาล่าง
                    </label>
                    <span style={{ fontSize: 11, fontWeight: 700, color: 'var(--pc-accent)', background: 'var(--pc-soft)', borderRadius: 999, padding: '4px 8px' }}>
                      ขั้นตอนสำคัญ
                    </span>
                  </div>
                  <div style={{ marginTop: 6, fontSize: 12, lineHeight: 1.5, color: 'var(--pc-muted)' }}>
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
                      style={{ ...btnStyle, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', minWidth: 170, background: 'var(--pc-accent)', color: '#fff', borderColor: 'var(--pc-accent-dark)', cursor: 'pointer' }}
                    >
                      📂 เลือกรูปภาพสินค้า
                    </label>
                    <div style={{ flex: '1 1 220px', minHeight: 38, display: 'flex', alignItems: 'center', padding: '8px 10px', borderRadius: 6, border: '1px solid var(--pc-line)', background: '#fff', fontSize: 12, color: editImageFileName ? 'var(--pc-ink)' : 'var(--pc-muted)' }}>
                      {editImageFileName || 'ยังไม่ได้เลือกรูป (รองรับ PNG / JPEG / WebP)'}
                    </div>
                  </div>
                  <div style={{ marginTop: 10, fontSize: 12, color: 'var(--pc-muted)' }}>
                    ใส่เลขที่ใบอนุญาต (เลขอย่างเดียว) — ระบบจะพิมพ์คำว่า <b style={{ color: 'var(--pc-muted)' }}>"เลขที่ใบอนุญาต"</b> นำหน้าให้อัตโนมัติ เช่น เลขที่ใบอนุญาต 65-2-3-2-0003417
                  </div>
                  <div style={{ marginTop: 8, display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--pc-muted)' }}>ข้อความบนรูป:</div>
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
                    style={{ ...btnStyle, background: 'var(--pc-accent)', color: '#fff', borderColor: 'var(--pc-accent-dark)' }}
                   className="pc-popup-action"><WorkspaceIcon name="download" />
                    Download รูป
                  </button>
                  <button
                    type="button"
                    onClick={cleanGreenBorderAndText}
                    disabled={cleaningImage || (!editImageSrc && !editProduct?.['*รูปภาพสินค้า'])}
                    style={{ ...btnStyle, background: 'var(--pc-accent)', color: '#fff', borderColor: 'var(--pc-accent)', opacity: cleaningImage ? 0.6 : 1 }}
                   className="pc-popup-action"><WorkspaceIcon name="clock" />
                    {cleaningImage ? 'กำลังลบกรอบ...' : 'ลบกรอบ+ข้อความ'}
                  </button>
                  <button
                    type="button"
                    onClick={uploadEditImageToDrive}
                    disabled={!editImageSrc || uploadingDrive}
                    style={{ ...btnStyle, background: 'var(--pc-accent)', color: '#fff', borderColor: 'var(--pc-accent-dark)', opacity: uploadingDrive ? 0.6 : 1 }}
                   className="pc-popup-action"><WorkspaceIcon name="upload" />
                    {uploadingDrive ? 'กำลังอัพโหลด...' : 'อัพโหลด Drive'}
                  </button>
                  </div>
                  <div style={{ marginTop: 10, textAlign: 'center', background: '#fff', borderRadius: 8, padding: 10, border: '1px solid var(--pc-line)' }}>
                    {editImageSrc ? (
                      <canvas
                        ref={editCanvasRef}
                        style={{ maxWidth: '100%', maxHeight: 220, borderRadius: 4, background: '#fff' }}
                      />
                    ) : (
                      <div style={{ padding: '20px 12px', fontSize: 12, color: 'var(--pc-muted)' }}>
                        Preview รูปจะขึ้นที่นี่หลังจากเลือกรูปภาพ
                      </div>
                    )}
                  </div>
                </div>
                {editImageMsg && (
                  <div style={{ marginTop: 6, fontSize: 12, color: editImageMsg.includes('สำเร็จ') ? 'var(--pc-success)' : 'var(--pc-danger)' }}>
                    {editImageMsg}
                  </div>
                )}
                <label style={{ ...labelStyle, marginTop: 10, display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                  รูปภาพสินค้า (Google Drive Link)
                  {driveLinkOwner?.type === 'anin' && (
                    <span title={driveLinkOwner.email} style={{ background: 'var(--pc-success-soft)', color: 'var(--pc-success)', padding: '2px 8px', borderRadius: 10, fontSize: 10, fontWeight: 600 }}>🏢 Anin (องค์กร)</span>
                  )}
                  {driveLinkOwner?.type === 'external' && (
                    <span title={driveLinkOwner.email} style={{ background: 'var(--pc-soft)', color: 'var(--pc-accent-dark)', padding: '2px 8px', borderRadius: 10, fontSize: 10, fontWeight: 600 }}>🌐 External (เมล์นอก)</span>
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
                      style={{ ...btnStyle, background: 'var(--pc-soft)', color: 'var(--pc-accent-dark)', borderColor: 'var(--pc-line)', textDecoration: 'none', display: 'inline-block', padding: '4px 12px', fontSize: 12, borderRadius: 4 }}
                    >
                      ⬇️ Download รูปจาก Drive
                    </a>
                  </div>
                )}
                {editProduct['*รูปภาพสินค้า'] && (
                  <div style={{ marginTop: 8, textAlign: 'center', background: 'var(--pc-bg)', borderRadius: 4, padding: 8, border: '1px solid var(--pc-line)' }}>
                    <img
                      src={convertDriveLink(editProduct['*รูปภาพสินค้า'])}
                      alt="preview"
                      style={{ maxWidth: '100%', maxHeight: 160, borderRadius: 4, objectFit: 'contain' }}
                      onError={e => { e.currentTarget.style.display = 'none' }}
                    />
                  </div>
                )}
              </div>
              <div style={{ gridColumn: '1 / -1' }}>
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
            <div className="pc-modal-footer">
              <button onClick={() => { setEditProduct(null); setIsAdding(false) }} style={btnStyle}>ยกเลิก</button>
              <button onClick={() => { setSaveBtnAnim(true); setTimeout(() => setSaveBtnAnim(false), 200); saveProduct(); }}
                className={`pc-popup-action ${saveBtnAnim ? 'btn-press-anim' : ''}`}
                style={{ ...btnStyle, background: 'var(--pc-accent)', color: '#fff', borderColor: 'var(--pc-accent)' }}><WorkspaceIcon name="check" />
                {isAdding ? 'เพิ่มสินค้า' : 'บันทึก'}
              </button>
            </div>
          </div>
        </StudioModal>
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
          if (!r) return <span style={{ color: 'var(--pc-muted)', fontSize: 11 }}>—</span>
          if (r.status === 'updated') return <span style={{ background: 'var(--pc-success-soft)', color: 'var(--pc-success)', padding: '2px 6px', borderRadius: 10, fontSize: 10 }}>✅</span>
          if (r.status === 'notFound') return <span style={{ background: 'var(--pc-bg)', color: 'var(--pc-muted)', padding: '2px 6px', borderRadius: 10, fontSize: 10 }}>⚠</span>
          return <span style={{ background: 'var(--pc-danger-soft)', color: 'var(--pc-danger)', padding: '2px 6px', borderRadius: 10, fontSize: 10 }}>❌</span>
        }

        return (
          <StudioModal label="ตรวจสอบ GM MME vs GRAB Menu" width={1100} onClose={() => setShowMmeCheckModal(false)}>
            <div className="pc-modal-panel">

              {/* Header */}
              <header className="pc-modal-header">
                <h2><WorkspaceIcon name="check" />ตรวจสอบ GM MME vs GRAB Menu</h2>
                <button onClick={() => setShowMmeCheckModal(false)} className="pc-modal-close" aria-label="ปิดหน้าต่าง"><WorkspaceIcon name="close" /></button>
              </header>

              {/* Upload section */}
              <div style={{ padding: '12px 20px', background: 'var(--pc-bg)', borderBottom: '1px solid var(--pc-line)' }} className="pc-modal-toolbar">
                <div style={{ display: 'flex', gap: 10, marginBottom: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                  <span style={{ fontSize: 12, color: 'var(--pc-ink)', whiteSpace: 'nowrap', fontWeight: 600 }}>Grab_menu (.csv):</span>
                  {branches.map(b => (
                    <button key={b.key} onClick={() => document.getElementById(b.inputId)?.click()}
                      style={{ ...btnStyle, fontSize: 11, padding: '5px 10px', background: b.file ? 'var(--pc-success-soft)' : '#fff', borderColor: b.file ? 'var(--pc-success)' : 'var(--pc-muted)', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                      title={b.file ? b.file.name : `เลือกไฟล์ Grab_menu สาขา ${b.label}`}
                     className="pc-popup-action"><WorkspaceIcon name="upload" />
                      {b.file ? `${b.label}` : `${b.label}`}
                    </button>
                  ))}
                </div>
                <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
                  <span style={{ fontSize: 12, color: 'var(--pc-ink)', whiteSpace: 'nowrap', fontWeight: 600 }}>GM MME (.xlsx):</span>
                  <button onClick={() => document.getElementById('mmeCheckMmeInput')?.click()}
                    style={{ ...btnStyle, fontSize: 11, padding: '5px 10px', background: mmeCheckMmeFile ? 'var(--pc-success-soft)' : '#fff', borderColor: mmeCheckMmeFile ? 'var(--pc-success)' : 'var(--pc-muted)', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                    title={mmeCheckMmeFile ? mmeCheckMmeFile.name : 'เลือกไฟล์ GM MME'}
                   className="pc-popup-action"><WorkspaceIcon name="upload" />
                    {mmeCheckMmeFile ? `${mmeCheckMmeFile.name}` : 'เลือกไฟล์'}
                  </button>
                  <button onClick={handleMmeVsGrabCheck} disabled={!canCheck}
                    style={{ ...btnStyle, background: canCheck ? 'var(--pc-accent)' : 'var(--pc-line)', color: canCheck ? '#fff' : 'var(--pc-muted)', borderColor: 'var(--pc-accent)', marginLeft: 'auto', opacity: mmeChecking ? 0.6 : 1 }}
                   className="pc-popup-action"><WorkspaceIcon name="check" />
                    {mmeChecking ? 'กำลังตรวจสอบ...' : 'ตรวจสอบ'}
                  </button>
                </div>
              </div>

              {/* Tabs */}
              {hasResults && (
                <div style={{ display: 'flex', borderBottom: '2px solid var(--pc-line)', background: 'var(--pc-bg)' }} className="pc-modal-tabs">
                  {checkedBranches.map(b => {
                    const notUpd = mmeCheckResults[b.key].filter((r: any) => r.status === 'notUpdated').length
                    const isActive = mmeActiveTab === b.key
                    return (
                      <button key={b.key} onClick={() => setMmeActiveTab(b.key)}
                        style={{ padding: '8px 18px', border: 'none', borderBottom: isActive ? '2px solid var(--pc-accent)' : '2px solid transparent', background: 'none', cursor: 'pointer', fontWeight: isActive ? 700 : 400, color: isActive ? 'var(--pc-accent)' : 'var(--pc-muted)', fontSize: 13, marginBottom: -2 }}
                       className="pc-popup-action"><WorkspaceIcon name="branch" />
                        {b.label}
                        {notUpd > 0 && <span style={{ marginLeft: 5, background: 'var(--pc-danger)', color: '#fff', borderRadius: 10, fontSize: 10, padding: '1px 6px' }}>{notUpd}</span>}
                      </button>
                    )
                  })}
                  {checkedBranches.length >= 2 && (
                    <button onClick={() => setMmeActiveTab('compare')}
                      style={{ padding: '8px 18px', border: 'none', borderBottom: mmeActiveTab === 'compare' ? '2px solid var(--pc-accent)' : '2px solid transparent', background: 'none', cursor: 'pointer', fontWeight: mmeActiveTab === 'compare' ? 700 : 400, color: mmeActiveTab === 'compare' ? 'var(--pc-accent)' : 'var(--pc-muted)', fontSize: 13, marginBottom: -2, marginLeft: 'auto' }}
                     className="pc-popup-action"><WorkspaceIcon name="grid" />
                      เปรียบเทียบ
                      {compareNotUpdatedCount > 0 && <span style={{ marginLeft: 5, background: 'var(--pc-danger)', color: '#fff', borderRadius: 10, fontSize: 10, padding: '1px 6px' }}>{compareNotUpdatedCount}</span>}
                    </button>
                  )}
                </div>
              )}

              {/* Summary bar */}
              {mmeActiveTab !== 'compare' && tabResults.length > 0 && (
                <div style={{ padding: '8px 20px', background: 'var(--pc-bg)', borderBottom: '1px solid var(--pc-line)', display: 'flex', gap: 20, fontSize: 13 }} className="pc-modal-toolbar">
                  <span>ทั้งหมด: <strong>{tabResults.length}</strong></span>
                  <span style={{ color: 'var(--pc-success)' }}>✅ แก้ไขแล้ว: <strong>{tabResults.filter((r: any) => r.status === 'updated').length}</strong></span>
                  <span style={{ color: 'var(--pc-danger)' }}>❌ ยังไม่แก้ไข: <strong>{tabResults.filter((r: any) => r.status === 'notUpdated').length}</strong></span>
                  {tabResults.filter((r: any) => r.status === 'notFound').length > 0 && (
                    <span style={{ color: 'var(--pc-accent-dark)' }}>⚠ ไม่พบใน GRAB: <strong>{tabResults.filter((r: any) => r.status === 'notFound').length}</strong></span>
                  )}
                </div>
              )}
              {mmeActiveTab === 'compare' && compareRows.length > 0 && (
                <div style={{ padding: '8px 20px', background: 'var(--pc-soft)', borderBottom: '1px solid var(--pc-line)', display: 'flex', gap: 20, fontSize: 13 }} className="pc-modal-toolbar">
                  <span>SKU ทั้งหมด: <strong>{compareRows.length}</strong></span>
                  <span style={{ color: 'var(--pc-danger)' }}>❌ มีสาขาที่ยังไม่แก้ไข: <strong>{compareNotUpdatedCount}</strong></span>
                  <span style={{ color: 'var(--pc-success)' }}>✅ แก้ไขครบทุกสาขา: <strong>{compareRows.filter(r => checkedBranches.every(b => r[b.key as 'src'|'kkl'|'sss']?.status === 'updated')).length}</strong></span>
                </div>
              )}

              {/* Table area */}
              <div style={{ overflowY: 'auto', flex: 1 }} className="pc-modal-body">
                {!hasResults ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-muted)', fontSize: 13 }}>
                    เลือกไฟล์ Grab_menu อย่างน้อย 1 สาขา และ GM MME (.xlsx) แล้วกด &quot;ตรวจสอบ&quot;
                  </div>
                ) : mmeActiveTab === 'compare' ? (
                  /* Compare table */
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: 'var(--pc-soft)', position: 'sticky', top: 0 }}>
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
                          <tr key={r.sku} style={{ background: anyNotUpdated ? 'var(--pc-danger-soft)' : allUpdated ? 'var(--pc-success-soft)' : 'transparent' }}>
                            <td style={td}>{i + 1}</td>
                            <td style={{ ...td, color: 'var(--pc-ink)' }}><strong>{r.sku}</strong></td>
                            <td style={{ ...td, color: 'var(--pc-ink)' }}>{r.name || '-'}</td>
                            <td style={{ ...td, textAlign: 'center', fontWeight: 600 }}>{r.mmePrice?.toLocaleString() ?? '-'}</td>
                            {checkedBranches.map(b => {
                              const br = r[b.key as 'src'|'kkl'|'sss']
                              return (
                                <td key={b.key} style={{ ...td, textAlign: 'center' }}>
                                  {statusBadge(br)}
                                  {br?.grabPrice != null && (
                                    <div style={{ fontSize: 10, color: 'var(--pc-muted)', marginTop: 2 }}>{br.grabPrice.toLocaleString()}</div>
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
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-muted)', fontSize: 13 }}>
                    {tabHasFile ? 'ไม่มีข้อมูล' : 'ไม่ได้อัพโหลดไฟล์สาขานี้'}
                  </div>
                ) : (
                  /* Single branch table */
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: 'var(--pc-bg)', position: 'sticky', top: 0 }}>
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
                        <tr key={i} style={{ background: r.status === 'updated' ? 'transparent' : r.status === 'notFound' ? 'var(--pc-bg)' : 'var(--pc-danger-soft)' }}>
                          <td style={td}>{i + 1}</td>
                          <td style={{ ...td, color: 'var(--pc-ink)' }}><strong>{r.sku}</strong></td>
                          <td style={{ ...td, color: 'var(--pc-ink)' }}>{r.name || '-'}</td>
                          <td style={{ ...td, fontWeight: 600 }}>{r.mmePrice?.toLocaleString() ?? '-'}</td>
                          <td style={{ ...td, fontWeight: 600 }}>{r.grabPrice?.toLocaleString() ?? '-'}</td>
                          <td style={td}>
                            {r.status === 'updated'
                              ? <span style={{ background: 'var(--pc-success-soft)', color: 'var(--pc-success)', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✅ แก้ไขแล้ว</span>
                              : r.status === 'notFound'
                                ? <span style={{ background: 'var(--pc-bg)', color: 'var(--pc-muted)', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>⚠ ไม่พบใน GRAB</span>
                                : <span style={{ background: 'var(--pc-danger-soft)', color: 'var(--pc-danger)', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>❌ ยังไม่แก้ไข</span>
                            }
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>

              {/* Footer */}
              <div className="pc-modal-footer">
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
                        style={{ ...btnStyle, background: 'var(--pc-soft)', borderColor: 'var(--pc-line)', color: 'var(--pc-accent-dark)', opacity: notUpdated.length === 0 ? 0.5 : 1 }}
                       className="pc-popup-action"><WorkspaceIcon name="download" />
                        Export {b.label} ({notUpdated.length})
                      </button>
                    )
                  })}
                </div>
                <button onClick={() => setShowMmeCheckModal(false)} style={btnStyle}>ปิด</button>
              </div>
            </div>
          </StudioModal>
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
          <StudioModal label="สินค้าที่ขาดใน GRAB" width={1100} onClose={() => setShowMissingBranchModal(false)}>
            <div className="pc-modal-panel">

              {/* Header */}
              <header className="pc-modal-header">
                <h2><WorkspaceIcon name="branch" />สินค้าที่ขาดใน GRAB</h2>
                <button onClick={() => setShowMissingBranchModal(false)} className="pc-modal-close" aria-label="ปิดหน้าต่าง"><WorkspaceIcon name="close" /></button>
              </header>

              {/* Tabs */}
              <div style={{ display: 'flex', borderBottom: '2px solid var(--pc-line)', background: 'var(--pc-bg)' }} className="pc-modal-tabs">
                {branches.map(b => {
                  const isActive = missingBranchTab === b.key
                  const count = missingBranchLoaded[b.key] ? missingBranchResults[b.key].length : null
                  return (
                    <button key={b.key}
                      onClick={() => { setMissingBranchTab(b.key); setMissingBranchExportCat('all'); setMissingBranchSearch(''); loadMissingBranch(b.key) }}
                      style={{ padding: '8px 24px', border: 'none', borderBottom: isActive ? '2px solid var(--pc-accent)' : '2px solid transparent', background: 'none', cursor: 'pointer', fontWeight: isActive ? 700 : 400, color: isActive ? 'var(--pc-accent)' : 'var(--pc-muted)', fontSize: 13, marginBottom: -2 }}
                     className="pc-popup-action"><WorkspaceIcon name="download" />
                      {b.label}
                      {count !== null && (
                        <span style={{ marginLeft: 6, background: count > 0 ? 'var(--pc-accent)' : 'var(--pc-success)', color: '#fff', borderRadius: 10, fontSize: 10, padding: '1px 6px' }}>{count}</span>
                      )}
                    </button>
                  )
                })}
                <div style={{ marginLeft: 'auto', padding: '8px 16px', fontSize: 11, color: 'var(--pc-muted)', alignSelf: 'center' }}>
                  ข้อมูลอ้างอิงจาก branch field ใน DB (อัพเดทเมื่ออัพโหลด GRAB CSV)
                </div>
              </div>

              {/* Summary + Search */}
              {isLoaded && (
                <div style={{ padding: '8px 20px', background: 'var(--pc-bg)', borderBottom: '1px solid var(--pc-line)', display: 'flex', gap: 16, fontSize: 13, alignItems: 'center', flexWrap: 'wrap' }} className="pc-modal-toolbar">
                  <span>ทั้งหมด: <strong style={{ color: tabData.length > 0 ? 'var(--pc-accent)' : 'var(--pc-success)' }}>{tabData.length} รายการ</strong></span>
                  <input
                    type="text"
                    placeholder="ค้นหาชื่อ / SKU..."
                    value={missingBranchSearch}
                    onChange={e => setMissingBranchSearch(e.target.value)}
                    style={{ padding: '3px 10px', fontSize: 12, borderRadius: 4, border: '1px solid var(--pc-line)', width: 200 }}
                  />
                  {missingBranchSearch && (
                    <span style={{ color: 'var(--pc-muted)' }}>พบ {tabData.filter((p: any) => {
                      const q = missingBranchSearch.toLowerCase()
                      return (p['*ชื่อสินค้า (NAME)'] || '').toLowerCase().includes(q) || (p['รหัสสินค้า (SKU NUMBER)'] || '').toLowerCase().includes(q)
                    }).length} รายการ</span>
                  )}
                </div>
              )}

              {/* Table */}
              <div style={{ overflowY: 'auto', flex: 1 }} className="pc-modal-body">
                {!isLoaded || missingBranchLoading ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-muted)', fontSize: 13 }}>⏳ กำลังโหลด...</div>
                ) : tabData.length === 0 ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-success)', fontSize: 14 }}>
                    ✅ ไม่มีสินค้าที่ขาดใน GRAB {missingBranchTab.toUpperCase()}
                  </div>
                ) : (
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: 'var(--pc-bg)', position: 'sticky', top: 0 }}>
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
                        <tr key={p['รหัสสินค้า (SKU NUMBER)']} style={{ background: 'var(--pc-bg)' }}>
                          <td style={td}>{i + 1}</td>
                          <td style={{ ...td, color: 'var(--pc-ink)' }}><strong>{p['รหัสสินค้า (SKU NUMBER)']}</strong></td>
                          <td style={{ ...td, color: 'var(--pc-ink)' }}>{p['*ชื่อสินค้า (NAME)'] || '-'}</td>
                          <td style={{ ...td, fontWeight: 600 }}>{p['*ราคาสินค้า'] || '-'}</td>
                          <td style={td}>{p['หมวดหมู่สินค้า (CATEGORIES)'] || '-'}</td>
                          {(['src','kkl','sss'] as const).map(b => (
                            <td key={b} style={{ ...td, textAlign: 'center' }}>
                              {p[b]
                                ? <span style={{ background: 'var(--pc-success-soft)', color: 'var(--pc-success)', padding: '2px 6px', borderRadius: 10, fontSize: 10 }}>✅</span>
                                : <span style={{ background: 'var(--pc-danger-soft)', color: 'var(--pc-danger)', padding: '2px 6px', borderRadius: 10, fontSize: 10 }}>❌</span>
                              }
                            </td>
                          ))}
                          {(['src','kkl','sss'] as const).map(b => {
                            const val = p[`${b}_price`]
                            const dbPrice = parsePriceRobust(String(p['*ราคาสินค้า']))
                            const mismatch = val !== null && val !== undefined && dbPrice !== null && Math.abs(Number(val) - dbPrice) >= 0.01
                            return (
                              <td key={`${b}-price`} style={{ ...td, textAlign: 'center', background: mismatch ? 'var(--pc-soft)' : 'transparent' }}>
                                {val !== null && val !== undefined ? (
                                  <span style={{ fontWeight: 600, color: mismatch ? 'var(--pc-accent-dark)' : 'var(--pc-success)', fontSize: 11 }}>
                                    {Number(val).toLocaleString()}
                                    {mismatch && <span style={{ display: 'block', fontSize: 9, color: 'var(--pc-accent-dark)' }}>≠ DB</span>}
                                  </span>
                                ) : (
                                  <span style={{ color: 'var(--pc-line)', fontSize: 11 }}>-</span>
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
              <div className="pc-modal-footer">
                <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                  {(() => {
                    const cats = isLoaded ? [...new Set(tabData.map((p: any) => p['หมวดหมู่สินค้า (CATEGORIES)'] || '').filter(Boolean))].sort((a, b) => a.localeCompare(b, 'th')) : []
                    const exportData = missingBranchExportCat === 'all' ? tabData : tabData.filter((p: any) => (p['หมวดหมู่สินค้า (CATEGORIES)'] || '') === missingBranchExportCat)
                    return (<>
                      <select
                        value={missingBranchExportCat}
                        onChange={e => setMissingBranchExportCat(e.target.value)}
                        disabled={!isLoaded || tabData.length === 0}
                        style={{ padding: '5px 8px', fontSize: 12, borderRadius: 4, border: '1px solid var(--pc-line)', cursor: 'pointer' }}
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
                        style={{ ...btnStyle, background: 'var(--pc-soft)', borderColor: 'var(--pc-line)', color: 'var(--pc-accent-dark)', opacity: (!isLoaded || exportData.length === 0) ? 0.5 : 1 }}
                       className="pc-popup-action"><WorkspaceIcon name="download" />
                        Export ({exportData.length})
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
                   className="pc-popup-action"><WorkspaceIcon name="refresh" />
                    รีโหลด
                  </button>
                </div>
                <button onClick={() => setShowMissingBranchModal(false)} style={btnStyle}>ปิด</button>
              </div>
            </div>
          </StudioModal>
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
            background: has ? 'var(--pc-success-soft)' : 'var(--pc-danger-soft)',
            color: has ? 'var(--pc-success)' : 'var(--pc-danger)'
          }}>{has ? '✓' : '✗'}</span>
        )
        return (
          <StudioModal label="Product Master vs สาขา" width={1000} onClose={() => setShowMasterBranchModal(false)}>
            <div className="pc-modal-panel">
              <header className="pc-modal-header">
                <div>
                  <h2><WorkspaceIcon name="grid" />Product Master vs สาขา</h2>
                  <span style={{ marginLeft: 12, fontSize: 12, color: 'var(--pc-muted)' }}>
                    {masterBranchResults.length} SKU
                    {missingCounts.src + missingCounts.kkl + missingCounts.sss > 0 && (
                      <> — ขาด SRC: <strong style={{ color: 'var(--pc-danger)' }}>{missingCounts.src}</strong>{' '}
                      KKL: <strong style={{ color: 'var(--pc-danger)' }}>{missingCounts.kkl}</strong>{' '}
                      SSS: <strong style={{ color: 'var(--pc-danger)' }}>{missingCounts.sss}</strong></>
                    )}
                  </span>
                </div>
                <button onClick={() => setShowMasterBranchModal(false)} className="pc-modal-close" aria-label="ปิดหน้าต่าง"><WorkspaceIcon name="close" /></button>
              </header>
              <div style={{ display: 'flex', borderBottom: '2px solid var(--pc-line)', background: 'var(--pc-bg)' }} className="pc-modal-tabs">
                {branchTabs.map(t => {
                  const isActive = masterBranchTab === t.key
                  const cnt = t.key === 'all' ? masterBranchResults.length : missingCounts[t.key]
                  return (
                    <button key={t.key}
                      onClick={() => setMasterBranchTab(t.key)}
                      style={{
                        padding: '8px 16px', border: 'none', cursor: 'pointer', fontSize: 12, fontWeight: isActive ? 700 : 400,
                        background: isActive ? '#fff' : 'transparent',
                        borderBottom: isActive ? '2px solid var(--pc-accent)' : '2px solid transparent',
                        color: isActive ? 'var(--pc-accent)' : (t.key !== 'all' && cnt > 0 ? 'var(--pc-danger)' : 'var(--pc-muted)')
                      }}
                    >{t.label} ({cnt})</button>
                  )
                })}
              </div>
              <div style={{ padding: '8px 16px', display: 'flex', gap: 8, alignItems: 'center', borderBottom: '1px solid var(--pc-line)' }} className="pc-modal-toolbar">
                <input
                  type="text"
                  placeholder="ค้นหา SKU / ชื่อสินค้า..."
                  value={masterBranchSearch}
                  onChange={e => setMasterBranchSearch(e.target.value)}
                  style={{ padding: '4px 10px', fontSize: 12, borderRadius: 4, border: '1px solid var(--pc-line)', width: 220 }}
                />
                {masterBranchSearch && <span style={{ color: 'var(--pc-muted)', fontSize: 12 }}>พบ {filtered.length} รายการ</span>}
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
                  style={{ ...btnStyle, fontSize: 11, padding: '4px 10px', background: 'var(--pc-success-soft)', borderColor: 'var(--pc-success)', color: 'var(--pc-success)', opacity: filtered.length === 0 ? 0.5 : 1 }}
                 className="pc-popup-action"><WorkspaceIcon name="download" />Export</button>
              </div>
              <div style={{ overflowY: 'auto', flex: 1 }} className="pc-modal-body">
                {masterBranchLoading ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-muted)', fontSize: 13 }}>⏳ กำลังโหลด...</div>
                ) : filtered.length === 0 ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-success)', fontSize: 14 }}>
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
                        <tr key={r.sku} style={{ background: (!r.src || !r.kkl || !r.sss) ? 'var(--pc-danger-soft)' : 'transparent' }}>
                          <td style={{ ...td, color: 'var(--pc-muted)', textAlign: 'center' }}>{i + 1}</td>
                          <td style={{ ...td, fontFamily: 'monospace', fontWeight: 600 }}>{r.sku}</td>
                          <td style={td}>{r.name || <span style={{ color: 'var(--pc-muted)' }}>—</span>}</td>
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
          </StudioModal>
        )
      })()}

      {popupNotice && <StudioModal label={popupNotice.confirm ? 'ยืนยันการลบหมวดหมู่' : 'แจ้งเตือน'} width={520} onClose={() => dismissPopup(false)}>
        <div className="pc-modal-panel">
          <header className="pc-modal-header"><h2><WorkspaceIcon name="alert" />{popupNotice.confirm ? 'ยืนยันการลบหมวดหมู่' : 'แจ้งเตือน'}</h2><button className="pc-modal-close" aria-label="ปิดแจ้งเตือน" onClick={() => dismissPopup(false)}><WorkspaceIcon name="close" /></button></header>
          <div className="pc-modal-body pc-modal-message">{popupNotice.message}</div>
          <footer className="pc-modal-footer">
            {popupNotice.confirm && <button className="pc-button" autoFocus onClick={() => dismissPopup(false)}>ยกเลิก</button>}
            <button className={`pc-button ${popupNotice.confirm ? 'pc-modal-danger' : 'pc-button-primary'}`} autoFocus={!popupNotice.confirm} onClick={() => dismissPopup(true)}>{popupNotice.confirm ? 'ยืนยันลบ' : 'รับทราบ'}</button>
          </footer>
        </div>
      </StudioModal>}

      {/* Image Tooltip */}
      {tooltip && (
        <div style={{
          position: 'fixed',
          left: tooltip.x + 15,
          top: tooltip.y + 15,
          zIndex: 9999,
          background: '#fff',
          border: '2px solid var(--pc-accent)',
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
            <div style={{ padding: 40, color: 'var(--pc-ink)', fontSize: 12 }}>ไม่มีรูป</div>
          )}
        </div>
      )}


      {/* Promaxx vs DB Modal */}
      {showPromaxxModal && (() => {
        const mismatchList = promaxxResults.filter((r: any) => !r.notFound && !r.matched)
        const matchedList = promaxxResults.filter((r: any) => !r.notFound && r.matched)
        const notFoundList = promaxxResults.filter((r: any) => r.notFound)
        return (
          <StudioModal label="ผลตรวจสอบ Promaxx (R05.103) vs DB" width={1000} onClose={() => setShowPromaxxModal(false)}>
            <div className="pc-modal-panel">
              <header className="pc-modal-header">
                <h2><WorkspaceIcon name="tag" />ผลตรวจสอบ Promaxx (R05.103) vs DB</h2>
                <button onClick={() => setShowPromaxxModal(false)} className="pc-modal-close" aria-label="ปิดหน้าต่าง"><WorkspaceIcon name="close" /></button>
              </header>
              {/* Progress bar */}
              {promaxxProgress.total > 0 && (
                <div style={{ padding: '10px 20px', background: 'var(--pc-bg)', borderBottom: '1px solid var(--pc-line)' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11, color: 'var(--pc-accent-dark)', marginBottom: 4 }}>
                    <span>{promaxxProgress.phase === 'fetch' ? '🔄 กำลังดึงข้อมูล...' : '💾 กำลังบันทึกราคา...'}</span>
                    <span>{promaxxProgress.current.toLocaleString()} / {promaxxProgress.total.toLocaleString()}</span>
                  </div>
                  <div style={{ background: 'var(--pc-soft)', borderRadius: 999, height: 6, overflow: 'hidden' }}>
                    <div style={{ background: 'var(--pc-accent)', height: '100%', width: `${Math.min(100, (promaxxProgress.current / promaxxProgress.total) * 100)}%`, transition: 'width 0.3s' }} />
                  </div>
                </div>
              )}
              {promaxxResults.length > 0 && !promaxxResults[0]?.error && (
                <div style={{ padding: '10px 20px', background: 'var(--pc-bg)', borderBottom: '1px solid var(--pc-line)', display: 'flex', gap: 20, fontSize: 13 }} className="pc-modal-toolbar">
                  <span>ทั้งหมด: <strong>{promaxxResults.length}</strong></span>
                  <span style={{ color: 'var(--pc-success)' }}>✓ ตรง: <strong>{matchedList.length}</strong></span>
                  <span style={{ color: 'var(--pc-danger)' }}>✗ ราคาไม่ตรง: <strong>{mismatchList.length}</strong></span>
                  {notFoundList.length > 0 && <span style={{ color: 'var(--pc-accent-dark)' }}>⚠ ไม่พบใน DB: <strong>{notFoundList.length}</strong></span>}
                </div>
              )}
              <div style={{ overflowY: 'auto', flex: 1 }} className="pc-modal-body">
                {promaxxChecking ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-muted)' }}>⏳ กำลังโหลด...</div>
                ) : promaxxResults[0]?.error ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-danger)' }}>{promaxxResults[0].error}</div>
                ) : promaxxResults.length === 0 ? (
                  <div style={{ padding: 40, textAlign: 'center', color: 'var(--pc-muted)' }}>ไม่มีข้อมูล</div>
                ) : (
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: 'var(--pc-bg)', position: 'sticky', top: 0 }}>
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
                        <tr key={i} style={{ background: r.notFound ? 'var(--pc-bg)' : r.matched ? 'transparent' : 'var(--pc-soft)' }}>
                          <td style={td}>{i + 1}</td>
                          <td style={td}><strong>{r.sku}</strong></td>
                          <td style={td}>{r.name || '-'}</td>
                          <td style={td}>{r.notFound ? '-' : r.promaxxPrice?.toLocaleString()}</td>
                          <td style={td}>{r.notFound ? '-' : r.dbPrice?.toLocaleString()}</td>
                          <td style={td}>
                            {r.notFound
                              ? <span style={{ background: 'var(--pc-bg)', color: 'var(--pc-muted)', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>⚠ ไม่พบใน DB</span>
                              : r.matched
                                ? <span style={{ background: 'var(--pc-success-soft)', color: 'var(--pc-success)', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✓ ตรง</span>
                                : <span style={{ background: 'var(--pc-soft)', color: 'var(--pc-accent-dark)', padding: '2px 8px', borderRadius: 10, fontSize: 11 }}>✗ ราคาไม่ตรง</span>
                            }
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
              <div className="pc-modal-footer">
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
                  style={{ ...btnStyle, background: 'var(--pc-soft)', borderColor: 'var(--pc-line)', color: 'var(--pc-accent-dark)', opacity: mismatchList.length === 0 ? 0.5 : 1 }}
                 className="pc-popup-action"><WorkspaceIcon name="download" />
                  Export ราคาไม่ตรง ({mismatchList.length})
                </button>
                <button onClick={() => setShowPromaxxModal(false)} style={btnStyle}>ปิด</button>
              </div>
            </div>
          </StudioModal>
        )
      })()}

    </div>
  )
}

const btnStyle: React.CSSProperties = {
  padding: '10px 14px', border: '1px solid var(--pc-line)', borderRadius: 8,
  fontSize: 11, fontWeight: 700, cursor: 'pointer', background: '#fff', color: 'var(--pc-ink)', minHeight: 40
}
const th: React.CSSProperties = {
  padding: '12px 16px', textAlign: 'left', borderBottom: '1px solid var(--pc-line)',
  fontWeight: 600, position: 'sticky', top: 0, background: 'var(--pc-bg)'
}
const td: React.CSSProperties = {
  padding: '12px 16px', borderBottom: '1px solid var(--pc-line)', color: 'var(--pc-ink)'
}
const labelStyle: React.CSSProperties = {
  fontSize: 11, color: 'var(--pc-ink)', display: 'block', marginBottom: 4
}
const inputStyle: React.CSSProperties = {
  width: '100%', padding: '6px 10px', border: '1px solid var(--pc-muted)',
  borderRadius: 8, fontSize: 13, fontFamily: 'inherit', color: 'var(--pc-ink)',
  boxSizing: 'border-box'
}

