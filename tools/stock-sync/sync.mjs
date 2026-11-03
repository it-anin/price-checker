// Copies on-hand stock (R01.102 CF_QUANTITY) from MSSQL to Supabase.
// Usage: node sync.mjs                one sync of every branch, then exit
//        node sync.mjs --dry-run      read MSSQL only, print totals, write nothing
//        node sync.mjs --serve        stay running; sync live whenever the web app asks
//        node sync.mjs --compare src.csv [sss.csv kkl.csv]   read-only check against R01.102 CSV exports
import fs from 'node:fs'
import path from 'node:path'
import sql from 'mssql'
import { createClient } from '@supabase/supabase-js'
import { buildConnectionConfig, buildStockQuery, chunkRows, parseBranchMap, toStockRows } from './query.mjs'
import { branchFromFileName, compareStock, decodeBytes, differencesToCsv, readStockCsv } from './compare.mjs'
import { HEARTBEAT_MS, POLL_MS, claimNextRequest, finishRequest, recoverStuckRequests, sendHeartbeat, summarizeResults } from './serve.mjs'

const dryRun = process.argv.includes('--dry-run')
const serveMode = process.argv.includes('--serve')
const compareIndex = process.argv.indexOf('--compare')
const compareFiles = compareIndex < 0 ? [] : process.argv.slice(compareIndex + 1).filter(arg => !arg.startsWith('--'))
const CHUNK = 1000
const CONCURRENCY = 4

function env(name) {
  const value = process.env[name]
  if (!value) throw new Error(`ยังไม่ได้ตั้ง ${name}`)
  return value
}

const makeSupabase = () => createClient(env('SUPABASE_URL'), env('SUPABASE_SERVICE_ROLE_KEY'), { auth: { persistSession: false } })
const connectMssql = () => new sql.ConnectionPool(buildConnectionConfig(process.env)).connect()

async function readBranch(pool, { branchId, warehouseIds }) {
  const request = pool.request().input('branchId', sql.Int, branchId)
  warehouseIds.forEach((id, i) => request.input(`w${i}`, sql.Int, id))
  const result = await request.query(buildStockQuery(warehouseIds))
  return result.recordset
}

async function writeBranch(supabase, branch, rows, syncedAt) {
  // A few chunks at a time keeps a live request fast without flooding Supabase.
  const chunks = chunkRows(rows, CHUNK)
  for (let i = 0; i < chunks.length; i += CONCURRENCY) {
    const results = await Promise.all(chunks.slice(i, i + CONCURRENCY)
      .map(chunk => supabase.from('stock_onhand').upsert(chunk, { onConflict: 'branch,sku' })))
    const failed = results.find(result => result.error)
    if (failed) throw new Error(`upsert ไม่สำเร็จ: ${failed.error.message}`)
  }
  // Only after every chunk succeeded: drop SKUs that no longer appear for this branch.
  const { error } = await supabase.from('stock_onhand').delete().eq('branch', branch).lt('synced_at', syncedAt)
  if (error) throw new Error(`ลบ SKU เก่าไม่สำเร็จ: ${error.message}`)
}

async function setStatus(supabase, branch, status) {
  const { error } = await supabase.from('stock_sync_status').upsert({ branch, ...status }, { onConflict: 'branch' })
  if (error) console.error(`[${branch}] เขียนสถานะไม่สำเร็จ: ${error.message}`)
}

// Reads one branch from MSSQL and (unless supabase is null) writes it to Supabase.
// Never throws: the outcome is returned so callers can report every branch.
async function syncBranch(pool, supabase, branch, config) {
  const syncedAt = new Date().toISOString()
  try {
    const rows = toStockRows(await readBranch(pool, config), branch, syncedAt)
    // An empty result almost always means a wrong branchId; never wipe good data for it.
    if (!rows.length) throw new Error('query ไม่พบสต็อกเลย ตรวจ branchId/warehouseIds')
    if (supabase) {
      await writeBranch(supabase, branch, rows, syncedAt)
      await setStatus(supabase, branch, { synced_at: syncedAt, row_count: rows.length, ok: true, message: null })
    }
    return { branch, ok: true, count: rows.length, rows }
  } catch (error) {
    if (supabase) await setStatus(supabase, branch, { synced_at: syncedAt, row_count: 0, ok: false, message: String(error.message).slice(0, 200) })
    return { branch, ok: false, error: String(error.message) }
  }
}

async function runOnce() {
  const branchMap = parseBranchMap(env('STOCK_BRANCH_MAP'))
  const supabase = dryRun ? null : makeSupabase()
  const pool = await connectMssql()
  try {
    for (const [branch, config] of Object.entries(branchMap)) {
      const result = await syncBranch(pool, supabase, branch, config)
      if (!result.ok) { console.error(`[${branch}] ล้มเหลว: ${result.error}`); process.exitCode = 1; continue }
      const total = result.rows.reduce((sum, row) => sum + row.qty, 0)
      console.log(`[${branch}] ${result.count} SKUs · รวม ${total.toLocaleString('en-US')}`)
      if (dryRun) console.table(result.rows.slice(0, 10).map(({ sku, qty }) => ({ sku, qty })))
    }
  } finally { await pool.close() }
}

// Long-running listener. The MSSQL connection is opened on the first request and reused;
// it is dropped after any failed request so the next one starts from a clean connection.
async function serve() {
  const branchMap = parseBranchMap(env('STOCK_BRANCH_MAP'))
  const supabase = makeSupabase()
  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms))
  let pool = null
  const dropPool = async () => { const old = pool; pool = null; if (old) await old.close().catch(() => {}) }
  console.log(`[serve] เริ่มรอคำขอ (poll ทุก ${POLL_MS / 1000} วินาที) สาขา: ${Object.keys(branchMap).join(', ')}`)
  await recoverStuckRequests(supabase).catch(error => console.error(`[serve] ${error.message}`))
  // Own timer, so the web app still sees the agent as online while a long request is running.
  const beat = () => sendHeartbeat(supabase, 'serve').catch(error => console.error(`[serve] ${error.message}`))
  await beat()
  setInterval(beat, HEARTBEAT_MS)
  for (;;) {
    try {
      const request = await claimNextRequest(supabase)
      if (request) {
        const started = Date.now()
        let outcome
        try {
          pool ??= await connectMssql()
          const results = await Promise.all(Object.entries(branchMap).map(([branch, config]) => syncBranch(pool, supabase, branch, config)))
          outcome = summarizeResults(results)
        } catch (error) {
          outcome = { status: 'error', message: `เชื่อมต่อ MSSQL ไม่สำเร็จ: ${error.message}`.slice(0, 300) }
        }
        if (outcome.status === 'error') await dropPool()
        await finishRequest(supabase, request.id, outcome)
        console.log(`[serve] ${outcome.status} ${((Date.now() - started) / 1000).toFixed(1)} วินาที · ${outcome.message}`)
      }
    } catch (error) {
      // Supabase hiccup, network loss, etc.: log and keep listening.
      console.error(`[serve] ${error.message}`)
      await dropPool()
    }
    await sleep(POLL_MS)
  }
}

async function runCompare() {
  if (!compareFiles.length) throw new Error('ใส่ไฟล์ CSV หลัง --compare เช่น --compare src.csv sss.csv kkl.csv')
  const branchMap = parseBranchMap(env('STOCK_BRANCH_MAP'))
  // Validate every file name first so a typo does not waste a database connection.
  const jobs = compareFiles.map(file => ({ file, branch: branchFromFileName(file) }))
  for (const { branch } of jobs) if (!branchMap[branch]) throw new Error(`STOCK_BRANCH_MAP ไม่มีสาขา ${branch}`)
  const pool = await connectMssql()
  try {
    for (const { file, branch } of jobs) {
      try {
        const csv = readStockCsv(decodeBytes(fs.readFileSync(file)))
        const db = new Map(toStockRows(await readBranch(pool, branchMap[branch]), branch, '').map(row => [row.sku, row.qty]))
        const result = compareStock(csv, db)
        console.log(`\n===== ${branch.toUpperCase()} · ${path.basename(file)} =====`)
        console.log(`SKU     CSV ${result.csvCount} | MSSQL ${result.dbCount}`)
        console.log(`ยอดรวม  CSV ${result.csvTotal.toLocaleString('en-US')} | MSSQL ${result.dbTotal.toLocaleString('en-US')} | ต่าง ${(result.dbTotal - result.csvTotal).toLocaleString('en-US', { maximumFractionDigits: 4 })}`)
        console.log(`ตรงกัน ${result.matched} · qty ต่างกัน ${result.differ.length} · มีแต่ใน CSV ${result.onlyCsv.length} (ไม่ใช่ศูนย์ ${result.onlyCsvNonZero}) · มีแต่ใน MSSQL ${result.onlyDb.length} (ไม่ใช่ศูนย์ ${result.onlyDbNonZero})`)
        if (result.identical) { console.log('✔ ตรงกันทุก SKU'); continue }
        process.exitCode = 1
        if (result.differ.length) { console.log('qty ต่างกัน (diff = MSSQL - CSV) 20 อันดับแรก:'); console.table(result.differ.slice(0, 20)) }
        if (result.onlyCsv.length) { console.log('มีแต่ใน CSV 10 อันดับแรก:'); console.table(result.onlyCsv.slice(0, 10)) }
        if (result.onlyDb.length) { console.log('มีแต่ใน MSSQL 10 อันดับแรก:'); console.table(result.onlyDb.slice(0, 10)) }
        const out = `compare-${branch}.csv`
        fs.writeFileSync(out, differencesToCsv(result))
        console.log(`บันทึกรายการต่างทั้งหมดไว้ที่ ${path.resolve(out)}`)
      } catch (error) {
        console.error(`[${branch}] เทียบไม่สำเร็จ (${file}): ${error.message}`)
        process.exitCode = 1
      }
    }
  } finally { await pool.close() }
}

if (compareIndex >= 0) await runCompare()
else if (serveMode) await serve()
else await runOnce()
