// Copies on-hand stock (R01.102 CF_QUANTITY) from MSSQL to Supabase.
// Usage: node sync.mjs [--dry-run] [--watch]
//        node sync.mjs --compare src.csv [sss.csv kkl.csv]   (read-only check against R01.102 CSV exports)
import fs from 'node:fs'
import path from 'node:path'
import sql from 'mssql'
import { createClient } from '@supabase/supabase-js'
import { buildConnectionConfig, buildStockQuery, parseBranchMap, toStockRows } from './query.mjs'
import { branchFromFileName, compareStock, decodeBytes, differencesToCsv, readStockCsv } from './compare.mjs'

const dryRun = process.argv.includes('--dry-run')
const watch = process.argv.includes('--watch')
const compareIndex = process.argv.indexOf('--compare')
const compareFiles = compareIndex < 0 ? [] : process.argv.slice(compareIndex + 1).filter(arg => !arg.startsWith('--'))
const CHUNK = 300

function env(name) {
  const value = process.env[name]
  if (!value) throw new Error(`ยังไม่ได้ตั้ง ${name}`)
  return value
}

async function readBranch(pool, { branchId, warehouseIds }) {
  const request = pool.request().input('branchId', sql.Int, branchId)
  warehouseIds.forEach((id, i) => request.input(`w${i}`, sql.Int, id))
  const result = await request.query(buildStockQuery(warehouseIds))
  return result.recordset
}

async function writeBranch(supabase, branch, rows, syncedAt) {
  for (let i = 0; i < rows.length; i += CHUNK) {
    const { error } = await supabase.from('stock_onhand').upsert(rows.slice(i, i + CHUNK), { onConflict: 'branch,sku' })
    if (error) throw new Error(`upsert ไม่สำเร็จ: ${error.message}`)
  }
  // Only after every chunk succeeded: drop SKUs that no longer appear for this branch.
  const { error } = await supabase.from('stock_onhand').delete().eq('branch', branch).lt('synced_at', syncedAt)
  if (error) throw new Error(`ลบ SKU เก่าไม่สำเร็จ: ${error.message}`)
}

async function setStatus(supabase, branch, status) {
  const { error } = await supabase.from('stock_sync_status').upsert({ branch, ...status }, { onConflict: 'branch' })
  if (error) console.error(`[${branch}] เขียนสถานะไม่สำเร็จ: ${error.message}`)
}

async function runOnce() {
  const branchMap = parseBranchMap(env('STOCK_BRANCH_MAP'))
  const supabase = dryRun ? null : createClient(env('SUPABASE_URL'), env('SUPABASE_SERVICE_ROLE_KEY'), { auth: { persistSession: false } })
  const pool = await new sql.ConnectionPool(buildConnectionConfig(process.env)).connect()
  try {
    for (const [branch, config] of Object.entries(branchMap)) {
      const syncedAt = new Date().toISOString()
      try {
        const rows = toStockRows(await readBranch(pool, config), branch, syncedAt)
        // An empty result almost always means a wrong branchId; never wipe good data for it.
        if (!rows.length) throw new Error('query ไม่พบสต็อกเลย ตรวจ branchId/warehouseIds')
        const total = rows.reduce((sum, row) => sum + row.qty, 0)
        console.log(`[${branch}] ${rows.length} SKUs · รวม ${total.toLocaleString('en-US')}`)
        if (dryRun) { console.table(rows.slice(0, 10).map(({ sku, qty }) => ({ sku, qty }))); continue }
        await writeBranch(supabase, branch, rows, syncedAt)
        await setStatus(supabase, branch, { synced_at: syncedAt, row_count: rows.length, ok: true, message: null })
      } catch (error) {
        console.error(`[${branch}] ล้มเหลว: ${error.message}`)
        if (supabase) await setStatus(supabase, branch, { synced_at: syncedAt, row_count: 0, ok: false, message: String(error.message).slice(0, 200) })
        process.exitCode = 1
      }
    }
  } finally { await pool.close() }
}

async function runCompare() {
  if (!compareFiles.length) throw new Error('ใส่ไฟล์ CSV หลัง --compare เช่น --compare src.csv sss.csv kkl.csv')
  const branchMap = parseBranchMap(env('STOCK_BRANCH_MAP'))
  // Validate every file name first so a typo does not waste a database connection.
  const jobs = compareFiles.map(file => ({ file, branch: branchFromFileName(file) }))
  for (const { branch } of jobs) if (!branchMap[branch]) throw new Error(`STOCK_BRANCH_MAP ไม่มีสาขา ${branch}`)
  const pool = await new sql.ConnectionPool(buildConnectionConfig(process.env)).connect()
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

if (compareIndex >= 0) {
  await runCompare()
} else if (watch) {
  const minutes = Number(process.env.SYNC_INTERVAL_MINUTES || 15)
  for (;;) {
    try { await runOnce() } catch (error) { console.error(`รอบนี้ล้มเหลว: ${error.message}`) }
    await new Promise(resolve => setTimeout(resolve, minutes * 60000))
  }
} else {
  await runOnce()
}
