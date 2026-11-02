// Copies on-hand stock (R01.102 CF_QUANTITY) from MSSQL to Supabase.
// Usage: node sync.mjs [--dry-run] [--watch]
import sql from 'mssql'
import { createClient } from '@supabase/supabase-js'
import { buildConnectionConfig, buildStockQuery, parseBranchMap, toStockRows } from './query.mjs'

const dryRun = process.argv.includes('--dry-run')
const watch = process.argv.includes('--watch')
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

if (watch) {
  const minutes = Number(process.env.SYNC_INTERVAL_MINUTES || 15)
  for (;;) {
    try { await runOnce() } catch (error) { console.error(`รอบนี้ล้มเหลว: ${error.message}`) }
    await new Promise(resolve => setTimeout(resolve, minutes * 60000))
  }
} else {
  await runOnce()
}
