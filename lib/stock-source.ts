import { supabase } from './supabase'
import { stockReportFromRows, stockSyncWarning, type EngineBranch, type StockReport, type StockRow } from './process-engine'
import { pollStockRequest, type StockRequestRow } from './stock-live'

const PAGE_SIZE = 1000 // Supabase returns at most 1,000 rows per request.

// Reads the stock that tools/stock-sync copied from MSSQL (R01.102 CF_QUANTITY).
export async function loadMssqlStock(branch: EngineBranch): Promise<{ report: StockReport; warning: string }> {
  const status = await supabase.from('stock_sync_status')
    .select('synced_at,ok,message').eq('branch', branch).maybeSingle()
  if (status.error) throw new Error(`อ่านสถานะ sync ของ ${branch.toUpperCase()} ไม่สำเร็จ: ${status.error.message}`)
  if (!status.data) throw new Error(`ยังไม่เคย sync สต็อก ${branch.toUpperCase()} จาก MSSQL กรุณารัน Stock Sync Agent ก่อน`)

  const rows: StockRow[] = []
  for (let from = 0; ; from += PAGE_SIZE) {
    // Stable order is required, otherwise pages can overlap or skip rows.
    const { data, error } = await supabase.from('stock_onhand')
      .select('sku,qty').eq('branch', branch).order('sku').range(from, from + PAGE_SIZE - 1)
    if (error) throw new Error(`อ่านสต็อก ${branch.toUpperCase()} ไม่สำเร็จ: ${error.message}`)
    rows.push(...((data ?? []) as StockRow[]))
    if (!data || data.length < PAGE_SIZE) break
  }
  return {
    report: stockReportFromRows(branch, rows, status.data.synced_at),
    warning: stockSyncWarning(branch, status.data),
  }
}

// ---- Live stock on demand (agent: tools/stock-sync --serve) ----

export async function loadAgentLastSeen(): Promise<string | null> {
  const { data, error } = await supabase.from('stock_agent_status').select('last_seen').eq('id', 1).maybeSingle()
  if (error) throw new Error(`อ่านสถานะ agent ไม่สำเร็จ: ${error.message}`)
  return data?.last_seen ?? null
}

// Asks the agent for fresh stock. Needs a signed-in user; the database enforces it.
export async function requestLiveStock(): Promise<string> {
  const { data, error } = await supabase.rpc('request_stock_sync')
  if (error) {
    if (error.code === '28000') throw new Error('ต้องเข้าสู่ระบบก่อนดึงสต็อกสด')
    if (error.code === 'PGRST202' || error.code === '42883') throw new Error('ฐานข้อมูลยังไม่มีฟังก์ชัน request_stock_sync กรุณารัน migration 202611030001_stock_requests.sql ใน Supabase')
    throw new Error(`สั่งดึงสต็อกสดไม่สำเร็จ: ${error.message}`)
  }
  if (typeof data !== 'string') throw new Error('สั่งดึงสต็อกสดไม่สำเร็จ: ไม่ได้รหัสคำขอ')
  return data
}

export function waitForStockRequest(id: string, onUpdate?: (row: StockRequestRow) => void) {
  return pollStockRequest(async () => {
    const { data, error } = await supabase.from('stock_requests')
      .select('id,status,message,requested_at,finished_at').eq('id', id).maybeSingle()
    if (error) throw new Error(`อ่านสถานะคำขอไม่สำเร็จ: ${error.message}`)
    return (data as StockRequestRow | null) ?? null
  }, { onUpdate })
}
