import { supabase } from './supabase'
import { stockReportFromRows, stockSyncWarning, type EngineBranch, type StockReport, type StockRow } from './process-engine'

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
