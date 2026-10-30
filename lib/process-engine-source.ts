import { supabase } from './supabase'
import { pendingSkuSets } from './process-engine'

export const PENDING_SOURCE = { table: 'app_settings', key: 'pending_skus' }

export async function loadPendingSkuSets() {
  const { data, error } = await supabase.from(PENDING_SOURCE.table)
    .select('value').eq('key', PENDING_SOURCE.key).maybeSingle()
  if (error) throw new Error(`อ่าน PENDING SKU UPDATE ไม่สำเร็จ: ${error.message}`)
  if (!data) throw new Error('ยังไม่พบ pending_skus ใน Supabase ปัจจุบัน กรุณาย้ายข้อมูล PENDING เดิมก่อน RUN')
  let raw: unknown = data.value
  if (typeof raw === 'string') {
    try { raw = JSON.parse(raw) } catch { throw new Error('pending_skus ใน Supabase ไม่ใช่ JSON ที่ถูกต้อง') }
  }
  return pendingSkuSets(raw)
}
