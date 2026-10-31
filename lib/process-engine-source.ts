import { supabase } from './supabase'
import { pendingSkuSets } from './process-engine'
import { mutatePendingEntries, nextPendingTimestamp, normalizePendingEntries, type PendingMutation, type PendingSnapshot } from './pending-sku'

export const PENDING_SOURCE = { table: 'app_settings', key: 'pending_skus' }

export async function loadPendingSnapshot(): Promise<PendingSnapshot> {
  const { data, error } = await supabase.from(PENDING_SOURCE.table)
    .select('value,updated_at').eq('key', PENDING_SOURCE.key).maybeSingle()
  if (error) throw new Error(`อ่าน PENDING SKU UPDATE ไม่สำเร็จ: ${error.message}`)
  if (!data) throw new Error('ยังไม่พบ pending_skus ใน Supabase ปัจจุบัน กรุณาย้ายข้อมูล PENDING เดิมก่อน RUN')
  if (typeof data.updated_at !== 'string' || !Number.isFinite(Date.parse(data.updated_at))) throw new Error('อ่านเวอร์ชัน PENDING ไม่สำเร็จ กรุณาตรวจ updated_at ก่อนบันทึก')
  return { entries: normalizePendingEntries(data.value), updatedAt: data.updated_at }
}

export async function loadPendingSkuSets() {
  return pendingSkuSets((await loadPendingSnapshot()).entries)
}

export async function savePendingMutation(mutation: PendingMutation): Promise<PendingSnapshot> {
  // Every write starts from current data. The short timestamp filter avoids
  // putting the entire JSON list into a URL, and protects simultaneous edits.
  for (let attempt = 0; attempt < 3; attempt++) {
    const snapshot = await loadPendingSnapshot()
    const next = mutatePendingEntries(snapshot.entries, mutation)
    if (!next.changed) return snapshot
    const updatedAt = nextPendingTimestamp(snapshot.updatedAt)
    const { data, error } = await supabase.from(PENDING_SOURCE.table)
      .update({ value: JSON.stringify(next.entries), updated_at: updatedAt })
      .eq('key', PENDING_SOURCE.key).eq('updated_at', snapshot.updatedAt)
      .select('value,updated_at').maybeSingle()
    if (error) throw new Error(`บันทึก PENDING ไม่สำเร็จ: ${error.message}`)
    if (data) return { entries: normalizePendingEntries(data.value), updatedAt: data.updated_at }
  }
  throw new Error('มีผู้ใช้อื่นแก้ไข PENDING พร้อมกัน กรุณาโหลดรายการใหม่แล้วลองอีกครั้ง ข้อมูลที่กรอกยังอยู่')
}
