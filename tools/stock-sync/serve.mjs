// Listener mode (`sync.mjs --serve`): wait for requests made by the web app and answer them
// with live MSSQL stock. Pure helpers + thin Supabase calls; the sync itself lives in sync.mjs.

export const POLL_MS = 2000
export const HEARTBEAT_MS = 30000

// results: [{ branch, ok, count?, error? }] -> how the request row should be closed.
export function summarizeResults(results) {
  const failed = results.filter(result => !result.ok)
  if (!results.length) return { status: 'error', message: 'ไม่มีสาขาให้ดึง ตรวจ STOCK_BRANCH_MAP' }
  if (!failed.length) {
    return { status: 'done', message: results.map(result => `${result.branch} ${result.count}`).join(', ') }
  }
  const message = failed.map(result => `${result.branch}: ${result.error}`).join(' | ')
  return { status: 'error', message: message.slice(0, 300) }
}

export async function recoverStuckRequests(supabase) {
  const { error } = await supabase.from('stock_requests')
    .update({ status: 'error', finished_at: new Date().toISOString(), message: 'agent รีสตาร์ทระหว่างดึงข้อมูล' })
    .in('status', ['pending', 'running'])
  if (error) throw new Error(`ล้างคำขอค้างไม่สำเร็จ: ${error.message}`)
}

// Takes the oldest pending request. The status filter makes the claim atomic: if another
// process took it first, no row matches and we simply get null.
export async function claimNextRequest(supabase) {
  const pending = await supabase.from('stock_requests')
    .select('id').eq('status', 'pending').order('requested_at').limit(1)
  if (pending.error) throw new Error(`อ่านคำขอไม่สำเร็จ: ${pending.error.message}`)
  const id = pending.data?.[0]?.id
  if (!id) return null
  const claimed = await supabase.from('stock_requests')
    .update({ status: 'running', started_at: new Date().toISOString() })
    .eq('id', id).eq('status', 'pending').select('id').maybeSingle()
  if (claimed.error) throw new Error(`รับคำขอไม่สำเร็จ: ${claimed.error.message}`)
  return claimed.data ?? null
}

export async function finishRequest(supabase, id, outcome) {
  const { error } = await supabase.from('stock_requests')
    .update({ status: outcome.status, message: outcome.message, finished_at: new Date().toISOString() })
    .eq('id', id)
  if (error) throw new Error(`ปิดคำขอไม่สำเร็จ: ${error.message}`)
}

export async function sendHeartbeat(supabase, info) {
  const { error } = await supabase.from('stock_agent_status')
    .upsert({ id: 1, last_seen: new Date().toISOString(), info }, { onConflict: 'id' })
  if (error) throw new Error(`เขียน heartbeat ไม่สำเร็จ: ${error.message}`)
}
