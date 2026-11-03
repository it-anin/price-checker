// Pure helpers for the live MSSQL stock request (no Supabase import, so they are unit-testable).

export const AGENT_ONLINE_SECONDS = 90
export const LIVE_TIMEOUT_MS = 90_000
export const LIVE_POLL_MS = 1_000

export type StockRequestStatus = 'pending' | 'running' | 'done' | 'error'

export interface StockRequestRow {
  id: string
  status: StockRequestStatus
  message: string | null
  requested_at: string
  finished_at: string | null
}

// The agent writes a heartbeat every 30 s; 90 s tolerates two missed beats.
export function agentOnline(lastSeen: string | null | undefined, now = Date.now()): boolean {
  if (!lastSeen) return false
  const time = Date.parse(lastSeen)
  if (!Number.isFinite(time)) return false
  return now - time <= AGENT_ONLINE_SECONDS * 1000
}

export function requestStatusText(row: Pick<StockRequestRow, 'status' | 'message'>): string {
  switch (row.status) {
    case 'pending': return 'รอ agent รับคำขอ'
    case 'running': return 'agent กำลังดึงสต็อกจาก MSSQL'
    case 'done': return `ดึงสำเร็จ${row.message ? ` (${row.message})` : ''}`
    default: return `ดึงไม่สำเร็จ${row.message ? `: ${row.message}` : ''}`
  }
}

// Calls `fetchRow` until the request is done/error or time runs out.
// Injected `sleep`/`now` keep this testable without real waiting.
export async function pollStockRequest(
  fetchRow: () => Promise<StockRequestRow | null>,
  options: { timeoutMs?: number; pollMs?: number; onUpdate?: (row: StockRequestRow) => void; sleep?: (ms: number) => Promise<void>; now?: () => number } = {},
): Promise<StockRequestRow> {
  const timeoutMs = options.timeoutMs ?? LIVE_TIMEOUT_MS
  const pollMs = options.pollMs ?? LIVE_POLL_MS
  const sleep = options.sleep ?? (ms => new Promise<void>(resolve => setTimeout(resolve, ms)))
  const now = options.now ?? Date.now
  const started = now()
  let lastStatus = ''
  for (;;) {
    const row = await fetchRow()
    if (!row) throw new Error('ไม่พบคำขอดึงสต็อก (อาจถูกลบ) กรุณาลองใหม่')
    if (row.status !== lastStatus) { lastStatus = row.status; options.onUpdate?.(row) }
    if (row.status === 'done') return row
    if (row.status === 'error') throw new Error(requestStatusText(row))
    if (now() - started >= timeoutMs) {
      throw new Error(`รอเกิน ${Math.round(timeoutMs / 1000)} วินาที ${row.status === 'pending' ? 'agent ไม่ตอบสนอง (เครื่อง BigyaMainPC อาจปิดอยู่)' : 'MSSQL ตอบช้า'}`)
    }
    await sleep(pollMs)
  }
}
