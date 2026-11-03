const { test } = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const ts = require('typescript')

const file = path.join(__dirname, '../lib/stock-live.ts')
const compiled = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText
const live = { exports: {} }
new Function('exports', 'require', 'module', compiled)(live.exports, require, live)
const { agentOnline, requestStatusText, pollStockRequest, AGENT_ONLINE_SECONDS } = live.exports

const row = (status, message = null) => ({ id: 'r1', status, message, requested_at: 't0', finished_at: null })

test('agentOnline: recent heartbeat is online, old/missing/invalid is not', () => {
  const now = Date.parse('2026-11-03T02:00:00.000Z')
  assert.equal(agentOnline('2026-11-03T01:59:40.000Z', now), true)
  assert.equal(agentOnline(new Date(now - AGENT_ONLINE_SECONDS * 1000).toISOString(), now), true)
  assert.equal(agentOnline(new Date(now - AGENT_ONLINE_SECONDS * 1000 - 1).toISOString(), now), false)
  assert.equal(agentOnline(null, now), false)
  assert.equal(agentOnline(undefined, now), false)
  assert.equal(agentOnline('not a date', now), false)
})

test('requestStatusText describes every status in Thai', () => {
  assert.match(requestStatusText(row('pending')), /รอ agent/)
  assert.match(requestStatusText(row('running')), /กำลังดึง/)
  assert.equal(requestStatusText(row('done', 'src 5461, sss 3287')), 'ดึงสำเร็จ (src 5461, sss 3287)')
  assert.equal(requestStatusText(row('error', 'login failed')), 'ดึงไม่สำเร็จ: login failed')
  assert.equal(requestStatusText(row('error')), 'ดึงไม่สำเร็จ')
})

function fakeClock() {
  let time = 0
  return { now: () => time, sleep: async ms => { time += ms } }
}

test('pollStockRequest returns when the request is done and reports each status change once', async () => {
  const states = [row('pending'), row('pending'), row('running'), row('running'), row('done', 'ok')]
  const seen = []
  const clock = fakeClock()
  const result = await pollStockRequest(async () => states.shift(), { ...clock, onUpdate: update => seen.push(update.status) })
  assert.equal(result.status, 'done')
  assert.deepEqual(seen, ['pending', 'running', 'done'])
})

test('pollStockRequest throws the agent message when the request failed', async () => {
  const clock = fakeClock()
  await assert.rejects(pollStockRequest(async () => row('error', 'Login failed'), clock), /ดึงไม่สำเร็จ: Login failed/)
})

test('pollStockRequest times out and says whether the agent never answered', async () => {
  const clock = fakeClock()
  await assert.rejects(pollStockRequest(async () => row('pending'), { ...clock, timeoutMs: 5000, pollMs: 1000 }), /agent ไม่ตอบสนอง/)
  const clock2 = fakeClock()
  await assert.rejects(pollStockRequest(async () => row('running'), { ...clock2, timeoutMs: 5000, pollMs: 1000 }), /MSSQL ตอบช้า/)
  assert.ok(clock2.now() >= 5000, 'must wait the full timeout before giving up')
})

test('pollStockRequest fails clearly when the request row disappears', async () => {
  await assert.rejects(pollStockRequest(async () => null, fakeClock()), /ไม่พบคำขอ/)
})
