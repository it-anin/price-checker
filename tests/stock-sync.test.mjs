import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import { createRequire } from 'node:module'
import ts from 'typescript'
import { buildConnectionConfig, buildStockQuery, parseBranchMap, toStockRows } from '../tools/stock-sync/query.mjs'
import { branchFromFileName, compareStock, decodeBytes, differencesToCsv, parseCsv, readStockCsv } from '../tools/stock-sync/compare.mjs'

test('query keeps every R01.102 CF_QUANTITY condition', () => {
  const query = buildStockQuery()
  for (const condition of ['H.FTaxInvoiceReplacement = 0', 'H.FVoidTaxInvoice = 0', 'V.FSTOCKMAIN <> 0', 'D.FPROCESS = 1', 'H.FCANCEL = 0', 'H.FAPPROVE = 1',
    "I.ITEMID NOT IN ('DEPOSIT', 'DEPOSIT-NV')", 'I.FEFFECTSTOCK = 1', 'I.FSHOWONREPORT = 1', 'I.FUSED = 1', 'SUM(D.BASEQUANTITY * V.FSTOCKMAIN)', 'H.SYSBRANCHID = @branchId']) {
    assert.ok(query.includes(condition), condition)
  }
  assert.ok(!query.includes('SYSWAREHOUSEID IN'))
  assert.ok(!query.includes('//'), 'PowerBuilder comments must not leak into T-SQL')
})

test('query binds warehouse ids as parameters', () => {
  assert.ok(buildStockQuery([4, 7]).includes('H.SYSWAREHOUSEID IN (@w0, @w1)'))
})

test('parseBranchMap validates the config', () => {
  assert.deepEqual(parseBranchMap('{"src":{"branchId":1},"sss":{"branchId":2,"warehouseIds":[3]}}'), { src: { branchId: 1, warehouseIds: [] }, sss: { branchId: 2, warehouseIds: [3] } })
  assert.throws(() => parseBranchMap('nope'), /JSON/)
  assert.throws(() => parseBranchMap('{}'), /อย่างน้อยหนึ่งสาขา/)
  assert.throws(() => parseBranchMap('{"src":{"branchId":"1"}}'), /branchId/)
  assert.throws(() => parseBranchMap('{"src":{"branchId":1,"warehouseIds":["x"]}}'), /warehouseIds/)
})

test('toStockRows trims, merges duplicates, rounds and skips null/invalid', () => {
  const rows = toStockRows([{ sku: ' a1 ', qty: 0.1 }, { sku: 'A1', qty: 0.2 }, { sku: '', qty: 5 }, { sku: 'B2', qty: null }, { sku: 'C3', qty: 'x' }], 'src', 't')
  assert.deepEqual(rows, [{ branch: 'src', sku: 'A1', qty: 0.3, synced_at: 't' }])
})

test('buildConnectionConfig: named instance replaces the port', () => {
  const base = { MSSQL_SERVER: 'BigyaMainPC', MSSQL_DATABASE: 'MSMAXX', MSSQL_USER: 'web_readonly', MSSQL_PASSWORD: 'x' }
  const named = buildConnectionConfig({ ...base, MSSQL_INSTANCE: 'SQLSERVER', MSSQL_PORT: '1433' })
  assert.equal(named.options.instanceName, 'SQLSERVER')
  assert.ok(!('port' in named), 'port must not be sent with instanceName')
  assert.equal(named.server, 'BigyaMainPC')
  assert.equal(named.database, 'MSMAXX')
})

test('buildConnectionConfig: default instance uses port, defaulting to 1433', () => {
  const base = { MSSQL_SERVER: 'h', MSSQL_DATABASE: 'd', MSSQL_USER: 'u', MSSQL_PASSWORD: 'p' }
  assert.equal(buildConnectionConfig(base).port, 1433)
  assert.equal(buildConnectionConfig({ ...base, MSSQL_PORT: '14330' }).port, 14330)
  assert.ok(!('instanceName' in buildConnectionConfig(base).options))
  assert.equal(buildConnectionConfig(base).options.trustServerCertificate, true)
  assert.equal(buildConnectionConfig({ ...base, MSSQL_ENCRYPT: 'true' }).options.encrypt, true)
})

test('buildConnectionConfig: missing required values are reported by name', () => {
  const base = { MSSQL_SERVER: 'h', MSSQL_DATABASE: 'd', MSSQL_USER: 'u', MSSQL_PASSWORD: 'p' }
  for (const name of Object.keys(base)) {
    assert.throws(() => buildConnectionConfig({ ...base, [name]: '' }), new RegExp(name))
  }
})

const csvText = [
  'WH,BR,ID,NAME,SKU,UNIT,CF_QUANTITY',
  '1,x,y,"ยา, ""ไทย""",ab-1,box,"1,200.50"',
  '1,x,y,UM Splint (M/2-3") NEW,AB-1,box,0.5',
  '1,x,y,plain,C2,box,0',
  '1,x,y,Total,TOTAL,box,999',
  ',,,,,,',
].join('\r\n') + '\r\n'

test('readStockCsv follows the web app rules (column E SKU, column G qty, duplicates summed)', () => {
  const stock = readStockCsv(csvText)
  assert.deepEqual([...stock], [['AB-1', 1201], ['C2', 0]])
})

test('readStockCsv works with TIS-620 bytes and other delimiters', () => {
  assert.equal(decodeBytes(Buffer.from([0xa1, 0xd2])), 'กา')
  assert.equal(decodeBytes(Buffer.from('\uFEFFabc', 'utf8')), 'abc')
  assert.deepEqual([...readStockCsv('a;b;c;d;sku1;f;5\n')], [['SKU1', 5]])
  assert.deepEqual(parseCsv('a\tb\n1\t2\n'), [['a', 'b'], ['1', '2']])
})

test('readStockCsv rejects non-numeric quantities and empty files', () => {
  assert.throws(() => readStockCsv('a,b,c,d,SKU9,f,abc\n'), /ไม่ใช่ตัวเลข/)
  assert.throws(() => readStockCsv('nothing,here\n'), /ไม่พบข้อมูล/)
})

test('branchFromFileName accepts only src/sss/kkl csv names', () => {
  assert.equal(branchFromFileName('C:\\price-checker\\SRC.CSV'), 'src')
  assert.equal(branchFromFileName('/tmp/kkl.csv'), 'kkl')
  assert.throws(() => branchFromFileName('stock.csv'), /src\.csv/)
})

test('compareStock reports identical, differing and one-sided SKUs', () => {
  const same = compareStock(new Map([['A', 1], ['B', 2.5]]), new Map([['A', 1], ['B', 2.5]]))
  assert.equal(same.identical, true)
  assert.equal(same.matched, 2)

  const result = compareStock(
    new Map([['A', 10], ['B', 5], ['C', 0], ['D', 3]]),
    new Map([['A', 10], ['B', 7], ['E', 4], ['F', 0]]),
  )
  assert.equal(result.identical, false)
  assert.equal(result.matched, 1)
  assert.deepEqual(result.differ, [{ sku: 'B', csv: 5, db: 7, diff: 2 }])
  assert.deepEqual(result.onlyCsv.map(row => row.sku), ['D', 'C'])
  assert.equal(result.onlyCsvNonZero, 1)
  assert.deepEqual(result.onlyDb.map(row => row.sku), ['E', 'F'])
  assert.equal(result.onlyDbNonZero, 1)
  assert.equal(result.csvTotal, 18)
  assert.equal(result.dbTotal, 21)
})

test('compareStock ignores float noise below tolerance', () => {
  assert.equal(compareStock(new Map([['A', 0.1 + 0.2]]), new Map([['A', 0.3]])).identical, true)
})

test('differencesToCsv lists every difference with a BOM', () => {
  const result = compareStock(new Map([['B', 5], ['D', 3]]), new Map([['B', 7], ['E', 4]]))
  const text = differencesToCsv(result)
  assert.ok(text.startsWith('\uFEFFsku,status,csv_qty,mssql_qty,diff'))
  assert.ok(text.includes('B,qty_differs,5,7,2'))
  assert.ok(text.includes('D,only_in_csv,3,,'))
  assert.ok(text.includes('E,only_in_mssql,,4,'))
})

test('compare CSV reader gives the same stock as the web app readStockReport', () => {
  const source = fs.readFileSync(new URL('../lib/process-engine.ts', import.meta.url), 'utf8')
  const compiled = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText
  const holder = { exports: {} }
  new Function('exports', 'require', 'module', compiled)(holder.exports, createRequire(import.meta.url), holder)
  const sample = [csvText, csvText.replace(/,/g, ';').replace('1;200.50', '1200.50')]
  for (const text of sample) {
    assert.deepEqual([...readStockCsv(text)], [...holder.exports.readStockReport('src.csv', text, 'UTF-8').stock])
  }
})

// ---- serve mode helpers ----
import { chunkRows } from '../tools/stock-sync/query.mjs'
import { claimNextRequest, finishRequest, recoverStuckRequests, summarizeResults } from '../tools/stock-sync/serve.mjs'

test('chunkRows splits evenly and rejects bad sizes', () => {
  assert.deepEqual(chunkRows([1, 2, 3, 4, 5], 2), [[1, 2], [3, 4], [5]])
  assert.deepEqual(chunkRows([], 3), [])
  assert.throws(() => chunkRows([1], 0), /chunk size/)
})

test('summarizeResults closes the request as done or error with a short message', () => {
  assert.deepEqual(
    summarizeResults([{ branch: 'src', ok: true, count: 5461 }, { branch: 'sss', ok: true, count: 3287 }]),
    { status: 'done', message: 'src 5461, sss 3287' })
  const failed = summarizeResults([{ branch: 'src', ok: true, count: 1 }, { branch: 'kkl', ok: false, error: 'timeout' }])
  assert.equal(failed.status, 'error')
  assert.equal(failed.message, 'kkl: timeout')
  assert.equal(summarizeResults([]).status, 'error')
  assert.ok(summarizeResults([{ branch: 'x', ok: false, error: 'e'.repeat(1000) }]).message.length <= 300)
})

// Minimal in-memory stand-in for the supabase-js calls serve.mjs makes on stock_requests.
function fakeSupabase(rows) {
  const builder = (filters = [], patch = null, order = null, max = Infinity) => {
    const matching = () => rows.filter(r => filters.every(f => f(r)))
      .sort((a, b) => (order ? String(a[order]).localeCompare(String(b[order])) : 0)).slice(0, max)
    const self = {
      select: () => self,
      eq: (key, value) => builder([...filters, r => r[key] === value], patch, order, max),
      in: (key, values) => builder([...filters, r => values.includes(r[key])], patch, order, max),
      order: key => builder(filters, patch, key, max),
      limit: n => builder(filters, patch, order, n),
      update: p => builder(filters, p, order, max),
      maybeSingle: async () => { const hit = matching()[0]; if (hit && patch) Object.assign(hit, patch); return { data: hit ? { id: hit.id } : null, error: null } },
      then: (resolve) => {
        const hits = matching()
        if (patch) hits.forEach(h => Object.assign(h, patch))
        return Promise.resolve({ data: hits.map(h => ({ id: h.id })), error: null }).then(resolve)
      },
    }
    return self
  }
  return { from: () => builder() }
}

test('claimNextRequest takes the oldest pending request exactly once', async () => {
  const rows = [
    { id: 'b', status: 'pending', requested_at: '2026-11-03T00:00:02Z' },
    { id: 'a', status: 'pending', requested_at: '2026-11-03T00:00:01Z' },
    { id: 'c', status: 'done', requested_at: '2026-11-03T00:00:00Z' },
  ]
  const db = fakeSupabase(rows)
  assert.deepEqual(await claimNextRequest(db), { id: 'a' })
  assert.equal(rows.find(r => r.id === 'a').status, 'running')
  assert.deepEqual(await claimNextRequest(db), { id: 'b' })
  assert.equal(await claimNextRequest(db), null)
})

test('finishRequest and recoverStuckRequests update the right rows', async () => {
  const rows = [{ id: 'x', status: 'running' }, { id: 'y', status: 'pending' }, { id: 'z', status: 'done' }]
  const db = fakeSupabase(rows)
  await finishRequest(db, 'x', { status: 'done', message: 'src 1' })
  assert.deepEqual([rows[0].status, rows[0].message], ['done', 'src 1'])
  assert.ok(rows[0].finished_at)
  await recoverStuckRequests(db)
  assert.equal(rows[1].status, 'error')
  assert.equal(rows[2].status, 'done')
})
