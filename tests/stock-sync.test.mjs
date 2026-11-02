import { test } from 'node:test'
import assert from 'node:assert/strict'
import { buildStockQuery, parseBranchMap, toStockRows } from '../tools/stock-sync/query.mjs'

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
