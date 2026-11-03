import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import ts from 'typescript'
import XLSX from 'xlsx'

const skuKey = 'รหัสสินค้า (SKU NUMBER)'
const headers = ['*ประเภทสินค้า', '*ชื่อสินค้า', '*เลขที่ใบอนุญาตโฆษณา', '*ราคาสินค้า', 'รหัสสินค้า', '*รูปภาพสินค้า', 'หมวดหมู่รายการสินค้า']
const source = ts.createSourceFile('page.tsx', fs.readFileSync(new URL('../app/page.tsx', import.meta.url), 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX)
const functions = []
function visit(node) {
  if (ts.isFunctionDeclaration(node) && ['exportPendingGmMmeXlsx', 'toExportRows', 'productPrice', 'parsePriceRobust', 'engineMessage'].includes(node.name?.text)) functions.push(node.getText(source))
  ts.forEachChild(node, visit)
}
visit(source)
assert.equal(functions.length, 5)
const code = ts.transpileModule(functions.join('\n'), { compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.CommonJS } }).outputText
const product = (sku, price = 10) => ({ [skuKey]: sku, '*ชื่อสินค้า (NAME)': `ชื่อสินค้า ${sku}`, 'ประเภทสินค้า': 'อัปเดตรายการสินค้า', '*เลขที่ใบอนุญาตโฆษณา': '000123', '*ราคาสินค้า': price, '*รูปภาพสินค้า': `https://drive.google.com/file/d/${sku}/view?usp=sharing`, 'หมวดหมู่สินค้า (CATEGORIES)': 'ยา', promaxx_price: 999, new_promaxx_price: 888, src_price: 777 })
const entry = (sku, branches = ['SRC']) => ({ sku, name: 'ชื่อเก่าใน PENDING', branches })

function harness(entries, products, options = {}) {
  const state = { snapshots: 0, calls: [], files: [], notice: null, busy: '', active: 0, maxActive: 0 }
  const runtime = {
    engineBusyRef: { current: Boolean(options.busy) },
    setEngineBusy: value => { state.busy = value },
    setPendingExportNotice: value => { state.notice = value },
    setStatus: value => { state.status = value },
    engineDate: () => '03-11-2026',
    normalizeStockSku: value => String(value ?? '').trim(),
    loadPendingSnapshot: async () => {
      state.snapshots++
      if (options.snapshotError) throw new Error('อ่าน PENDING ไม่สำเร็จ')
      return { entries: options.currentEntries ? options.currentEntries() : entries, updatedAt: '2026-11-03T00:00:00Z' }
    },
    fetch: async (url, init) => {
      assert.equal(url, '/api/products')
      assert.equal(init.method, 'POST')
      const body = JSON.parse(init.body)
      assert.deepEqual(Object.keys(body), ['skus'])
      assert.ok(body.skus.length <= 300)
      state.calls.push(body.skus)
      state.active++; state.maxActive = Math.max(state.maxActive, state.active)
      try {
        await new Promise(resolve => setImmediate(resolve))
        const failed = body.skus.includes(options.failSku)
        return { ok: !failed, json: async () => options.badBody ? { success: true, products: null } : { success: true, products: products.filter(row => body.skus.includes(row[skuKey])).reverse() } }
      } finally { state.active-- }
    },
    require: name => {
      assert.equal(name, 'xlsx')
      return { ...XLSX, writeFile: (wb, filename) => { state.files.push({ wb, filename }) } }
    },
  }
  const run = new Function(...Object.keys(runtime), code + '\nreturn exportPendingGmMmeXlsx')(...Object.values(runtime))
  return { run, state, busyRef: runtime.engineBusyRef }
}

test('exports only the active branch, unions saved sources once per SKU and uses current DB product fields', async () => {
  const entries = [{ ...entry('00001', ['SRC', 'KKL']), manualBranches: ['SRC'], autoPriceBranches: ['KKL'] }, { ...entry('1'), manualBranches: [], autoPriceBranches: ['SRC'] }, entry('OTHER', ['SSS']), entry('LEGACY', ['WRD'])]
  const h = harness(entries, [product('00001', 0), product('1', '฿ 1,234.50'), product('OTHER'), product('LEGACY')])
  await h.run('src')
  assert.equal(h.state.snapshots, 1)
  assert.deepEqual(h.state.calls, [['00001', '1']])
  const { wb, filename } = h.state.files[0]
  assert.deepEqual(wb.SheetNames, ['TEMPLATE'])
  assert.equal(filename, 'GM MME PENDING SRC 03-11-2026.xlsx')
  const ws = wb.Sheets.TEMPLATE
  assert.deepEqual(XLSX.utils.sheet_to_json(ws, { header: 1 })[0], headers)
  assert.equal(ws.D2.t, 'n'); assert.equal(ws.D2.v, 0)
  assert.equal(ws.D3.v, 1234.5)
  assert.equal(ws.E2.t, 's'); assert.equal(ws.E2.v, '00001')
  assert.equal(ws.E3.v, '1')
  assert.equal(ws.B2.v, 'ชื่อสินค้า 00001')
  assert.equal(ws.C2.v, '000123')
  assert.equal(ws.F2.v, product('00001')['*รูปภาพสินค้า'])
  assert.equal(ws.A2.v, 'อัปเดตรายการสินค้า')
  assert.equal(ws.D2.z, '#,##0.00')
  assert.equal(ws['!cols'].length, 7)
  assert.equal(ws['!rows'][1].hpt, 141)
  const reloaded = XLSX.read(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }))
  assert.equal(reloaded.Sheets.TEMPLATE.E2.v, '00001')
  assert.equal(reloaded.Sheets.TEMPLATE.D2.v, 0)
  assert.equal(h.state.notice.error, false)
  assert.equal(h.state.busy, ''); assert.equal(h.busyRef.current, false)
})

test('each export re-reads saved PENDING, and changing branches does not reuse stale membership', async () => {
  let current = [entry('SRC', ['SRC'])]
  const h = harness([], [product('SRC'), product('SSS')], { currentEntries: () => current })
  await h.run('src')
  current = [entry('SSS', ['SSS'])]
  await h.run('sss')
  assert.equal(h.state.snapshots, 2)
  assert.deepEqual(h.state.calls, [['SRC'], ['SSS']])
  assert.match(h.state.files[1].filename, /PENDING SSS/)
})

test('more than 1,000 pending products export completely with bounded 300-SKU lookups', async () => {
  const skus = Array.from({ length: 1501 }, (_, index) => String(index).padStart(5, '0'))
  const h = harness(skus.map(sku => entry(sku)), skus.map(sku => product(sku)))
  await h.run('src')
  assert.deepEqual(h.state.calls.map(chunk => chunk.length), [300, 300, 300, 300, 300, 1])
  assert.ok(h.state.maxActive <= 4)
  assert.equal(XLSX.utils.sheet_to_json(h.state.files[0].wb.Sheets.TEMPLATE).length, 1501)
})

test('missing product records are explicitly counted and zero is distinct from a missing DB price', async () => {
  const h = harness([entry('00001'), entry('MISSING')], [product('00001', null)])
  await h.run('src')
  assert.equal(h.state.files.length, 1)
  assert.equal(h.state.files[0].wb.Sheets.TEMPLATE.D2.v, '')
  assert.equal(h.state.notice.error, true)
  assert.match(h.state.notice.message, /1\/2/)
  assert.match(h.state.notice.message, /MISSING/)
})

test('empty branches and absent product data never download an empty or misleading file', async () => {
  for (const entries of [[], [entry('OTHER', ['KKL'])], [entry('MISSING')]]) {
    const h = harness(entries, [])
    await h.run('src')
    assert.equal(h.state.files.length, 0)
    assert.ok(h.state.notice.message)
    assert.equal(h.busyRef.current, false)
  }
})

test('snapshot errors, failed product batches and malformed replies prevent partial downloads and allow retry', async () => {
  const entries = Array.from({ length: 301 }, (_, index) => entry(String(index)))
  for (const options of [{ snapshotError: true }, { failSku: '300' }, { badBody: true }]) {
    const h = harness(entries, entries.map(row => product(row.sku)), options)
    await h.run('src')
    assert.equal(h.state.files.length, 0)
    assert.equal(h.state.notice.error, true)
    assert.equal(h.state.busy, '')
    assert.equal(h.busyRef.current, false)
  }
})

test('the shared Engine lock prevents a second export or reads during another operation', async () => {
  const h = harness([entry('001')], [product('001')], { busy: true })
  await h.run('src')
  assert.equal(h.state.snapshots, 0)
  assert.equal(h.state.calls.length, 0)
  assert.equal(h.state.files.length, 0)
})
