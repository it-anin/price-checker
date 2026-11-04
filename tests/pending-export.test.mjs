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
  if (ts.isFunctionDeclaration(node) && ['exportPendingGmMmeXlsx', 'loadPendingExportProducts', 'pendingPromaxxRows', 'reviewPendingPromaxxExport', 'closePendingPromaxxReview', 'hasProductBranch', 'pricesDiffer', 'priceState', 'toExportRows', 'productPrice', 'parsePriceRobust', 'engineMessage'].includes(node.name?.text)) functions.push(node.getText(source))
  ts.forEachChild(node, visit)
}
visit(source)
assert.equal(functions.length, 12)
const code = ts.transpileModule(functions.join('\n'), { compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.CommonJS } }).outputText
const product = (sku, price = 10) => ({ [skuKey]: sku, '*ชื่อสินค้า (NAME)': `ชื่อสินค้า ${sku}`, 'ประเภทสินค้า': 'อัปเดตรายการสินค้า', '*เลขที่ใบอนุญาตโฆษณา': '000123', '*ราคาสินค้า': price, '*รูปภาพสินค้า': `https://drive.google.com/file/d/${sku}/view?usp=sharing`, 'หมวดหมู่สินค้า (CATEGORIES)': 'ยา', promaxx_price: 999, new_promaxx_price: 888, src_price: 777 })
const entry = (sku, branches = ['SRC']) => ({ sku, name: 'ชื่อเก่าใน PENDING', branches })

function harness(entries, products, options = {}) {
  const state = { snapshots: 0, calls: [], files: [], notice: null, busy: '', active: 0, maxActive: 0, review: null }
  const runtime = {
    engineBusyRef: { current: Boolean(options.busy) },
    pendingPromaxxRequest: { current: 0 },
    inspectionBranch: options.branch || 'src',
    setPendingPromaxxReview: value => { state.review = value },
    setEngineBusy: value => { state.busy = value },
    setPendingExportNotice: value => { state.notice = value },
    setStatus: value => { state.status = value },
    engineDate: () => '03-11-2026',
    normalizeStockSku: value => String(value ?? '').trim(),
    loadPendingSnapshot: async () => {
      state.snapshots++
      if (options.waitForSnapshot) await options.waitForSnapshot()
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
        return { ok: !failed, json: async () => options.badBody ? { success: true, products: null } : { success: true, products: (options.currentProducts ? options.currentProducts() : products).filter(row => body.skus.includes(row[skuKey])).reverse() } }
      } finally { state.active-- }
    },
    require: name => {
      assert.equal(name, 'xlsx')
      return { ...XLSX, writeFile: (wb, filename) => { state.files.push({ wb, filename }) } }
    },
  }
  const actions = new Function(...Object.keys(runtime), code + '\nreturn { run: exportPendingGmMmeXlsx, preview: reviewPendingPromaxxExport, close: closePendingPromaxxReview, priceState }')(...Object.values(runtime))
  return { ...actions, state, busyRef: runtime.engineBusyRef }
}

test('exports only the active branch, unions saved sources once per SKU and uses current DB product fields', async () => {
  const entries = [{ ...entry('00001', ['SRC', 'KKL']), manualBranches: ['SRC'], autoPriceBranches: ['KKL'] }, { ...entry('1'), manualBranches: [], autoPriceBranches: ['SRC'] }, entry('OTHER', ['SSS']), entry('LEGACY', ['WRD'])]
  const h = harness(entries, [product('00001', 0), product('1', '฿ 1,234.50'), product('OTHER'), product('LEGACY')])
  await h.run('src')
  assert.equal(h.state.snapshots, 1)
  assert.deepEqual(h.state.calls, [['00001', '1']])
  const { wb, filename } = h.state.files[0]
  assert.deepEqual(wb.SheetNames, ['TEMPLATE'])
  assert.equal(filename, 'GM MME PENDING SRC DB Cloud 03-11-2026.xlsx')
  const ws = wb.Sheets.TEMPLATE
  assert.deepEqual(XLSX.utils.sheet_to_json(ws, { header: 1 })[0], headers)
  assert.equal(ws.D2.t, 'n'); assert.equal(ws.D2.v, 0)
  assert.equal(ws.D3.v, 1234.5)
  assert.equal(ws.E2.t, 's'); assert.equal(ws.E2.v, '00001')
  assert.equal(ws.E3.v, '1')
  assert.equal(ws.B2.v, 'ชื่อสินค้า 00001')
  assert.equal(ws.C2.v, '000123')
  assert.equal(ws.F2.v, product('00001')['*รูปภาพสินค้า'])
  assert.equal(ws.F2.l.Target, product('00001')['*รูปภาพสินค้า'])
  assert.equal(ws.A2.v, 'อัปเดตรายการสินค้า')
  assert.equal(ws.D2.z, '#,##0.00')
  assert.equal(ws['!cols'].length, 7)
  assert.equal(ws['!rows'][1].hpt, 141)
  const reloaded = XLSX.read(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }))
  assert.equal(reloaded.Sheets.TEMPLATE.E2.v, '00001')
  assert.equal(reloaded.Sheets.TEMPLATE.D2.v, 0)
  assert.equal(reloaded.Sheets.TEMPLATE.F2.l.Target, product('00001')['*รูปภาพสินค้า'])
  assert.equal(h.state.notice.error, false)
  assert.equal(h.state.busy, ''); assert.equal(h.busyRef.current, false)
})

test('column F preserves full Google Drive URLs as clickable hyperlinks and leaves missing URLs unlinked', async () => {
  const link = 'https://drive.google.com/file/d/example-id/view?usp=sharing&resourcekey=0-keep-this'
  const h = harness([entry('LINK'), entry('EMPTY')], [{ ...product('LINK'), '*รูปภาพสินค้า': link }, { ...product('EMPTY'), '*รูปภาพสินค้า': '' }])
  await h.run('src')
  const bytes = XLSX.write(h.state.files[0].wb, { type: 'buffer', bookType: 'xlsx' })
  const reloaded = XLSX.read(bytes)
  assert.equal(reloaded.Sheets.TEMPLATE.F2.v, link)
  // SheetJS 0.18 leaves XML entities in parsed hyperlink targets; verify the stored relationship.
  const zip = XLSX.CFB.read(bytes, { type: 'buffer' })
  const index = zip.FullPaths.findIndex(path => path.endsWith('worksheets/_rels/sheet1.xml.rels'))
  assert.ok(index >= 0)
  const relationships = Buffer.from(zip.FileIndex[index].content).toString('utf8')
  assert.ok(relationships.includes(`Target="${link.replace(/&/g, '&amp;')}"`))
  assert.equal(reloaded.Sheets.TEMPLATE.F2.f, undefined)
  assert.equal(reloaded.Sheets.TEMPLATE.F3?.l, undefined)
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

test('closed price acknowledgements are excluded from exports without excluding remaining active branches', async () => {
  const saved = [
    { ...entry('CLOSED', []), manualBranches: [], autoPriceBranches: [], resolvedPrices: { src: { grabCents: 100, promaxxCents: 0 } } },
    { ...entry('OTHER', ['SSS']), resolvedPrices: { src: { grabCents: 100, promaxxCents: 0 } } },
  ]
  const h = harness(saved, [product('CLOSED'), product('OTHER')])
  await h.run('src')
  assert.equal(h.state.files.length, 0)
  assert.equal(h.state.calls.length, 0)
  await h.run('sss')
  assert.deepEqual(XLSX.utils.sheet_to_json(h.state.files[0].wb.Sheets.TEMPLATE).map(row => row['รหัสสินค้า']), ['OTHER'])
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

test('price checking follows actual Promaxx even when the Cloud price agrees with Grab or is missing', () => {
  const h = harness([], [])
  assert.equal(h.priceState({ ...product('001', 100), branch: 'src', src_price: 100, promaxx_price: 120 }), 'mismatch')
  assert.equal(h.priceState({ ...product('001', 100), branch: 'src', src_price: 120, promaxx_price: 120 }), 'matched')
  assert.equal(h.priceState({ ...product('001', null), branch: 'src', src_price: 0, promaxx_price: 0 }), 'matched')
  assert.equal(h.priceState({ ...product('001'), branch: 'src', src_price: 10.004, promaxx_price: 10 }), 'matched')
  assert.equal(h.priceState({ ...product('001'), branch: 'src', promaxx_price: null }), 'unpriced')
  assert.equal(h.priceState({ ...product('001'), branch: 'src', src_price: null }), 'unpriced')
  assert.equal(h.priceState({ ...product('001'), branch: 'sss' }), 'absent')
})

test('Promaxx export requires review, writes the selling price to column D, preserves links/SKUs and leaves Cloud exports independent', async () => {
  const original = { ...product('00001', 100), branch: 'src', src_price: 100, promaxx_price: 120 }
  const h = harness([entry('00001')], [original])
  await h.preview('src')
  assert.equal(h.state.files.length, 0)
  assert.equal(h.state.review.rows[0].grab, 100)
  assert.equal(h.state.review.rows[0].promaxx, 120)
  assert.equal(h.state.review.rows[0].ready, true)
  await h.run('src', h.state.review)
  assert.equal(h.state.snapshots, 2)
  assert.equal(h.state.review, null)
  const first = h.state.files[0]
  assert.match(first.filename, /SRC PROMAXX/)
  const roundtrip = XLSX.read(XLSX.write(first.wb, { type: 'buffer', bookType: 'xlsx' })).Sheets.TEMPLATE
  assert.equal(roundtrip.D2.v, 120)
  assert.equal(roundtrip.E2.v, '00001')
  assert.equal(roundtrip.F2.l.Target, original['*รูปภาพสินค้า'])
  assert.deepEqual(XLSX.utils.sheet_to_json(first.wb.Sheets.TEMPLATE, { header: 1 })[0], headers)
  assert.equal(original['*ราคาสินค้า'], 100)
  assert.equal(original.src_price, 100)
  await h.run('src')
  assert.match(h.state.files[1].filename, /DB Cloud/)
  assert.equal(h.state.files[1].wb.Sheets.TEMPLATE.D2.v, 100)
})

test('each active branch uses its own Grab price and never includes another branch pending SKU', async () => {
  for (const branch of ['src', 'sss', 'kkl']) {
    const rows = [entry('00001', [branch.toUpperCase()]), entry('OTHER', ['WRD'])]
    const h = harness(rows, [{ ...product('00001'), branch, src_price: 10, sss_price: 20, kkl_price: 30, promaxx_price: 40 }, product('OTHER')])
    await h.preview(branch)
    assert.deepEqual(h.state.review.rows.map(row => row.sku), ['00001'])
    assert.equal(h.state.review.rows[0].grab, { src: 10, sss: 20, kkl: 30 }[branch])
    await h.run(branch, h.state.review)
    assert.match(h.state.files[0].filename, new RegExp(branch.toUpperCase() + ' PROMAXX'))
    assert.equal(h.state.files[0].wb.Sheets.TEMPLATE.D2.v, 40)
  }
})

test('missing/invalid targets, missing Grab prices, absent branches and matched rows are explained and skipped; zero is exportable', async () => {
  const variants = [
    ['ZERO', { promaxx_price: 0, src_price: 5 }],
    ['NO_TARGET', { promaxx_price: null }],
    ['BAD_TARGET', { promaxx_price: 'invalid' }],
    ['NEGATIVE', { promaxx_price: -1 }],
    ['NO_GRAB', { src_price: null }],
    ['ABSENT', { branch: 'sss' }],
    ['SAME', { promaxx_price: 10, src_price: 10.004 }],
  ]
  const products = variants.map(([sku, fields]) => ({ ...product(sku), branch: 'src', ...fields }))
  const h = harness([...variants.map(([sku]) => entry(sku)), entry('MISSING')], products)
  await h.preview('src')
  assert.deepEqual(h.state.review.rows.filter(row => row.ready).map(row => row.sku), ['ZERO'])
  assert.ok(h.state.review.rows.filter(row => !row.ready).every(row => row.reason))
  assert.match(h.state.review.rows.find(row => row.sku === 'NO_TARGET').reason, /PROMAXX/)
  assert.match(h.state.review.rows.find(row => row.sku === 'NO_GRAB').reason, /GRAB/)
  await h.run('src', h.state.review)
  const ws = h.state.files[0].wb.Sheets.TEMPLATE
  assert.equal(ws.D2.t, 'n'); assert.equal(ws.D2.v, 0)
  assert.deepEqual(XLSX.utils.sheet_to_json(ws).map(row => row['รหัสสินค้า']), ['ZERO'])
})

test('changed targets, branch prices or pending membership require another review before downloading', async () => {
  for (const change of ['promaxx', 'grab', 'membership']) {
    let entries = [entry('001')]
    let products = [{ ...product('001'), branch: 'src', src_price: 10, promaxx_price: 20 }]
    const h = harness([], [], { currentEntries: () => entries, currentProducts: () => products })
    await h.preview('src')
    const reviewed = h.state.review
    if (change === 'membership') entries = []
    else products = [{ ...products[0], [change === 'promaxx' ? 'promaxx_price' : 'src_price']: 30 }]
    await h.run('src', reviewed)
    assert.equal(h.state.files.length, 0)
    assert.match(h.state.review.error, /เปลี่ยนแล้ว/)
    assert.equal(h.state.review.loading, false)
    if (change !== 'membership') {
      await h.run('src', h.state.review)
      assert.equal(h.state.files.length, 1)
      assert.equal(h.state.files[0].wb.Sheets.TEMPLATE.D2.v, change === 'promaxx' ? 30 : 20)
    }
  }
})

test('failed preview/confirmation and cancellation cannot download a partial or unreviewed file', async () => {
  for (const options of [{ snapshotError: true }, { failSku: '001' }, { badBody: true }]) {
    const h = harness([entry('001')], [{ ...product('001'), branch: 'src' }], options)
    await h.preview('src')
    assert.ok(h.state.review.error)
    assert.equal(h.state.review.loading, false)
    assert.equal(h.state.files.length, 0)
    assert.equal(h.busyRef.current, false)
  }
  let release, hold = false
  const h = harness([entry('001')], [{ ...product('001'), branch: 'src' }], { waitForSnapshot: async () => { if (hold) await new Promise(resolve => { release = resolve }) } })
  await h.preview('src')
  hold = true
  const confirming = h.run('src', h.state.review)
  await Promise.resolve()
  h.close(); release(); await confirming
  assert.equal(h.state.review, null)
  assert.equal(h.state.files.length, 0)
  assert.equal(h.busyRef.current, false)
  const preparing = h.preview('src')
  await Promise.resolve()
  h.close(); release(); await preparing
  assert.equal(h.state.review, null)
  assert.equal(h.busyRef.current, false)
})

test('an all-matched or empty branch produces no price-adjustment file', async () => {
  for (const entries of [[], [entry('001')]]) {
    const h = harness(entries, [{ ...product('001'), branch: 'src', src_price: 10, promaxx_price: 10 }])
    await h.preview('src')
    assert.equal(h.state.review.rows.filter(row => row.ready).length, 0)
    assert.equal(h.state.files.length, 0)
  }
})
