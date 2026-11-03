const { test } = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const ts = require('typescript')

function loadModule(relative, supabase) {
  const cache = new Map()
  function load(filename) {
    if (cache.has(filename)) return cache.get(filename).exports
    const module = { exports: {} }; cache.set(filename, module)
    const code = ts.transpileModule(fs.readFileSync(filename, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText
    const requireLocal = name => name === './supabase' ? { supabase } : name.startsWith('.') ? load(path.resolve(path.dirname(filename), name + '.ts')) : require(name)
    new Function('require', 'exports', 'module', code)(requireLocal, module.exports, module)
    return module.exports
  }
  return load(path.resolve(__dirname, '..', relative))
}
const pending = loadModule('lib/pending-sku.ts')
const engine = loadModule('lib/process-engine.ts')
const parsePrice = value => {
  if (!value) return null
  const n = parseFloat(value.replace(/[฿$€£¥\s]/g, '').replace(/,/g, ''))
  return Number.isNaN(n) ? null : n
}
const normalize = pending.normalizePendingEntries

test('product Pending coverage requires all active branches and preserves SKU identity and legacy behavior', () => {
  const sets = engine.pendingSkuSets(normalize([
    { sku: '00001', branches: ['KKL', 'SRC', 'SSS'] },
    { sku: '00002', branches: ['SRC', 'WRD', 'ONN'] },
    { sku: 'ALPHA-01', branches: ['KKL'] },
    { sku: 'OLD', branches: ['WRD', 'ONN'] },
    { sku: 'LEGACY' },
  ]))
  const coverage = sku => pending.pendingBranchCoverage(sku, sets)
  assert.deepEqual(coverage('00001'), { present: ['src', 'sss', 'kkl'], missing: [], complete: true })
  assert.deepEqual(coverage('00002'), { present: ['src'], missing: ['sss', 'kkl'], complete: false })
  assert.deepEqual(coverage(' alpha-01\u200b '), { present: ['kkl'], missing: ['src', 'sss'], complete: false })
  for (const sku of ['OLD', 'MISSING', '1']) assert.deepEqual(coverage(sku), { present: [], missing: ['src', 'sss', 'kkl'], complete: false })
  assert.equal(coverage('LEGACY').complete, true)
})

function store(entries) {
  const state = { entries, updatedAt: '2026-10-30T00:00:00.000Z', reads: 0, writes: 0, failRead: false, failWrite: false, onWrite: null }
  const client = { from(table) {
    assert.equal(table, 'app_settings')
    let update, filters = {}
    return {
      select() { return this },
      update(value) { update = value; return this },
      eq(key, value) { filters[key] = value; return this },
      async maybeSingle() {
        assert.equal(filters.key, 'pending_skus')
        if (!update) {
          state.reads++
          return state.failRead ? { error: { message: 'read denied' } } : { data: { value: JSON.stringify(state.entries), updated_at: state.updatedAt } }
        }
        state.writes++
        assert.ok(filters.updated_at.length < 100)
        assert.deepEqual(Object.keys(filters).sort(), ['key', 'updated_at'])
        if (state.failWrite) return { error: { message: 'session expired / write denied' } }
        if (state.onWrite) state.onWrite(state)
        if (filters.updated_at !== state.updatedAt) return { data: null }
        assert.ok(Date.parse(update.updated_at) > Date.parse(state.updatedAt))
        state.entries = JSON.parse(update.value); state.updatedAt = update.updated_at
        return { data: { value: update.value, updated_at: update.updated_at } }
      },
    }
  } }
  return { state, source: loadModule('lib/process-engine-source.ts', client) }
}

test('ADD accepts all documented separators, normalizes keys, and keeps leading zeroes', () => {
  assert.deepEqual(pending.parsePendingInput(' 001\nabc,001;DEF\tghi|\uFEFFjkl\r\n'), ['001', 'ABC', 'DEF', 'GHI', 'JKL'])
})

test('ADD merges selected branches without duplication or losing names and metadata', () => {
  const entries = normalize([{ sku: '001', name: 'เดิม', branches: ['SRC', 'WRD'], note: { keep: true } }])
  const result = pending.mutatePendingEntries(entries, { kind: 'add', skus: ['001', '001', '002'], branches: ['sss'] })
  assert.deepEqual(result.entries[0], { sku: '001', name: 'เดิม', branches: ['SRC', 'SSS', 'WRD'], note: { keep: true } })
  assert.deepEqual(result.entries[1].branches, ['SSS'])
  assert.deepEqual(entries[0].branches, ['SRC', 'WRD'])
  assert.throws(() => pending.mutatePendingEntries(entries, { kind: 'add', skus: ['003'], branches: [] }), /เลือกสาขา/)
})

test('DELETE affects one branch, preserves WRD/ONN, and removes an entry only when no branch remains', () => {
  const entries = normalize([{ sku: '001', branches: ['SRC', 'SSS', 'ONN'], custom: 42 }, { sku: '002', branches: ['SRC'] }])
  const first = pending.mutatePendingEntries(entries, { kind: 'removeBranch', sku: '001', branch: 'src' }).entries
  assert.deepEqual(first[0].branches, ['SSS', 'ONN']); assert.equal(first[0].custom, 42)
  assert.equal(pending.mutatePendingEntries(first, { kind: 'removeBranch', sku: '002', branch: 'src' }).entries.length, 1)
  assert.deepEqual(normalize([{ sku: 'legacy' }])[0].branches, ['SRC', 'SSS', 'KKL'])
})

test('name enrichment fills only empty current names without reviving deleted entries', () => {
  const entries = normalize([{ sku: '001', name: '', branches: ['SRC'], note: true }, { sku: '002', name: 'keep', branches: ['SSS'] }])
  const result = pending.mutatePendingEntries(entries, { kind: 'enrichNames', names: new Map([['001', 'new'], ['002', 'overwrite'], ['003', 'deleted']]) })
  assert.equal(result.entries.length, 2); assert.equal(result.entries[0].name, 'new'); assert.equal(result.entries[0].note, true); assert.equal(result.entries[1].name, 'keep')
})

test('malformed pending JSON, duplicate keys and unknown branches block writes', () => {
  for (const raw of ['not json', {}, [{ sku: '' }], [{ sku: '001', branches: ['UNKNOWN'] }], [{ sku: '001' }, { sku: '001' }]]) assert.throws(() => normalize(raw))
})

test('timestamps advance past the existing version even within one millisecond or with a skewed clock', () => {
  const previous = '2026-10-30T00:00:00.999999Z'
  assert.ok(Date.parse(pending.nextPendingTimestamp(previous, 0)) > Date.parse(previous))
  assert.throws(() => pending.nextPendingTimestamp('bad'), /updated_at/)
})

test('stock-result freshness depends on active SKU/branch membership, not names, metadata or legacy-only edits', () => {
  const before = normalize([{ sku: '001', name: 'before', branches: ['SRC', 'WRD'] }, { sku: 'OLD', branches: ['ONN'] }])
  const namesOnly = normalize([{ sku: '001', name: 'after', branches: ['WRD', 'SRC'], note: true }, { sku: 'NEW-OLD', branches: ['WRD'] }])
  assert.equal(pending.samePendingStockMembership(before, namesOnly), true)
  assert.equal(pending.samePendingStockMembership(before, normalize([{ sku: '001', branches: ['SSS'] }])), false)
  assert.equal(pending.samePendingStockMembership(before, normalize([{ sku: '001', branches: ['SRC'] }, { sku: '002', branches: ['SRC'] }])), false)
})

test('CAS reload/retry preserves a peer ADD and legacy metadata', async () => {
  const { state, source } = store([{ sku: '001', name: 'original', branches: ['WRD'], note: 'keep' }])
  state.onWrite = s => { if (s.writes === 1) { s.entries.push({ sku: 'PEER', name: 'other user', branches: ['KKL'] }); s.updatedAt = '2026-10-30T00:00:00.001Z' } }
  const snapshot = await source.savePendingMutation({ kind: 'add', skus: ['001', '002'], branches: ['src'] })
  assert.equal(state.writes, 2); assert.equal(snapshot.entries.length, 3)
  assert.deepEqual(snapshot.entries.find(x => x.sku === '001').branches, ['SRC', 'WRD'])
  assert.equal(snapshot.entries.find(x => x.sku === '001').note, 'keep')
  assert.equal(snapshot.entries.find(x => x.sku === 'PEER').name, 'other user')
})

test('three CAS conflicts report failure and do not overwrite any peer edits', async () => {
  const { state, source } = store([])
  state.onWrite = s => { s.entries.push({ sku: 'PEER' + s.writes, branches: ['SSS'] }); s.updatedAt = new Date(Date.parse(s.updatedAt) + 1).toISOString() }
  await assert.rejects(source.savePendingMutation({ kind: 'add', skus: ['MINE'], branches: ['src'] }), /ผู้ใช้อื่น/)
  assert.equal(state.writes, 3); assert.equal(state.entries.length, 3); assert.ok(state.entries.every(x => x.sku.startsWith('PEER')))
})

test('read errors and expired sessions never turn into successful or destructive writes', async () => {
  const { state, source } = store([{ sku: '001', branches: ['SRC'] }])
  state.failRead = true
  await assert.rejects(source.savePendingMutation({ kind: 'removeBranch', sku: '001', branch: 'src' }), /อ่าน PENDING/)
  assert.equal(state.writes, 0)
  state.failRead = false; state.failWrite = true
  await assert.rejects(source.savePendingMutation({ kind: 'removeBranch', sku: '001', branch: 'src' }), /session expired/)
  assert.equal(state.entries.length, 1)
})

test('an intentional empty list stays valid, and an already-applied operation needs no write', async () => {
  const { state, source } = store([{ sku: '001', name: '', branches: ['SRC'] }])
  await source.savePendingMutation({ kind: 'add', skus: ['001'], branches: ['src'] }); assert.equal(state.writes, 0)
  const result = await source.savePendingMutation({ kind: 'removeBranch', sku: '001', branch: 'src' })
  assert.deepEqual(result.entries, []); assert.deepEqual(await source.loadPendingSkuSets(), { src: new Set(), sss: new Set(), kkl: new Set() })
})

function priceCsv(storeId = engine.ENGINE_STORES.src) {
  return 'Export instructions\r\nStoreID,Item Name,Price,ItemCode,E,F,G,H,SKU NUMBER\r\n' + storeId + ',ยา,0,WRONG-D,,,,,001\r\n' + storeId + ',ยา 2,"1,234.50",WRONG-D2,,,,,002\r\n'
}
test('GRAB prices require B/C/I headers, use I rather than D, and preserve zero and Thai-format numbers', () => {
  const result = pending.readPendingGrabPrices('grab_src.csv', priceCsv(), parsePrice)
  assert.equal(result.prices.get('001'), 0); assert.equal(result.prices.get('002'), 1234.5); assert.equal(result.prices.has('WRONG-D'), false)
  assert.throws(() => pending.readPendingGrabPrices('GRAB_SRC.CSV', priceCsv().replace('SKU NUMBER', 'BAD'), parsePrice), /หัวตาราง/)
  assert.throws(() => pending.readPendingGrabPrices('GRAB_SSS.CSV', priceCsv(), parsePrice), /StoreID/)
  assert.throws(() => pending.readPendingGrabPrices('OTHER_SRC.CSV', priceCsv(), parsePrice), /ชื่อ/)
})

test('actual stock templates provide names from C and never provide GRAB prices', () => {
  const text = 'StoreID,ItemID,ITEMNAME,ITEMCODE,CURRENTSTOCK,MAXSTOCK,G,H,I,J,K,L,M,N,O,ITEMSTATUS\r\n' + engine.ENGINE_STORES.src + ',ITEM-1,ชื่อจริง,001,1,1,,,,,,,,,,in stock'
  assert.equal(pending.readTemplateNames('GRAB_SRC.CSV', text).get('001'), 'ชื่อจริง')
  assert.throws(() => pending.readPendingGrabPrices('GRAB_SRC.CSV', text, parsePrice), /Template สต็อก/)
})

test('PROMAX selects B/E=5/G only; TIS-620 is decoded before reading prices', () => {
  const bytes = Buffer.concat([Buffer.from([0xc2, 0xd2]), Buffer.from(',SKU,C,D,FILTER,F,PRICE\r\n,001,,,5,,0\r\n,002,,,4,,99\r\n,003,,,5,,45')])
  const decoded = engine.decodeCsv(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))
  assert.equal(decoded.encoding, 'TIS-620')
  const result = pending.readPendingPromaxPrices('R05.103.csv', decoded.text, parsePrice)
  assert.deepEqual([...result.prices], [['001', 0], ['003', 45]])
  assert.throws(() => pending.readPendingPromaxPrices('bad.csv', 'A,B,C,D,E,F,G\n,001,,,4,,1', parsePrice), /ไม่พบราคา/)
})

test('comparison distinguishes missing values and zero, and compares rounded cents', () => {
  assert.equal(pending.comparePendingPrices(0, 0), 'same')
  assert.equal(pending.comparePendingPrices(undefined, 45), 'na')
  assert.equal(pending.comparePendingPrices(45, undefined), 'na')
  assert.equal(pending.comparePendingPrices(45.004, 45), 'same')
  assert.equal(pending.comparePendingPrices(45.01, 45), 'diff')
})

test('ADD/DELETE changes stock suppression for that branch only', () => {
  const report = { name: 'src.csv', branch: 'src', stock: new Map([['001', 10]]), rowCount: 1, encoding: 'UTF-8' }
  const template = 'StoreID,ItemID,ITEMNAME,ITEMCODE,CURRENTSTOCK,MAXSTOCK,G,H,I,J,K,L,M,N,O,ITEMSTATUS\n'+engine.ENGINE_STORES.src+',id,ยา,001,1,1,,,,,,,,,,in stock'
  let entries = pending.mutatePendingEntries([], { kind: 'add', skus: ['001'], branches: ['src'] }).entries
  assert.equal(engine.pendingSkuSets(entries).sss.size, 0)
  let output = engine.processGrabTemplate('GRAB_SRC.CSV', template, report, 5, engine.pendingSkuSets(entries).src)
  assert.equal(engine.parseEngineCsv(output.content).rows[1][4], '')
  entries = pending.mutatePendingEntries(entries, { kind: 'removeBranch', sku: '001', branch: 'src' }).entries
  output = engine.processGrabTemplate('GRAB_SRC.CSV', template, report, 5, engine.pendingSkuSets(entries).src)
  assert.equal(engine.parseEngineCsv(output.content).rows[1][4], '5')
})

const product = (sku, values = {}) => ({ 'รหัสสินค้า (SKU NUMBER)': sku, '*ชื่อสินค้า (NAME)': 'สินค้า '+sku, branch: 'src,kkl,sss', promaxx_price: '100', src_price: '100', kkl_price: '100', sss_price: '100', ...values })
const syncPrices = (entries, products) => pending.mutatePendingEntries(entries, { kind: 'syncPrices', checks: pending.buildPendingPriceChecks(products, parsePrice) })

test('automatic price checks compare all three offered branches to Promaxx, using rounded cents and valid zero prices', () => {
  const checks = pending.buildPendingPriceChecks([
    product('001', { src_price: '฿1,000.00', kkl_price: '99.99', sss_price: '100.004', '*ราคาสินค้า': '900' }),
    product('002', { promaxx_price: '0', src_price: '0', kkl_price: '1', sss_price: '' }),
    product('003', { branch: 'src', src_price: '', kkl_price: '10', sss_price: 'Infinity' }),
    product('004', { promaxx_price: null, src_price: '50' }),
  ], parsePrice)
  assert.deepEqual(checks.get('001').states, { src: 'diff', sss: 'same', kkl: 'diff' })
  assert.deepEqual(checks.get('002').states, { src: 'same', sss: 'na', kkl: 'diff' })
  assert.deepEqual(checks.get('003').states, { src: 'na', sss: 'na', kkl: 'na' })
  assert.deepEqual(checks.get('004').states, { src: 'na', sss: 'na', kkl: 'na' })
  assert.equal(checks.get('001').prices.src, 1000)
  const entries = syncPrices([], [...checks.keys()].map(sku => product(sku, { branch: 'src', src_price: '101' })) ).entries
  assert.equal(entries.length, 4)
  assert.ok(entries.every(entry => entry.branches.join(',') === 'SRC' && entry.manualBranches.length === 0))
})

test('automatic and manual origins coexist for one SKU and resolve independently without losing legacy metadata', () => {
  const original = normalize([{ sku: '001', name: 'เดิม', branches: ['SRC', 'WRD', 'ONN'], note: { keep: true } }])
  let entries = syncPrices(original, [product('001', { src_price: '110', sss_price: '120' }), product('002', { kkl_price: '90' })]).entries
  const mixed = entries.find(entry => entry.sku === '001')
  assert.deepEqual(mixed.manualBranches, ['SRC', 'WRD', 'ONN']); assert.deepEqual(mixed.autoPriceBranches, ['SRC', 'SSS'])
  assert.deepEqual(mixed.branches, ['SRC', 'SSS', 'WRD', 'ONN']); assert.equal(mixed.name, 'เดิม'); assert.deepEqual(mixed.note, { keep: true })
  assert.deepEqual(original[0].branches, ['SRC', 'WRD', 'ONN'])
  entries = pending.mutatePendingEntries(entries, { kind: 'removeBranch', sku: '001', branch: 'src' }).entries
  assert.deepEqual(entries[0].manualBranches, ['WRD', 'ONN']); assert.ok(entries[0].branches.includes('SRC'))
  entries = pending.mutatePendingEntries(entries, { kind: 'add', skus: ['001', '002'], branches: ['sss', 'kkl'] }).entries
  entries = syncPrices(entries, [product('001'), product('002')]).entries
  assert.deepEqual(entries[0].autoPriceBranches, []); assert.deepEqual(entries[0].branches, ['SSS', 'KKL', 'WRD', 'ONN'])
  assert.deepEqual(entries[1].manualBranches, ['SSS', 'KKL']); assert.deepEqual(entries[1].autoPriceBranches, [])
  assert.deepEqual(normalize(JSON.stringify(entries)), entries)
})

test('resolved auto-only entries disappear, while missing prices, absent branches and missing products never release pending', () => {
  let entries = syncPrices([], [product('001', { src_price: 110, sss_price: 120 }), product('002', { kkl_price: 80 })]).entries
  assert.equal(syncPrices(entries, [product('001', { src_price: 110, sss_price: 120 }), product('002', { kkl_price: 80 })]).changed, false)
  const missing = syncPrices(entries, [product('001', { branch: 'sss', src_price: null, promaxx_price: '' })])
  assert.equal(missing.changed, false); assert.deepEqual(missing.entries, entries)
  entries = syncPrices(entries, [product('001', { src_price: 100, sss_price: 120 }), product('002')]).entries
  assert.equal(entries.length, 1); assert.deepEqual(entries[0].branches, ['SSS'])
  assert.deepEqual(syncPrices(entries, [product('001')]).entries, [])
})

test('auto reconciliation retries a concurrent manual addition without deleting its source or metadata', async () => {
  const { state, source } = store([{ sku: '001', branches: ['WRD'], legacy: true }])
  state.onWrite = s => {
    if (s.writes !== 1) return
    s.entries = pending.mutatePendingEntries(normalize(s.entries), { kind: 'add', skus: ['001', 'PEER'], branches: ['src'] }).entries
    s.updatedAt = '2026-10-30T00:00:00.001Z'
  }
  const checks = pending.buildPendingPriceChecks([product('001', { src_price: '101' })], parsePrice)
  let snapshot = await source.savePendingMutation({ kind: 'syncPrices', checks })
  assert.equal(state.writes, 2); assert.ok(snapshot.entries.some(entry => entry.sku === 'PEER'))
  assert.deepEqual(snapshot.entries[0].manualBranches, ['SRC', 'WRD']); assert.equal(snapshot.entries[0].legacy, true)
  snapshot = await source.savePendingMutation({ kind: 'syncPrices', checks: pending.buildPendingPriceChecks([product('001')], parsePrice) })
  assert.deepEqual(snapshot.entries[0].branches, ['SRC', 'WRD']); assert.deepEqual(snapshot.entries[0].autoPriceBranches, [])
})

test('invalid source metadata blocks writes instead of treating automatic entries as manual', () => {
  for (const item of [
    { sku: '001', branches: ['SRC'], manualBranches: [], autoPriceBranches: 'SRC' },
    { sku: '001', branches: ['SRC'], manualBranches: [], autoPriceBranches: ['SSS'] },
    { sku: '001', branches: ['ONN'], manualBranches: [], autoPriceBranches: ['ONN'] },
  ]) assert.throws(() => normalize([item]), /ที่มา/)
})

test('automatic pending suppresses only mismatched branches and manual suppression survives an equal price', () => {
  const template = branch => 'StoreID,ItemID,ITEMNAME,ITEMCODE,CURRENTSTOCK,MAXSTOCK,G,H,I,J,K,L,M,N,O,ITEMSTATUS\n'+engine.ENGINE_STORES[branch]+',id,ยา,001,1,1,,,,,,,,,,in stock'
  const report = branch => ({ name: branch+'.csv', branch, stock: new Map([['001', 10]]), rowCount: 1, encoding: 'UTF-8' })
  let entries = syncPrices([], [product('001', { src_price: 101 })]).entries
  const stock = branch => engine.parseEngineCsv(engine.processGrabTemplate('GRAB_'+branch.toUpperCase()+'.CSV', template(branch), report(branch), 2, engine.pendingSkuSets(entries)[branch]).content).rows[1][4]
  assert.equal(stock('src'), ''); assert.equal(stock('sss'), '8'); assert.equal(stock('kkl'), '8')
  entries = pending.mutatePendingEntries(entries, { kind: 'add', skus: ['001'], branches: ['src'] }).entries
  entries = syncPrices(entries, [product('001')]).entries
  assert.equal(stock('src'), '')
  entries = pending.mutatePendingEntries(entries, { kind: 'removeBranch', sku: '001', branch: 'src' }).entries
  assert.equal(stock('src'), '8')
})

test('closing a branch removes both sources, keeps other branches/metadata and survives saved reloads and identical price scans', async () => {
  const checks = pending.buildPendingPriceChecks([product('00001', { src_price: 110, sss_price: 120 })], parsePrice)
  const initial = pending.mutatePendingEntries(normalize([{ sku: '00001', name: 'keep', branches: ['SRC', 'WRD'], note: true }]), { kind: 'syncPrices', checks }).entries
  const { state, source } = store(initial)
  const closed = await source.savePendingMutation({ kind: 'resolveBranch', sku: '00001', branch: 'src', prices: pending.pendingPricePair(checks.get('00001'), 'src') })
  assert.deepEqual(closed.entries[0].branches, ['SSS', 'WRD'])
  assert.deepEqual(closed.entries[0].manualBranches, ['WRD'])
  assert.deepEqual(closed.entries[0].autoPriceBranches, ['SSS'])
  assert.deepEqual(closed.entries[0].resolvedPrices.src, { grabCents: 11000, promaxxCents: 10000 })
  assert.equal(closed.entries[0].note, true)
  const writes = state.writes
  await source.savePendingMutation({ kind: 'syncPrices', checks })
  assert.equal(state.writes, writes)
  const sets = await source.loadPendingSkuSets()
  assert.equal(sets.src.has('00001'), false)
  assert.equal(sets.sss.has('00001'), true)
  assert.equal(initial[0].branches.includes('SRC'), true)
  assert.equal(initial[0].resolvedPrices, undefined)
})

test('a fully closed SKU remains outside stock suppression after JSON reload, while legacy empty branches still mean all branches', () => {
  let entries = syncPrices([], [product('00001', { src_price: 110 })]).entries
  const check = pending.buildPendingPriceChecks([product('00001', { src_price: 110 })], parsePrice).get('00001')
  entries = pending.mutatePendingEntries(entries, { kind: 'resolveBranch', sku: '00001', branch: 'src', prices: pending.pendingPricePair(check, 'src') }).entries
  assert.deepEqual(entries[0].branches, [])
  const reloaded = normalize(JSON.stringify(entries))
  const sets = engine.pendingSkuSets(reloaded)
  assert.ok(Object.values(sets).every(set => set.size === 0))
  assert.equal(pending.samePendingStockMembership(entries, reloaded), true)
  assert.equal(engine.pendingSkuSets([{ sku: 'LEGACY', branches: [] }]).src.has('LEGACY'), true)
  assert.equal(syncPrices(reloaded, [product('00001', { src_price: 110 })]).changed, false)
})

test('changed GRAB or Promaxx prices reopen only differing branches; a verified match clears the old resolution', () => {
  const check = pending.buildPendingPriceChecks([product('001', { src_price: 110 })], parsePrice).get('001')
  const initial = syncPrices([], [product('001', { src_price: 110 })]).entries
  const closed = pending.mutatePendingEntries(initial, { kind: 'resolveBranch', sku: '001', branch: 'src', prices: pending.pendingPricePair(check, 'src') }).entries
  for (const prices of [{ src_price: 111 }, { src_price: 110, promaxx_price: 101 }]) {
    const reopened = syncPrices(closed, [product('001', prices)]).entries[0]
    assert.ok(reopened.branches.includes('SRC'))
    assert.equal(reopened.resolvedPrices?.src, undefined)
  }
  const matched = syncPrices(closed, [product('001')])
  assert.deepEqual(matched.entries, [])
  assert.ok(syncPrices(matched.entries, [product('001', { src_price: 110 })]).entries[0].branches.includes('SRC'))
})

test('closing uses rounded cents, preserves valid zero and handles missing price pairs until prices become available', () => {
  const check = pending.buildPendingPriceChecks([product('001', { src_price: 0, promaxx_price: 0.1 })], parsePrice).get('001')
  let entries = syncPrices([], [product('001', { src_price: 0, promaxx_price: 0.1 })]).entries
  entries = pending.mutatePendingEntries(entries, { kind: 'resolveBranch', sku: '001', branch: 'src', prices: pending.pendingPricePair(check, 'src') }).entries
  assert.deepEqual(entries[0].resolvedPrices.src, { grabCents: 0, promaxxCents: 10 })
  assert.equal(syncPrices(entries, [product('001', { src_price: 0.001, promaxx_price: 0.1001 })]).changed, false)
  let missing = normalize([{ sku: 'UNKNOWN', branches: ['SRC'] }])
  missing = pending.mutatePendingEntries(missing, { kind: 'resolveBranch', sku: 'UNKNOWN', branch: 'src', prices: pending.pendingPricePair(undefined, 'src') }).entries
  assert.equal(syncPrices(missing, []).changed, false)
  assert.equal(engine.pendingSkuSets(missing).src.size, 0)
  assert.ok(syncPrices(missing, [product('UNKNOWN', { src_price: 110 })]).entries[0].branches.includes('SRC'))
})

test('explicit ADD reopens a closed branch and clears only its acknowledged price pair', () => {
  let entries = normalize([{ sku: '001', branches: [], manualBranches: [], autoPriceBranches: [], resolvedPrices: { src: { grabCents: 11000, promaxxCents: 10000 }, sss: { grabCents: 12000, promaxxCents: 10000 } } }])
  entries = pending.mutatePendingEntries(entries, { kind: 'add', skus: ['001'], branches: ['src'] }).entries
  assert.deepEqual(entries[0].branches, ['SRC'])
  assert.equal(entries[0].resolvedPrices.src, undefined)
  assert.ok(entries[0].resolvedPrices.sss)
  entries = syncPrices(entries, [product('001', { src_price: 110, sss_price: 120 })]).entries
  assert.deepEqual(entries[0].branches, ['SRC'])
})

test('malformed resolution metadata and contradictory active/closed branches block writes', () => {
  for (const resolvedPrices of [[], null, { invalid: {} }, { src: {} }, { src: { grabCents: NaN, promaxxCents: 0 } }, { src: { grabCents: '100', promaxxCents: 0 } }]) {
    assert.throws(() => normalize([{ sku: '001', branches: [], resolvedPrices }]), /ปิดงาน/)
  }
  assert.throws(() => normalize([{ sku: '001', branches: ['SRC'], resolvedPrices: { src: { grabCents: 100, promaxxCents: 0 } } }]), /ปิดงาน/)
})

test('a CAS retry after a peer closes a price mismatch cannot resurrect it from the identical scan', async () => {
  const checks = pending.buildPendingPriceChecks([product('001', { src_price: 110 })], parsePrice)
  const { state, source } = store([])
  state.onWrite = s => {
    if (s.writes !== 1) return
    s.entries = [{ sku: '001', name: 'peer', branches: [], manualBranches: [], autoPriceBranches: [], resolvedPrices: { src: { grabCents: 11000, promaxxCents: 10000 } } }]
    s.updatedAt = '2026-10-30T00:00:00.001Z'
  }
  const snapshot = await source.savePendingMutation({ kind: 'syncPrices', checks })
  assert.deepEqual(snapshot.entries[0].branches, [])
  assert.ok(snapshot.entries[0].resolvedPrices.src)
  assert.equal(state.writes, 1)
})

test('closing retries preserve peer branches and metadata, and failed writes leave pending active', async () => {
  const { state, source } = store([{ sku: '001', branches: ['SRC'], note: 'keep' }])
  state.onWrite = s => {
    if (s.writes !== 1) return
    s.entries[0].branches.push('SSS', 'WRD')
    s.updatedAt = '2026-10-30T00:00:00.001Z'
  }
  const mutation = { kind: 'resolveBranch', sku: '001', branch: 'src', prices: { grabCents: 11000, promaxxCents: 10000 } }
  const snapshot = await source.savePendingMutation(mutation)
  assert.deepEqual(snapshot.entries[0].branches, ['SSS', 'WRD'])
  assert.equal(snapshot.entries[0].note, 'keep')
  assert.equal(state.writes, 2)
  await source.savePendingMutation(mutation)
  assert.equal(state.writes, 2, 'repeating a close is idempotent')
  const failed = store([{ sku: '001', branches: ['SRC'] }])
  failed.state.failWrite = true
  await assert.rejects(failed.source.savePendingMutation(mutation), /write denied/)
  assert.deepEqual(failed.state.entries[0].branches, ['SRC'])
  assert.equal(failed.state.entries[0].resolvedPrices, undefined)
})

test('product Pending coverage follows the union of manual and automatic branches after edits and price resolution', () => {
  let entries = syncPrices(normalize([{ sku: '001', branches: ['SRC', 'WRD'] }]), [product('001', { src_price: 110, sss_price: 110, kkl_price: 110 })]).entries
  const coverage = () => pending.pendingBranchCoverage('001', engine.pendingSkuSets(entries))
  assert.equal(coverage().complete, true)
  entries = syncPrices(entries, [product('001', { src_price: 110, sss_price: 110 })]).entries
  assert.deepEqual(coverage(), { present: ['src', 'sss'], missing: ['kkl'], complete: false })
  entries = pending.mutatePendingEntries(entries, { kind: 'add', skus: ['001'], branches: ['kkl'] }).entries
  assert.equal(coverage().complete, true)
  entries = pending.mutatePendingEntries(entries, { kind: 'removeBranch', sku: '001', branch: 'src' }).entries
  assert.equal(coverage().complete, true, 'Automatic SRC membership remains after manual SRC removal')
})
