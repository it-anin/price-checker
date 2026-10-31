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
