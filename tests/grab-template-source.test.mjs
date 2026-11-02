import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import ts from 'typescript'
import { demo, csv } from '../public/design-previews/grab-manager/review.js'

const directory = path.dirname(fileURLToPath(import.meta.url)), require = createRequire(import.meta.url)
function loadModules(supabase) {
  const cache = new Map()
  function load(name) {
    if (cache.has(name)) return cache.get(name)
    const source = fs.readFileSync(path.join(directory, '../lib', name + '.ts'), 'utf8')
    const code = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText
    const output = { exports: {} }; cache.set(name, output.exports)
    new Function('require', 'exports', 'module', code)(name => name === './supabase' ? { supabase } : name.startsWith('./') ? load(name.slice(2)) : require(name), output.exports, output)
    return output.exports
  }
  return { source: load('grab-template-source'), review: load('grab-template-review'), engine: load('process-engine') }
}

function store(initial = demo().src.baseline) {
  const state = { files: new Map(initial === null ? [] : [['GRAB_SRC.CSV', { name: 'GRAB_SRC.CSV', content: initial, updated_at: '2026-11-02T00:00:00.000Z' }]]), reads: 0, writes: 0, onWrite: null, readError: null, writeError: null }
  const supabase = { from(table) {
    assert.equal(table, 'grab_files')
    let mode = 'read', payload
    const filters = {}
    return {
      select() { return this },
      eq(column, value) { filters[column] = value; return this },
      update(values) { mode = 'update'; payload = values; return this },
      insert(values) { mode = 'insert'; payload = values; return this },
      async maybeSingle() {
        const name = mode === 'insert' ? payload.name : filters.name
        if (mode === 'read') { state.reads++; return state.readError ? { error: state.readError } : { data: state.files.has(name) ? { ...state.files.get(name) } : null } }
        state.writes++; state.onWrite?.(state)
        if (state.writeError) return { error: state.writeError }
        if (mode === 'insert' && state.files.has(name)) return { error: { code: '23505', message: 'duplicate key' } }
        if (mode === 'update' && state.files.get(name)?.updated_at !== filters.updated_at) return { data: null }
        state.files.set(name, { name, ...payload })
        return { data: { name, updated_at: payload.updated_at } }
      },
    }
  } }
  return { state, ...loadModules(supabase) }
}

test('saving a changed CSV compares the pre-write baseline and returns only a confirmed saved version', async () => {
  const sample = demo().src, { state, source } = store(sample.baseline)
  const saved = await source.saveGrabTemplate('ชื่ออะไรก็ได้', sample.candidate)
  assert.equal(saved.name, 'GRAB_SRC.CSV'); assert.equal(saved.sourceName, 'ชื่ออะไรก็ได้')
  assert.equal(saved.changed, true); assert.equal(saved.isNew, false)
  assert.deepEqual(saved.counts, { updated: 3, added: 1, removed: 1, unchanged: 2 })
  assert.equal(state.files.get(saved.name).content, sample.candidate)
  assert.equal(saved.updatedAt, state.files.get(saved.name).updated_at)
  assert.equal(state.reads, 1); assert.equal(state.writes, 1)
})

test('unchanged data and first uploads save successfully without reporting a baseline change', async () => {
  const sample = demo().src
  const same = store(sample.baseline)
  const unchanged = await same.source.saveGrabTemplate('same.csv', sample.baseline.replace(/\r\n/g, '\n'))
  assert.equal(unchanged.changed, false); assert.equal(unchanged.isNew, false); assert.equal(unchanged.counts.unchanged, 6)
  const first = store(null)
  const initial = await first.source.saveGrabTemplate('first.CSV', sample.candidate)
  assert.equal(initial.isNew, true); assert.equal(initial.changed, false); assert.equal(initial.counts.added, 6)
  assert.equal(first.state.files.get('GRAB_SRC.CSV').content, sample.candidate)
})

test('read failures, invalid versions, invalid columns, duplicate IDs and write denial never become successful changes', async () => {
  const sample = demo().src
  const read = store(); read.state.readError = { message: 'read unavailable' }
  await assert.rejects(read.source.saveGrabTemplate('file', sample.candidate), /อ่านไฟล์เดิม/)
  assert.equal(read.state.writes, 0)
  const version = store(); version.state.files.get('GRAB_SRC.CSV').updated_at = 'bad'
  await assert.rejects(version.source.saveGrabTemplate('file', sample.candidate), /เวอร์ชัน/)
  assert.equal(version.state.writes, 0)
  const invalid = store()
  await assert.rejects(invalid.source.saveGrabTemplate('file', sample.candidate.replace('CURRENTSTOCK', 'BAD')), /คอลัมน์ E/)
  assert.equal(invalid.state.reads, 0); assert.equal(invalid.state.writes, 0)
  const duplicate = store(), rows = duplicate.engine.parseEngineCsv(sample.candidate).rows
  rows[4][1] = rows[3][1]
  await assert.rejects(duplicate.source.saveGrabTemplate('file', csv(rows)), /ซ้ำ/)
  assert.equal(duplicate.state.writes, 0)
  const denied = store(); denied.state.writeError = { code: '42501', message: 'policy denied' }
  await assert.rejects(denied.source.saveGrabTemplate('file', sample.candidate), /policy denied/)
  assert.equal(denied.state.files.get('GRAB_SRC.CSV').content, sample.baseline)
})

test('a simultaneous update is re-read and the returned before-values use the peer version', async () => {
  const sample = demo().src, { state, source, engine } = store()
  const rows = engine.parseEngineCsv(sample.baseline).rows; rows[4][4] = '99'
  const peer = csv(rows)
  state.onWrite = state => { if (state.writes === 1) state.files.set('GRAB_SRC.CSV', { name: 'GRAB_SRC.CSV', content: peer, updated_at: '2026-11-02T00:00:00.001Z' }) }
  const saved = await source.saveGrabTemplate('file', sample.candidate)
  assert.equal(state.writes, 2); assert.equal(state.reads, 2)
  assert.equal(saved.rows.find(row => row.sku === '0002').changes.find(change => change.name === 'CURRENTSTOCK').before, '99')
  assert.equal(state.files.get('GRAB_SRC.CSV').content, sample.candidate)
})

test('three conflicting updates stop without overwriting the latest peer file', async () => {
  const sample = demo().src, { state, source } = store()
  state.onWrite = state => { state.files.set('GRAB_SRC.CSV', { name: 'GRAB_SRC.CSV', content: sample.baseline, updated_at: `2026-11-02T00:00:00.00${state.writes}Z` }) }
  await assert.rejects(source.saveGrabTemplate('file', sample.candidate), /พร้อมกัน/)
  assert.equal(state.writes, 3); assert.equal(state.files.get('GRAB_SRC.CSV').content, sample.baseline)
})

test('a concurrent first insertion is re-read instead of blindly replacing it', async () => {
  const sample = demo().src, { state, source } = store(null)
  state.onWrite = state => { if (state.writes === 1) state.files.set('GRAB_SRC.CSV', { name: 'GRAB_SRC.CSV', content: sample.baseline, updated_at: '2026-11-02T00:00:00.001Z' }) }
  const saved = await source.saveGrabTemplate('file', sample.candidate)
  assert.equal(state.writes, 2); assert.equal(saved.isNew, false); assert.equal(saved.changed, true)
})

test('stock review matches stable ItemIDs, ignores record order and exports exact saved values with one BOM', () => {
  const { review, engine } = store(), sample = demo().src
  const rows = engine.parseEngineCsv(sample.baseline).rows
  const reordered = csv([...rows.slice(0, 3), ...rows.slice(3).reverse()])
  assert.equal(review.compareGrabTemplates(sample.baseline, reordered).changed, false)
  rows[3][2] = 'ยา, "ใหม่"\nสองบรรทัด'; rows[3][3] = '0099'
  const compared = review.compareGrabTemplates(sample.baseline, csv(rows))
  assert.deepEqual(compared.counts, { updated: 1, added: 0, removed: 0, unchanged: 5 })
  assert.deepEqual(compared.rows[0].changes.map(change => change.name), ['ITEMNAME', 'ITEMCODE'])
  const report = engine.parseEngineCsv(review.grabChangesCsv(compared)).rows
  assert.equal(report[1][8], 'ยา, "ใหม่"\nสองบรรทัด')
  assert.equal(report[2][8], '0099')
  assert.equal(review.grabCsvWithBom('\uFEFF\uFEFFa,b\r\n'), '\uFEFFa,b\r\n')
})

test('stock review falls back to unique SKUs for empty ItemIDs and blocks ambiguous matches', () => {
  const { review, engine } = store()
  const rows = engine.parseEngineCsv(demo().src.baseline).rows
  rows.slice(3).forEach(row => { row[1] = '' })
  const baseline = csv(rows)
  rows[3][4] = '25'
  const changed = review.compareGrabTemplates(baseline, csv(rows))
  assert.deepEqual(changed.counts, { updated: 1, added: 0, removed: 0, unchanged: 5 })
  assert.equal(changed.rows[0].sku, '0001')
  rows[4][3] = rows[3][3]
  assert.throws(() => review.compareGrabTemplates(baseline, csv(rows)), /ซ้ำ/)
})
