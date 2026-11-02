import { test } from 'node:test'
import assert from 'node:assert/strict'
import { compare, csv, demo, reportCsv } from '../public/design-previews/grab-manager/review.js'
import { parseEngineCsv } from '../public/design-previews/grab-manager/engine.js'

const serialize = rows => '\uFEFF' + csv(rows)

test('preview summaries reflect the actual changed, added, removed and unchanged demo rows', () => {
  const reviews = demo()
  assert.deepEqual(reviews.src.counts, { updated: 3, added: 1, removed: 1, unchanged: 2 })
  assert.deepEqual(reviews.sss.counts, { updated: 0, added: 0, removed: 0, unchanged: 6 })
  assert.deepEqual(reviews.kkl.counts, { updated: 2, added: 0, removed: 0, unchanged: 4 })
  assert.ok(Object.values(reviews).every(review => review.total === 6))
  assert.equal(reviews.src.rows.find(row => row.sku === '0002').before.row, 5)
  assert.equal(reviews.src.rows.find(row => row.sku === '0002').after.row, 5)
})

test('reordered CSV records and formatting changes alone do not count as data edits', () => {
  const { baseline } = demo().src
  const rows = parseEngineCsv(baseline).rows
  const reordered = serialize([...rows.slice(0, 3), ...rows.slice(3).reverse()]).replace(/\r\n/g, '\n')
  assert.deepEqual(compare(baseline, reordered).counts, { updated: 0, added: 0, removed: 0, unchanged: 6 })
})

test('stable ItemID keeps SKU, Thai name and multiline edits on the same row', () => {
  const { baseline } = demo().src
  const rows = parseEngineCsv(baseline).rows
  rows[3][2] = 'ยา, "ทดสอบ"\nฉลากใหม่'
  rows[3][3] = '0099'
  const candidate = serialize(rows)
  const review = compare(baseline, candidate)
  assert.deepEqual(review.counts, { updated: 1, added: 0, removed: 0, unchanged: 5 })
  const edited = review.rows.find(row => row.type === 'updated')
  assert.equal(edited.sku, '0099')
  assert.deepEqual(edited.changes.map(change => change.name), ['ITEMNAME', 'ITEMCODE'])
  const report = parseEngineCsv(reportCsv(review)).rows
  assert.equal(report[1][8], 'ยา, "ทดสอบ"\nฉลากใหม่')
  assert.equal(review.candidate, candidate)
})

test('duplicate matching keys, invalid columns and cross-branch comparisons are blocked', () => {
  const samples = demo(), { baseline } = samples.src
  const rows = parseEngineCsv(baseline).rows
  rows[4][1] = rows[3][1]
  assert.throws(() => compare(baseline, serialize(rows)), /ซ้ำ/)
  assert.throws(() => compare(baseline, baseline.replace('CURRENTSTOCK', 'CHANGED')), /คอลัมน์ E/)
  assert.throws(() => compare(baseline, samples.sss.baseline), /สาขาเดียวกัน/)
})

test('empty ItemIDs use unique SKUs, while ambiguous duplicate SKUs are blocked', () => {
  const rows = parseEngineCsv(demo().src.baseline).rows
  rows.slice(3).forEach(row => { row[1] = '' })
  const baseline = serialize(rows)
  rows[3][4] = '25'
  assert.deepEqual(compare(baseline, serialize(rows)).counts, { updated: 1, added: 0, removed: 0, unchanged: 5 })
  rows[4][3] = rows[3][3]
  assert.throws(() => compare(baseline, serialize(rows)), /ซ้ำ/)
})

test('change-report exports contain only changed cells with old/new row references and one BOM', () => {
  const reviews = demo(), report = reportCsv(reviews.src)
  assert.ok(report.startsWith('\uFEFF') && !report.startsWith('\uFEFF\uFEFF'))
  const rows = parseEngineCsv(report).rows
  assert.deepEqual(rows[0], ['FILE', 'STATUS', 'ITEMID', 'SKU', 'OLD_ROW', 'NEW_ROW', 'COLUMN', 'BEFORE', 'AFTER'])
  assert.equal(rows.length, 39)
  assert.ok(rows.slice(1).every(row => row[0] === 'GRAB_SRC.CSV' && row[1] !== 'ไม่เปลี่ยน' && row[3] !== '0001'))
  const stock = rows.find(row => row[3] === '0002' && row[6] === 'CURRENTSTOCK')
  assert.deepEqual(stock.slice(4), ['5', '5', 'CURRENTSTOCK', '20', '35'])
  assert.equal(parseEngineCsv(reportCsv(reviews.sss)).rows.length, 1)
})
