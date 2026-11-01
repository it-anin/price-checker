import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import ts from 'typescript'

const directory = path.dirname(fileURLToPath(import.meta.url))
const compilerOptions = { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 }
const engine = { exports: {} }
const code = ts.transpileModule(fs.readFileSync(path.join(directory, '../lib/process-engine.ts'), 'utf8'), { compilerOptions }).outputText
new Function('exports', code)(engine.exports)
const { validateGrabTemplateUpload, parseEngineCsv, decodeCsv, processGrabTemplate } = engine.exports

const headers = 'STOREID,ITEMID,ITEMNAME,ITEMCODE,CURRENTSTOCK,MAXSTOCK,RESTOCKLOCALTIME,RESTOCKFREQ,ISMONDAYRESTOCK,ISTUESDAYRESTOCK,ISWEDNESDAYRESTOCK,ISTHURSDAYRESTOCK,ISFRIDAYRESTOCK,ISSATURDAYRESTOCK,ISSUNDAYRESTOCK,ITEMSTATUS'.split(',')
const stores = { sss: '3-C7MKLKLYL8LYAJ', src: '3-C3JVCCKYNU23TA', kkl: '3-C7AAJUMCGAE3RE' }
const quote = value => /[",\r\n;]/.test(value) ? '"' + value.replace(/"/g, '""') + '"' : value
const csv = (rows, delimiter = ',') => rows.map(row => row.map(quote).join(delimiter)).join('\r\n') + '\r\n'
function stockRows(branch = 'src') {
  return [
    ['StoreID', 'ItemID', 'Item Name', 'Item Code', 'Current Stock', 'Max Stock', 'Local Time', 'Restock Settings', '', '', '', '', '', '', '', 'Item Status'],
    ['IMPORTANT: do not change StoreIDs', 'Keep item IDs', 'ชื่อสินค้า, ห้ามแก้ไข\nบรรทัดที่สอง'],
    [...headers],
    [stores[branch], 'ITEM-123', 'ยาทดสอบ', '00001', '20', '20', '', '', '', '', '', '', '', '', '', 'in stock'],
    [stores[branch], 'ITEM-456', 'ยาอีกชนิด', '002-A', '', '', '', '', '', '', '', '', '', '', '', 'out-of-stock indefinitely'],
  ]
}

test('upload determines all three canonical filenames from StoreID, regardless of source filename', () => {
  for (const branch of Object.keys(stores)) {
    const text = '\uFEFF' + csv(stockRows(branch))
    for (const source of ['Grab stock export (1).CSV', 'export.csv', 'GRAB_SRC.CSV', 'GRAB_SSS.CSV', 'GRAB_KKL.CSV']) {
      const result = validateGrabTemplateUpload(source, text)
      assert.equal(result.branch, branch)
      assert.equal(result.name, `GRAB_${branch.toUpperCase()}.CSV`)
      assert.equal(result.headerIndex, 2)
      assert.deepEqual(result.dataIndices, [3, 4])
      assert.deepEqual(result.document.rows, parseEngineCsv(text).rows)
    }
  }
})

test('every machine header column is required verbatim in its approved position', () => {
  for (let column = 0; column < headers.length; column++) {
    for (const wrong of ['BAD', headers[column].toLowerCase(), ` ${headers[column]} `, '']) {
      const rows = stockRows()
      rows[2][column] = wrong
      assert.throws(() => validateGrabTemplateUpload('export.csv', csv(rows)), /หัว|โครงสร้าง|คอลัมน์/, `column ${column}: ${JSON.stringify(wrong)}`)
    }
  }
  const swapped = stockRows()
  ;[swapped[2][8], swapped[2][9]] = [swapped[2][9], swapped[2][8]]
  assert.throws(() => validateGrabTemplateUpload('export.csv', csv(swapped)), /คอลัมน์ I/)
  const duplicate = stockRows()
  duplicate[2][14] = duplicate[2][13]
  assert.throws(() => validateGrabTemplateUpload('export.csv', csv(duplicate)), /คอลัมน์ O/)
})

test('missing and extra columns in either the header or any data row block uploads', () => {
  for (const rowIndex of [2, 3, 4]) {
    for (const column of [0, 6, 15]) {
      const rows = stockRows()
      rows[rowIndex].splice(column, 1)
      assert.throws(() => validateGrabTemplateUpload('export.csv', csv(rows)))
    }
    for (const extra of ['', 'EXTRA']) {
      const rows = stockRows()
      rows[rowIndex].push(extra)
      assert.throws(() => validateGrabTemplateUpload('export.csv', csv(rows)), /16 คอลัมน์/)
    }
  }
})

test('unknown, empty, and mixed StoreIDs are blocked in the first or later record', () => {
  for (const rowIndex of [3, 4]) {
    for (const storeId of ['', '3-UNKNOWN', 'UNKNOWN', stores.sss, stores.kkl]) {
      const rows = stockRows()
      rows[rowIndex][0] = storeId
      assert.throws(() => validateGrabTemplateUpload('export.csv', csv(rows)), /StoreID/)
    }
  }
  for (const field of [2, 3]) {
    const rows = stockRows()
    rows[3][field] = ''
    assert.throws(() => validateGrabTemplateUpload('export.csv', csv(rows)), /ก่อนแถวเริ่มต้น|ไม่ถูกต้อง/)
  }
  const hiddenRecord = stockRows()
  hiddenRecord.splice(3, 0, ['UNKNOWN', ...Array(15).fill('')])
  assert.throws(() => validateGrabTemplateUpload('export.csv', csv(hiddenRecord)), /StoreID/)
})

test('non-CSV, empty, malformed, price-menu, and legacy price-bearing exports cannot replace approved stock templates', () => {
  const valid = csv(stockRows())
  assert.throws(() => validateGrabTemplateUpload('export.xlsx', valid), /\.CSV/)
  for (const text of ['', csv(stockRows().slice(0, 3)), '"unclosed', 'StoreID,NAME,PRICE,D,E,F,G,H,SKU\n' + stores.src + ',ยา,10,,,,,,001']) {
    assert.throws(() => validateGrabTemplateUpload('export.csv', text))
  }
  const priceBearing = stockRows()
  priceBearing[2][2] = 'PRICE'
  priceBearing[3][2] = '45.00'; priceBearing[4][2] = '50.00'
  assert.throws(() => validateGrabTemplateUpload('export.csv', csv(priceBearing)), /ITEMNAME/)
})

test('BOM, UTF-8/TIS-620, multiline instructions, blank lines and delimiters remain usable by PROCESS ENGINE', () => {
  for (const branch of Object.keys(stores)) {
    const rows = stockRows(branch)
    rows.push([''])
    for (const delimiter of [',', ';', '\t']) {
      const text = '\uFEFF' + csv(rows, delimiter)
      const template = validateGrabTemplateUpload('export.csv', text)
      const report = { branch, stock: new Map([['00001', 12]]), name: `${branch}.csv`, rowCount: 1, encoding: 'UTF-8' }
      const output = processGrabTemplate(template.name, text, report, 2, new Set())
      assert.equal(output.total, 2)
      const after = parseEngineCsv(output.content)
      assert.equal(after.rows[3][4], '10')
      assert.equal(after.rows[3][5], '10')
      assert.equal(after.rows[3][15], 'in stock')
      assert.deepEqual(after.rawRows.slice(0, 3), template.document.rawRows.slice(0, 3))
    }
  }
  const thaiRows = stockRows()
  thaiRows.splice(0, 2)
  thaiRows[1][2] = thaiRows[2][2] = 'ยา'
  const bytes = Uint8Array.from([...csv(thaiRows)].map(char => char === 'ย' ? 0xc2 : char === 'า' ? 0xd2 : char.charCodeAt(0)))
  const decoded = decodeCsv(bytes.buffer)
  assert.equal(decoded.encoding, 'TIS-620')
  const template = validateGrabTemplateUpload('export.csv', decoded.text)
  assert.equal(template.document.rows[1][2], 'ยา')
})

// Exercise the actual page handler with a fake Supabase table, so the upload
// guard, stored filename, cache key and mixed valid/invalid batches are covered.
function uploadHandler(context) {
  const source = ts.createSourceFile('page.tsx', fs.readFileSync(path.join(directory, '../app/page.tsx'), 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX)
  let declaration
  const visit = node => {
    if (ts.isFunctionDeclaration(node) && node.name?.text === 'uploadGrabTemplates') declaration = node
    ts.forEachChild(node, visit)
  }
  visit(source)
  assert.ok(declaration)
  const compiled = ts.transpileModule(declaration.getText(source), { compilerOptions }).outputText
  return new Function(...Object.keys(context), compiled + '\nreturn uploadGrabTemplates;')(...Object.values(context))
}

test('the manager writes only validated files under canonical names and keeps rejected templates intact', async () => {
  const original = csv(stockRows())
  const rows = stockRows(); rows[2][7] = 'WRONG'
  const invalid = csv(rows)
  const writes = [], logs = [], alerts = [], busy = []
  const cache = new Map(Object.keys(stores).map(branch => [`GRAB_${branch.toUpperCase()}.CSV`, 'cached names']))
  let invalidations = 0, refreshes = 0
  const engineBusyRef = { current: false }
  const handler = uploadHandler({
    grabManagerUnlocked: true, engineBusyRef, decodeCsv, validateGrabTemplateUpload,
    setEngineBusy: value => busy.push(value),
    supabase: { from(table) {
      assert.equal(table, 'grab_files')
      return { async upsert(row, options) { writes.push({ row, options }); return { error: null } } }
    } },
    pendingNameCache: { current: cache }, logEngine: message => logs.push(message),
    invalidateEngineResults: () => invalidations++, setEngineChecks: () => {},
    refreshGrabFiles: async () => { refreshes++ }, engineMessage: error => error.message,
    window: { alert: message => alerts.push(message) },
  })
  const file = (name, text) => ({ name, arrayBuffer: async () => new TextEncoder().encode(text).buffer })
  await handler([file('GRAB_SSS.CSV', original), file('bad.csv', invalid), file('another export.csv', csv(stockRows('kkl')))])
  assert.deepEqual(writes.map(write => write.row.name), ['GRAB_SRC.CSV', 'GRAB_KKL.CSV'])
  assert.equal(writes[0].row.content, original)
  assert.ok(writes.every(write => write.options.onConflict === 'name' && Number.isFinite(Date.parse(write.row.updated_at))))
  assert.deepEqual([...cache.keys()], ['GRAB_SSS.CSV'])
  assert.ok(logs.some(message => message.includes('GRAB_SSS.CSV → GRAB_SRC.CSV')))
  assert.equal(alerts.length, 1); assert.match(alerts[0], /BLOCKED 1/)
  assert.equal(invalidations, 1); assert.equal(refreshes, 1)
  assert.deepEqual(busy, ['templates', '']); assert.equal(engineBusyRef.current, false)

  await handler([file('all rejected.csv', invalid)])
  assert.equal(writes.length, 2)
  assert.equal(invalidations, 1)
})
