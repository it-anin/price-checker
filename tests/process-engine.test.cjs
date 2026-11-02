const { test } = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const os = require('node:os')
const ts = require('typescript')
const { execFileSync } = require('node:child_process')

const file = path.join(__dirname, '../lib/process-engine.ts')
const compiled = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText
const engine = { exports: {} }
new Function('exports', 'require', 'module', compiled)(engine.exports, require, engine)
const { ENGINE_STORES, parseEngineCsv, processGrabTemplate, readStockReport, validateGrabTemplate, decodeCsv, pendingSkuSets, createResultsZip, engineDate } = engine.exports

const quote = value => /[",\r\n;]/.test(String(value)) ? '"' + String(value).replace(/"/g, '""') + '"' : String(value)
function template(branch = 'src', rows = []) {
  const header = ['StoreID', 'NAME', 'PRICE', 'ITEMCODE', 'CURRENTSTOCK', 'MAXSTOCK', 'IMAGE', 'DESCRIPTION', 'BARCODE', 'CATEGORY', 'J', 'K', 'L', 'M', 'N', 'STOCK', 'EXTRA']
  return ['"Approved template, keep this line"', header.join(','), ...rows.map(row => [ENGINE_STORES[branch], 'ยา, "ไทย"\nบรรทัดสอง', '45.00', row.sku, '99', '100', 'https://example.com/p.png', '  unchanged  ', 'DIFFERENT-BARCODE', '', 'a', 'b', 'c', 'd', 'e', 'old status', 'keep this too'].map(quote).join(','))].join('\r\n') + '\r\n'
}
function report(branch = 'src', rows = []) {
  return readStockReport(`${branch}.csv`, ['A,B,C,D,SKU,F,QTY', ...rows.map(([sku, qty]) => `,,,,${sku},,${quote(qty)}`)].join('\r\n'), 'UTF-8')
}

test('bare inch marks preserve product names and column boundaries for every delimiter', () => {
  const names = ['UM Finger Cot Splint E01 (M/2-3") NEW 08-26', 'Gauze 3" x 4"', 'Cotton 2"']
  for (const delimiter of [',', ';', '\t']) {
    const input = names.map((name, index) => ['item', name, `000${index}`, '12'].join(delimiter)).join('\r\n') + '\r\n'
    const parsed = parseEngineCsv(input)
    assert.equal(parsed.delimiter, delimiter)
    assert.deepEqual(parsed.rows, names.map((name, index) => ['item', name, `000${index}`, '12']))
    assert.deepEqual(parsed.rawRows.map(row => row[1]), names)
    assert.deepEqual(parsed.endings, ['\r\n', '\r\n', '\r\n'])
  }
})

test('bare inch marks coexist with escaped quotes, embedded delimiters and multiline quoted fields', () => {
  const name = 'UM Finger Cot Splint E01 (M/2-3") NEW 08-26'
  for (const delimiter of [',', ';', '\t']) {
    const description = 'ยา, "ไทย"; description\tcontinued\nบรรทัดสอง'
    const escaped = '"' + description.replace(/"/g, '""') + '"'
    const parsed = parseEngineCsv([name, escaped, '00001', '12'].join(delimiter) + '\r\n')
    assert.equal(parsed.delimiter, delimiter)
    assert.deepEqual(parsed.rows, [[name, description, '00001', '12']])
    assert.equal(parsed.rawRows[0][1], escaped)
    for (const invalid of ['"unclosed', '"a"bad', '"a" "b"']) {
      assert.throws(() => parseEngineCsv(['item', invalid, '12'].join(delimiter)), /เครื่องหมายคำพูด/)
    }
  }
})

test('R01.102 reads row 1893 with an inch-mark name in ordinary TIS-620 and UTF-8 CSVs', () => {
  const name = 'ยา UM Finger Cot Splint E01 (M/2-3") NEW 08-26'
  const lines = ['A,B,NAME,D,SKU,F,QTY', ...Array.from({ length: 1892 }, (_, index) => [
    '', '', index === 1891 ? name : 'product', '', String(index + 1).padStart(5, '0'), '', index === 1891 ? '17' : '1',
  ].join(','))]
  const text = lines.join('\r\n') + '\r\n'
  const tis620 = Uint8Array.from([...text], ch => ch.charCodeAt(0) > 127 ? ch.charCodeAt(0) - 0xd60 : ch.charCodeAt(0))
  for (const bytes of [new TextEncoder().encode(text), tis620]) {
    const decoded = decodeCsv(bytes.buffer)
    const result = readStockReport('src.CSV', decoded.text, decoded.encoding)
    assert.equal(result.rowCount, 1892)
    assert.equal(result.stock.get('01892'), 17)
    assert.equal(parseEngineCsv(decoded.text).rows[1892][2], name)
  }
})

test('pending has priority; subtract once; zero, negative and missing stock are empty', () => {
  const input = template('src', [{ sku: '001' }, { sku: '002' }, { sku: '003' }, { sku: '004' }, { sku: '005' }])
  const stock = report('src', [['001', 12], ['002', 20], ['003', 5], ['004', -2]])
  const result = processGrabTemplate('GRAB_SRC.CSV', input, stock, 5, new Set(['002']))
  const rows = parseEngineCsv(result.content).rows.slice(2)
  assert.deepEqual(rows.map(row => [row[4], row[5], row[15]]), [['7', '7', 'in stock'], ['', '', 'out-of-stock indefinitely'], ['', '', 'out-of-stock indefinitely'], ['', '', 'out-of-stock indefinitely'], ['', '', 'out-of-stock indefinitely']])
  assert.equal(result.matched, 4)
  assert.equal(result.pendingCount, 1)
  assert.equal(result.inStock, 1)
})

test('only E/F/P change; quoted headers, all other raw cells, CRLF and multiline Thai stay identical', () => {
  const input = template('src', [{ sku: '00001' }])
  const output = processGrabTemplate('GRAB_SRC.CSV', input, report('src', [['00001', 10]]), 5, new Set()).content
  const before = parseEngineCsv(input), after = parseEngineCsv(output)
  assert.equal(output.charCodeAt(0), 0xfeff)
  assert.deepEqual(after.rawRows.slice(0, 2), before.rawRows.slice(0, 2))
  before.rawRows[2].forEach((cell, column) => { if (![4, 5, 15].includes(column)) assert.equal(after.rawRows[2][column], cell) })
  assert.deepEqual(after.endings, before.endings)
  assert.equal(after.rows[2][3], '00001')
  assert.equal(after.rows[2][8], 'DIFFERENT-BARCODE')
})

test('duplicate report SKUs are summed before subtracting; alphanumeric and leading zero keys survive', () => {
  const stock = report('src', [['001-A', 4], ['001-A', 7], ['002', '1,234']])
  const result = processGrabTemplate('GRAB_SRC.CSV', template('src', [{ sku: '001-A' }]), stock, 5, new Set())
  assert.equal(parseEngineCsv(result.content).rows[2][4], '6')
  assert.equal(stock.stock.get('002'), 1234)
})

test('invalid template names, headers and mixed StoreIDs are blocked; 2 or 3 header rows work', () => {
  const valid = template('src', [{ sku: '001' }])
  assert.equal(validateGrabTemplate('GRAB_SRC.CSV', valid).dataStart, 2)
  assert.equal(validateGrabTemplate('GRAB_SRC.CSV', 'Instructions\r\n' + valid).dataStart, 3)
  assert.throws(() => validateGrabTemplate('grab_src.csv', valid), /ตัวพิมพ์ใหญ่/)
  assert.throws(() => validateGrabTemplate('GRAB_SRC.CSV', valid.replace('CURRENTSTOCK', 'BAD')), /โครงสร้าง/)
  assert.throws(() => validateGrabTemplate('GRAB_SSS.CSV', valid), /StoreID/)
  assert.throws(() => validateGrabTemplate('GRAB_SRC.CSV', valid + template('sss', [{ sku: '002' }]).split('\r\n').slice(2).join('\r\n')), /StoreID/)
  assert.throws(() => processGrabTemplate('GRAB_SRC.CSV', valid, report('sss', [['001', 10]]), 2, new Set()), /สาขา/)
})

test('bad CSV and invalid report quantities do not produce stock outputs', () => {
  assert.throws(() => parseEngineCsv('"unclosed'), /ยังไม่ปิด/)
  assert.throws(() => parseEngineCsv('"a"bad,b'), /หลังปิด/)
  assert.throws(() => report('src', [['001', '12bad']]), /คอลัมน์ G/)
  assert.throws(() => report('src', []), /ไม่พบข้อมูล/)
  const input = template('src', [{ sku: '001' }])
  assert.throws(() => processGrabTemplate('GRAB_SRC.CSV', input, report('src', [['001', 10]]), -1, new Set()), /SUBTRACT/)
})

test('actual Grab stock-export headers use C=ITEMNAME and D=ITEMCODE without a price column', () => {
  const rows = [
    ['StoreID', 'ItemID', 'Item Name', 'Item Code', 'Current Stock', 'Max Stock', 'Local Time', 'Restock Settings', '', '', '', '', '', '', '', 'Item Status'],
    ['IMPORTANT: do not change store IDs', 'IMPORTANT: item IDs', 'IMPORTANT: names', '', 'Actual stock\nquantity', 'Maximum stock'],
    ['STOREID', 'ITEMID', 'ITEMNAME', 'ITEMCODE', 'CURRENTSTOCK', 'MAXSTOCK', 'RESTOCKLOCALTIME', 'RESTOCKFREQ', 'ISMONDAYRESTOCK', 'ISTUESDAYRESTOCK', 'ISWEDNESDAYRESTOCK', 'ISTHURSDAYRESTOCK', 'ISFRIDAYRESTOCK', 'ISSATURDAYRESTOCK', 'ISSUNDAYRESTOCK', 'ITEMSTATUS'],
    [ENGINE_STORES.src, 'ITEM-123', 'ยาทดสอบ', '001', '20', '20', '', '', '', '', '', '', '', '', '', 'in stock'],
  ]
  const input = rows.map(row => row.map(quote).join(',')).join('\r\n')
  const result = processGrabTemplate('GRAB_SRC.CSV', input, report('src', [['001', 10]]), 5, new Set())
  assert.equal(validateGrabTemplate('GRAB_SRC.CSV', input).dataStart, 3)
  assert.equal(parseEngineCsv(result.content).rows[3][4], '5')
  assert.deepEqual(parseEngineCsv(result.content).rawRows.slice(0, 3), parseEngineCsv(input).rawRows.slice(0, 3))
})

test('zero SKU matches is reported and every stock field is empty', () => {
  const result = processGrabTemplate('GRAB_SRC.CSV', template('src', [{ sku: '001' }]), report('src', [['999', 100]]), 5, new Set())
  assert.equal(result.matched, 0)
  assert.equal(parseEngineCsv(result.content).rows[2][4], '')
})

test('legacy pending JSON keeps branch membership and missing-branch defaults', () => {
  const pending = pendingSkuSets([{ sku: '001', branches: ['SRC'] }, { sku: '002', branches: ['SSS', 'KKL'] }, { sku: '003' }, { sku: '004', branches: ['WRD', 'ONN'] }])
  assert.deepEqual([...pending.src], ['001', '003'])
  assert.deepEqual([...pending.sss], ['002', '003'])
  assert.deepEqual([...pending.kkl], ['002', '003'])
  assert.throws(() => pendingSkuSets({}), /JSON array/)
  assert.throws(() => pendingSkuSets([{ sku: '001', branches: ['UNKNOWN'] }]), /สาขา/)
})

test('UTF-8 and TIS-620 Thai decode correctly; date follows Bangkok', () => {
  assert.equal(decodeCsv(new TextEncoder().encode('\uFEFFยา').buffer).text, 'ยา')
  assert.deepEqual(decodeCsv(Uint8Array.from([0xc2, 0xd2]).buffer), { text: 'ยา', encoding: 'TIS-620' })
  assert.equal(engineDate(new Date('2026-10-30T18:00:00Z')), '31-10-2026')
})

test('semicolon CSV remains semicolon-separated and preserves every other field', () => {
  const input = template('src', [{ sku: '001' }])
  const document = parseEngineCsv(input)
  const semicolon = document.rows.map(row => row.map(quote).join(';')).join('\r\n')
  const output = processGrabTemplate('GRAB_SRC.CSV', semicolon, report('src', [['001', 10]]), 5, new Set())
  assert.equal(parseEngineCsv(output.content).delimiter, ';')
  assert.equal(parseEngineCsv(output.content).rows[2][4], '5')
})

test('DOWNLOAD ALL creates a valid ZIP with intact BOM, Thai and all CSV filenames', () => {
  const files = [{ name: 'GRAB_SRC_30-10-2026.csv', content: '\uFEFFยา,7\r\n' }, { name: 'GRAB_SSS_30-10-2026.csv', content: '\uFEFFไทย,2\r\n' }]
  const zip = createResultsZip(files)
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'grab-zip-test-'))
  try {
    const archive = path.join(directory, 'results.zip')
    fs.writeFileSync(archive, zip)
    const result = execFileSync('python', ['-c', 'import sys,zipfile,json; z=zipfile.ZipFile(sys.argv[1]); assert z.testzip() is None; print(json.dumps([{ "name":n, "content":z.read(n).decode("utf-8")} for n in z.namelist()]))', archive], { timeout: 10000 })
    assert.deepEqual(JSON.parse(result), files)
  } finally { fs.rmSync(directory, { recursive: true, force: true }) }
})
