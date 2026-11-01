const { test } = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const ts = require('typescript')

function loadRoute(pages) {
  const reads = []
  const supabase = { from(table) {
    assert.equal(table, 'products')
    const read = {}; reads.push(read)
    return {
      select(columns) { read.columns = columns; return this },
      order(column) { read.order = column; return this },
      async range(from, to) { read.range = [from, to]; return pages[reads.length - 1] ?? { data: [] } },
    }
  } }
  const filename = path.resolve(__dirname, '../app/api/products/route.ts')
  const code = ts.transpileModule(fs.readFileSync(filename, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS } }).outputText
  const module = { exports: {} }
  const requireMock = name => name === '../../../lib/supabase' ? { supabase } : name === 'next/server' ? { NextResponse: { json: (body, options = {}) => ({ body, ...options }) } } : require(name)
  new Function('require', 'exports', 'module', code)(requireMock, module.exports, module)
  return { GET: module.exports.GET, reads }
}

test('automatic pending scan fetches beyond the first 1000 SKUs in stable order, independently of filters', async () => {
  const page = Array.from({ length: 1000 }, (_, index) => ({ 'รหัสสินค้า (SKU NUMBER)': String(index).padStart(5, '0'), branch: 'src', src_price: 100, promaxx_price: 90 }))
  const { GET, reads } = loadRoute([{ data: page }, { data: [{ 'รหัสสินค้า (SKU NUMBER)': 'LAST', branch: 'kkl,sss', kkl_price: 50, sss_price: 60, promaxx_price: 50 }] }])
  const response = await GET({ url: 'https://app.test/api/products?pendingPriceCheck=true&q=unrelated&sheet=hidden' })
  assert.equal(response.body.success, true)
  assert.equal(response.body.products.length, 1001)
  assert.deepEqual(reads.map(read => read.range), [[0, 999], [1000, 1999]])
  assert.ok(reads.every(read => read.order === '"รหัสสินค้า (SKU NUMBER)"' && read.columns.includes('promaxx_price')))
  assert.equal(response.body.products[1000].sss, true)
  assert.equal(response.body.products[1000].kkl, true)
  assert.equal(response.headers['Cache-Control'], 'no-store')
})

test('a failed scan page returns an error without exposing a partial list as a successful scan', async () => {
  const { GET } = loadRoute([{ data: Array.from({ length: 1000 }, (_, index) => ({ 'รหัสสินค้า (SKU NUMBER)': String(index) })) }, { error: { message: 'Read failed' } }])
  const response = await GET({ url: 'https://app.test/api/products?pendingPriceCheck=true' })
  assert.equal(response.status, 503)
  assert.equal(response.body.success, false)
  assert.equal(response.body.products, undefined)
})
