import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import ts from 'typescript'

const requireDependency = createRequire(import.meta.url)

const skuColumn = 'รหัสสินค้า (SKU NUMBER)'
const product = (sku, price = null) => ({ [skuColumn]: sku, '*ราคาสินค้า': 122, promaxx_price: 100, new_promaxx_price: price, branch: 'src' })

function loadRoute(products, options = {}) {
  const state = { products: new Map(products.map(row => [row[skuColumn], { ...row }])), writes: [], reads: [], active: 0, maxActive: 0 }
  const supabase = { from(table) {
    assert.equal(table, 'products')
    let patch, sku
    return {
      select(columns) { state.reads.push(columns); return this },
      update(value) { patch = value; return this },
      eq(column, value) { assert.equal(column, `"${skuColumn}"`); sku = value; return this },
      async limit() { return { data: [...state.products.values()] } },
      async in(column, skus) { assert.equal(column, `"${skuColumn}"`); return { data: skus.flatMap(value => state.products.has(value) ? [state.products.get(value)] : []) } },
      async maybeSingle() {
        state.active++; state.maxActive = Math.max(state.maxActive, state.active)
        try {
          await new Promise(resolve => setImmediate(resolve))
          state.writes.push({ sku, patch })
          if (sku === options.failSku) return { error: { message: 'write denied' } }
          if (!state.products.has(sku) || sku === options.hiddenSku) return { data: null }
          Object.assign(state.products.get(sku), patch)
          return { data: { [skuColumn]: sku } }
        } finally { state.active-- }
      },
    }
  } }
  const filename = fileURLToPath(new URL('../app/api/products/route.ts', import.meta.url))
  const code = ts.transpileModule(fs.readFileSync(filename, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS } }).outputText
  const routeModule = { exports: {} }
  const requireMock = name => name === '../../../lib/supabase' ? { supabase } : name === 'next/server' ? { NextResponse: { json: (body, options = {}) => ({ body, ...options }) } } : requireDependency(name)
  new Function('require', 'exports', 'module', code)(requireMock, routeModule.exports, routeModule)
  return { ...routeModule.exports, state, save: priceMap => routeModule.exports.POST({ json: async () => ({ updateNewPromaxxPrice: true, priceMap }) }) }
}

test('New Promaxx persists separately from Promaxx/base prices and is returned by a subsequent product load', async () => {
  const route = loadRoute([product('001'), product('002'), product('003', 70)])
  const saved = await route.save({ '001': 122, '002': 0 })
  assert.deepEqual(saved.body, { success: true, updated: 2 })
  const reloaded = await route.GET({ url: 'https://app.test/api/products' })
  assert.equal(reloaded.body.products.find(row => row[skuColumn] === '001').new_promaxx_price, 122)
  assert.equal(reloaded.body.products.find(row => row[skuColumn] === '002').new_promaxx_price, 0)
  assert.equal(reloaded.body.products.find(row => row[skuColumn] === '003').new_promaxx_price, 70)
  assert.ok(reloaded.body.products.every(row => row['*ราคาสินค้า'] === 122 && row.promaxx_price === 100))
  assert.ok(route.state.reads.every(columns => columns.includes('new_promaxx_price') || columns === `"${skuColumn}"`))
  assert.ok(route.state.writes.every(write => Object.keys(write.patch).join() === 'new_promaxx_price'))
})

test('SKU lookup also includes persisted New Promaxx, preserving leading zeroes and missing values', async () => {
  const route = loadRoute([product('001', 122), product('1')])
  const result = await route.POST({ json: async () => ({ skus: ['001', '1'] }) })
  assert.deepEqual(result.body.products.map(row => [row[skuColumn], row.new_promaxx_price]), [['001', 122], ['1', null]])
  assert.equal(route.state.writes.length, 0)
})

test('300 calculations are saved with at most ten concurrent updates', async () => {
  const rows = Array.from({ length: 300 }, (_, index) => product(String(index).padStart(4, '0')))
  const route = loadRoute(rows)
  const result = await route.save(Object.fromEntries(rows.map((row, index) => [row[skuColumn], index])))
  assert.equal(result.body.updated, 300)
  assert.equal(route.state.writes.length, 300)
  assert.equal(route.state.maxActive, 10)
})

test('invalid maps, values and oversized batches are rejected before any write', async () => {
  const route = loadRoute([product('001')])
  const invalid = [null, [], 1, {}, { '': 1 }, { ' ': 1 }, { '001': null }, { '001': '122' }, { '001': -1 }, { '001': Infinity }, { '001': NaN }, Object.fromEntries(Array.from({ length: 301 }, (_, index) => [String(index), 1]))]
  for (const priceMap of invalid) {
    const response = await route.save(priceMap)
    assert.equal(response.status, 400)
    assert.equal(response.body.success, false)
  }
  assert.equal(route.state.writes.length, 0)
})

test('missing or RLS-hidden products never become successful saves or inserted rows', async () => {
  for (const hiddenSku of [undefined, '001']) {
    const route = loadRoute(hiddenSku ? [product('001')] : [], { hiddenSku })
    const response = await route.save({ '001': 122 })
    assert.equal(response.status, 503)
    assert.equal(response.body.success, false)
    assert.equal(response.body.updated, 0)
    if (hiddenSku) assert.equal(route.state.products.get('001').new_promaxx_price, null)
    else assert.equal(route.state.products.size, 0)
  }
})

test('a partially failed write reports its verified count and stops subsequent batches', async () => {
  const rows = Array.from({ length: 15 }, (_, index) => product(`SKU${index}`))
  const route = loadRoute(rows, { failSku: 'SKU3' })
  const result = await route.save(Object.fromEntries(rows.map(row => [row[skuColumn], 122])))
  assert.equal(result.status, 503)
  assert.equal(result.body.success, false)
  assert.equal(result.body.updated, 9)
  assert.match(result.body.error, /write denied/)
  assert.equal(route.state.writes.length, 10)
  assert.equal(route.state.products.get('SKU3').new_promaxx_price, null)
  assert.equal(route.state.products.get('SKU10').new_promaxx_price, null)
})
