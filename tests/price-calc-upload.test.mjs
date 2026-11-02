import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import { createRequire } from 'node:module'
import ts from 'typescript'

const requireDependency = createRequire(import.meta.url)
const compile = code => ts.transpileModule(code, { compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.CommonJS } }).outputText
const engine = { exports: {} }
new Function('exports', 'require', 'module', compile(fs.readFileSync(new URL('../lib/process-engine.ts', import.meta.url), 'utf8')))(engine.exports, requireDependency, engine)
const source = ts.createSourceFile('page.tsx', fs.readFileSync(new URL('../app/page.tsx', import.meta.url), 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX)
const functions = []
function visit(node) {
  if (ts.isFunctionDeclaration(node) && ['handlePriceCalcUpload', 'parsePriceRobust'].includes(node.name?.text)) functions.push(node.getText(source))
  ts.forEachChild(node, visit)
}
visit(source)
assert.equal(functions.length, 2)
const uploadCode = compile(functions.join('\n'))
const skuKey = 'รหัสสินค้า (SKU NUMBER)'
const product = sku => ({ [skuKey]: sku, '*ชื่อสินค้า (NAME)': `Item ${sku}`, '*ราคาสินค้า': 122, promaxx_price: 100, new_promaxx_price: null })
const header = ['CF_ITEMS_ORDINARY', 'CF_ITEMID', 'CF_ITEMNAME', 'CF_BASEMULTIPLE', 'CF_UNITNAME', 'CF_LEVELNO', 'CF_ITEMPRICELEVEL_PRICE']

function harness(skus, failSave = false) {
  const state = { products: skus.map(product), error: '', results: [], calls: [], saving: false }
  const runtime = {
    priceCalcBusyRef: { current: false },
    setSavingNewPromaxx: value => { state.saving = value },
    setNewPromaxxSaveError: value => { state.error = value },
    setStatus: () => {},
    setShowPriceCalcModal: () => {},
    setPriceCalcResults: value => { state.results = value },
    setProducts: update => { state.products = update(state.products) },
    decodeCsv: engine.exports.decodeCsv,
    parseEngineCsv: engine.exports.parseEngineCsv,
    fetch: async (url, options) => {
      assert.equal(url, '/api/products')
      assert.equal(options.method, 'POST')
      const body = JSON.parse(options.body)
      state.calls.push(body)
      if (Array.isArray(body.skus)) return { ok: true, json: async () => ({ success: true, products: state.products.filter(row => body.skus.includes(row[skuKey])) }) }
      assert.equal(body.updateNewPromaxxPrice, true)
      return { ok: !failSave, json: async () => failSave ? { success: false, error: 'write denied' } : { success: true, updated: Object.keys(body.priceMap).length } }
    },
  }
  const run = new Function(...Object.keys(runtime), uploadCode + '\nreturn handlePriceCalcUpload;')(...Object.values(runtime))
  return { state, upload: bytes => run(new File([bytes], '05105.CSV')) }
}

test('raw R05.105 uses item IDs and base-unit level-0 prices instead of ordinary IDs or numeric unit names', async () => {
  const { state, upload } = harness(['000123', '000124', '777'])
  await upload('\uFEFF' + [header.join(','),
    '777,000123,Gauze 3" x 4",1.0000,10 เม็ด,0,100.0000',
    '777,000123,Gauze 3" x 4",12.0000,กล่อง,0,9999',
    '777,000123,Gauze 3" x 4",1.0000,10 เม็ด,4,0.01',
    '888,000124,Zero,1.0000,แผง,0.0000,0',
  ].join('\r\n'))
  assert.equal(state.error, '')
  assert.deepEqual(state.calls[0].skus, ['000123', '000124'])
  assert.deepEqual(state.calls[1].priceMap, { '000123': 122, '000124': 0 })
  assert.equal(state.results[0].changed, false)
  assert.deepEqual(state.products.map(row => row.new_promaxx_price), [122, 0, null])
  assert.ok(state.products.every(row => row['*ราคาสินค้า'] === 122 && row.promaxx_price === 100))
  assert.equal(state.saving, false)
})

test('raw columns follow their CF headers when reordered, including after an introductory record', async () => {
  const { state, upload } = harness(['000123'])
  await upload('R05.105\nCF_ITEMPRICELEVEL_PRICE,CF_LEVELNO,CF_UNITNAME,CF_ITEMID,CF_BASEMULTIPLE\n100,0,แผง,000123,1\n')
  assert.equal(state.error, '')
  assert.deepEqual(state.calls[1].priceMap, { '000123': 122 })
})

test('quoted names with commas, quotes and embedded newlines retain correct prices for each supported delimiter', async () => {
  for (const delimiter of [',', ';', '\t']) {
    const { state, upload } = harness(['000123'])
    const name = '"Gauze, 3""\nบรรทัดสอง"'
    await upload(header.join(delimiter) + '\n' + ['777', '000123', name, '1', 'แผง', '0', '9'].join(delimiter))
    assert.equal(state.error, '')
    assert.deepEqual(state.calls[1].priceMap, { '000123': 11 })
  }
})

test('TIS-620 reports use the same reader and keep zero-padded SKUs', async () => {
  const { state, upload } = harness(['000123'])
  const bytes = Buffer.concat([Buffer.from(header.join(',') + '\n777,000123,'), Buffer.from([0xc2, 0xd2]), Buffer.from(' 3",1,unit,0,45\n')])
  await upload(bytes)
  assert.equal(state.error, '')
  assert.deepEqual(state.calls[1].priceMap, { '000123': 55 })
})

test('the explicit compact SKU / level-0 format continues to work', async () => {
  const { state, upload } = harness(['000123'])
  await upload('SKU,B,C,D,0\n000123,,,,100\n')
  assert.equal(state.error, '')
  assert.deepEqual(state.calls[1].priceMap, { '000123': 122 })
})

test('missing raw columns, wrong headers, invalid quotes and rows outside base 1 / level 0 never reach the API', async () => {
  const invalid = [
    'CF_ITEMID,CF_ITEMPRICELEVEL_PRICE\n000123,100',
    'A,B,C,D,E\n777,000123,Name,1,100',
    header.join(',') + '\n777,000123,"unterminated,1,unit,0,100',
    header.join(',') + '\n777,000123,Name,2,unit,0,100\n777,000123,Name,1,unit,5,100\n777,000123,Name,1,unit,,100',
    header.join(',') + '\n777,000123,Name,1,unit,0,-10\n777,000123,Name,1,unit,0,invalid',
  ]
  for (const text of invalid) {
    const { state, upload } = harness(['000123'])
    await upload(text)
    assert.notEqual(state.error, '')
    assert.equal(state.calls.length, 0)
    assert.equal(state.saving, false)
    assert.equal(state.products[0].new_promaxx_price, null)
  }
})

test('a rejected New Promaxx save retains its preview and reports failure without updating displayed prices', async () => {
  const { state, upload } = harness(['000123'], true)
  await upload(header.join(',') + '\n777,000123,Name,1,unit,0,100')
  assert.match(state.error, /write denied/)
  assert.equal(state.results[0].newPrice, 122)
  assert.equal(state.products[0].new_promaxx_price, null)
  assert.equal(state.saving, false)
})
