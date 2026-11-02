// Pure helpers (no DB access) so they can be unit-tested without MSSQL.

export const BRANCHES = ['src', 'sss', 'kkl']

// CF_QUANTITY of R01.102 (d_01102_classic_p22112018), rewritten as one GROUP BY
// instead of a correlated subquery per item x warehouse. Every condition of the
// original report is kept; rows with no movement (NULL in the report) never appear.
export function buildStockQuery(warehouseIds = []) {
  const warehouseFilter = warehouseIds.length
    ? `\n  AND H.SYSWAREHOUSEID IN (${warehouseIds.map((_, i) => `@w${i}`).join(', ')})`
    : ''
  return `SELECT I.ITEMID AS sku, SUM(D.BASEQUANTITY * V.FSTOCKMAIN) AS qty
FROM TRANS H
JOIN VOUCHER V ON H.SYSVOUCHERID = V.SYSVOUCHERID
JOIN TRANDETAILITEM D ON H.SYSTRANNO = D.SYSTRANNO
JOIN ITEMS I ON I.SYSITEMID = D.SYSITEMID
JOIN WAREHOUSE W ON W.SYSBRANCHID = H.SYSBRANCHID AND W.SYSWAREHOUSEID = H.SYSWAREHOUSEID
WHERE H.FTaxInvoiceReplacement = 0
  AND H.FVoidTaxInvoice = 0
  AND V.FSTOCKMAIN <> 0
  AND D.FPROCESS = 1
  AND H.FCANCEL = 0
  AND H.FAPPROVE = 1
  AND I.ITEMID NOT IN ('DEPOSIT', 'DEPOSIT-NV')
  AND I.FEFFECTSTOCK = 1
  AND I.FSHOWONREPORT = 1
  AND I.FUSED = 1
  AND H.SYSBRANCHID = @branchId${warehouseFilter}
GROUP BY I.ITEMID`
}

// STOCK_BRANCH_MAP example: {"src":{"branchId":1},"sss":{"branchId":2,"warehouseIds":[1,3]},"kkl":{"branchId":3}}
export function parseBranchMap(text) {
  let raw
  try { raw = JSON.parse(text) } catch { throw new Error('STOCK_BRANCH_MAP ไม่ใช่ JSON ที่ถูกต้อง') }
  const map = {}
  for (const branch of BRANCHES) {
    const entry = raw?.[branch]
    if (!entry) continue
    if (!Number.isInteger(entry.branchId)) throw new Error(`STOCK_BRANCH_MAP.${branch}.branchId ต้องเป็นจำนวนเต็ม`)
    const warehouseIds = entry.warehouseIds ?? []
    if (!Array.isArray(warehouseIds) || !warehouseIds.every(Number.isInteger)) throw new Error(`STOCK_BRANCH_MAP.${branch}.warehouseIds ต้องเป็นอาร์เรย์ของจำนวนเต็ม`)
    map[branch] = { branchId: entry.branchId, warehouseIds }
  }
  if (!Object.keys(map).length) throw new Error('STOCK_BRANCH_MAP ต้องมีอย่างน้อยหนึ่งสาขา (src, sss, kkl)')
  return map
}

// Trim SKUs and round away float noise; skip empty SKUs and non-finite quantities.
export function toStockRows(records, branch, syncedAt) {
  const bySku = new Map()
  for (const record of records) {
    const sku = String(record.sku ?? '').trim().toUpperCase()
    const qty = record.qty == null ? Number.NaN : Number(record.qty)
    if (!sku || !Number.isFinite(qty)) continue
    bySku.set(sku, (bySku.get(sku) ?? 0) + qty)
  }
  return [...bySku].map(([sku, qty]) => ({ branch, sku, qty: Number(qty.toFixed(4)), synced_at: syncedAt }))
}

// Builds the mssql connection config from environment variables.
// A named instance (e.g. BigyaMainPC\SQLSERVER) must NOT be combined with a port:
// tedious rejects both together, so MSSQL_INSTANCE replaces MSSQL_PORT.
export function buildConnectionConfig(env) {
  for (const name of ['MSSQL_SERVER', 'MSSQL_DATABASE', 'MSSQL_USER', 'MSSQL_PASSWORD']) {
    if (!env[name]) throw new Error(`ยังไม่ได้ตั้ง ${name}`)
  }
  const config = {
    server: env.MSSQL_SERVER,
    database: env.MSSQL_DATABASE,
    user: env.MSSQL_USER,
    password: env.MSSQL_PASSWORD,
    options: { encrypt: env.MSSQL_ENCRYPT === 'true', trustServerCertificate: env.MSSQL_TRUST_CERT !== 'false' },
    requestTimeout: 120000,
  }
  if (env.MSSQL_INSTANCE) config.options.instanceName = env.MSSQL_INSTANCE
  else config.port = Number(env.MSSQL_PORT || 1433)
  return config
}
