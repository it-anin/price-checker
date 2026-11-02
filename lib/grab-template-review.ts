import { normalizeStockSku, validateGrabTemplate, validateGrabTemplateUpload, type EngineBranch } from './process-engine'

export const GRAB_REVIEW_COLUMNS = [
  'STOREID', 'ITEMID', 'ITEMNAME', 'ITEMCODE', 'CURRENTSTOCK', 'MAXSTOCK',
  'RESTOCKLOCALTIME', 'RESTOCKFREQ', 'ISMONDAYRESTOCK', 'ISTUESDAYRESTOCK',
  'ISWEDNESDAYRESTOCK', 'ISTHURSDAYRESTOCK', 'ISFRIDAYRESTOCK',
  'ISSATURDAYRESTOCK', 'ISSUNDAYRESTOCK', 'ITEMSTATUS',
] as const
export const GRAB_CHANGE_LABELS = { updated: 'แก้ไข', added: 'เพิ่มใหม่', removed: 'นำออก', unchanged: 'ไม่เปลี่ยน' } as const
export type GrabChangeType = keyof typeof GRAB_CHANGE_LABELS
export interface GrabReviewRecord { cells: string[]; row: number }
export interface GrabCellChange { column: number; name: string; before: string | undefined; after: string | undefined }
export interface GrabRowChange {
  key: string
  itemId: string
  sku: string
  name: string
  before: GrabReviewRecord | undefined
  after: GrabReviewRecord | undefined
  changes: GrabCellChange[]
  type: GrabChangeType
}
export interface GrabTemplateReview {
  branch: EngineBranch
  name: string
  candidate: string
  rows: GrabRowChange[]
  counts: Record<GrabChangeType, number>
  total: number
  isNew: boolean
  changed: boolean
}
export interface SavedGrabReview extends GrabTemplateReview { sourceName: string; updatedAt: string }

function keyedRows(template: ReturnType<typeof validateGrabTemplate>) {
  const result = new Map<string, GrabReviewRecord>()
  const itemIdHeader = template.document.rows[template.headerIndex][1]?.trim().toUpperCase() === 'ITEMID'
  for (const index of template.dataIndices) {
    const cells = template.document.rows[index]
    const key = itemIdHeader && cells[1].trim() ? `item:${cells[1].trim()}` : `sku:${normalizeStockSku(cells[3])}`
    if (result.has(key)) throw new Error(`แถว ${index + 1}: พบ ItemID หรือ SKU ซ้ำ จึงยังเทียบการเปลี่ยนแปลงไม่ได้`)
    result.set(key, { cells, row: index + 1 })
  }
  return result
}

export function compareGrabTemplates(baseline: string | null, candidate: string): GrabTemplateReview {
  const next = validateGrabTemplateUpload(candidate)
  const before = baseline === null ? new Map<string, GrabReviewRecord>() : keyedRows(validateGrabTemplate(next.name, baseline))
  const after = keyedRows(next)
  const rows: GrabRowChange[] = []
  const counts: Record<GrabChangeType, number> = { updated: 0, added: 0, removed: 0, unchanged: 0 }
  for (const key of new Set([...after.keys(), ...before.keys()])) {
    const oldRow = before.get(key), newRow = after.get(key), record = newRow ?? oldRow!
    const changes = GRAB_REVIEW_COLUMNS.flatMap((name, column) => {
      const oldValue = oldRow?.cells[column], newValue = newRow?.cells[column]
      return oldValue === newValue ? [] : [{ column, name, before: oldValue, after: newValue }]
    })
    const type: GrabChangeType = !oldRow ? 'added' : !newRow ? 'removed' : changes.length ? 'updated' : 'unchanged'
    counts[type]++
    rows.push({ key, itemId: record.cells[1], sku: record.cells[3], name: record.cells[2], before: oldRow, after: newRow, changes, type })
  }
  return { branch: next.branch, name: next.name, candidate, rows, counts, total: after.size, isNew: baseline === null, changed: baseline !== null && counts.updated + counts.added + counts.removed > 0 }
}

export function grabCsvWithBom(text: string) {
  return '\uFEFF' + text.replace(/^\uFEFF+/, '')
}

export function grabChangesCsv(review: GrabTemplateReview) {
  const rows: (string | number)[][] = [['FILE', 'STATUS', 'ITEMID', 'SKU', 'OLD_ROW', 'NEW_ROW', 'COLUMN', 'BEFORE', 'AFTER']]
  for (const row of review.rows) {
    if (row.type === 'unchanged') continue
    for (const change of row.changes) rows.push([review.name, GRAB_CHANGE_LABELS[row.type], row.itemId, row.sku, row.before?.row ?? '', row.after?.row ?? '', change.name, change.before ?? '', change.after ?? ''])
  }
  const quote = (value: string | number) => /[",\r\n]/.test(String(value)) ? '"' + String(value).replace(/"/g, '""') + '"' : String(value)
  return '\uFEFF' + rows.map(row => row.map(quote).join(',')).join('\r\n') + '\r\n'
}
