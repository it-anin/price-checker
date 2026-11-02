import { ENGINE_STORES, decodeCsv, validateGrabTemplateUpload, createResultsZip } from './engine.js'

export const columns = ['STOREID', 'ITEMID', 'ITEMNAME', 'ITEMCODE', 'CURRENTSTOCK', 'MAXSTOCK', 'RESTOCKLOCALTIME', 'RESTOCKFREQ', 'ISMONDAYRESTOCK', 'ISTUESDAYRESTOCK', 'ISWEDNESDAYRESTOCK', 'ISTHURSDAYRESTOCK', 'ISFRIDAYRESTOCK', 'ISSATURDAYRESTOCK', 'ISSUNDAYRESTOCK', 'ITEMSTATUS']
export const branches = ['src', 'sss', 'kkl']
export const labels = { updated: 'แก้ไข', added: 'เพิ่มใหม่', removed: 'นำออก', unchanged: 'ไม่เปลี่ยน' }
export const filename = branch => `GRAB_${branch.toUpperCase()}.CSV`
export const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char])
const quote = value => /[",\r\n]/.test(String(value)) ? '"' + String(value).replace(/"/g, '""') + '"' : String(value)
export const csv = rows => rows.map(row => row.map(quote).join(',')).join('\r\n') + '\r\n'

function keyedRows(template) {
  const result = new Map()
  for (const index of template.dataIndices) {
    const cells = template.document.rows[index]
    // ItemID remains stable when a SKU or its name is edited.
    const key = cells[1].trim() ? `item:${cells[1].trim()}` : `sku:${cells[3].trim()}`
    if (result.has(key)) throw new Error(`พบ ItemID หรือ SKU ซ้ำที่แถว ${index + 1} จึงยังเทียบแถวไม่ได้`)
    result.set(key, { cells, row: index + 1 })
  }
  return result
}

export function compare(baseline, candidate) {
  const oldTemplate = validateGrabTemplateUpload(baseline), newTemplate = validateGrabTemplateUpload(candidate)
  if (oldTemplate.branch !== newTemplate.branch) throw new Error('StoreID ของไฟล์เดิมและไฟล์ใหม่ต้องเป็นสาขาเดียวกัน')
  const oldRows = keyedRows(oldTemplate), newRows = keyedRows(newTemplate)
  const rows = []
  for (const key of new Set([...newRows.keys(), ...oldRows.keys()])) {
    const before = oldRows.get(key), after = newRows.get(key)
    const changes = columns.flatMap((name, column) => {
      const oldValue = before?.cells[column], newValue = after?.cells[column]
      return oldValue === newValue ? [] : [{ column, name, before: oldValue, after: newValue }]
    })
    rows.push({ key, itemId: (after ?? before).cells[1], sku: (after ?? before).cells[3], name: (after ?? before).cells[2], before, after, changes, type: !before ? 'added' : !after ? 'removed' : changes.length ? 'updated' : 'unchanged' })
  }
  const counts = Object.fromEntries(Object.keys(labels).map(type => [type, rows.filter(row => row.type === type).length]))
  return { branch: newTemplate.branch, name: newTemplate.name, baseline, candidate, rows, counts, total: newRows.size, changed: rows.some(row => row.type !== 'unchanged') }
}

export async function readFile(file) {
  const { text } = decodeCsv(await file.arrayBuffer())
  return { text, branch: validateGrabTemplateUpload(text).branch }
}

const names = ['พาราเซตามอล 500 มก.', 'วิตามินซี 1,000 มก.', 'หน้ากากอนามัย 3 ชั้น', 'แอลกอฮอล์ 70%', 'น้ำเกลือ 100 มล.', 'พลาสเตอร์กันน้ำ']
function demoRows(branch) {
  return names.map((name, index) => [ENGINE_STORES[branch], `ITEM-${branch.toUpperCase()}-${index + 1}`, name, `000${index + 1}`, '20', '20', '', 'DAILY', '', '', '', '', '', '', '', 'in stock'])
}
function template(rows) {
  return '\uFEFF' + csv([['StoreID', 'ItemID', 'Item Name', 'Item Code', 'Current Stock', 'Max Stock', 'Local Time', 'Restock Settings', '', '', '', '', '', '', '', 'Item Status'], ['IMPORTANT: approved stock template. Keep StoreID and ItemID unchanged.'], columns, ...rows])
}
export function demo() {
  return Object.fromEntries(branches.map(branch => {
    const original = demoRows(branch), next = original.map(row => [...row])
    if (branch === 'src') {
      next[1][4] = next[1][5] = '35'
      next[2][2] = 'หน้ากากอนามัย 3 ชั้น / 50 ชิ้น'
      next[3][4] = next[3][5] = ''; next[3][15] = 'out-of-stock indefinitely'
      next.pop()
      next.push([ENGINE_STORES.src, 'ITEM-SRC-7', 'เจลล้างมือ 50 มล.', '0007', '12', '12', '', 'DAILY', '', '', '', '', '', '', '', 'in stock'])
    }
    if (branch === 'kkl') { next[0][4] = next[0][5] = '28'; next[2][2] = 'หน้ากากอนามัย 3 ชั้น / 50 ชิ้น' }
    return [branch, compare(template(original), template(next))]
  }))
}

export function reportCsv(review) {
  const rows = [['FILE', 'STATUS', 'ITEMID', 'SKU', 'OLD_ROW', 'NEW_ROW', 'COLUMN', 'BEFORE', 'AFTER']]
  for (const row of review.rows.filter(row => row.type !== 'unchanged')) {
    for (const change of row.changes) rows.push([review.name, labels[row.type], row.itemId, row.sku, row.before?.row ?? '', row.after?.row ?? '', change.name, change.before ?? '', change.after ?? ''])
  }
  return '\uFEFF' + csv(rows)
}

export function download(content, name, type = 'text/csv;charset=utf-8') {
  const url = URL.createObjectURL(new Blob([content], { type }))
  const link = document.createElement('a')
  link.href = url; link.download = name; document.body.append(link); link.click(); link.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}
export function exportFile(review, report = false) {
  download(report ? reportCsv(review) : '\uFEFF' + review.candidate.replace(/^\uFEFF/, ''), report ? review.name.replace('.CSV', '_CHANGES.csv') : review.name)
}
export function exportZip(reviews) {
  download(createResultsZip(reviews.map(review => ({ name: review.name, content: '\uFEFF' + review.candidate.replace(/^\uFEFF/, '') }))), 'GRAB_TEMPLATES.zip', 'application/zip')
}
