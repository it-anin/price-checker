import { branches, columns, labels, demo, compare, readFile, escapeHtml as esc, exportFile, exportZip, download } from './review.js'

const id = Number(document.body.dataset.design)
const designs = await fetch('../concepts.json').then(response => response.json())
const design = designs.find(item => item.id === id)
let reviews = demo()
const state = { branch: 'src', row: null, filter: 'all', search: '', step: 2, selected: new Set(branches), problems: [] }
const root = document.querySelector('#app')
const paths = {
  upload: '<path d="M12 16V3m-5 5 5-5 5 5M4 16v5h16v-5"/>',
  download: '<path d="M12 3v13m-5-5 5 5 5-5M4 17v4h16v-4"/>',
  file: '<path d="M14 2H4v20h16V8l-6-6Zm0 0v6h6M8 13h8M8 17h6"/>',
  check: '<path d="m5 12 4 4L19 6"/>',
  change: '<path d="M4 7h15m-4-4 4 4-4 4M20 17H5m4-4-4 4 4 4"/>',
  search: '<circle cx="10" cy="10" r="6"/><path d="m15 15 6 6"/>',
  arrow: '<path d="M4 12h16m-6-6 6 6-6 6"/>',
  grid: '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v6l4 2"/>',
}
const icon = name => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${paths[name] ?? paths.file}</svg>`
const button = (label, action, extra = '', symbol = 'download', classes = '') => `<button class="button ${classes}" data-action="${action}" ${extra}>${icon(symbol)}<span>${label}</span></button>`
const upload = () => button('อัปไฟล์เพื่อเปรียบเทียบ', 'upload', '', 'upload', 'primary')
const exportOne = branch => button('Export CSV', 'csv', `data-branch="${branch}"`)
const exportReport = branch => button('Export รายงาน', 'report', `data-branch="${branch}"`, 'file', 'quiet')
const exportAll = () => button('Export ทั้งหมด .ZIP', 'zip', state.selected.size ? '' : 'disabled', 'download', 'primary')
const title = (eyebrow, heading = 'GRAB CSV MANAGER', subtitle = 'ตรวจทุกแถวก่อนส่งต่อไฟล์ ให้ทุกการเปลี่ยนแปลงมองเห็นได้') => `<div class="title"><div><div class="eyebrow">${eyebrow}</div><h1>${heading}</h1><p>${subtitle}</p></div><div class="actions">${upload()}${exportAll()}</div></div>`
const badge = review => `<span class="badge ${review.changed ? 'updated' : 'unchanged'}">${icon(review.changed ? 'change' : 'check')}${review.changed ? 'มีข้อมูลเปลี่ยน' : 'ข้อมูลไม่เปลี่ยน'}</span>`
const typeBadge = type => `<span class="type ${type}">${labels[type]}</span>`
const counts = review => `<div class="counts"><span><b>${review.counts.updated}</b> แก้ไข</span><span><b>${review.counts.added}</b> เพิ่ม</span><span><b>${review.counts.removed}</b> นำออก</span><span><b>${review.counts.unchanged}</b> ไม่เปลี่ยน</span></div>`
function totals() {
  const values = Object.values(reviews)
  return { files: values.length, changed: values.filter(review => review.changed).length, rows: values.reduce((sum, review) => sum + review.counts.updated + review.counts.added + review.counts.removed, 0), same: values.filter(review => !review.changed).length }
}
function metrics() {
  const t = totals()
  return `<div class="metrics">${[[t.files, 'ไฟล์ที่ตรวจแล้ว', 'file'], [t.changed, 'ไฟล์มีการเปลี่ยนแปลง', 'change'], [t.rows, 'แถวที่ต่างจากเดิม', 'grid'], [t.same, 'ไฟล์ข้อมูลตรงกัน', 'check']].map(([value, label, symbol]) => `<section>${icon(symbol)}<div><strong>${value.toString().padStart(2, '0')}</strong><span>${label}</span></div></section>`).join('')}</div>`
}
function tabs() {
  return `<div class="branch-tabs" role="group" aria-label="เลือกสาขา">${branches.map(branch => `<button data-action="branch" data-branch="${branch}" aria-pressed="${branch === state.branch}"><span class="branch-dot ${reviews[branch].changed ? 'is-changed' : ''}"></span>${branch.toUpperCase()}<small>${reviews[branch].changed ? reviews[branch].counts.updated + reviews[branch].counts.added + reviews[branch].counts.removed : '✓'}</small></button>`).join('')}</div>`
}
function filters() {
  return `<div class="filters"><div class="filter-pills" role="group" aria-label="กรองแถว">${[['all', 'ทุกแถว'], ['changed', 'เฉพาะที่เปลี่ยน'], ['added', 'เพิ่มใหม่'], ['removed', 'นำออก']].map(([value, label]) => `<button data-action="filter" data-filter="${value}" aria-pressed="${state.filter === value}">${label}</button>`).join('')}</div><label class="search">${icon('search')}<input id="row-search" type="search" aria-label="ค้นหา SKU หรือชื่อสินค้า" placeholder="ค้นหา SKU หรือชื่อสินค้า" value="${esc(state.search)}"></label></div>`
}
function visibleRows(review) {
  return review.rows.filter(row => (state.filter === 'all' || (state.filter === 'changed' ? row.type !== 'unchanged' : row.type === state.filter)) && `${row.sku} ${row.name} ${row.itemId}`.toLowerCase().includes(state.search.toLowerCase()))
}
function selectedRow(review) {
  return review.rows.find(row => row.key === state.row) ?? review.rows.find(row => row.type === 'updated') ?? review.rows[0]
}
const value = text => text === undefined ? '<span class="empty-value">ไม่มีแถวนี้</span>' : text === '' ? '<span class="empty-value">(ว่าง)</span>' : esc(text)
function changes(row) {
  if (!row.changes.length) return `<div class="all-same">${icon('check')}ข้อมูลทั้ง 16 คอลัมน์ตรงกัน</div>`
  return `<div class="change-cells">${row.changes.map(change => `<article><label>${String.fromCharCode(65 + change.column)} / ${change.name}</label><div><span class="old-value">${value(change.before)}</span>${icon('arrow')}<strong class="new-value">${value(change.after)}</strong></div></article>`).join('')}</div>`
}
function detail(review) {
  const row = selectedRow(review)
  return `<section class="detail"><div class="panel-heading"><div><div class="eyebrow">ROW INSPECTOR</div><h3>SKU ${esc(row.sku)}</h3><p>${esc(row.name)}</p></div>${typeBadge(row.type)}</div><div class="row-reference">แถวเดิม ${row.before?.row ?? '—'} ${icon('arrow')} แถวใหม่ ${row.after?.row ?? '—'} <span>${row.changes.length} คอลัมน์ต่างกัน</span></div>${changes(row)}</section>`
}
function rowTable(review) {
  const rows = visibleRows(review), selected = selectedRow(review)
  return `<div class="table-scroll"><table class="row-table"><thead><tr><th>สถานะ</th><th>SKU / สินค้า</th><th>แถวเดิม → ใหม่</th><th>สิ่งที่เปลี่ยน</th><th><span class="sr-only">ดูรายละเอียด</span></th></tr></thead><tbody>${rows.length ? rows.map(row => `<tr class="${row.key === selected.key ? 'selected' : ''}" data-type="${row.type}"><td>${typeBadge(row.type)}</td><td><strong>${esc(row.sku)}</strong><small>${esc(row.name)}</small></td><td class="mono">${row.before?.row ?? '—'} → ${row.after?.row ?? '—'}</td><td><span class="column-summary">${row.type === 'unchanged' ? 'ตรงกันทุกคอลัมน์' : row.type === 'added' ? 'รายการสินค้าใหม่' : row.type === 'removed' ? 'ไม่พบในไฟล์ใหม่' : row.changes.map(change => change.name).join(', ')}</span><small>${row.changes.length ? `${row.changes.length} คอลัมน์` : 'ไม่มีการเปลี่ยนแปลง'}</small></td><td><button class="row-open" data-action="row" data-row="${esc(row.key)}" aria-label="ดูรายละเอียด SKU ${esc(row.sku)}">${icon('arrow')}</button></td></tr>`).join('') : '<tr><td colspan="5" class="empty-state">ไม่พบแถวที่ตรงกับตัวกรอง ลองค้นหาหรือเลือกสถานะอื่น</td></tr>'}</tbody></table></div>`
}
function fileList() {
  return branches.map(branch => `<button class="file-link ${state.branch === branch ? 'active' : ''}" data-action="branch" data-branch="${branch}"><span class="file-glyph">${icon('file')}</span><span><strong>${reviews[branch].name}</strong><small>${reviews[branch].total} แถว · ${reviews[branch].changed ? 'มีการเปลี่ยนแปลง' : 'ข้อมูลตรงกัน'}</small></span><b>${reviews[branch].counts.updated + reviews[branch].counts.added + reviews[branch].counts.removed || '✓'}</b></button>`).join('')
}
function fileCard(review) {
  return `<article class="file-card ${review.branch === state.branch ? 'active' : ''}"><div class="card-top"><span class="file-glyph">${icon('file')}</span>${badge(review)}</div><button class="card-select" data-action="branch" data-branch="${review.branch}"><span class="eyebrow">${review.branch.toUpperCase()} BRANCH</span><h2>${review.name}</h2><div class="big-count">${review.counts.updated + review.counts.added + review.counts.removed}<small>แถวที่เปลี่ยน</small></div></button>${counts(review)}<div class="card-bottom">${exportOne(review.branch)}${button('ดูรายละเอียด', 'branch', `data-branch="${review.branch}"`, 'arrow', 'quiet')}</div></article>`
}
function ledger(review) {
  return `${title('01 / EVERY CHANGE, IN ONE PLACE')}${metrics()}<section class="panel"><div class="panel-heading"><h2>ไฟล์ของคุณ</h2><span class="subtle">เทียบกับไฟล์เดิมที่เก็บไว้</span></div><div class="table-scroll"><table class="file-table"><thead><tr><th>ชื่อไฟล์</th><th>ผลการตรวจ</th><th>สรุปแถว</th><th>ส่งออกไฟล์</th></tr></thead><tbody>${branches.map(branch => `<tr class="${branch === state.branch ? 'active' : ''}"><td><button class="text-button" data-action="branch" data-branch="${branch}">${icon('file')}<strong>${reviews[branch].name}</strong></button><small>${reviews[branch].total} แถว · โครงสร้างถูกต้อง</small></td><td>${badge(reviews[branch])}</td><td>${counts(reviews[branch])}</td><td><div class="actions">${exportOne(branch)}${exportReport(branch)}</div></td></tr>`).join('')}</tbody></table></div></section><section class="panel"><div class="panel-heading"><div><div class="eyebrow">CHANGE DETAILS</div><h2>${review.name}</h2></div>${tabs()}</div>${filters()}${rowTable(review)}${detail(review)}</section>`
}
function deck(review) {
  return `${title('02 / THREE BRANCHES. ONE CLEAR PICTURE.', 'ดูการเปลี่ยนแปลง แยกทีละสาขา', 'GRAB CSV MANAGER · เลือกการ์ดเพื่อดูแถวที่เปลี่ยนแปลง')}<div class="branch-deck">${branches.map(branch => fileCard(reviews[branch])).join('')}</div><section class="deck-detail"><div class="panel"><div class="panel-heading"><h2>${review.name}</h2>${exportReport(review.branch)}</div>${filters()}${rowTable(review)}</div>${detail(review)}</section>`
}
function inspector(review) {
  return `<div class="inspector"><aside class="inspector-files"><div class="eyebrow">03 / REVIEW INSPECTOR</div><h1>GRAB CSV<br>MANAGER<span>.</span></h1><p>ทุกไฟล์ ทุกแถว<br>ชัดเจนในพื้นที่เดียว</p><nav>${fileList()}</nav><div class="sidebar-actions">${upload()}${exportAll()}</div></aside><section class="inspector-work"><div class="panel-heading"><div><div class="eyebrow">FILE / ${review.branch.toUpperCase()}</div><h2>${review.name}</h2></div>${badge(review)}</div>${counts(review)}${filters()}${rowTable(review)}<div class="inspector-export">${exportOne(review.branch)}${exportReport(review.branch)}</div></section><aside class="inspector-detail">${detail(review)}<div class="inspector-hint">${icon('change')}เลือกแถวทางซ้ายเพื่ออ่านค่าเดิมและค่าใหม่</div></aside></div>`
}
function journal() {
  return `<div class="journal-intro"><div><div class="eyebrow">04 / THE REVISION JOURNAL</div><h1>ทุกไฟล์มีเรื่องราว.<br>ทุกการเปลี่ยนมีหลักฐาน.</h1><p>GRAB CSV MANAGER / บันทึกผลการเทียบข้อมูลรายแถว</p></div><div class="journal-stamp"><strong>${totals().rows}</strong><span>ROW CHANGES<br>ACROSS THREE FILES</span></div></div><div class="journal-layout"><aside><h2>รอบตรวจสอบนี้</h2><p>เปรียบเทียบไฟล์เดิม<br>กับไฟล์ที่เตรียมใช้งาน</p>${upload()}${exportAll()}<div class="journal-key"><span class="type updated">แก้ไข</span><span class="type added">เพิ่มใหม่</span><span class="type removed">นำออก</span></div></aside><div class="journal-entries">${branches.map((branch, index) => `<article class="journal-entry"><div class="entry-mark">${String(index + 1).padStart(2, '0')}</div><div class="entry-content"><div class="panel-heading"><h2>${reviews[branch].name}</h2>${badge(reviews[branch])}</div>${counts(reviews[branch])}${reviews[branch].changed ? `<div class="journal-changes">${reviews[branch].rows.filter(row => row.type !== 'unchanged').map(row => `<details><summary>${typeBadge(row.type)}<strong>${esc(row.sku)}</strong><span>${esc(row.name)}</span><b>+</b></summary>${changes(row)}</details>`).join('')}</div>` : '<div class="journal-same">✓ ตรวจครบทุกแถวแล้ว ข้อมูลตรงกับไฟล์เดิมทั้งหมด</div>'}<div class="actions">${exportOne(branch)}${exportReport(branch)}</div></div></article>`).join('')}</div></div>`
}
function beforeAfter(review) {
  const row = selectedRow(review), changed = new Set(row.changes.map(change => change.column))
  return `${title('05 / COMPARE WITH CONFIDENCE', 'ไฟล์เดิม. ไฟล์ใหม่. เห็นชัดทั้งสองฝั่ง.', 'GRAB CSV MANAGER · ดูการเปลี่ยนแปลงทีละแถวก่อนส่งออก')}<div class="compare-toolbar">${tabs()}<label>เลือกแถว <select id="row-select" aria-label="เลือก SKU เพื่อเปรียบเทียบ">${review.rows.map(item => `<option value="${esc(item.key)}" ${item.key === row.key ? 'selected' : ''}>${esc(item.sku)} · ${labels[item.type]}</option>`).join('')}</select></label>${typeBadge(row.type)}</div><div class="compare-panes">${['before', 'after'].map(side => `<section class="compare-pane ${side}"><header><span>${side === 'before' ? '01 / BASELINE' : '02 / INCOMING'}</span><h2>${side === 'before' ? 'ไฟล์เดิมที่เก็บไว้' : 'ไฟล์ใหม่ที่อัปโหลด'}</h2><p>${review.name} · แถว ${row[side]?.row ?? '—'}</p></header><div>${columns.map((column, index) => `<article class="${changed.has(index) ? 'cell-changed' : ''}"><label>${String.fromCharCode(65 + index)} / ${column}</label><span>${value(row[side]?.cells[index])}</span>${changed.has(index) ? `<small>${side === 'before' ? 'เดิม' : 'ใหม่'}</small>` : ''}</article>`).join('')}</div></section>`).join('')}</div><div class="compare-bottom"><span>${row.changes.length} คอลัมน์ต่างกันในแถวนี้ · SKU ${esc(row.sku)}</span><div class="actions">${exportReport(review.branch)}${exportOne(review.branch)}</div></div>`
}
function board(review) {
  return `${title('06 / ORGANIZE THE REVIEW', 'จัดไฟล์ให้ชัด ก่อนส่งออก', 'GRAB CSV MANAGER · บอร์ดตรวจสอบตามสถานะของแต่ละไฟล์')}<div class="review-board"><section class="lane changed-lane"><header><span>01 / NEEDS A LOOK</span><h2>ข้อมูลเปลี่ยน <b>${totals().changed}</b></h2></header>${branches.filter(branch => reviews[branch].changed).map(branch => fileCard(reviews[branch])).join('')}</section><section class="lane same-lane"><header><span>02 / ALL CLEAR</span><h2>ข้อมูลตรงกัน <b>${totals().same}</b></h2></header>${branches.filter(branch => !reviews[branch].changed).map(branch => fileCard(reviews[branch])).join('')}<div class="lane-note">${icon('check')}ไฟล์เหล่านี้ไม่มีแถวที่เพิ่ม ลบ หรือแก้ไข</div></section><section class="lane export-lane"><header><span>03 / READY TO EXPORT</span><h2>ชุดส่งออก</h2></header>${selection()}${exportAll()}<p>CSV ใหม่ของสาขาที่เลือก<br>รวมอยู่ในไฟล์ ZIP เดียว</p><div class="board-orbit">${icon('download')}<span>REVIEW<br>THEN DELIVER</span></div></section></div><div class="board-inspection"><div class="panel"><div class="panel-heading"><h2>รายละเอียด / ${review.name}</h2><div class="actions">${tabs()}${exportReport(review.branch)}</div></div>${rowTable(review)}</div>${detail(review)}</div>`
}
function atlas(review) {
  const fields = [2, 3, 4, 5, 7, 15]
  return `${title('07 / MAP THE DIFFERENCE', 'การเปลี่ยนแปลงที่มองเห็นได้.', 'GRAB CSV MANAGER · คลิกเซลล์เพื่อเปิดรายละเอียดของแถวนั้น')}${metrics()}<section class="atlas-layout"><div class="matrix-panel"><div class="panel-heading">${tabs()}<div class="matrix-legend"><span>● เปลี่ยน</span><span>+ เพิ่ม</span><span>− นำออก</span></div></div>${filters()}<div class="table-scroll"><table class="matrix"><thead><tr><th>SKU / แถว</th>${fields.map(column => `<th><b>${String.fromCharCode(65 + column)}</b><small>${columns[column]}</small></th>`).join('')}<th>สถานะ</th></tr></thead><tbody>${visibleRows(review).map(row => `<tr><th>${esc(row.sku)}<small>แถว ${row.after?.row ?? row.before?.row}</small></th>${fields.map(column => { const changed = row.changes.some(change => change.column === column); return `<td><button class="heat ${changed ? row.type : 'same'}" data-action="row" data-row="${esc(row.key)}" aria-label="SKU ${esc(row.sku)} ${columns[column]} ${changed ? labels[row.type] : 'ไม่เปลี่ยน'}" title="${columns[column]}">${changed ? row.type === 'added' ? '+' : row.type === 'removed' ? '−' : '●' : '·'}</button></td>` }).join('')}<td>${typeBadge(row.type)}</td></tr>`).join('')}</tbody></table></div><div class="matrix-footer">${exportOne(review.branch)}${exportReport(review.branch)}</div></div>${detail(review)}</section>`
}
function sheet(review) {
  const fields = [1, 2, 3, 4, 5, 15]
  return `<div class="sheet-title"><div>${icon('grid')}<div><div class="eyebrow">08 / SHEET STUDIO</div><h1>GRAB CSV MANAGER</h1></div></div><div class="actions">${upload()}${exportOne(review.branch)}</div></div><div class="sheet-menu"><span>ตรวจข้อมูล</span><span>เปรียบเทียบแถว</span><span>ส่งออก</span><b>${badge(review)}</b></div><div class="sheet-formula"><strong>${esc(selectedRow(review).sku)}</strong><span>ƒx</span><p>${esc(selectedRow(review).name)}</p>${typeBadge(selectedRow(review).type)}</div>${filters()}<div class="sheet-body table-scroll"><table><thead><tr><th>#</th>${fields.map(column => `<th>${String.fromCharCode(65 + column)}</th>`).join('')}<th>ผลตรวจ</th></tr><tr><th>แถว</th>${fields.map(column => `<th>${columns[column]}</th>`).join('')}<th>STATUS</th></tr></thead><tbody>${visibleRows(review).map(row => `<tr class="${selectedRow(review).key === row.key ? 'sheet-selected' : ''}"><th><button data-action="row" data-row="${esc(row.key)}" aria-label="ตรวจแถว ${row.after?.row ?? row.before?.row}">${row.after?.row ?? '−' + row.before?.row}</button></th>${fields.map(column => `<td class="${row.changes.some(change => change.column === column) ? 'sheet-cell-' + row.type : ''}"><button data-action="row" data-row="${esc(row.key)}">${value((row.after ?? row.before).cells[column])}</button></td>`).join('')}<td>${typeBadge(row.type)}</td></tr>`).join('')}</tbody></table></div><div class="sheet-tabs">${tabs()}<span>${review.total} แถวในไฟล์ใหม่ · 16 คอลัมน์</span>${exportReport(review.branch)}</div><div class="sheet-detail">${detail(review)}</div>`
}
function guided(review) {
  return `<div class="guided-shell"><div class="guided-intro"><div class="eyebrow">09 / ONE STEP AT A TIME</div><h1>ตรวจให้มั่นใจ.<br>ส่งออกได้ตรงไฟล์.</h1><p>GRAB CSV MANAGER</p><div class="guide-index">${[['เลือกไฟล์', 'อัปโหลดไฟล์ของแต่ละสาขา'], ['ตรวจการเปลี่ยนแปลง', 'เปิดดูค่าเดิมและค่าใหม่'], ['ส่งออกไฟล์', 'CSV รายสาขา หรือ ZIP รวม']].map(([label, description], index) => `<button data-action="step" data-step="${index + 1}" aria-current="${state.step === index + 1 ? 'step' : 'false'}"><b>${index + 1}</b><span><strong>${label}</strong><small>${description}</small></span>${icon('arrow')}</button>`).join('')}</div><div class="guided-seal">${icon('check')}ชัดเจนใน 3 ขั้นตอน</div></div><section class="guided-work"><div class="eyebrow">STEP 0${state.step} OF 03</div>${state.step === 1 ? `<h2>เริ่มจากเลือกไฟล์</h2><p>เลือกไฟล์ CSV ชื่อใดก็ได้ ระบบแยกสาขาจาก StoreID และตรวจคอลัมน์ก่อนเทียบกับไฟล์เดิมตัวอย่าง</p><div class="upload-zone">${icon('upload')}<h3>เตรียมไฟล์ใหม่ของคุณ</h3><p>รองรับ UTF-8 / TIS-620</p>${upload()}</div><div class="guide-checklist"><p>✓ โครงสร้าง A–P ต้องตรงกัน</p><p>✓ StoreID ต้องอยู่ในสาขาเดียวกัน</p><p>✓ เทียบแถวด้วย ItemID</p></div>${button('ลองดูผลจากข้อมูลตัวอย่าง', 'step', 'data-step="2"', 'arrow', 'primary')}` : state.step === 2 ? `<div class="panel-heading"><h2>ตรวจแถวที่เปลี่ยน</h2>${badge(review)}</div>${tabs()}${counts(review)}${filters()}${rowTable(review)}${detail(review)}<div class="guide-next">${button('เลือกไฟล์', 'step', 'data-step="1"', 'upload', 'quiet')}${button('ไปส่งออกไฟล์', 'step', 'data-step="3"', 'arrow', 'primary')}</div>` : `<h2>เลือกชุดไฟล์ที่ต้องการ</h2><p>ตรวจเสร็จแล้ว เลือกส่งออก CSV ใหม่เป็นรายไฟล์ หรือรวมไฟล์ทั้งหมดใน ZIP</p>${selection()}<div class="guide-deliver">${exportAll()}${exportOne(review.branch)}${exportReport(review.branch)}</div>${button('กลับไปตรวจแถว', 'step', 'data-step="2"', 'arrow', 'quiet')}`}</section></div>`
}
function selection() {
  return `<div class="export-selection">${branches.map(branch => `<label><input type="checkbox" data-select="${branch}" ${state.selected.has(branch) ? 'checked' : ''}><span>${icon('file')}<strong>${reviews[branch].name}</strong><small>${reviews[branch].total} แถว · ${reviews[branch].changed ? 'มีข้อมูลเปลี่ยน' : 'ข้อมูลตรงกัน'}</small></span><b>${state.selected.has(branch) ? '✓' : ''}</b></label>`).join('')}</div>`
}
function dock(review) {
  return `<div class="dock-top"><div class="eyebrow">10 / FROM REVIEW TO DELIVERY</div><h1>GRAB CSV MANAGER</h1><p>การเปลี่ยนแปลงอยู่ทางซ้าย ไฟล์พร้อมส่งอยู่ทางขวา</p></div><div class="dock-layout"><main class="dock-review"><div class="panel-heading"><h2>พื้นที่ตรวจสอบ</h2>${upload()}</div>${tabs()}<div class="dock-review-title"><h3>${review.name}</h3>${badge(review)}</div>${counts(review)}${filters()}${rowTable(review)}${detail(review)}</main><aside class="export-dock"><div class="dock-icon">${icon('download')}</div><div class="eyebrow">YOUR EXPORT PACKET</div><h2>ชุดไฟล์ส่งออก</h2><p>เลือกสาขาที่ต้องการดาวน์โหลด</p>${selection()}<div class="dock-total"><span>ไฟล์ที่เลือก</span><strong>${state.selected.size}<small>ไฟล์ CSV</small></strong></div>${exportAll()}<div class="dock-divider"><span>หรือส่งออกเฉพาะไฟล์ที่เปิด</span></div>${exportOne(review.branch)}${exportReport(review.branch)}<div class="dock-note">${icon('check')}ชื่อไฟล์ส่งออกใช้ GRAB_XXX.CSV<br>ตาม StoreID ของแต่ละสาขา</div></aside></div>`
}
const views = [ledger, deck, inspector, journal, beforeAfter, board, atlas, sheet, guided, dock]
function render() {
  const focused = document.activeElement?.id, cursor = document.activeElement?.selectionStart
  root.innerHTML = `<header class="preview-nav"><a href="../index.html" class="wordmark">p<span>.</span><small>PRICE CHECKER</small></a><span class="preview-label">ตัวอย่าง ${String(id).padStart(2, '0')} / 10 · ${esc(design.name)} · ข้อมูลจำลอง</span><a href="../index.html">${icon('grid')}ดูทั้ง 10 แบบ</a></header>${state.problems.length ? `<div class="upload-errors" role="alert"><strong>ตรวจไฟล์ไม่ผ่าน</strong><p>ไฟล์ตัวอย่างเดิมยังอยู่ครบ กรุณาตรวจรายละเอียดแล้วลองใหม่</p>${state.problems.map(problem => `<pre>${esc(problem)}</pre>`).join('')}</div>` : ''}<div class="workspace v${String(id).padStart(2, '0')}">${views[id - 1](reviews[state.branch])}</div><footer class="preview-footer"><span>ข้อมูลจำลอง · เปรียบเทียบในหน้านี้ · Export ดาวน์โหลดได้จริง</span><div>${button('โหลดตัวอย่างเดิม', 'reset', '', 'clock', 'quiet')}${button('ดาวน์โหลดไฟล์เดิมเพื่อลองเทียบ', 'baseline', '', 'file', 'quiet')}</div></footer>`
  if (focused === 'row-search') { const input = document.querySelector('#row-search'); input?.focus(); if (input && cursor !== null) input.setSelectionRange(cursor, cursor) }
}
let toastTimer
function notify(message) {
  const toast = document.querySelector('#toast')
  toast.textContent = message; toast.hidden = false
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { toast.hidden = true }, 4000)
}
document.addEventListener('click', event => {
  const control = event.target.closest('[data-action]')
  if (!control) return
  const action = control.dataset.action, branch = control.dataset.branch ?? state.branch
  if (action === 'upload') { document.querySelector('#file-input').click(); return }
  if (action === 'csv' || action === 'report') { exportFile(reviews[branch], action === 'report'); notify(`ส่งออก ${action === 'report' ? 'รายงานการเปลี่ยนแปลง' : reviews[branch].name} แล้ว`); return }
  if (action === 'zip') { if (state.selected.size) { exportZip([...state.selected].map(key => reviews[key])); notify(`ส่งออกชุดไฟล์ ${state.selected.size} สาขาแล้ว`) } return }
  if (action === 'branch') { state.branch = branch; state.row = null; state.filter = 'all'; state.search = '' }
  if (action === 'row') state.row = control.dataset.row
  if (action === 'filter') state.filter = control.dataset.filter
  if (action === 'step') state.step = Number(control.dataset.step)
  if (action === 'reset') { reviews = demo(); state.problems = []; state.branch = 'src'; state.row = null; state.search = ''; state.filter = 'all'; state.step = 2; state.selected = new Set(branches); notify('โหลดข้อมูลตัวอย่างเดิมแล้ว') }
  if (action === 'baseline') { download(reviews[state.branch].baseline, reviews[state.branch].name); notify('ดาวน์โหลดไฟล์เดิมแล้ว ลองอัปไฟล์นี้กลับเพื่อดูสถานะไม่เปลี่ยน') }
  render()
})
document.addEventListener('input', event => {
  if (event.target.id === 'row-search') { state.search = event.target.value; render() }
})
document.addEventListener('change', event => {
  if (event.target.id === 'row-select') { state.row = event.target.value; render() }
  if (event.target.dataset.select) { const branch = event.target.dataset.select; if (event.target.checked) state.selected.add(branch); else state.selected.delete(branch); render() }
})
document.querySelector('#file-input').addEventListener('change', async event => {
  const files = Array.from(event.target.files ?? []); event.target.value = ''
  if (!files.length) return
  state.problems = []
  let loaded = 0
  for (const file of files) {
    try {
      const { text, branch } = await readFile(file)
      reviews[branch] = compare(reviews[branch].baseline, text)
      state.branch = branch; state.row = null; state.filter = 'all'; state.search = ''; state.step = 2; loaded++
    } catch (error) { state.problems.push(`${file.name}: ${error.message}`) }
  }
  render(); notify(`เทียบไฟล์แล้ว ${loaded} ไฟล์${state.problems.length ? ` · ตรวจไม่ผ่าน ${state.problems.length} ไฟล์` : ''}`)
})
render()
