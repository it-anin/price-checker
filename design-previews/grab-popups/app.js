import { demo, compare, readFile, branches, labels, escapeHtml as esc, exportFile } from './review.js'

const designs = await fetch('../concepts.json').then(response => response.json())
const id = Number(document.body.dataset.design), design = designs.find(item => item.id === id)
const samples = demo()
const state = { uploaded: [], branch: 'src', row: null, search: '', errors: [], notice: '', source: '', busy: false }
const dialog = document.querySelector('#change-popup'), popup = document.querySelector('#popup-content')
let opener, toastTimer
const icons = {
  alert: '<path d="m12 3 10 18H2L12 3Zm0 5v6m0 3v1"/>',
  file: '<path d="M14 2H4v20h16V8l-6-6Zm0 0v6h6M8 13h8M8 17h6"/>',
  arrow: '<path d="M4 12h16m-6-6 6 6-6 6"/>',
  check: '<path d="m5 12 4 4L19 6"/>',
  upload: '<path d="M12 16V3m-5 5 5-5 5 5M4 16v5h16v-5"/>',
  download: '<path d="M12 3v13m-5-5 5 5 5-5M4 17v4h16v-4"/>',
  close: '<path d="m6 6 12 12M6 18 18 6"/>',
  search: '<circle cx="10" cy="10" r="6"/><path d="m15 15 6 6"/>',
}
const icon = name => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[name] ?? icons.file}</svg>`
const closeButton = () => `<button class="close" data-action="close" aria-label="ปิดป๊อปอัป">${icon('close')}</button>`
const changed = () => state.uploaded.filter(review => review.changed)
const current = () => state.uploaded.find(review => review.branch === state.branch) ?? changed()[0]
const changedRows = review => review.rows.filter(row => row.type !== 'unchanged')
const count = review => review.counts.updated + review.counts.added + review.counts.removed
function totals() {
  return state.uploaded.reduce((sum, review) => ({ updated: sum.updated + review.counts.updated, added: sum.added + review.counts.added, removed: sum.removed + review.counts.removed }), { updated: 0, added: 0, removed: 0 })
}
function selectedRow() {
  return changedRows(current()).find(row => row.key === state.row) ?? changedRows(current())[0]
}
const badge = type => `<span class="badge ${type}">${labels[type]}</span>`
const showValue = text => text === undefined ? '<span class="empty-value">ไม่มีแถวนี้</span>' : text === '' ? '<span class="empty-value">(ว่าง)</span>' : esc(text)
const exportButton = (report = false) => `<button class="button export" data-action="${report ? 'report' : 'export'}">${icon(report ? 'file' : 'download')}${report ? 'Export รายงาน' : 'Export CSV'}</button>`
const acknowledge = () => '<button class="button primary" data-action="close" autofocus>รับทราบ</button>'
const fileTag = review => `<div class="file-tag">${icon('file')}<strong>${review.name}</strong><span>${review.total} แถว</span></div>`
function stats(review) {
  const counts = review?.counts ?? totals()
  return `<div class="stats">${['updated', 'added', 'removed'].map(type => `<div class="${type}"><strong>${counts[type]}</strong><span>${labels[type]}</span></div>`).join('')}</div>`
}
function header(title = 'อัปโหลดแล้ว พบข้อมูลเปลี่ยนแปลง') {
  return `<header class="popup-header"><span class="alert-icon">${icon('alert')}</span><div><div class="eyebrow">UPLOAD COMPLETE / DATA CHANGED</div><h2 id="popup-title">${title}</h2><p id="popup-description">${changed().length} ไฟล์มีข้อมูลต่างจากไฟล์เดิมที่เก็บไว้</p></div>${closeButton()}</header>`
}
function footer() {
  return `<footer class="popup-footer"><div class="export-actions">${exportButton()}${exportButton(true)}</div>${acknowledge()}</footer>`
}
function fileTabs() {
  const files = id === 9 ? state.uploaded : changed()
  if (files.length < 2) return ''
  return `<div class="file-tabs" role="group" aria-label="ผลตรวจแต่ละสาขา">${files.map(review => `<button data-action="branch" data-branch="${review.branch}" aria-pressed="${state.branch === review.branch}"><strong>${review.branch.toUpperCase()}</strong><small>${review.changed ? count(review) + ' แถวเปลี่ยน' : 'ไม่เปลี่ยน'}</small></button>`).join('')}</div>`
}
function cellChanges(row, limit = Infinity) {
  if (!row) return '<div class="no-changes">✓ ข้อมูลทุกแถวตรงกับไฟล์เดิม</div>'
  return `<div class="cell-changes">${row.changes.slice(0, limit).map(change => `<div><label>${change.name}</label><p><span class="before-value">${showValue(change.before)}</span>${icon('arrow')}<strong class="after-value">${showValue(change.after)}</strong></p></div>`).join('')}${row.changes.length > limit ? `<small>และอีก ${row.changes.length - limit} คอลัมน์ · ดูครบได้ในรายงาน Export</small>` : ''}</div>`
}
function rowCards(review, expanded = false) {
  return `<div class="row-cards">${changedRows(review).map((row, index) => `<details ${expanded && index === 0 ? 'open' : ''}><summary>${badge(row.type)}<div><strong>SKU ${esc(row.sku)}</strong><span>${esc(row.name)}</span></div><small>${row.changes.length} คอลัมน์</small><b>+</b></summary>${cellChanges(row)}</details>`).join('') || '<div class="no-changes">✓ ไม่พบข้อมูลเปลี่ยนแปลงในไฟล์นี้</div>'}</div>`
}
function diffTable(review) {
  const rows = changedRows(review).flatMap(row => row.changes.map(change => ({ row, change })))
  return `<div class="diff-scroll"><table><thead><tr><th>SKU / สถานะ</th><th>คอลัมน์</th><th>ค่าเดิม</th><th>ค่าใหม่</th></tr></thead><tbody>${rows.map(({ row, change }) => `<tr><td><strong>${esc(row.sku)}</strong>${badge(row.type)}</td><td>${change.name}</td><td class="before-value">${showValue(change.before)}</td><td class="after-value">${showValue(change.after)}</td></tr>`).join('') || '<tr><td colspan="4" class="no-changes">ข้อมูลตรงกับไฟล์เดิมทุกแถว</td></tr>'}</tbody></table></div>`
}
function softSummary(review) {
  const row = selectedRow()
  return `<div class="summary-top">${closeButton()}<span class="summary-icon">${icon('alert')}</span><div class="eyebrow">UPLOAD COMPLETE</div><h2 id="popup-title">อัปโหลดแล้ว<br>มีข้อมูลเปลี่ยนแปลง</h2><p id="popup-description">พบ ${count(review)} แถวที่ต่างจากไฟล์เดิม</p></div><div class="popup-body">${fileTag(review)}${fileTabs()}${stats(review)}<div class="sample-change"><div><strong>ตัวอย่างที่เปลี่ยน</strong>${badge(row.type)}</div><p>SKU ${esc(row.sku)} · ${esc(row.name)}</p>${cellChanges(row, 2)}</div><details class="more-details"><summary>ดูครบทั้ง ${count(review)} แถวที่เปลี่ยน</summary>${rowCards(review)}</details></div>${footer()}`
}
function tablePopup(review) {
  return `${header()}<div class="popup-body"><div class="body-heading">${fileTag(review)}<span class="result-count">${count(review)} แถวเปลี่ยน</span></div>${fileTabs()}${stats(review)}${diffTable(review)}</div>${footer()}`
}
function beforeAfter(review) {
  const rows = changedRows(review), row = selectedRow()
  const fields = row.changes.slice(0, 8)
  return `${header('ข้อมูลใหม่ต่างจากเดิมตรงไหนบ้าง?')}<div class="popup-body">${fileTag(review)}${fileTabs()}<label class="row-picker">เลือกแถวที่เปลี่ยน<select id="row-picker">${rows.map(item => `<option value="${esc(item.key)}" ${item.key === row.key ? 'selected' : ''}>SKU ${esc(item.sku)} · ${labels[item.type]}</option>`).join('')}</select>${badge(row.type)}</label><div class="compare-panes">${['before', 'after'].map(side => `<section class="${side}"><header><span>${side === 'before' ? 'BEFORE' : 'AFTER'}</span><h3>${side === 'before' ? 'ข้อมูลเดิม' : 'ข้อมูลหลังอัปโหลด'}</h3></header>${fields.map(change => `<article><label>${change.name}</label><strong>${showValue(change[side])}</strong></article>`).join('')}</section>`).join('')}</div><div class="row-context">${esc(row.name)} · แถวเดิม ${row.before?.row ?? '—'} → แถวใหม่ ${row.after?.row ?? '—'}</div>${rows.length > 1 ? `<button class="text-button" data-action="next-row">ดูแถวถัดไป ${icon('arrow')}</button>` : ''}</div>${footer()}`
}
function sideNotice(review) {
  return `${header('แจ้งการเปลี่ยนแปลง')}<div class="side-success">${icon('check')}ตรวจไฟล์ที่อัปโหลดเรียบร้อยแล้ว</div><div class="popup-body">${fileTag(review)}${fileTabs()}<div class="drawer-total"><strong>${count(review)}</strong><span>แถวมีข้อมูลเปลี่ยนแปลง</span></div>${stats(review)}<h3 class="section-label">รายการที่เปลี่ยน</h3>${rowCards(review, true)}</div>${footer()}`
}
function bottomReview(review) {
  return `<div class="sheet-handle"></div>${header('ไฟล์ที่อัปโหลดมีข้อมูลเปลี่ยน')}<div class="popup-body"><div class="sheet-summary">${fileTag(review)}${stats(review)}</div>${fileTabs()}<div class="horizontal-cards">${changedRows(review).map(row => `<article>${badge(row.type)}<h3>SKU ${esc(row.sku)}</h3><p>${esc(row.name)}</p>${cellChanges(row, 3)}</article>`).join('')}</div></div>${footer()}`
}
function receipt(review) {
  return `<div class="receipt-top">${closeButton()}<span class="receipt-icon">${icon('check')}</span><div class="eyebrow">FILE UPLOAD RECEIPT</div><h2 id="popup-title">รับไฟล์แล้ว</h2><p id="popup-description">ตรวจพบการเปลี่ยนแปลงในข้อมูล</p></div><div class="popup-body">${fileTag(review)}${fileTabs()}<div class="receipt-lines">${['updated', 'added', 'removed'].map(type => `<div><span>${labels[type]}</span><strong>${review.counts[type]}<small>แถว</small></strong></div>`).join('')}<div class="receipt-total"><span>รวมที่เปลี่ยน</span><strong>${count(review)}<small>แถว</small></strong></div></div><div class="receipt-example"><span>ตัวอย่าง / SKU ${esc(selectedRow().sku)}</span>${cellChanges(selectedRow(), 1)}</div><details class="more-details"><summary>เปิดรายการเปลี่ยนแปลงทั้งหมด</summary>${rowCards(review)}</details><p class="receipt-note">เทียบกับไฟล์เดิมก่อนอัปโหลด<br>Export รายงานเพื่อเก็บรายละเอียดได้</p></div>${footer()}`
}
function impactSplit(review) {
  return `<div class="impact-layout"><aside><div class="eyebrow">UPLOAD COMPLETE</div><span class="impact-icon">${icon('alert')}</span><strong>${count(review)}</strong><p>แถวเปลี่ยนแปลง<br>หลังอัปโหลดไฟล์</p><div class="impact-file">${review.name}</div>${stats(review)}</aside><section><header class="impact-heading"><div><h2 id="popup-title">ข้อมูลในไฟล์มีการเปลี่ยนแปลง</h2><p id="popup-description">ตรวจสอบค่าเดิมกับค่าใหม่ได้ที่นี่</p></div>${closeButton()}</header><div class="popup-body">${fileTabs()}${rowCards(review, true)}</div>${footer()}</section></div>`
}
function timeline(review) {
  return `${header('สรุปสิ่งที่เปลี่ยนหลังอัปโหลด')}<div class="popup-body">${fileTag(review)}${fileTabs()}<div class="change-timeline">${['updated', 'added', 'removed'].map(type => `<section class="${type}"><span class="timeline-dot">${type === 'updated' ? '↻' : type === 'added' ? '+' : '−'}</span><header><h3>${labels[type]}</h3><b>${review.counts[type]} แถว</b></header>${review.rows.filter(row => row.type === type).map(row => `<details><summary><strong>${esc(row.sku)}</strong><span>${esc(row.name)}</span><b>+</b></summary>${cellChanges(row)}</details>`).join('') || '<p>ไม่มีรายการ</p>'}</section>`).join('')}</div></div>${footer()}`
}
function branchTabs(review) {
  const all = totals(), same = state.uploaded.filter(item => !item.changed).length
  return `${header('อัปโหลดเสร็จแล้ว สรุปผลแต่ละไฟล์')}<div class="batch-summary"><div><strong>${state.uploaded.length}</strong><span>ไฟล์ที่อัปโหลด</span></div><div><strong>${changed().length}</strong><span>ไฟล์ข้อมูลเปลี่ยน</span></div><div><strong>${all.updated + all.added + all.removed}</strong><span>แถวที่เปลี่ยนทั้งหมด</span></div><div><strong>${same}</strong><span>ไฟล์ข้อมูลตรงกัน</span></div></div><div class="popup-body">${fileTabs()}<div class="body-heading">${fileTag(review)}<span class="result-count ${review.changed ? '' : 'same'}">${review.changed ? count(review) + ' แถวเปลี่ยน' : 'ข้อมูลไม่เปลี่ยน'}</span></div>${rowCards(review, true)}</div>${footer()}`
}
function searchable(review) {
  const rows = changedRows(review).filter(row => `${row.sku} ${row.name}`.toLowerCase().includes(state.search.toLowerCase()))
  const row = rows.find(item => item.key === state.row) ?? rows[0]
  return `${header('ค้นหาข้อมูลที่เปลี่ยนหลังอัปโหลด')}<div class="popup-body">${fileTag(review)}${fileTabs()}<label class="search-box">${icon('search')}<input id="popup-search" type="search" placeholder="ค้นหา SKU หรือชื่อสินค้าในแถวที่เปลี่ยน" aria-label="ค้นหาแถวที่เปลี่ยน" value="${esc(state.search)}"><span>${rows.length} แถว</span></label><div class="search-review"><div class="row-list">${rows.map(item => `<button data-action="row" data-row="${esc(item.key)}" aria-pressed="${item.key === row?.key}"><span><strong>${esc(item.sku)}</strong>${badge(item.type)}</span><small>${esc(item.name)}</small></button>`).join('') || '<p class="no-changes">ไม่พบแถวที่ตรงกับคำค้นหา</p>'}</div><section class="row-detail">${row ? `<div class="row-detail-heading"><span class="eyebrow">SELECTED ROW</span><h3>SKU ${esc(row.sku)}</h3><p>${esc(row.name)}</p></div>${cellChanges(row)}` : '<div class="no-changes">ลองค้นหาด้วย SKU หรือชื่ออื่น</div>'}</section></div></div>${footer()}`
}
const views = [softSummary, tablePopup, beforeAfter, sideNotice, bottomReview, receipt, impactSplit, timeline, branchTabs, searchable]
function renderPopup() {
  const focused = document.activeElement?.id, cursor = document.activeElement?.selectionStart
  popup.innerHTML = views[id - 1](current())
  if (focused === 'popup-search') { const input = document.querySelector('#popup-search'); input?.focus(); if (input && cursor !== null) input.setSelectionRange(cursor, cursor) }
}
function openPopup() {
  if (!changed().length) return
  if (!current()?.changed && id !== 9) state.branch = changed()[0].branch
  renderPopup()
  if (!dialog.open) { opener = document.activeElement; dialog.showModal(); document.body.classList.add('modal-open') }
}
function renderBackground() {
  const table = document.querySelector('#files')
  table.innerHTML = branches.map(branch => {
    const uploaded = state.uploaded.find(review => review.branch === branch)
    return `<tr><td><span class="file-mark">${icon('file')}</span><div><strong>GRAB_${branch.toUpperCase()}.CSV</strong><small>สาขา ${branch.toUpperCase()}</small></div></td><td><span class="file-status ${uploaded ? 'just-uploaded' : ''}">${uploaded ? 'เพิ่งอัปโหลด' : 'ไฟล์เดิม'}</span></td><td>${uploaded?.total ?? 6} แถว</td></tr>`
  }).join('')
  const notice = document.querySelector('#notice')
  notice.hidden = !state.notice; notice.textContent = state.notice
  const errors = document.querySelector('#errors')
  errors.hidden = !state.errors.length; errors.innerHTML = state.errors.map(error => `<pre>${esc(error)}</pre>`).join('')
}
function afterUpload(reviews, source) {
  state.uploaded = reviews; state.source = source; state.row = null; state.search = ''
  state.branch = reviews.find(review => review.changed)?.branch ?? reviews[0]?.branch ?? 'src'
  state.notice = changed().length ? `อัปโหลด ${reviews.length} ไฟล์แล้ว · พบข้อมูลเปลี่ยนใน ${changed().length} ไฟล์` : reviews.length ? `อัปโหลด ${reviews.length} ไฟล์แล้ว · ข้อมูลตรงกับไฟล์เดิมทุกแถว` : ''
  renderBackground()
  if (changed().length) openPopup()
  else if (dialog.open) dialog.close()
}
function simulate(mode = 'changed') {
  state.errors = []
  const data = demo()
  const result = mode === 'batch' ? Object.values(data) : mode === 'same' ? [compare(data.src.baseline, data.src.baseline)] : [data.src]
  afterUpload(result, 'ข้อมูลตัวอย่าง')
}
function notify(message) {
  const toast = document.querySelector('#toast'); toast.textContent = message; toast.hidden = false
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { toast.hidden = true }, 3500)
}
document.addEventListener('click', event => {
  const control = event.target.closest('[data-action]')
  if (!control || state.busy) return
  const action = control.dataset.action
  if (action === 'close') { dialog.close(); return }
  if (action === 'simulate') { simulate(control.dataset.mode); return }
  if (action === 'upload') { document.querySelector('#file-input').click(); return }
  if (action === 'replay') { openPopup(); return }
  if (action === 'export' || action === 'report') { exportFile(current(), action === 'report'); notify(`ส่งออก${action === 'report' ? 'รายงานการเปลี่ยนแปลง' : 'ไฟล์ CSV'}แล้ว`); return }
  if (action === 'branch') { state.branch = control.dataset.branch; state.row = null; renderPopup() }
  if (action === 'row') { state.row = control.dataset.row; renderPopup() }
  if (action === 'next-row') {
    const rows = changedRows(current()), index = rows.findIndex(row => row.key === selectedRow().key)
    state.row = rows[(index + 1) % rows.length].key; renderPopup()
  }
})
document.addEventListener('change', event => { if (event.target.id === 'row-picker') { state.row = event.target.value; renderPopup() } })
document.addEventListener('input', event => { if (event.target.id === 'popup-search') { state.search = event.target.value; renderPopup() } })
dialog.addEventListener('close', () => { document.body.classList.remove('modal-open'); opener?.focus() })
document.querySelector('#file-input').addEventListener('change', async event => {
  const files = Array.from(event.target.files ?? []); event.target.value = ''
  if (!files.length || state.busy) return
  state.busy = true; state.errors = []
  const next = new Map()
  for (const file of files) {
    try { const { text, branch } = await readFile(file); next.set(branch, compare(samples[branch].baseline, text)) }
    catch (error) { state.errors.push(`${file.name}: ${error.message}`) }
  }
  state.busy = false; afterUpload([...next.values()], 'ไฟล์ที่เลือก')
})
document.querySelector('#design-number').textContent = `แบบ ${String(id).padStart(2, '0')} / ${design.name}`
document.querySelector('#design-description').textContent = design.description
renderBackground()
if (new URLSearchParams(location.search).has('demo')) simulate(id === 9 ? 'batch' : 'changed')
