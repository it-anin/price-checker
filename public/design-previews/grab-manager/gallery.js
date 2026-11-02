const designs = await fetch('concepts.json').then(response => response.json())
const dialog = document.querySelector('#preview'), frame = document.querySelector('#frame')
let current = 1, opener, timer
function show(id) {
  current = id < 1 ? 10 : id > 10 ? 1 : id
  const design = designs.find(item => item.id === current), number = String(current).padStart(2, '0')
  document.querySelector('#preview-number').textContent = `แบบ ${number} / ${design.thai}`
  document.querySelector('#preview-title').textContent = design.name
  document.querySelector('#preview-description').textContent = design.description
  document.querySelector('#counter').textContent = `${number} / 10`
  const url = `concepts/${number}.html`
  frame.src = url; frame.title = `ตัวอย่าง ${number} ${design.name}`
  document.querySelector('#full-view').href = url
  if (!dialog.open) { opener = document.activeElement; dialog.showModal(); document.body.style.overflow = 'hidden' }
}
function close() { dialog.close() }
dialog.addEventListener('close', () => { frame.removeAttribute('src'); document.body.style.overflow = ''; opener?.focus() })
document.addEventListener('click', event => { const button = event.target.closest('[data-open]'); if (button) show(Number(button.dataset.open)) })
document.querySelector('#close').addEventListener('click', close)
document.querySelector('#next').addEventListener('click', () => show(current + 1))
document.querySelector('#previous').addEventListener('click', () => show(current - 1))
document.addEventListener('keydown', event => { if (dialog.open && !['INPUT', 'SELECT'].includes(document.activeElement.tagName)) { if (event.key === 'ArrowRight') show(current + 1); if (event.key === 'ArrowLeft') show(current - 1) } })
for (const mode of ['desktop', 'mobile']) document.querySelector('#' + mode).addEventListener('click', () => {
  document.querySelector('.preview-stage').classList.toggle('mobile', mode === 'mobile')
  for (const other of ['desktop', 'mobile']) document.querySelector('#' + other).setAttribute('aria-pressed', String(other === mode))
})
function notify(message) {
  const toast = document.querySelector('#gallery-toast'); toast.textContent = message; toast.hidden = false
  clearTimeout(timer); timer = setTimeout(() => { toast.hidden = true }, 4000)
}
document.querySelector('#choose').addEventListener('click', async () => {
  const design = designs.find(item => item.id === current)
  const text = `GRAB CSV MANAGER แบบ ${String(current).padStart(2, '0')} — ${design.name}`
  try { await navigator.clipboard.writeText(text); notify('คัดลอกแล้ว: ' + text) } catch { notify(text) }
})
const search = document.querySelector('#design-search')
function filter() {
  const query = search.value.trim().toLowerCase(); let count = 0
  for (const card of document.querySelectorAll('.design-card')) { card.hidden = !card.dataset.search.includes(query); if (!card.hidden) count++ }
  document.querySelector('#result-count').textContent = `${count} แบบ`
  document.querySelector('#no-results').hidden = count > 0
}
search.addEventListener('input', filter)
document.querySelector('#reset-search').addEventListener('click', () => { search.value = ''; filter(); search.focus() })
search.addEventListener('keydown', event => { if (event.key === 'Escape') { search.value = ''; filter() } })
