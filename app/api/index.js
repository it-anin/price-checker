import { useState, useEffect } from 'react'

export default function Home() {
  const [products, setProducts] = useState([])
  const [stats, setStats] = useState({ totalProducts: 0, sheets: [] })
  const [search, setSearch] = useState('')
  const [selectedSheet, setSelectedSheet] = useState('all')
  const [status, setStatus] = useState('พร้อมใช้งาน')
  const [loading, setLoading] = useState(false)

  // โหลด stats ตอนเปิดหน้า
  useEffect(() => {
    loadStats()
  }, [])

  async function loadStats() {
    const res = await fetch('/api/stats')
    const data = await res.json()
    if (data.success) setStats(data)
  }

  async function loadProducts(sheet = 'all') {
    setLoading(true)
    setStatus('กำลังโหลด...')
    const url = sheet === 'all' ? '/api/products' : `/api/products?sheet=${encodeURIComponent(sheet)}`
    const res = await fetch(url)
    const data = await res.json()
    if (data.success) {
      setProducts(data.products)
      setStatus(`พร้อมใช้งาน (${data.products.length} รายการ)`)
    }
    setLoading(false)
  }

  async function searchProducts(sku) {
    if (sku.length < 3) return
    setLoading(true)
    const res = await fetch(`/api/products?sku=${encodeURIComponent(sku)}`)
    const data = await res.json()
    if (data.success) setProducts(data.products)
    setLoading(false)
  }

  function handleSheetClick(name) {
    setSelectedSheet(name)
    loadProducts(name)
  }

  return (
    <div style={{ fontFamily: 'sans-serif', height: '100vh', display: 'flex', flexDirection: 'column' }}>

      {/* Header */}
      <div style={{ background: '#f5f7fa', padding: '12px 20px', borderBottom: '1px solid #ddd', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <h1 style={{ fontSize: 18, color: '#333' }}>📦 ระบบตรวจสอบราคาสินค้า</h1>
      </div>

      {/* Toolbar */}
      <div style={{ background: '#f8f8f8', borderBottom: '1px solid #bbb', padding: '10px 20px', display: 'flex', gap: 12, alignItems: 'center' }}>
        <input
          type="text"
          placeholder="ค้นหารหัสสินค้า..."
          value={search}
          onChange={e => { setSearch(e.target.value); searchProducts(e.target.value) }}
          style={{ padding: '6px 10px', border: '1px solid #999', borderRadius: 3, fontSize: 13, width: 220 }}
        />
        <button onClick={() => loadProducts(selectedSheet)} style={btnStyle}>🔄 รีเฟรช</button>
        <button onClick={checkPrices} style={btnStyle}>💰 ตรวจสอบราคา</button>
      </div>

      {/* Main Layout */}
      <div style={{ display: 'grid', gridTemplateColumns: '260px 1fr', flex: 1, overflow: 'hidden' }}>

        {/* Left Panel */}
        <div style={{ background: '#f0f0f0', borderRight: '1px solid #bbb', overflowY: 'auto' }}>
          <div style={{ padding: '6px 10px', background: '#e0e0e0', borderBottom: '1px solid #ccc', fontWeight: 700, fontSize: 13 }}>📊 สถิติ</div>
          <div style={{ padding: 15 }}>
            <div style={statCard}>
              <div style={{ fontSize: 11, color: '#666' }}>สินค้าทั้งหมด</div>
              <div style={{ fontSize: 24, fontWeight: 700, color: '#c74634' }}>{stats.totalProducts.toLocaleString()}</div>
            </div>
            <div style={statCard}>
              <div style={{ fontSize: 11, color: '#666' }}>หมวดหมู่</div>
              <div style={{ fontSize: 24, fontWeight: 700, color: '#c74634' }}>{stats.sheets.length}</div>
            </div>
          </div>

          <div style={{ padding: '6px 10px', background: '#e0e0e0', borderBottom: '1px solid #ccc', fontWeight: 700, fontSize: 13 }}>📁 หมวดหมู่</div>
          {stats.sheets.map(s => (
            <div
              key={s.name}
              onClick={() => handleSheetClick(s.name)}
              style={{
                padding: '8px 12px', borderBottom: '1px solid #ddd', fontSize: 12,
                cursor: 'pointer', display: 'flex', justifyContent: 'space-between',
                background: selectedSheet === s.name ? '#cce5ff' : 'transparent'
              }}
            >
              <span>{s.name}</span>
              <span style={{ background: '#4a8bc4', color: '#fff', padding: '2px 8px', borderRadius: 10, fontSize: 10 }}>{s.count}</span>
            </div>
          ))}
        </div>

        {/* Right Panel */}
        <div style={{ overflowY: 'auto', background: '#fff' }}>
          {loading ? (
            <div style={{ textAlign: 'center', padding: 40, color: '#666' }}>กำลังโหลด...</div>
          ) : products.length === 0 ? (
            <div style={{ textAlign: 'center', padding: 40, color: '#888' }}>เลือกหมวดหมู่หรือค้นหาสินค้า</div>
          ) : (
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr style={{ background: '#f5f5f5' }}>
                  <th style={th}>#</th>
                  <th style={th}>หมวดหมู่</th>
                  <th style={th}>รหัสสินค้า</th>
                  <th style={th}>ชื่อสินค้า</th>
                  <th style={th}>ใบอนุญาต</th>
                  <th style={th}>ราคา</th>
                </tr>
              </thead>
              <tbody>
                {products.map((p, i) => (
                  <tr key={i} style={{ borderBottom: '1px solid #eee' }}
                    onMouseEnter={e => e.currentTarget.style.background = '#e8f4fc'}
                    onMouseLeave={e => e.currentTarget.style.background = 'transparent'}
                  >
                    <td style={td}>{i + 1}</td>
                    <td style={td}>{p['หมวดหมู่สินค้า (CATEGORIES)'] || '-'}</td>
                    <td style={td}><strong>{p['รหัสสินค้า (SKU NUMBER)'] || '-'}</strong></td>
                    <td style={td}>{p['*ชื่อสินค้า (NAME)'] || '-'}</td>
                    <td style={td}>{p['*เลขที่ใบอนุญาตโฆษณา'] || '-'}</td>
                    <td style={{ ...td, color: '#c74634', fontWeight: 600 }}>{p['*ราคาสินค้า'] || '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>

      {/* Status Bar */}
      <div style={{ background: '#f0f0f0', borderTop: '1px solid #bbb', padding: '4px 15px', fontSize: 11, color: '#666' }}>
        {status}
      </div>
    </div>
  )

  async function checkPrices() {
    setLoading(true)
    setStatus('กำลังตรวจสอบราคา...')
    const res = await fetch('/api/prices')
    const data = await res.json()
    if (data.success) {
      setProducts(data.mismatched)
      setStatus(`พบ ${data.mismatchCount} รายการราคาไม่ตรง`)
    }
    setLoading(false)
  }
}

// Styles
const btnStyle = {
  padding: '7px 16px', border: '1px solid #777', borderRadius: 20,
  fontSize: 12, fontWeight: 600, cursor: 'pointer', background: '#f5f5f5'
}
const statCard = {
  background: '#fff', border: '1px solid #ccc', borderRadius: 4,
  padding: 15, marginBottom: 10
}
const th = {
  padding: '8px 12px', textAlign: 'left', border: '1px solid #ccc',
  fontWeight: 600, position: 'sticky', top: 0, background: '#f5f5f5'
}
const td = { padding: '8px 12px', border: '1px solid #eee', color: '#444' }