'use client'

import { useState, useEffect } from 'react'

interface Product {
  [key: string]: any
}

interface SheetStat {
  name: string
  count: number
}

interface Stats {
  totalProducts: number
  sheets: SheetStat[]
}

export default function Home() {
  const [products, setProducts] = useState<Product[]>([])
  const [stats, setStats] = useState<Stats>({ totalProducts: 0, sheets: [] })
  const [search, setSearch] = useState('')
  const [selectedSheet, setSelectedSheet] = useState('all')
  const [status, setStatus] = useState('พร้อมใช้งาน')
  const [loading, setLoading] = useState(false)
  const [tooltip, setTooltip] = useState<{ x: number; y: number; url: string } | null>(null)
  const [editProduct, setEditProduct] = useState<Product | null>(null)

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
      preloadImages(data.products)
      setStatus(`พร้อมใช้งาน (${data.products.length} รายการ)`)
    }
    setLoading(false)
  }

  async function searchProducts(sku: string) {
    if (sku.length < 3) return
    setLoading(true)
    const res = await fetch(`/api/products?sku=${encodeURIComponent(sku)}`)
    const data = await res.json()
    if (data.success) {
      setProducts(data.products)
      preloadImages(data.products)
    }
    setLoading(false)
  }

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

  async function saveProduct() {
    if (!editProduct) return
    const res = await fetch('/api/products', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sku: editProduct['_originalSku'],
        '*ชื่อสินค้า (NAME)': editProduct['*ชื่อสินค้า (NAME)'],
        '*เลขที่ใบอนุญาตโฆษณา': editProduct['*เลขที่ใบอนุญาตโฆษณา'],
        '*ราคาสินค้า': editProduct['*ราคาสินค้า'],
        'หมวดหมู่สินค้า (CATEGORIES)': editProduct['หมวดหมู่สินค้า (CATEGORIES)'],
        'รหัสสินค้า (SKU NUMBER)': editProduct['รหัสสินค้า (SKU NUMBER)']
      })
    })
    const data = await res.json()
    if (data.success) {
      setProducts(prev => prev.map(p =>
        p['รหัสสินค้า (SKU NUMBER)'] === editProduct['_originalSku']
          ? { ...editProduct } : p
      ))
      setEditProduct(null)
      setStatus('บันทึกสำเร็จ ✅')
    } else {
      setStatus('เกิดข้อผิดพลาด: ' + data.error)
    }
  }

  function convertDriveLink(url: string): string {
    if (!url) return ''
    const match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
      url.match(/[?&]id=([a-zA-Z0-9_-]+)/)
    if (match) return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w300`
    return url.startsWith('http') ? url : ''
  }

  function preloadImages(productList: Product[]) {
    productList.forEach(p => {
      const url = convertDriveLink(p['*รูปภาพสินค้า'] || '')
      if (url) {
        const img = new Image()
        img.src = url
      }
    })
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
              onClick={() => { setSelectedSheet(s.name); loadProducts(s.name) }}
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
                  <th style={th}>ดำเนินการ</th>
                </tr>
              </thead>
              <tbody>
                {products.map((p, i) => (
                  <tr key={i}
                    onMouseEnter={e => {
                      e.currentTarget.style.background = '#e8f4fc'
                      const imgUrl = convertDriveLink(p['*รูปภาพสินค้า'] || '')
                      setTooltip({ x: e.clientX, y: e.clientY, url: imgUrl })
                    }}
                    onMouseMove={e => {
                      setTooltip(prev => prev ? { ...prev, x: e.clientX, y: e.clientY } : null)
                    }}
                    onMouseLeave={e => {
                      e.currentTarget.style.background = 'transparent'
                      setTooltip(null)
                    }}
                  >
                    <td style={td}>{i + 1}</td>
                    <td style={td}>{p['หมวดหมู่สินค้า (CATEGORIES)'] || '-'}</td>
                    <td style={td}><strong>{p['รหัสสินค้า (SKU NUMBER)'] || '-'}</strong></td>
                    <td style={td}>{p['*ชื่อสินค้า (NAME)'] || '-'}</td>
                    <td style={td}>{p['*เลขที่ใบอนุญาตโฆษณา'] || '-'}</td>
                    <td style={{ ...td, color: '#c74634', fontWeight: 600 }}>{p['*ราคาสินค้า'] || '-'}</td>
                    <td style={td}>
                      <button
                        onClick={() => setEditProduct({ ...p, _originalSku: p['รหัสสินค้า (SKU NUMBER)'] })}
                        style={{ ...btnStyle, fontSize: 11, padding: '4px 10px' }}
                      >
                        ✏️ แก้ไข
                      </button>
                    </td>
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

      {/* Edit Modal */}
      {editProduct && (
        <div style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)',
          display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000
        }}>
          <div style={{
            background: '#fff', borderRadius: 6, width: 480,
            boxShadow: '0 10px 40px rgba(0,0,0,0.3)', overflow: 'hidden'
          }}>
            {/* Modal Header */}
            <div style={{ background: '#4a4a4a', color: '#fff', padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <strong>✏️ แก้ไขสินค้า</strong>
              <button onClick={() => setEditProduct(null)}
                style={{ background: 'none', border: 'none', color: '#fff', fontSize: 20, cursor: 'pointer' }}>×</button>
            </div>

            {/* Modal Body */}
            <div style={{ padding: 20, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
              <div>
                <label style={labelStyle}>หมวดหมู่</label>
                <input style={inputStyle}
                  value={editProduct['หมวดหมู่สินค้า (CATEGORIES)'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, 'หมวดหมู่สินค้า (CATEGORIES)': e.target.value })}
                />
              </div>
              <div>
                <label style={labelStyle}>รหัสสินค้า</label>
                <input style={inputStyle}
                  value={editProduct['รหัสสินค้า (SKU NUMBER)'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, 'รหัสสินค้า (SKU NUMBER)': e.target.value })}
                />
              </div>
              <div style={{ gridColumn: 'span 2' }}>
                <label style={labelStyle}>ชื่อสินค้า</label>
                <input style={inputStyle}
                  value={editProduct['*ชื่อสินค้า (NAME)'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, '*ชื่อสินค้า (NAME)': e.target.value })}
                />
              </div>
              <div>
                <label style={labelStyle}>ใบอนุญาต</label>
                <input style={inputStyle}
                  value={editProduct['*เลขที่ใบอนุญาตโฆษณา'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, '*เลขที่ใบอนุญาตโฆษณา': e.target.value })}
                />
              </div>
              <div>
                <label style={labelStyle}>ราคา</label>
                <input style={inputStyle} type="number"
                  value={editProduct['*ราคาสินค้า'] || ''}
                  onChange={e => setEditProduct({ ...editProduct, '*ราคาสินค้า': e.target.value })}
                />
              </div>
            </div>

            {/* Modal Footer */}
            <div style={{ padding: '0 20px 20px', display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              <button onClick={() => setEditProduct(null)} style={btnStyle}>ยกเลิก</button>
              <button onClick={saveProduct}
                style={{ ...btnStyle, background: '#4a8bc4', color: '#fff', borderColor: '#3d7ab3' }}>
                💾 บันทึก
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Image Tooltip */}
      {tooltip && (
        <div style={{
          position: 'fixed',
          left: tooltip.x + 15,
          top: tooltip.y + 15,
          zIndex: 9999,
          background: '#fff',
          border: '2px solid #4a8bc4',
          borderRadius: 8,
          padding: 8,
          boxShadow: '0 10px 40px rgba(0,0,0,0.3)',
          pointerEvents: 'none'
        }}>
          {tooltip.url ? (
            <img
              src={tooltip.url}
              alt=""
              style={{ maxWidth: 280, maxHeight: 280, borderRadius: 4, display: 'block' }}
              onError={e => (e.currentTarget.style.display = 'none')}
            />
          ) : (
            <div style={{ padding: 40, color: '#999', fontSize: 12 }}>ไม่มีรูป</div>
          )}
        </div>
      )}

    </div>
  )
}

const btnStyle: React.CSSProperties = {
  padding: '7px 16px', border: '1px solid #777', borderRadius: 20,
  fontSize: 12, fontWeight: 600, cursor: 'pointer', background: '#f5f5f5'
}
const statCard: React.CSSProperties = {
  background: '#fff', border: '1px solid #ccc', borderRadius: 4,
  padding: 15, marginBottom: 10
}
const th: React.CSSProperties = {
  padding: '8px 12px', textAlign: 'left', border: '1px solid #ccc',
  fontWeight: 600, position: 'sticky', top: 0, background: '#f5f5f5'
}
const td: React.CSSProperties = {
  padding: '8px 12px', border: '1px solid #eee', color: '#444'
}
const labelStyle: React.CSSProperties = {
  fontSize: 11, color: '#666', display: 'block', marginBottom: 4
}
const inputStyle: React.CSSProperties = {
  width: '100%', padding: '6px 10px', border: '1px solid #999',
  borderRadius: 3, fontSize: 13, fontFamily: 'sans-serif',
  boxSizing: 'border-box'
}