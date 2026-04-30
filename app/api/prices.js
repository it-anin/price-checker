import { supabase } from '../../lib/supabase'

export default async function handler(req, res) {
  // GET - ตรวจสอบราคาทั้งหมด เปรียบเทียบกับ sku_reference
  if (req.method === 'GET') {
    const { data, error } = await supabase
      .from('products')
      .select(`
        *,
        sku_reference!inner(ref_price)
      `)

    if (error) return res.status(500).json({ success: false, error: error.message })

    const mismatched = data.filter(p => {
      const current = parseFloat(p['*ราคาสินค้า']) || 0
      const correct = parseFloat(p.sku_reference?.ref_price) || 0
      return Math.abs(current - correct) > 0.01
    }).map(p => ({
      sku: p['รหัสสินค้า (SKU NUMBER)'],
      name: p['*ชื่อสินค้า (NAME)'],
      sheet: p['หมวดหมู่สินค้า (CATEGORIES)'],
      currentPrice: p['*ราคาสินค้า'],
      correctPrice: p.sku_reference.ref_price,
      difference: parseFloat(p.sku_reference.ref_price) - parseFloat(p['*ราคาสินค้า'])
    }))

    return res.status(200).json({
      success: true,
      mismatchCount: mismatched.length,
      mismatched
    })
  }

  // POST - อัพเดทราคา
  if (req.method === 'POST') {
    const { items } = req.body
    let updated = 0

    for (const item of items) {
      const { error } = await supabase
        .from('products')
        .update({ '*ราคาสินค้า': item.correctPrice })
        .eq('"รหัสสินค้า (SKU NUMBER)"', item.sku)

      if (!error) updated++
    }

    return res.status(200).json({ success: true, updated })
  }
}
```

---

### Step 2 — ทดสอบ API

รัน dev server แล้วเปิด browser ทดสอบ:
```
http://localhost:3000/api/stats
http://localhost:3000/api/products
http://localhost:3000/api/products?sku=SKU001