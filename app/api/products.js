import { supabase } from '../../lib/supabase'

export default async function handler(req, res) {
  const { method, query } = req

  // GET /api/products - ดึงทั้งหมด
  // GET /api/products?sku=xxx - ค้นหา
  // GET /api/products?sheet=xxx - กรองตาม category

  if (method === 'GET') {
    let q = supabase.from('products').select('*')

    if (query.sku) {
      q = q.ilike('"รหัสสินค้า (SKU NUMBER)"', `%${query.sku}%`)
    }

    if (query.sheet) {
      q = q.eq('"หมวดหมู่สินค้า (CATEGORIES)"', query.sheet)
    }

    const { data, error } = await q.limit(500)
    if (error) return res.status(500).json({ success: false, error: error.message })
    return res.status(200).json({ success: true, products: data })
  }

  // PATCH /api/products - อัพเดทสินค้า
  if (method === 'PATCH') {
    const { sku, ...updates } = req.body
    const { error } = await supabase
      .from('products')
      .update(updates)
      .eq('"รหัสสินค้า (SKU NUMBER)"', sku)

    if (error) return res.status(500).json({ success: false, error: error.message })
    return res.status(200).json({ success: true })
  }
}