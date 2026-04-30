import { supabase } from '../../lib/supabase'

export default async function handler(req, res) {
  // นับสินค้าแยกตาม category
  const { data, error } = await supabase
    .from('products')
    .select('"หมวดหมู่สินค้า (CATEGORIES)"')

  if (error) return res.status(500).json({ success: false, error: error.message })

  // Group by category
  const categoryCount = {}
  data.forEach(p => {
    const cat = p['หมวดหมู่สินค้า (CATEGORIES)'] || 'ไม่ระบุ'
    categoryCount[cat] = (categoryCount[cat] || 0) + 1
  })

  const sheets = Object.entries(categoryCount).map(([name, count]) => ({ name, count }))

  return res.status(200).json({
    success: true,
    totalProducts: data.length,
    sheets
  })
}