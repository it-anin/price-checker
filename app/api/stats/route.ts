import { supabase } from '../../../lib/supabase'
import { NextResponse } from 'next/server'

export async function GET() {
  const { data, error } = await supabase
    .from('products')
    .select('"หมวดหมู่สินค้า (CATEGORIES)"')

  if (error) return NextResponse.json({ success: false, error: error.message })

  const categoryCount: Record<string, number> = {}
  data.forEach((p: any) => {
    const cat = p['หมวดหมู่สินค้า (CATEGORIES)'] || 'ไม่ระบุ'
    categoryCount[cat] = (categoryCount[cat] || 0) + 1
  })

  const sheets = Object.entries(categoryCount).map(([name, count]) => ({ name, count }))

  return NextResponse.json({
    success: true,
    totalProducts: data.length,
    sheets
  })
}