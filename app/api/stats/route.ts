import { supabase } from '../../../lib/supabase'
import { NextResponse, NextRequest } from 'next/server'

export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url)
  const branch = searchParams.get('branch')

  let q = supabase.from('products').select('"หมวดหมู่สินค้า (CATEGORIES)"')
  if (branch && branch !== 'all') {
    q = q.or(`branch.eq.${branch},branch.is.null`)
  }

  const { data, error } = await q
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
