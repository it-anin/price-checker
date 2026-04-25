import { supabase } from '../../../lib/supabase'
import { NextResponse, NextRequest } from 'next/server'

function aggregateCategories(rows: any[]) {
  const grouped = new Map<string, string>()

  for (const row of rows || []) {
    const sku = row['รหัสสินค้า (SKU NUMBER)']
    if (!sku || grouped.has(sku)) continue
    grouped.set(sku, row['หมวดหมู่สินค้า (CATEGORIES)'] || 'ไม่ระบุ')
  }

  return Array.from(grouped.values())
}

export async function GET(req: NextRequest) {
  let q = supabase.from('products').select('"รหัสสินค้า (SKU NUMBER)","หมวดหมู่สินค้า (CATEGORIES)"')

  const { data, error } = await q
  if (error) return NextResponse.json({ success: false, error: error.message })

  const categoryCount: Record<string, number> = {}
  const categories = aggregateCategories(data || [])
  categories.forEach((cat: string) => {
    categoryCount[cat] = (categoryCount[cat] || 0) + 1
  })

  const sheets = Object.entries(categoryCount).map(([name, count]) => ({ name, count }))

  return NextResponse.json({
    success: true,
    totalProducts: categories.length,
    sheets
  })
}
