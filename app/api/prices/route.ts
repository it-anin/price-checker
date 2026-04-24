import { supabase } from '../../../lib/supabase'
import { NextResponse,NextRequest } from 'next/server'

export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url)
  const branch = searchParams.get('branch')

  let q = supabase.from('products').select('*, sku_reference!inner(ref_price)')
  if (branch && branch !== 'all') {
    q = q.eq('branch', branch)
  }

  const { data, error } = await q

  if (error) return NextResponse.json({ success: false, error: error.message })

  const mismatched = (data || []).filter((p: any) => {
    const current = parseFloat(p['*ราคาสินค้า']) || 0
    const correct = parseFloat(p.sku_reference?.ref_price) || 0
    return Math.abs(current - correct) > 0.01
  }).map((p: any) => ({
    sku: p['รหัสสินค้า (SKU NUMBER)'],
    name: p['*ชื่อสินค้า (NAME)'],
    sheet: p['หมวดหมู่สินค้า (CATEGORIES)'],
    currentPrice: p['*ราคาสินค้า'],
    correctPrice: p.sku_reference.ref_price,
    difference: parseFloat(p.sku_reference.ref_price) - parseFloat(p['*ราคาสินค้า'])
  }))

  return NextResponse.json({
    success: true,
    mismatchCount: mismatched.length,
    mismatched
  })
}

export async function POST(req: NextRequest) {
  const { items } = await req.json()
  let updated = 0

  for (const item of items) {
    const { error } = await supabase
      .from('products')
      .update({ '*ราคาสินค้า': item.correctPrice })
      .eq('"รหัสสินค้า (SKU NUMBER)"', item.sku)
    if (!error) updated++
  }

  return NextResponse.json({ success: true, updated })
}
