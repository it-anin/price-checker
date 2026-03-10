import { supabase } from '../../../lib/supabase'
import { NextResponse, NextRequest } from 'next/server'

export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url)
  const sku = searchParams.get('sku')
  const sheet = searchParams.get('sheet')

  let q = supabase.from('products').select('*')

  if (sku) q = q.ilike('"รหัสสินค้า (SKU NUMBER)"', `%${sku}%`)
  if (sheet) q = q.eq('"หมวดหมู่สินค้า (CATEGORIES)"', sheet)

  const { data, error } = await q.limit(500)
  if (error) return NextResponse.json({ success: false, error: error.message })
  return NextResponse.json({ success: true, products: data })
}

export async function PATCH(req: NextRequest) {
  const body = await req.json()
  const { sku, ...updates } = body

  const { error } = await supabase
    .from('products')
    .update(updates)
    .eq('"รหัสสินค้า (SKU NUMBER)"', sku)

  if (error) return NextResponse.json({ success: false, error: error.message })
  return NextResponse.json({ success: true })
}