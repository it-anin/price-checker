import { supabase } from '../../../lib/supabase'
import { NextResponse, NextRequest } from 'next/server'

export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url)
  const sku = searchParams.get('sku')
  const sheet = searchParams.get('sheet')
  const skus = searchParams.get('skus')

  let q = supabase.from('products').select('*')

  if (sku) q = q.ilike('"รหัสสินค้า (SKU NUMBER)"', `%${sku}%`)
  if (sheet) q = q.eq('"หมวดหมู่สินค้า (CATEGORIES)"', sheet)
  if (skus) q = q.in('"รหัสสินค้า (SKU NUMBER)"', skus.split(','))

  const { data, error } = await q.limit(500)
  if (error) return NextResponse.json({ success: false, error: error.message })
  return NextResponse.json({ success: true, products: data })
}

export async function POST(req: NextRequest) {
  const body = await req.json()

  const { data, error } = await supabase
    .from('products')
    .insert([body])
    .select()
    .single()

  if (error) return NextResponse.json({ success: false, error: error.message })
  return NextResponse.json({ success: true, product: data })
}

export async function PUT(req: NextRequest) {
  const { rows } = await req.json()
  if (!rows || rows.length === 0) return NextResponse.json({ success: true, count: 0 })

  const results = await Promise.all(
    rows.map(async (row: any) => {
      const sku = row['รหัสสินค้า (SKU NUMBER)']
      const { 'รหัสสินค้า (SKU NUMBER)': _sku, ...updates } = row

      const { error: insertErr } = await supabase.from('products').insert(row)
      if (!insertErr) return null

      // unique violation → update existing row
      if (insertErr.code === '23505') {
        const { error: updateErr } = await supabase
          .from('products')
          .update(updates)
          .eq('"รหัสสินค้า (SKU NUMBER)"', sku)
        return updateErr ? updateErr.message : null
      }

      return insertErr.message
    })
  )

  const failed = results.find(r => r !== null)
  if (failed) return NextResponse.json({ success: false, error: failed })

  return NextResponse.json({ success: true, count: rows.length })
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

