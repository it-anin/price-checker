import { supabase } from '../../../lib/supabase'
import { NextResponse, NextRequest } from 'next/server'

export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url)
  const sku = searchParams.get('sku')
  const sheet = searchParams.get('sheet')
  const skus = searchParams.get('skus')
  const branch = searchParams.get('branch')

  let q = supabase.from('products').select('*')

  if (sku) q = q.ilike('"รหัสสินค้า (SKU NUMBER)"', `%${sku}%`)
  if (sheet) q = q.eq('"หมวดหมู่สินค้า (CATEGORIES)"', sheet)
  if (skus) q = q.in('"รหัสสินค้า (SKU NUMBER)"', skus.split(','))
  if (branch && branch !== 'all') q = q.or(`branch.eq.${branch},branch.is.null`)

  const { data, error } = await q.limit(500)
  if (error) return NextResponse.json({ success: false, error: error.message })
  return NextResponse.json({ success: true, products: data })
}

export async function POST(req: NextRequest) {
  const body = await req.json()

  // query by skus array (batch 300 per request to avoid Supabase limits)
  if (Array.isArray(body.skus)) {
    const chunkSize = 300
    const branch = body.branch
    const all: any[] = []
    for (let i = 0; i < body.skus.length; i += chunkSize) {
      const chunk = body.skus.slice(i, i + chunkSize)
      let chunkQ = supabase
        .from('products')
        .select('*')
        .in('"รหัสสินค้า (SKU NUMBER)"', chunk)
      if (branch && branch !== 'all') chunkQ = chunkQ.or(`branch.eq.${branch},branch.is.null`)
      const { data, error } = await chunkQ
      if (error) return NextResponse.json({ success: false, error: error.message })
      all.push(...(data || []))
    }
    return NextResponse.json({ success: true, products: all })
  }

  // insert single product
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

export async function DELETE(req: NextRequest) {
  const body = await req.json()
  const sheet = body?.sheet
  const branch = body?.branch
  const deleteAll = body?.deleteAll

  if (!branch || branch === 'all') {
    return NextResponse.json({
      success: false,
      error: 'กรุณาระบุสาขาที่ต้องการลบให้ถูกต้อง'
    })
  }

  if (deleteAll) {
    const { data, error } = await supabase
      .from('products')
      .delete()
      .eq('branch', branch)
      .select('"รหัสสินค้า (SKU NUMBER)"')

    if (error) return NextResponse.json({ success: false, error: error.message })
    return NextResponse.json({ success: true, deleted: data?.length || 0 })
  }

  if (!sheet) {
    return NextResponse.json({
      success: false,
      error: 'กรุณาระบุหมวดหมู่ที่ต้องการลบ'
    })
  }

  const { data, error } = await supabase
    .from('products')
    .delete()
    .eq('"หมวดหมู่สินค้า (CATEGORIES)"', sheet)
    .eq('branch', branch)
    .select('"รหัสสินค้า (SKU NUMBER)"')

  if (error) return NextResponse.json({ success: false, error: error.message })
  return NextResponse.json({ success: true, deleted: data?.length || 0 })
}

