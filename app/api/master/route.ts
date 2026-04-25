import { supabase } from '../../../lib/supabase'
import { NextResponse, NextRequest } from 'next/server'

async function fetchAllMasterSkus(): Promise<{ skus: string[]; error: string | null }> {
  const skus: string[] = []
  let from = 0
  while (true) {
    const { data, error } = await supabase
      .from('product_master')
      .select('sku')
      .range(from, from + 999)
    if (error) return { skus: [], error: error.message }
    skus.push(...(data || []).map((r: any) => String(r.sku)))
    if (!data || data.length < 1000) break
    from += 1000
  }
  return { skus, error: null }
}

async function fetchAllGrabSkus(): Promise<{ skus: Set<string>; error: string | null }> {
  const skus = new Set<string>()
  let from = 0
  while (true) {
    const { data, error } = await supabase
      .from('products')
      .select('"รหัสสินค้า (SKU NUMBER)"')
      .range(from, from + 999)
    if (error) return { skus, error: error.message }
    for (const row of data || []) {
      const sku = row['รหัสสินค้า (SKU NUMBER)']
      if (sku) skus.add(String(sku))
    }
    if (!data || data.length < 1000) break
    from += 1000
  }
  return { skus, error: null }
}

export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url)
  const compare = searchParams.get('compare') === 'true'

  const { skus: masterSkus, error: masterError } = await fetchAllMasterSkus()
  if (masterError) return NextResponse.json({ success: false, error: masterError })

  if (compare) {
    const { skus: grabSkus, error: grabError } = await fetchAllGrabSkus()
    if (grabError) return NextResponse.json({ success: false, error: grabError })
    const missingSkus = masterSkus.filter(sku => !grabSkus.has(sku))
    return NextResponse.json({
      success: true,
      masterCount: masterSkus.length,
      grabCount: grabSkus.size,
      missingSkus,
      missingCount: missingSkus.length
    })
  }

  return NextResponse.json({ success: true, count: masterSkus.length })
}

export async function POST(req: NextRequest) {
  const { skus } = await req.json()
  if (!Array.isArray(skus)) return NextResponse.json({ success: false, error: 'skus required' })

  const uniqueSkus: string[] = Array.from(new Set(skus.map((s: any) => String(s ?? '').trim()).filter(Boolean)))

  // ลบทั้งหมดแล้ว insert ใหม่
  const { error: deleteError } = await supabase
    .from('product_master')
    .delete()
    .gte('sku', '')

  if (deleteError) return NextResponse.json({ success: false, error: `delete: ${deleteError.message}` })

  const CHUNK = 500
  for (let i = 0; i < uniqueSkus.length; i += CHUNK) {
    const chunk = uniqueSkus.slice(i, i + CHUNK).map(sku => ({ sku }))
    const { error } = await supabase.from('product_master').insert(chunk)
    if (error) return NextResponse.json({ success: false, error: `insert: ${error.message}` })
  }

  return NextResponse.json({ success: true, count: uniqueSkus.length })
}
