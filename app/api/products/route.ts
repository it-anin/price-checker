import { supabase } from '../../../lib/supabase'
import { NextResponse, NextRequest } from 'next/server'

const BRANCH_KEYS = ['src', 'kkl', 'sss'] as const

function normalizeBranchName(value: any) {
  return String(value ?? '').trim().toLowerCase()
}

function aggregateProducts(rows: any[]) {
  const grouped = new Map<string, any>()

  for (const row of rows || []) {
    const sku = row['รหัสสินค้า (SKU NUMBER)']
    if (!sku) continue

    if (!grouped.has(sku)) {
      grouped.set(sku, {
        ...row,
        src: false,
        kkl: false,
        sss: false,
        _branches: [] as string[]
      })
    }

    const current = grouped.get(sku)
    const normalizedBranch = normalizeBranchName(row.branch)
    if (BRANCH_KEYS.includes(normalizedBranch as (typeof BRANCH_KEYS)[number])) {
      current[normalizedBranch] = true
      if (!current._branches.includes(normalizedBranch)) current._branches.push(normalizedBranch)
    }
  }

  return Array.from(grouped.values())
}

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
  return NextResponse.json({ success: true, products: aggregateProducts(data || []) })
}

export async function POST(req: NextRequest) {
  const body = await req.json()

  if (body?.availabilityBranch && Array.isArray(body.skus)) {
    const availabilityBranch = normalizeBranchName(body.availabilityBranch)
    if (!BRANCH_KEYS.includes(availabilityBranch as (typeof BRANCH_KEYS)[number])) {
      return NextResponse.json({ success: false, error: 'Invalid branch' })
    }

    const skuSet: string[] = Array.from(new Set(body.skus.map((sku: any) => String(sku ?? '').trim()).filter(Boolean)))
    const skuColumn = '"รหัสสินค้า (SKU NUMBER)"'

    const { data: targetRows, error: targetError } = await supabase
      .from('products')
      .select('*')
      .eq('branch', availabilityBranch)

    if (targetError) return NextResponse.json({ success: false, error: targetError.message })

    const { data: sourceRows, error: sourceError } = await supabase
      .from('products')
      .select('*')
      .in(skuColumn, skuSet)

    if (sourceError) return NextResponse.json({ success: false, error: sourceError.message })

    const existingTargetSkus = new Set((targetRows || []).map((row: any) => row['รหัสสินค้า (SKU NUMBER)']).filter(Boolean))
    const sourceBySku = new Map<string, any>()
    for (const row of sourceRows || []) {
      const sku = row['รหัสสินค้า (SKU NUMBER)']
      if (sku && !sourceBySku.has(sku)) sourceBySku.set(sku, row)
    }

    const toDelete = (targetRows || [])
      .map((row: any) => row['รหัสสินค้า (SKU NUMBER)'])
      .filter((sku: string) => sku && !skuSet.includes(sku))

    const toInsert = skuSet
      .filter((sku: string) => !existingTargetSkus.has(sku))
      .map((sku: string) => {
        const base = sourceBySku.get(sku)
        if (!base) return null
        const { id, created_at, updated_at, ...clone } = base
        return { ...clone, branch: availabilityBranch }
      })
      .filter(Boolean)

    const missingBase = skuSet.filter((sku: string) => !sourceBySku.has(sku))

    if (toDelete.length > 0) {
      const { error: deleteError } = await supabase
        .from('products')
        .delete()
        .eq('branch', availabilityBranch)
        .in(skuColumn, toDelete)
      if (deleteError) return NextResponse.json({ success: false, error: deleteError.message })
    }

    if (toInsert.length > 0) {
      const { error: insertError } = await supabase
        .from('products')
        .insert(toInsert)
      if (insertError) return NextResponse.json({ success: false, error: insertError.message })
    }

    return NextResponse.json({
      success: true,
      syncedBranch: availabilityBranch,
      inserted: toInsert.length,
      deleted: toDelete.length,
      missingBase: missingBase.length
    })
  }

  // query by skus array (batch 300 per request to avoid Supabase limits)
  if (Array.isArray(body.skus)) {
    const chunkSize = 300
    const all: any[] = []
    for (let i = 0; i < body.skus.length; i += chunkSize) {
      const chunk = body.skus.slice(i, i + chunkSize)
      let chunkQ = supabase
        .from('products')
        .select('*')
        .in('"รหัสสินค้า (SKU NUMBER)"', chunk)
      const { data, error } = await chunkQ
      if (error) return NextResponse.json({ success: false, error: error.message })
      all.push(...(data || []))
    }
    return NextResponse.json({ success: true, products: aggregateProducts(all) })
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

  // deleteAll requires a specific branch (not 'all')
  if (deleteAll) {
    if (!branch || branch === 'all') {
      return NextResponse.json({ success: false, error: 'กรุณาระบุสาขาที่ต้องการลบให้ถูกต้อง' })
    }

    const { data, error } = await supabase
      .from('products')
      .delete()
      .eq('branch', branch)
      .select('"รหัสสินค้า (SKU NUMBER)"')

    if (error) return NextResponse.json({ success: false, error: error.message })
    return NextResponse.json({ success: true, deleted: data?.length || 0 })
  }

  // delete by category — branch='all' means delete across all branches
  if (!sheet) {
    return NextResponse.json({ success: false, error: 'กรุณาระบุหมวดหมู่ที่ต้องการลบ' })
  }

  let q = supabase
    .from('products')
    .delete()
    .eq('"หมวดหมู่สินค้า (CATEGORIES)"', sheet)

  if (branch && branch !== 'all') {
    q = q.eq('branch', branch)
  }

  const { data, error } = await q.select('"รหัสสินค้า (SKU NUMBER)"')

  if (error) return NextResponse.json({ success: false, error: error.message })
  return NextResponse.json({ success: true, deleted: data?.length || 0 })
}

