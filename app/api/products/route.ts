import { supabase } from '../../../lib/supabase'
import { NextResponse, NextRequest } from 'next/server'

const BRANCH_KEYS = ['src', 'kkl', 'sss'] as const

function normalizeBranchName(value: any) {
  return String(value ?? '').trim().toLowerCase()
}

function parseBranches(branch: any): string[] {
  if (!branch) return []
  return String(branch).split(',').map((b: string) => b.trim()).filter(Boolean)
}

function aggregateProducts(rows: any[]) {
  const grouped = new Map<string, any>()

  for (const row of rows || []) {
    const sku = row['รหัสสินค้า (SKU NUMBER)']
    if (!sku) continue

    if (!grouped.has(sku)) {
      grouped.set(sku, { ...row, src: false, kkl: false, sss: false, _branches: [] as string[] })
    }

    const current = grouped.get(sku)
    for (const b of parseBranches(row.branch)) {
      const nb = normalizeBranchName(b)
      if (BRANCH_KEYS.includes(nb as (typeof BRANCH_KEYS)[number])) {
        current[nb] = true
        if (!current._branches.includes(nb)) current._branches.push(nb)
      }
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
    const skuSetSet = new Set(skuSet)
    const skuColumn = '"รหัสสินค้า (SKU NUMBER)"'
    const CHUNK = 200

    // ดึงสินค้าทั้งหมดพร้อม branch ปัจจุบัน (paginate)
    const allProducts: any[] = []
    let from = 0
    while (true) {
      const { data, error } = await supabase
        .from('products')
        .select(`${skuColumn}, branch`)
        .range(from, from + 999)
      if (error) return NextResponse.json({ success: false, error: `fetch: ${error.message}` })
      allProducts.push(...(data || []))
      if (!data || data.length < 1000) break
      from += 1000
    }

    // สร้าง map SKU → branches ปัจจุบัน
    const branchMap = new Map<string, string[]>()
    for (const p of allProducts) {
      const sku = p['รหัสสินค้า (SKU NUMBER)']
      if (sku) branchMap.set(sku, parseBranches(p.branch))
    }

    const existingSkus = new Set(branchMap.keys())
    const missingBase = skuSet.filter(sku => !existingSkus.has(sku))

    // SKU ที่ต้องเพิ่ม branch (อยู่ใน CSV, มีใน DB, ยังไม่มี branch นี้)
    const toAdd = skuSet.filter(sku => {
      const b = branchMap.get(sku) || []
      return existingSkus.has(sku) && !b.includes(availabilityBranch)
    })

    // SKU ที่ต้องถอน branch (ไม่อยู่ใน CSV, มีใน DB, มี branch นี้อยู่)
    const toRemove = Array.from(existingSkus).filter(sku => {
      const b = branchMap.get(sku) || []
      return !skuSetSet.has(sku) && b.includes(availabilityBranch)
    })

    // จัดกลุ่มตาม new branch value เพื่อ bulk UPDATE รอบเดียวต่อกลุ่ม
    const addGroups = new Map<string, string[]>()
    for (const sku of toAdd) {
      const newVal = [...(branchMap.get(sku) || []), availabilityBranch].sort().join(',')
      if (!addGroups.has(newVal)) addGroups.set(newVal, [])
      addGroups.get(newVal)!.push(sku)
    }

    const removeGroups = new Map<string, string[]>()
    for (const sku of toRemove) {
      const newVal = (branchMap.get(sku) || []).filter(b => b !== availabilityBranch).sort().join(',')
      if (!removeGroups.has(newVal)) removeGroups.set(newVal, [])
      removeGroups.get(newVal)!.push(sku)
    }

    // UPDATE เพิ่ม branch
    for (const [newVal, skus] of addGroups) {
      for (let i = 0; i < skus.length; i += CHUNK) {
        const { error } = await supabase
          .from('products')
          .update({ branch: newVal })
          .in(skuColumn, skus.slice(i, i + CHUNK))
        if (error) return NextResponse.json({ success: false, error: `add: ${error.message}` })
      }
    }

    // UPDATE ถอน branch
    for (const [newVal, skus] of removeGroups) {
      for (let i = 0; i < skus.length; i += CHUNK) {
        const { error } = await supabase
          .from('products')
          .update({ branch: newVal || null })
          .in(skuColumn, skus.slice(i, i + CHUNK))
        if (error) return NextResponse.json({ success: false, error: `remove: ${error.message}` })
      }
    }

    return NextResponse.json({
      success: true,
      syncedBranch: availabilityBranch,
      inserted: toAdd.length,
      deleted: toRemove.length,
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

