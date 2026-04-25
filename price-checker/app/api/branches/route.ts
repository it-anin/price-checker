import { supabase } from '../../../lib/supabase'
import { NextResponse, NextRequest } from 'next/server'

export async function GET() {
  const { data, error } = await supabase
    .from('branches')
    .select('id, name, created_at')
    .order('created_at', { ascending: true })

  if (error) return NextResponse.json({ success: false, error: error.message })
  return NextResponse.json({ success: true, branches: data })
}

export async function POST(req: NextRequest) {
  const { name } = await req.json()
  if (!name?.trim()) {
    return NextResponse.json({ success: false, error: 'ชื่อสาขาห้ามว่าง' })
  }

  const { data, error } = await supabase
    .from('branches')
    .insert([{ name: name.trim() }])
    .select()
    .single()

  if (error) return NextResponse.json({ success: false, error: error.message })
  return NextResponse.json({ success: true, branch: data })
}
