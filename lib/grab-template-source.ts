import { supabase } from './supabase'
import { validateGrabTemplateUpload } from './process-engine'
import { compareGrabTemplates, type SavedGrabReview } from './grab-template-review'

export async function saveGrabTemplate(sourceName: string, text: string): Promise<SavedGrabReview> {
  const template = validateGrabTemplateUpload(text)
  for (let attempt = 0; attempt < 3; attempt++) {
    const { data: previous, error: readError } = await supabase.from('grab_files')
      .select('content,updated_at').eq('name', template.name).maybeSingle()
    if (readError) throw new Error(`อ่านไฟล์เดิมก่อนบันทึกไม่สำเร็จ: ${readError.message}`)
    if (previous && (typeof previous.content !== 'string' || typeof previous.updated_at !== 'string' || !Number.isFinite(Date.parse(previous.updated_at)))) throw new Error('ไฟล์เดิมหรือเวอร์ชันไฟล์ไม่ถูกต้อง จึงยังเทียบและบันทึกไม่ได้')
    const review = compareGrabTemplates(previous?.content ?? null, text)
    const updatedAt = new Date(Math.max(Date.now(), previous ? Date.parse(previous.updated_at) + 1 : 0)).toISOString()
    // Match the version used for comparison. A simultaneous edit must be
    // re-read and compared again, rather than overwritten with a stale diff.
    const write = previous
      ? supabase.from('grab_files').update({ content: text, updated_at: updatedAt }).eq('name', template.name).eq('updated_at', previous.updated_at)
      : supabase.from('grab_files').insert({ name: template.name, content: text, updated_at: updatedAt })
    const { data: saved, error: writeError } = await write.select('name,updated_at').maybeSingle()
    if (writeError) {
      if (!previous && writeError.code === '23505') continue
      throw new Error(`บันทึกไม่สำเร็จ: ${writeError.message}`)
    }
    if (saved) return { ...review, sourceName, updatedAt: saved.updated_at }
  }
  throw new Error('ไฟล์ถูกแก้ไขพร้อมกันหรือบัญชีไม่มีสิทธิ์บันทึก กรุณาตรวจสิทธิ์แล้วลองอัปโหลดอีกครั้งเพื่อเทียบกับไฟล์ล่าสุด')
}
