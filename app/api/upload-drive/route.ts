import { NextRequest, NextResponse } from 'next/server'

export const runtime = 'nodejs'

async function getAccessToken() {
  const clientId = process.env.GOOGLE_OAUTH_CLIENT_ID
  const clientSecret = process.env.GOOGLE_OAUTH_CLIENT_SECRET
  const refreshToken = process.env.GOOGLE_OAUTH_REFRESH_TOKEN
  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error('GOOGLE_OAUTH_CLIENT_ID / GOOGLE_OAUTH_CLIENT_SECRET / GOOGLE_OAUTH_REFRESH_TOKEN ยังไม่ได้ตั้งใน .env.local')
  }
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      refresh_token: refreshToken,
      grant_type: 'refresh_token'
    })
  })
  const data = await res.json()
  if (!data.access_token) {
    throw new Error('ขอ access token จาก Google ไม่สำเร็จ: ' + JSON.stringify(data))
  }
  return data.access_token as string
}

async function uploadToDrive(buf: Buffer, filename: string, mime: string, accessToken: string, folderId: string) {
  const metadata = { name: filename, parents: [folderId] }
  const boundary = '-------price-checker-' + Date.now()
  const delim = `\r\n--${boundary}\r\n`
  const closeDelim = `\r\n--${boundary}--`
  const head = delim
    + 'Content-Type: application/json; charset=UTF-8\r\n\r\n'
    + JSON.stringify(metadata)
    + delim
    + `Content-Type: ${mime}\r\n\r\n`
  const body = Buffer.concat([Buffer.from(head, 'utf8'), buf, Buffer.from(closeDelim, 'utf8')])
  const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': `multipart/related; boundary=${boundary}`,
      'Content-Length': String(body.length)
    },
    body
  })
  const data = await res.json()
  if (!res.ok || !data.id) {
    throw new Error('อัพโหลดไป Drive ไม่สำเร็จ: ' + JSON.stringify(data))
  }
  return data as { id: string; name: string }
}

async function makePublic(fileId: string, accessToken: string) {
  await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}/permissions`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ role: 'reader', type: 'anyone' })
  })
}

export async function POST(req: NextRequest) {
  try {
    const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID
    if (!folderId) {
      return NextResponse.json({ success: false, error: 'GOOGLE_DRIVE_FOLDER_ID ยังไม่ได้ตั้งใน .env.local' }, { status: 500 })
    }
    const formData = await req.formData()
    const file = formData.get('file') as File | null
    const filenameInput = (formData.get('filename') as string | null) || ''
    if (!file) {
      return NextResponse.json({ success: false, error: 'ไม่พบไฟล์' }, { status: 400 })
    }
    const buf = Buffer.from(await file.arrayBuffer())
    const mime = file.type || 'image/png'
    const ext = mime.includes('jpeg') ? 'jpg' : mime.includes('png') ? 'png' : mime.includes('webp') ? 'webp' : 'png'
    const safeBase = filenameInput.replace(/[\\/:*?"<>|]/g, '_').trim() || `product-${Date.now()}`
    const filename = /\.(png|jpe?g|webp)$/i.test(safeBase) ? safeBase : `${safeBase}.${ext}`
    const token = await getAccessToken()
    const result = await uploadToDrive(buf, filename, mime, token, folderId)
    await makePublic(result.id, token)
    const link = `https://drive.google.com/file/d/${result.id}/view?usp=sharing`
    return NextResponse.json({ success: true, link, id: result.id, name: result.name })
  } catch (e: any) {
    return NextResponse.json({ success: false, error: e?.message || String(e) }, { status: 500 })
  }
}

export async function GET(req: NextRequest) {
  try {
    const fileId = new URL(req.url).searchParams.get('fileId')
    if (!fileId) {
      return NextResponse.json({ success: false, error: 'กรุณาระบุ fileId' }, { status: 400 })
    }
    const token = await getAccessToken()
    const res = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?fields=owners(emailAddress),name`, {
      headers: { Authorization: `Bearer ${token}` }
    })
    if (!res.ok) {
      const err = await res.json().catch(() => ({}))
      return NextResponse.json({ success: false, error: err?.error?.message || 'ดึงข้อมูลไฟล์ไม่สำเร็จ' }, { status: res.status })
    }
    const data = await res.json()
    const owner = data?.owners?.[0]?.emailAddress || ''
    const type = owner.toLowerCase().endsWith('@anin.co.th') ? 'anin' : 'external'
    return NextResponse.json({ success: true, owner, name: data?.name || '', type })
  } catch (e: any) {
    return NextResponse.json({ success: false, error: e?.message || String(e) }, { status: 500 })
  }
}
