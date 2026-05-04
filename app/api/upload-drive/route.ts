import { NextRequest, NextResponse } from 'next/server'
import crypto from 'crypto'

export const runtime = 'nodejs'

function base64url(input: Buffer | string) {
  return Buffer.from(input).toString('base64')
    .replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_')
}

function signJWT(claims: Record<string, any>, privateKey: string) {
  const header = { alg: 'RS256', typ: 'JWT' }
  const headerB64 = base64url(JSON.stringify(header))
  const claimsB64 = base64url(JSON.stringify(claims))
  const data = `${headerB64}.${claimsB64}`
  const signer = crypto.createSign('RSA-SHA256')
  signer.update(data)
  const sig = signer.sign(privateKey)
  return `${data}.${base64url(sig)}`
}

async function getAccessToken() {
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL
  const rawKey = process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY
  if (!email || !rawKey) {
    throw new Error('GOOGLE_SERVICE_ACCOUNT_EMAIL/GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY ยังไม่ได้ตั้งใน .env.local')
  }
  const privateKey = rawKey.replace(/\\n/g, '\n')
  const now = Math.floor(Date.now() / 1000)
  const claims = {
    iss: email,
    scope: 'https://www.googleapis.com/auth/drive.file',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600
  }
  const jwt = signJWT(claims, privateKey)
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
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
