const SESSION_COOKIE_NAME = "earthsoft_admin_session"

function parsePayload(value: string) {
  try {
    return JSON.parse(decodeURIComponent(value))
  } catch {
    return null
  }
}

async function signValue(value: string, secret: string) {
  const encoder = new TextEncoder()
  const key = await crypto.subtle.importKey(
    "raw",
    encoder.encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  )
  const signature = await crypto.subtle.sign("HMAC", key, encoder.encode(value))
  return Array.from(new Uint8Array(signature), (byte) => byte.toString(16).padStart(2, "0")).join("")
}

async function isAuthorized(token: string, secret: string) {
  if (!token || !token.includes(".")) return false

  const separatorIndex = token.lastIndexOf(".")
  const payload = token.slice(0, separatorIndex)
  const signature = token.slice(separatorIndex + 1)
  const expectedSignature = await signValue(payload, secret)

  if (signature !== expectedSignature) return false

  const parsed = parsePayload(payload)
  return Boolean(parsed?.user && parsed?.exp && Number(parsed.exp) >= Math.floor(Date.now() / 1000))
}

function readCookie(header: string, name: string) {
  for (const part of header.split(";")) {
    const [key, ...rest] = part.trim().split("=")
    if (key === name) {
      return rest.join("=")
    }
  }
  return ""
}

export default async function middleware(request: Request) {
  const secret = process.env.SESSION_SECRET
  if (!secret) {
    return Response.json({ error: "SESSION_SECRET is not configured." }, { status: 500 })
  }

  const url = new URL(request.url)
  const token = readCookie(request.headers.get("cookie") || "", SESSION_COOKIE_NAME)
  const authorized = await isAuthorized(token, secret)

  if (authorized) {
    return
  }

  if (url.pathname.startsWith("/api/")) {
    return Response.json({ error: "Unauthorized" }, { status: 401 })
  }

  const loginUrl = new URL("/", request.url)
  return Response.redirect(loginUrl, 307)
}

export const config = {
  matcher: ["/workspace/index.html", "/api/extract-data"],
}
