const SESSION_COOKIE_NAME = "earthsoft_admin_session";
const SESSION_TTL_SECONDS = 60 * 60 * 12;

function getSessionSecret() {
  const secret = process.env.SESSION_SECRET;
  if (!secret) {
    throw new Error("SESSION_SECRET is not configured.");
  }
  return secret;
}

function getAdminUsername() {
  return String(process.env.ADMIN_USERNAME || "admin").trim().toLowerCase();
}

function getAdminPassword() {
  const password = process.env.ADMIN_PASSWORD;
  if (!password) {
    throw new Error("ADMIN_PASSWORD is not configured.");
  }
  return String(password);
}

function serializePayload(payload) {
  return encodeURIComponent(JSON.stringify(payload));
}

function parsePayload(value) {
  return JSON.parse(decodeURIComponent(value));
}

async function signValue(value, secret = getSessionSecret()) {
  const encoder = new TextEncoder();
  const key = await crypto.subtle.importKey(
    "raw",
    encoder.encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const signature = await crypto.subtle.sign("HMAC", key, encoder.encode(value));
  return Array.from(new Uint8Array(signature), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function getExpiryDate(ttlSeconds = SESSION_TTL_SECONDS) {
  return new Date(Date.now() + ttlSeconds * 1000);
}

async function createSessionToken(user, ttlSeconds = SESSION_TTL_SECONDS) {
  const payload = serializePayload({
    user,
    exp: Math.floor(Date.now() / 1000) + ttlSeconds,
  });
  const signature = await signValue(payload);
  return `${payload}.${signature}`;
}

async function verifySessionToken(token, secret = getSessionSecret()) {
  if (!token || typeof token !== "string" || !token.includes(".")) {
    return null;
  }

  const separatorIndex = token.lastIndexOf(".");
  const payload = token.slice(0, separatorIndex);
  const signature = token.slice(separatorIndex + 1);
  const expectedSignature = await signValue(payload, secret);

  if (signature !== expectedSignature) {
    return null;
  }

  const parsed = parsePayload(payload);
  if (!parsed?.user || !parsed?.exp || parsed.exp < Math.floor(Date.now() / 1000)) {
    return null;
  }

  return {
    user: String(parsed.user),
    exp: Number(parsed.exp),
  };
}

function buildSessionCookie(token, expiresAt = getExpiryDate()) {
  const parts = [
    `${SESSION_COOKIE_NAME}=${token}`,
    "Path=/",
    "HttpOnly",
    "SameSite=Strict",
    `Expires=${expiresAt.toUTCString()}`,
  ];

  if (process.env.NODE_ENV === "production") {
    parts.push("Secure");
  }

  return parts.join("; ");
}

function buildExpiredSessionCookie() {
  const parts = [
    `${SESSION_COOKIE_NAME}=`,
    "Path=/",
    "HttpOnly",
    "SameSite=Strict",
    "Expires=Thu, 01 Jan 1970 00:00:00 GMT",
    "Max-Age=0",
  ];

  if (process.env.NODE_ENV === "production") {
    parts.push("Secure");
  }

  return parts.join("; ");
}

export {
  SESSION_COOKIE_NAME,
  SESSION_TTL_SECONDS,
  buildExpiredSessionCookie,
  buildSessionCookie,
  createSessionToken,
  getAdminPassword,
  getAdminUsername,
  getExpiryDate,
  getSessionSecret,
  verifySessionToken,
};
