import { SESSION_COOKIE_NAME, verifySessionToken } from "../lib/auth.js";

function readCookie(req, name) {
  const rawCookies = String(req.headers.cookie || "");
  for (const part of rawCookies.split(";")) {
    const [key, ...rest] = part.trim().split("=");
    if (key === name) {
      return rest.join("=");
    }
  }
  return "";
}

export default async function handler(req, res) {
  if (req.method !== "GET") {
    res.setHeader("Allow", "GET");
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const token = readCookie(req, SESSION_COOKIE_NAME);
    const session = await verifySessionToken(token);

    if (!session) {
      return res.status(401).json({ authenticated: false });
    }

    return res.status(200).json({
      authenticated: true,
      user: session.user,
      expiresAt: new Date(session.exp * 1000).toISOString(),
    });
  } catch (error) {
    return res.status(500).json({ error: error.message || "Session check failed." });
  }
}
