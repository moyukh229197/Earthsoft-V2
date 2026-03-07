import {
  buildSessionCookie,
  createSessionToken,
  getAdminPassword,
  getAdminUsername,
  getExpiryDate,
} from "../lib/auth.js";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  try {
    const username = String(req.body?.username || "").trim().toLowerCase();
    const password = String(req.body?.password || "");

    if (!username || !password) {
      return res.status(400).json({ error: "Enter both username and password to continue." });
    }

    const adminUsername = getAdminUsername();
    const adminPassword = getAdminPassword();

    if (username !== adminUsername || password !== adminPassword) {
      return res.status(401).json({ error: "Invalid username or password." });
    }

    const expiresAt = getExpiryDate();
    const token = await createSessionToken(adminUsername);
    res.setHeader("Set-Cookie", buildSessionCookie(token, expiresAt));
    return res.status(200).json({
      authenticated: true,
      user: adminUsername,
      expiresAt: expiresAt.toISOString(),
    });
  } catch (error) {
    return res.status(500).json({ error: error.message || "Login failed." });
  }
}
