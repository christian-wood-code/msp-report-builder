"use strict";
const { createHmac, timingSafeEqual } = require("crypto");

// Staff session: an HS256 JWT in an httpOnly cookie, issued by auth-session.js after the Entra ID token
// has been verified server-side. Separate cookie name and secret from the dashboard's, so a session on
// one site is never valid on the other. Fails closed if SESSION_JWT_SECRET is unset or short.

const COOKIE = "msprb_session";
const SESSION_HOURS = 8;
const b64u = s => Buffer.from(s, "base64url");

function signSession(claims, secret, now = Date.now()) {
  const h = Buffer.from('{"alg":"HS256","typ":"JWT"}').toString("base64url");
  const p = Buffer.from(JSON.stringify({ ...claims, iat: Math.floor(now / 1000), exp: Math.floor(now / 1000) + SESSION_HOURS * 3600 })).toString("base64url");
  return h + "." + p + "." + createHmac("sha256", secret).update(h + "." + p).digest("base64url");
}
const sessionCookie = token => `${COOKIE}=${token}; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=${SESSION_HOURS * 3600}`;
const clearCookie = `${COOKIE}=; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=0`;

// Returns { claims } for a valid session, { error: 'not-configured' | 'none' } otherwise.
function readSession(event, now = Date.now()) {
  const secret = process.env.SESSION_JWT_SECRET;
  if (!secret || secret.length < 32) return { error: "not-configured" };
  const h = event.headers || {};
  const cookie = h.cookie || h.Cookie || "";
  const m = new RegExp("(?:^|;\\s*)" + COOKIE + "=([^;]+)").exec(cookie);
  if (!m) return { error: "none" };
  const parts = m[1].split(".");
  if (parts.length !== 3) return { error: "none" };
  try {
    const hdr = JSON.parse(b64u(parts[0]).toString());
    if (hdr.alg !== "HS256") throw new Error("alg");
    const sig = createHmac("sha256", secret).update(parts[0] + "." + parts[1]).digest();
    const given = b64u(parts[2]);
    if (given.length !== sig.length || !timingSafeEqual(sig, given)) throw new Error("sig");
    const claims = JSON.parse(b64u(parts[1]).toString());
    if (!claims.exp || claims.exp * 1000 < now) throw new Error("exp");
    return { claims };
  } catch (e) { return { error: "none" }; }
}

module.exports = { COOKIE, SESSION_HOURS, signSession, sessionCookie, clearCookie, readSession };
