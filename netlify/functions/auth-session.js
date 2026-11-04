"use strict";
// POST {idToken}: verify the Entra ID token server-side, then issue the staff session cookie.
// GET: who am I (from the session cookie). DELETE: sign out.
const { verifyEntraIdToken } = require("./lib/entra");
const { readSession, signSession, sessionCookie, clearCookie } = require("./lib/staff-session");
const { makeLimiter } = require("./lib/rate-limit");

const limiter = makeLimiter(20, 5 * 60 * 1000);
const allowedDomains = () => String(process.env.ALLOWED_EMAIL_DOMAINS || "").toLowerCase().split(",").map(s => s.trim()).filter(Boolean);
const J = (statusCode, obj, extra) => ({ statusCode, headers: { "Content-Type": "application/json", "Cache-Control": "no-store", ...(extra || {}) }, body: JSON.stringify(obj) });

exports.handler = async (event) => {
  if (event.httpMethod === "GET") {
    const s = readSession(event);
    return s.claims ? J(200, { name: s.claims.name, upn: s.claims.upn }) : J(401, { error: "Not signed in" });
  }
  if (event.httpMethod === "DELETE") return J(200, { ok: true }, { "Set-Cookie": clearCookie });
  if (event.httpMethod !== "POST") return J(405, { error: "Method not allowed" });
  const h = event.headers || {};
  if (!/application\/json/i.test(h["content-type"] || h["Content-Type"] || "")) return J(415, { error: "JSON only" });

  const ip = h["x-nf-client-connection-ip"] || "unknown";
  if (!limiter(ip)) return J(429, { error: "Too many sign-in attempts. Wait a few minutes." });

  const secret = process.env.SESSION_JWT_SECRET;
  if (!secret || secret.length < 32) return J(500, { error: "Sign-in not configured" });
  let body;
  try { body = JSON.parse(event.body || "{}"); } catch (e) { return J(400, { error: "Invalid JSON" }); }
  try {
    const who = await verifyEntraIdToken(body.idToken, { tenantId: process.env.ENTRA_TENANT_ID, clientId: process.env.ENTRA_CLIENT_ID, allowedDomains: allowedDomains() });
    const token = signSession({ sub: who.oid, name: who.name, upn: who.upn }, secret);
    return J(200, { ok: true, name: who.name, upn: who.upn }, { "Set-Cookie": sessionCookie(token) });
  } catch (e) {
    const notConfigured = /not configured/.test(e.message);
    if (!notConfigured) console.warn("Sign-in refused:", e.message); // detail stays in the function log, not the response
    return J(notConfigured ? 500 : 401, { error: notConfigured ? "Sign-in not configured" : "Sign-in refused" });
  }
};
