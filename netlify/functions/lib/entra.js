"use strict";
const { createPublicKey, createVerify } = require("crypto");

// Verifies a Microsoft Entra ID token (from MSAL.js in the browser) server-side. Single tenant:
// signature, issuer, tenant (tid), audience (our app's client id), time window and a staff-account
// check are ALL pinned; anything unexpected throws. Never trust the browser's claim about who it is.
// Ported from the MSP Client Health Dashboard (same Entra app registration).

const b64u = s => Buffer.from(s, "base64url");
const parse = s => JSON.parse(b64u(s).toString("utf8"));

let jwksCache = { at: 0, tenant: "", keys: [] };
async function fetchKeys(tenantId, force = false) {
  if (!force && jwksCache.tenant === tenantId && Date.now() - jwksCache.at < 3600000 && jwksCache.keys.length) return jwksCache.keys;
  const r = await fetch(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`);
  if (!r.ok) throw new Error("Could not fetch Microsoft signing keys");
  const { keys } = await r.json();
  jwksCache = { at: Date.now(), tenant: tenantId, keys };
  return keys;
}

async function verifyEntraIdToken(token, { tenantId, clientId, allowedDomains, getKeys = fetchKeys, now = Date.now() }) {
  if (!tenantId || !clientId || !allowedDomains || !allowedDomains.length) throw new Error("Sign-in not configured");
  if (typeof token !== "string" || token.length > 8192) throw new Error("Bad token");
  const parts = token.split(".");
  if (parts.length !== 3) throw new Error("Bad token");
  const header = parse(parts[0]);
  if (header.alg !== "RS256" || !header.kid) throw new Error("Bad token header");

  let keys = await getKeys(tenantId);
  let jwk = keys.find(k => k.kid === header.kid);
  if (!jwk) { keys = await getKeys(tenantId, true); jwk = keys.find(k => k.kid === header.kid); } // key rotation
  if (!jwk) throw new Error("Unknown signing key");
  const ok = createVerify("RSA-SHA256").update(parts[0] + "." + parts[1]).verify(createPublicKey({ key: jwk, format: "jwk" }), b64u(parts[2]));
  if (!ok) throw new Error("Bad signature");

  const c = parse(parts[1]);
  if (c.iss !== `https://login.microsoftonline.com/${tenantId}/v2.0`) throw new Error("Wrong issuer");
  if (c.tid !== tenantId) throw new Error("Wrong tenant");
  if (c.aud !== clientId) throw new Error("Wrong audience");
  const t = Math.floor(now / 1000);
  if (!c.exp || c.exp < t) throw new Error("Token expired");
  if (c.nbf && c.nbf > t + 60) throw new Error("Token not yet valid");
  if (!c.iat || c.iat < t - 300) throw new Error("Token too old, sign in again"); // limits replay of a captured token

  const upn = String(c.preferred_username || "").toLowerCase();
  if (!upn || upn.includes("#ext#")) throw new Error("Guest accounts are not allowed");
  const domain = upn.split("@")[1] || "";
  if (!allowedDomains.includes(domain)) throw new Error("Account domain not allowed");
  if (!c.oid) throw new Error("Token has no object id");
  return { oid: c.oid, name: c.name || upn, upn };
}

module.exports = { verifyEntraIdToken, fetchKeys };
