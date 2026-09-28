"use strict";

// Talks to Integricity's OWN SharePoint (a separate Graph identity from the
// client's tenant -- an app registration in Integricity's tenant, scoped to
// exactly one site via Sites.Selected) to store/retrieve small,
// metrics-only snapshots for month-on-month trend commentary. Every
// exported function is soft-fail by design: a problem here (missing env
// vars, an outage, a permission lapse) must never break report generation,
// only silently skip commentary for that run. See
// docs/IT_Health_Snapshot_Data_Access_Authorization.pdf for what this
// stores and why, and the Architecture doc for the full setup steps.
//
// Retention: rolling 13 months per tenant -- enough for one full
// year-over-year comparison plus a spare month. Pruning runs best-effort
// after each successful save; a failed prune never fails the save itself.

const ARCHIVE_TIMEOUT_MS = 5000;

function withTimeout(promise, ms, fallback) {
  return Promise.race([promise, new Promise(resolve => setTimeout(() => resolve(fallback), ms))]);
}

let cachedToken = null, cachedTokenExp = 0;

async function getArchiveToken() {
  const now = Date.now();
  if (cachedToken && now < cachedTokenExp - 60000) return cachedToken;
  const tenantId = process.env.ARCHIVE_TENANT_ID;
  const clientId = process.env.ARCHIVE_CLIENT_ID;
  const clientSecret = process.env.ARCHIVE_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) throw new Error("Archive credentials not configured");
  const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "client_credentials", client_id: clientId, client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
    }),
  });
  const d = await res.json();
  if (d.error) throw new Error(d.error_description || d.error);
  cachedToken = d.access_token;
  cachedTokenExp = now + (d.expires_in || 3600) * 1000;
  return cachedToken;
}

function fileName(period) { return `${period}-metrics.json`; } // period = "YYYY-MM"

function driveBase() {
  const siteId = process.env.ARCHIVE_SITE_ID, driveId = process.env.ARCHIVE_DRIVE_ID;
  if (!siteId || !driveId) return null;
  return `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}`;
}

// Returns the prior calendar-month key for a "YYYY-MM" period string.
function prevPeriod(period) {
  const [y, m] = period.split("-").map(Number);
  const d = new Date(Date.UTC(y, m - 2, 1)); // m is 1-indexed; -2 = previous month, 0-indexed
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
}

async function getLastMetrics(tenantId, period) {
  const base = driveBase();
  if (!base || !tenantId) return null;
  try {
    const run = (async () => {
      const token = await getArchiveToken();
      const path = `/${tenantId}/${fileName(period)}`;
      const res = await fetch(`${base}/root:${path}:/content`, { headers: { Authorization: `Bearer ${token}` } });
      if (res.status === 404) return null;
      if (!res.ok) return null;
      return await res.json();
    })();
    return await withTimeout(run, ARCHIVE_TIMEOUT_MS, null);
  } catch (e) {
    return null; // soft-fail -- no prior period available, commentary is just omitted
  }
}

async function saveMetrics(tenantId, period, metrics) {
  const base = driveBase();
  if (!base || !tenantId) return false;
  try {
    const run = (async () => {
      const token = await getArchiveToken();
      const path = `/${tenantId}/${fileName(period)}`;
      const res = await fetch(`${base}/root:${path}:/content`, {
        method: "PUT",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify(metrics),
      });
      if (res.ok) pruneOldMetrics(tenantId, base, token).catch(() => {}); // best-effort, not awaited past this
      return res.ok;
    })();
    return await withTimeout(run, ARCHIVE_TIMEOUT_MS, false);
  } catch (e) {
    return false;
  }
}

async function pruneOldMetrics(tenantId, base, token) {
  const res = await fetch(`${base}/root:/${tenantId}:/children`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) return;
  const d = await res.json();
  const cutoff = new Date();
  cutoff.setUTCMonth(cutoff.getUTCMonth() - 13);
  for (const item of (d.value || [])) {
    const m = /^(\d{4})-(\d{2})-metrics\.json$/.exec(item.name || "");
    if (!m) continue;
    const fileDate = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, 1));
    if (fileDate < cutoff) {
      await fetch(`${base}/items/${item.id}`, { method: "DELETE", headers: { Authorization: `Bearer ${token}` } }).catch(() => {});
    }
  }
}

module.exports = { getLastMetrics, saveMetrics, prevPeriod };
