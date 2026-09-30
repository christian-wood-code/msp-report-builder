"use strict";

// Stores/retrieves small, metrics-only snapshots (see lib/metrics-commentary.js)
// for month-on-month trend commentary, using Netlify Blobs -- scoped
// automatically to this site by the Functions runtime, no separate
// credentials to create, store, or rotate (unlike the SharePoint-based
// design this replaced: no app registration, no client secret, no external
// write-capable identity to leak). Blobs are encrypted at rest and in
// transit, and reachable only through this site. See
// docs/IT_Health_Snapshot_Data_Access_Authorization.pdf for what this
// stores and why.
//
// Every exported function is soft-fail by design: a problem here (an
// outage, a bad key, anything) must never break report generation, only
// silently skip commentary for that run.
//
// Keys: "{tenantId}/{yyyy-mm}" -- tenantId is the client's own Entra
// directory ID (always present on every report call, wizard or report-hub,
// so this works for any repeat run without needing our internal
// clients-store.js identifiers). Retention: rolling 13 months per tenant,
// pruned best-effort after each successful save.

const { getStore } = require("@netlify/blobs");

const STORE_NAME = "msp-report-metrics";
const RETENTION_MONTHS = 13;

function store() {
  return getStore(STORE_NAME);
}

function key(tenantId, period) { return `${tenantId}/${period}`; } // period = "YYYY-MM"

// Returns the prior calendar-month key for a "YYYY-MM" period string.
function prevPeriod(period) {
  const [y, m] = period.split("-").map(Number);
  const d = new Date(Date.UTC(y, m - 2, 1)); // m is 1-indexed; -2 = previous month, 0-indexed
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
}

async function getLastMetrics(tenantId, period) {
  if (!tenantId) return null;
  try {
    return await store().get(key(tenantId, period), { type: "json" });
  } catch (e) {
    return null; // soft-fail -- no prior period available, commentary is just omitted
  }
}

async function saveMetrics(tenantId, period, metrics) {
  if (!tenantId) return false;
  try {
    await store().setJSON(key(tenantId, period), metrics);
    pruneOldMetrics(tenantId).catch(() => {}); // best-effort, not awaited past this
    return true;
  } catch (e) {
    return false;
  }
}

async function pruneOldMetrics(tenantId) {
  const { blobs } = await store().list({ prefix: `${tenantId}/` });
  const cutoff = new Date();
  cutoff.setUTCMonth(cutoff.getUTCMonth() - RETENTION_MONTHS);
  for (const b of blobs) {
    const m = /\/(\d{4})-(\d{2})$/.exec(b.key);
    if (!m) continue;
    const fileDate = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, 1));
    if (fileDate < cutoff) await store().delete(b.key).catch(() => {});
  }
}

// Deletes every stored snapshot for a tenant -- used both for the "one-off,
// don't retain" wizard option (delete anything from before too, not just
// skip saving this run) and for the standalone admin deletion endpoint.
// Returns the number of entries deleted.
async function deleteAllMetrics(tenantId) {
  if (!tenantId) return 0;
  try {
    const { blobs } = await store().list({ prefix: `${tenantId}/` });
    for (const b of blobs) await store().delete(b.key).catch(() => {});
    return blobs.length;
  } catch (e) {
    return 0;
  }
}

module.exports = { getLastMetrics, saveMetrics, deleteAllMetrics, prevPeriod };
