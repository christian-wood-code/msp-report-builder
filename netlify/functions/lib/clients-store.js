/**
 * Server-side per-client Entra credential store.
 *
 * Until now, this app only ever took tenantId/clientId/clientSecret typed
 * into the browser per session (see index.html) -- no server-side storage at
 * all. report-hub needs to call this app headlessly (no human present to
 * type a secret), so each client's credentials now also live here, read from
 * one JSON-blob env var -- same pattern as google-workspace-report's
 * GOOGLE_SERVICE_ACCOUNT_KEY_JSON (a deployed Netlify function has no
 * filesystem for a committed secrets file).
 *
 * This is ADDITIVE: the existing browser flow (paste credentials, nothing
 * stored) still works completely unchanged via intune.js. This store only
 * backs the two new endpoints (clients.js, report-data.js) that report-hub
 * calls.
 *
 * INTUNE_CLIENTS_JSON shape:
 * [
 *   { "id": "acme", "displayName": "Acme Corp", "tenantId": "...", "clientId": "...", "clientSecret": "..." },
 *   ...
 * ]
 */

function loadClients() {
  const raw = process.env.INTUNE_CLIENTS_JSON;
  if (!raw) return [];
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    throw new Error('INTUNE_CLIENTS_JSON is not valid JSON: ' + e.message);
  }
  if (!Array.isArray(parsed)) throw new Error('INTUNE_CLIENTS_JSON must be a JSON array');
  return parsed;
}

function findClient(id) {
  return loadClients().find(c => c.id === id) || null;
}

// Strips credentials -- this is what the public clients.js endpoint returns.
function publicClientList() {
  return loadClients().map(c => ({ id: c.id, displayName: c.displayName }));
}

module.exports = { loadClients, findClient, publicClientList };
