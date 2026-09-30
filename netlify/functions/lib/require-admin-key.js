"use strict";

// Shared-secret gate for internal admin-only endpoints (currently just
// delete-client-metrics.js) -- these take a bare tenantId with no proof the
// caller controls that tenant, so they must never be reachable by the
// public wizard. This is a manual/internal tool: Christian runs it via curl
// when a client asks for their metrics history to be deleted outside the
// normal report-generation flow, not something linked from index.html.
//
// Fails CLOSED: if ARCHIVE_ADMIN_KEY isn't set, every call is rejected.

function requireAdminKey(event) {
  const expected = process.env.ARCHIVE_ADMIN_KEY;
  const headers = event.headers || {};
  const provided = headers['x-admin-key'] || headers['X-Admin-Key'];

  if (!expected) {
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: 'ARCHIVE_ADMIN_KEY is not configured on this site -- refusing all calls until it is set.' }),
    };
  }
  if (!provided || provided !== expected) {
    return {
      statusCode: 401,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: 'Missing or invalid x-admin-key header.' }),
    };
  }
  return null; // authorized
}

module.exports = { requireAdminKey };
