"use strict";

// Shared-secret gate for the two endpoints that serve report-hub
// (clients.js, report-data.js) rather than this app's own browser wizard.
// Those two endpoints previously had NO auth at all: clients.js publicly
// listed every configured client id, and report-data.js would run a full
// Graph pull for any clientId with zero credential of its own -- anyone who
// found the URL could pull any configured client's complete Intune/Entra
// report. See security review, 2026-09.
//
// This does NOT touch intune.js/export-docx.js/export-pdf.js -- those stay
// open, since the browser wizard's own security model is "you must already
// have that client's real Tenant ID/Client ID/Client Secret to get anything
// back", which this file has nothing to do with.
//
// Fails CLOSED: if REPORT_HUB_SHARED_KEY isn't set in this site's env vars,
// every call is rejected (misconfiguration should never silently mean "wide
// open"). Set the same value in msp-report-builder's and report-hub's
// Netlify env vars (Site settings -> Environment variables) -- one shared
// secret, not one per client.

function requireReportHubKey(event) {
  const expected = process.env.REPORT_HUB_SHARED_KEY;
  const headers = event.headers || {};
  // Netlify normalizes incoming header names to lowercase, but check both
  // just in case this is ever invoked from a test harness that doesn't.
  const provided = headers['x-report-hub-key'] || headers['X-Report-Hub-Key'];

  if (!expected) {
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: 'REPORT_HUB_SHARED_KEY is not configured on this site -- refusing all calls until it is set.' }),
    };
  }
  if (!provided || provided !== expected) {
    return {
      statusCode: 401,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: 'Missing or invalid x-report-hub-key header.' }),
    };
  }
  return null; // authorized
}

module.exports = { requireReportHubKey };
