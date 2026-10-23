"use strict";

const { safeEqual } = require('./safe-equal');

// Access gate for the three endpoints the browser wizard uses (intune, export-docx, export-pdf).
// They used to be completely open: anyone who found the URL could use the site as a Microsoft Graph
// proxy and burn function time / credits. Two ways in:
//   x-app-key         staff browser; value must equal APP_ACCESS_KEY (entered once, kept in localStorage)
//   x-report-hub-key  report-hub server-to-server; value must equal REPORT_HUB_SHARED_KEY
// Fails CLOSED: with neither env var set every call is refused. This is a stop-gap until Microsoft
// sign-in (Entra) replaces it.
function requireAppAccess(event, corsHeaders) {
  const h = event.headers || {};
  const appKey = h['x-app-key'] || h['X-App-Key'];
  const hubKey = h['x-report-hub-key'] || h['X-Report-Hub-Key'];
  const expectedApp = process.env.APP_ACCESS_KEY;
  const expectedHub = process.env.REPORT_HUB_SHARED_KEY;
  const headers = { ...(corsHeaders || {}), 'Content-Type': 'application/json' };

  if (!expectedApp && !expectedHub) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: 'APP_ACCESS_KEY is not configured on this site - refusing all calls until it is set.' }) };
  }
  if (appKey && expectedApp && safeEqual(appKey, expectedApp)) return null;
  if (hubKey && expectedHub && safeEqual(hubKey, expectedHub)) return null;
  return { statusCode: 401, headers, body: JSON.stringify({ error: 'Access key missing or incorrect.', code: 'ACCESS_KEY' }) };
}

module.exports = { requireAppAccess };
