"use strict";

const { safeEqual } = require('./safe-equal');
const { readSession } = require('./staff-session');

// Access gate for the three endpoints the browser wizard uses (intune, export-docx, export-pdf).
// Two ways in:
//   staff session     httpOnly cookie issued by auth-session.js after Microsoft (Entra) sign-in
//   x-report-hub-key  report-hub server-to-server; value must equal REPORT_HUB_SHARED_KEY
// Fails CLOSED: with sign-in unconfigured and no hub key set, every call is refused.
// Temporary switch: APP_ACCESS_DISABLED=true leaves the endpoints open (remove the variable to enforce sign-in).
function requireAppAccess(event, corsHeaders) {
  if (process.env.APP_ACCESS_DISABLED === 'true') return null;
  const h = event.headers || {};
  const hubKey = h['x-report-hub-key'] || h['X-Report-Hub-Key'];
  const expectedHub = process.env.REPORT_HUB_SHARED_KEY;
  const headers = { ...(corsHeaders || {}), 'Content-Type': 'application/json' };

  if (hubKey && expectedHub && safeEqual(hubKey, expectedHub)) return null;

  const s = readSession(event);
  if (s.claims) return null;

  if (s.error === 'not-configured' && !expectedHub) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: 'Sign-in is not configured on this site - refusing all calls until it is.' }) };
  }
  return { statusCode: 401, headers, body: JSON.stringify({ error: 'Sign in with your Integricity Microsoft account.', code: 'SIGN_IN' }) };
}

module.exports = { requireAppAccess };
