"use strict";

// Internal admin tool -- deletes all archived month-on-month metrics for one
// tenant, without needing to run another report first. For when a client
// asks for their history to be removed after the fact (e.g. by phone or
// email), rather than via the "don't retain metrics" checkbox at report
// generation time (which handles the common case: opting out up front).
//
// Not linked from index.html or any public UI -- invoke manually, e.g.:
//   curl -X POST https://integricitymsp.netlify.app/.netlify/functions/delete-client-metrics \
//     -H "x-admin-key: <ARCHIVE_ADMIN_KEY value>" -H "Content-Type: application/json" \
//     -d '{"tenantId":"<their Entra tenant/directory ID>"}'
//
// Deliberately gated by a bare admin key rather than that tenant's own Entra
// credentials: by the time a deletion request like this comes in, the
// client has typically already deleted the app registration used to prove
// tenant ownership (see the Setup Instructions doc's cleanup step), so that
// credential usually no longer exists to re-check.

const { requireAdminKey } = require("./lib/require-admin-key");
const { deleteAllMetrics } = require("./lib/report-archive");

const CORS = { "Access-Control-Allow-Origin": "*", "Content-Type": "application/json" };

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers: CORS, body: "" };
  if (event.httpMethod !== "POST") return { statusCode: 405, headers: CORS, body: JSON.stringify({ error: "Method not allowed" }) };

  const denied = requireAdminKey(event);
  if (denied) return { ...denied, headers: { ...CORS, ...denied.headers } };

  let body;
  try { body = JSON.parse(event.body); }
  catch { return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: "Invalid JSON" }) }; }

  const { tenantId } = body;
  const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (!tenantId || !GUID_RE.test(tenantId)) {
    return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: "tenantId (a GUID) is required" }) };
  }

  const deletedCount = await deleteAllMetrics(tenantId);
  return { statusCode: 200, headers: CORS, body: JSON.stringify({ tenantId, deletedCount }) };
};
