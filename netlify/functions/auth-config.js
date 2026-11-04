"use strict";
// Public, non-secret settings the browser needs to start Microsoft sign-in (a tenant/app id is not a
// secret). `required` tells the page whether to insist on sign-in before use.
exports.handler = async () => {
  const tenantId = process.env.ENTRA_TENANT_ID, clientId = process.env.ENTRA_CLIENT_ID;
  const required = process.env.APP_ACCESS_DISABLED !== "true";
  const body = (tenantId && clientId) ? { configured: true, tenantId, clientId, required } : { configured: false, required };
  return { statusCode: 200, headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }, body: JSON.stringify(body) };
};
