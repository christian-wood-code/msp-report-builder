"use strict";

// Headless report-data endpoint for report-hub -- looks up a stored client's
// Entra credentials (see lib/clients-store.js) and calls the existing
// intune.js handler internally with them, so intune.js itself (used live by
// this app's own browser UI, which still pastes credentials per session)
// needed zero changes.
//
// Query params: clientId (our own internal id, required), from, to
// (YYYY-MM-DD, optional -- defaults to intune.js's own rolling-30-days
// fallback if omitted).

const { findClient } = require('./lib/clients-store');
const intuneHandler = require('./intune').handler;

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Content-Type': 'application/json',
};

exports.handler = async (event) => {
  const q = event.queryStringParameters || {};
  const id = q.clientId;
  if (!id) return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: 'clientId is required' }) };

  let client;
  try {
    client = findClient(id);
  } catch (err) {
    return { statusCode: 500, headers: CORS, body: JSON.stringify({ error: err.message }) };
  }
  if (!client) return { statusCode: 404, headers: CORS, body: JSON.stringify({ error: `No stored credentials for clientId "${id}"` }) };

  const innerEvent = {
    httpMethod: 'POST',
    headers: event.headers,
    body: JSON.stringify({
      tenantId: client.tenantId,
      clientId: client.clientId,
      clientSecret: client.clientSecret,
      reportFrom: q.from || undefined,
      reportTo: q.to || undefined,
    }),
  };

  const result = await intuneHandler(innerEvent);
  // intune.js's own CORS/content-type headers already cover this; pass its
  // response straight through (including its status code on auth/graph
  // failures) rather than re-wrapping it.
  return result;
};
