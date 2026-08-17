"use strict";

// Public client list for report-hub's picker -- strips credentials.
// Only clients someone has added to INTUNE_CLIENTS_JSON (see
// config/clients.example.json and lib/clients-store.js) appear here; the
// existing paste-your-own-credentials browser flow (index.html) is
// unaffected and doesn't require a client to be listed here at all.

const { publicClientList } = require('./lib/clients-store');

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Content-Type': 'application/json',
};

exports.handler = async () => {
  try {
    return { statusCode: 200, headers: CORS, body: JSON.stringify({ clients: publicClientList() }) };
  } catch (err) {
    return { statusCode: 500, headers: CORS, body: JSON.stringify({ error: err.message }) };
  }
};
