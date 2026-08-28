#!/usr/bin/env node
// Safely add (or update) one client entry in .env's INTUNE_CLIENTS_JSON --
// the repeatable process for onboarding each new client's credentials,
// since every client's Entra tenant/app registration is different and
// hand-editing a single-line JSON blob by hand is easy to get wrong.
//
// Usage:
//   node scripts/add-client.js --id acme --name "Acme Corp" \
//     [--tenant <guid>] [--client <guid>] [--secret <value>]
//
// Any of --tenant/--client/--secret can be omitted -- the entry is written
// with a "REPLACE_ME" placeholder so you can safely paste this command
// (and the whole session transcript) without a real secret ever appearing in
// it, then fill the placeholder(s) in directly by editing .env yourself.
// Re-running with the same --id updates that entry in place rather than
// duplicating it.
//
// After adding a client here for local testing, remember: this only updates
// your LOCAL .env. Production (Netlify) has its own separate copy of
// INTUNE_CLIENTS_JSON under Site settings -> Environment variables --
// this script prints the final JSON at the end specifically so you can
// paste it there too when you're ready to deploy.

const fs = require('fs');
const path = require('path');

const ENV_PATH = path.join(__dirname, '..', '.env');
const PLACEHOLDER = 'REPLACE_ME';

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith('--')) { out[a.slice(2)] = argv[i + 1]; i++; }
  }
  return out;
}

function loadEnvLines() {
  if (!fs.existsSync(ENV_PATH)) return [];
  return fs.readFileSync(ENV_PATH, 'utf8').split('\n');
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.id || !args.name) {
    console.error('Usage: node scripts/add-client.js --id <short-id> --name "<Display Name>" [--tenant <guid>] [--client <guid>] [--secret <value>]');
    process.exit(1);
  }
  if (!/^[a-z0-9-]+$/.test(args.id)) {
    console.error('--id must be lowercase letters, numbers, and hyphens only (e.g. "acme-corp").');
    process.exit(1);
  }

  const lines = loadEnvLines();
  const lineIdx = lines.findIndex(l => l.trim().startsWith('INTUNE_CLIENTS_JSON='));
  let clients = [];
  if (lineIdx !== -1) {
    const raw = lines[lineIdx].slice(lines[lineIdx].indexOf('=') + 1).trim();
    try { clients = raw ? JSON.parse(raw) : []; }
    catch (e) { console.error('Existing INTUNE_CLIENTS_JSON in .env is not valid JSON -- fix that by hand first:', e.message); process.exit(1); }
  }

  const entry = {
    id: args.id,
    displayName: args.name,
    tenantId: args.tenant || PLACEHOLDER,
    clientId: args.client || PLACEHOLDER,
    clientSecret: args.secret || PLACEHOLDER,
  };

  const existingIdx = clients.findIndex(c => c.id === args.id);
  if (existingIdx !== -1) {
    clients[existingIdx] = { ...clients[existingIdx], ...entry };
    console.log(`Updated existing entry for "${args.id}".`);
  } else {
    clients.push(entry);
    console.log(`Added new entry for "${args.id}".`);
  }

  const newLine = `INTUNE_CLIENTS_JSON=${JSON.stringify(clients)}`;
  if (lineIdx !== -1) lines[lineIdx] = newLine;
  else lines.push(newLine);

  fs.writeFileSync(ENV_PATH, lines.join('\n'));

  const placeholders = Object.entries(entry).filter(([, v]) => v === PLACEHOLDER).map(([k]) => k);
  if (placeholders.length) {
    console.log(`\n${args.id} still has placeholder value(s) for: ${placeholders.join(', ')}.`);
    console.log(`Open .env and replace ${PLACEHOLDER} for "${args.id}" with the real value(s) from that client's Entra app registration (see docs/Setup_Guide.docx) before testing.`);
  } else {
    console.log(`\n${args.id} is fully configured with real credentials.`);
  }

  console.log('\nFull current INTUNE_CLIENTS_JSON (paste into Netlify env vars when deploying):');
  console.log(JSON.stringify(clients));
}

main();
