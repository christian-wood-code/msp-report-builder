# MSP Report Builder

A web application that generates monthly IT health reports for client Microsoft 365 and Intune tenants, built for Integricity Technology. It pulls live data directly from the Microsoft Graph API and produces an on-screen report plus Word (.docx) and PDF exports, with month-on-month commentary once a tenant has been run in consecutive months.

**Live app:** https://integricitymsp.netlify.app

## What it does

- Pulls 50+ metrics from Microsoft Graph across six areas: Device & Asset Management, Patch Status, Security Posture, User Data (including upcoming licence renewals), SharePoint / MS Teams, and an auto-generated Security Risk Register
- Adds a **Month-on-Month Changes** section comparing against the previous month's report for the same tenant (deterministic template sentences — no AI-generated text)
- Runs as a static front end plus six Netlify serverless functions. There is no database; the only stored data is a small Netlify Blobs store of per-tenant summary figures (device counts, Secure Score and similar — no raw tenant data, no credentials) kept for up to 13 months
- Credentials (Tenant ID, Client ID, Client Secret) are entered in the browser and sent to the server per request only — never logged or stored server-side. If you save a client in the wizard, it is kept in that browser's local storage (secret in plain text) until you delete it; edit and delete buttons are on each saved client
- A **One-off report** tick box skips the month-on-month store and deletes any history already held for that tenant (use it for prospects)
- The wizard builds the report straight after **Pull data** (no review step). A collapsed *SharePoint template diagnostics* panel above the report shows the raw SharePoint template values and how each was bucketed; it is never exported. The SharePoint section shows a **Teams** count (Teams-enabled M365 Groups, one Graph call) alongside site-type counts
- Exports a fully formatted Word document, a PDF, or use your browser's print-to-PDF

## Stack

| Layer | Technology |
|---|---|
| Front end | Single HTML file — no framework, no build step |
| Back end | Six Netlify serverless functions (Node.js) |
| Data source | Microsoft Graph API v1.0 (app-only, read-only) |
| Month-on-month store | [Netlify Blobs](https://docs.netlify.com/build/data-and-storage/netlify-blobs/) (`@netlify/blobs`, pinned to an exact version) |
| Word export | [`docx`](https://www.npmjs.com/package/docx) npm package |
| PDF export | `jspdf` + `jspdf-autotable` |
| Hosting | Netlify (free tier) |

## Project structure

```
.
├── index.html                       Front-end SPA (all HTML/CSS/JS inline; builds its own on-screen report)
├── server.js                        Local dev server (node server.js)
├── netlify.toml                     Netlify config, security headers
├── package.json                     Dependencies (docx, jspdf, jspdf-autotable, @netlify/blobs)
├── .env.example                     Environment variables (see below)
├── config/clients.example.json      Shape of INTUNE_CLIENTS_JSON entries
├── scripts/add-client.js            Local helper to add a client to .env (not deployed)
├── netlify/
│   └── functions/
│       ├── intune.js                Graph API proxy — gathers all metrics, month-on-month commentary
│       ├── export-docx.js           Word document generator
│       ├── export-pdf.js            PDF generator (full mirror of the Word export)
│       ├── clients.js               Client picker list for report-hub (key-protected)
│       ├── report-data.js           Headless report data for report-hub (key-protected)
│       ├── delete-client-metrics.js Admin-only: delete a tenant's stored history (key-protected)
│       └── lib/
│           ├── clients-store.js     Reads INTUNE_CLIENTS_JSON
│           ├── report-archive.js    Netlify Blobs read/write/prune/delete for month-on-month figures
│           ├── metrics-commentary.js  Snapshot extraction + commentary sentences
│           ├── site-templates.js    SharePoint site-type classification
│           ├── require-report-hub-key.js  Shared-secret gate (clients, report-data)
│           └── require-admin-key.js       Shared-secret gate (delete-client-metrics)
├── fulllogo_*.b64, logo.b64, icon.b64   Logo assets (base64), used in the report cover
└── docs/
    ├── Intune_Report_Architecture.docx   System architecture and design decisions
    ├── architecture-diagram.png
    ├── User_Guide.docx / .pdf            How to use the app day to day
    ├── Setup_Guide.docx / .pdf           Per-tenant Entra ID app registration steps
    ├── IT_Health_Snapshot_Setup_Instructions.pdf       Client-facing setup steps (one-off snapshot)
    └── IT_Health_Snapshot_Data_Access_Authorization.pdf  What is accessed, retention, disclaimer and sign-off
```

## Setup — per client tenant

Before the report builder can pull data for a client, an app registration must be created in **that client's** Microsoft Entra ID tenant with 12 read-only Graph API permissions and admin consent granted. Full step-by-step instructions are in `docs/Setup_Guide.docx`.

Once registered, you need three values per client:

- Tenant ID
- Client ID
- Client Secret

These are entered directly into the app at runtime. Nothing per-client is configured in code. (Only the optional report-hub integration stores client credentials server-side, via `INTUNE_CLIENTS_JSON`.)

## Environment variables

Set in Netlify (Site settings → Environment variables). The browser wizard works with none of them set.

| Variable | Purpose |
|---|---|
| `REPORT_HUB_SHARED_KEY` | Shared secret report-hub sends as `x-report-hub-key` to `clients` and `report-data`. Every call is refused while unset |
| `INTUNE_CLIENTS_JSON` | JSON array of client credentials used by report-hub (see `config/clients.example.json`) |
| `ARCHIVE_ADMIN_KEY` | Optional. Enables `delete-client-metrics` (sent as `x-admin-key`). Every call is refused while unset |

Netlify Blobs (the month-on-month store) needs no variables or credentials — access is scoped to the site automatically.

## Deployment

Pushing to `main` on GitHub deploys automatically (the repo is linked to the Netlify site). No build step runs beyond `npm install`.

> **A push is not proof of a deploy.** A failed Netlify build leaves the site serving the last good version. After pushing, confirm the deploy reaches `ready` (`netlify api listSiteDeploys --data '{"site_id":"<id>"}'`, or the Deploys tab). `package-lock.json` is gitignored, so Netlify installs from `package.json` alone — which is why `@netlify/blobs` is pinned to an exact version: a caret range on the newest patch once failed to resolve (`ETARGET`) and silently blocked two deploys.

Stage specific files when committing (`git add <files>`, then check `git status`) rather than `git add -A`, so stray local files never get committed.

Manual alternatives:

### Netlify CLI

```bash
npm install -g netlify-cli
netlify login
npm install
netlify deploy --prod
```

### Netlify UI drag-and-drop

Drag the project folder onto the **Deploys** tab in the Netlify dashboard.

> **Important:** the drag-and-drop method does **not** run `npm install` automatically. If deploying this way, run `npm install` locally first and include the resulting `node_modules` folder in the zip you drop — otherwise the export functions will fail at runtime with `Cannot find module 'docx'` (or `jspdf`, `@netlify/blobs`).

## Local development

```bash
npm install
node server.js          # or: netlify dev
```

`node server.js` serves `index.html` and the functions on http://localhost:8791. Locally the Netlify Blobs store is unavailable, so month-on-month commentary silently skips (it never blocks a report).

## Documentation

| Document | Covers |
|---|---|
| `docs/Intune_Report_Architecture.docx` | System design, data flow, security model, metrics collected, key implementation decisions, known limitations |
| `docs/User_Guide.docx` / `.pdf` | Day-to-day usage — generating a report, understanding each section, troubleshooting |
| `docs/Setup_Guide.docx` / `.pdf` | Per-tenant Microsoft Entra ID app registration and Graph API permission setup |
| `docs/IT_Health_Snapshot_Setup_Instructions.pdf` | Client-facing, step-by-step setup for a one-off snapshot |
| `docs/IT_Health_Snapshot_Data_Access_Authorization.pdf` | Plain-language access, retention and disclaimer, with sign-off |
| `DEPLOY.md` | Headless GitHub + Netlify deploy runbook, written for a Claude Code session with local git/netlify CLI access |

## Known limitations

- Rate limiting is in-memory and resets on Netlify cold starts — not a durable limit; only `intune` and `export-docx` have it
- Netlify free tier has a 10-second function timeout; very large tenants may see the sign-in log query fall back to a shorter window
- Antivirus detection is Windows Defender-aware only; third-party AV is inferred, not identified by name
- macOS devices are excluded from the encryption count — Intune does not reliably report FileVault state
- SharePoint figures come from Microsoft's usage report, which runs 24–48 hours behind and includes hidden Teams-channel and system sites, so totals differ from the SharePoint admin center's Active sites list. Storage is in binary gigabytes, matching the admin center
- Usage-report entries with no template, no recorded activity and 0 GB are left out of every SharePoint count (they don't appear in the admin center's Active sites); the number excluded is shown in the diagnostics panel
- Licence renewal Monthly vs Annual/Multi-year labels are inferred from days to renewal (Graph does not expose billing frequency)
- The manual-entry sections (Executive Summary, Ticketing & Support, Recommendations & Opportunities) are not currently collected, because the wizard's Review step was removed; they are omitted from the report
- Saved clients' secrets sit in browser local storage in plain text on that device
- `npm audit` reports advisories in `jspdf`'s `dompurify` dependency; the fix is a breaking `jspdf` major upgrade (not yet done)
- Netlify's built-in site password protection blocks serverless function calls — use an in-app password gate or Netlify Identity instead if access control is needed

## License

Proprietary — for internal use by Integricity Technology and authorised developers only.
