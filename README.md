# MSP Report Builder

A web application that generates monthly IT health reports for client Microsoft 365 and Intune tenants, built for Integricity Technology. It pulls live data directly from the Microsoft Graph API and produces an on-screen report plus a Word (.docx) export.

**Live app:** https://integricitymsp.netlify.app

## What it does

- Pulls 52 metrics from Microsoft Graph across six areas: Device & Asset Management, Patch Status, Security Posture, User Data, SharePoint / MS Teams, and an auto-generated Security Risk Register
- Runs entirely as a static front end plus two Netlify serverless functions — no database, no persistent server
- Credentials (Tenant ID, Client ID, Client Secret) are entered per session and held in browser memory only — never stored server-side
- Exports a fully formatted Word document, or use your browser's print-to-PDF for a PDF copy

## Stack

| Layer | Technology |
|---|---|
| Front end | Single HTML file — no framework, no build step |
| Back end | Two Netlify serverless functions (Node.js) |
| Data source | Microsoft Graph API v1.0 (app-only, read-only) |
| Word export | [`docx`](https://www.npmjs.com/package/docx) npm package |
| Hosting | Netlify (free tier) |

## Project structure

```
.
├── index.html                       Front-end SPA (all HTML/CSS/JS inline)
├── netlify.toml                     Netlify config, security headers
├── package.json                     Declares the `docx` dependency
├── netlify/
│   └── functions/
│       ├── intune.js                Graph API proxy — gathers all 52 metrics
│       └── export-docx.js           Word document generator
├── fulllogo_transparent.b64         Logo assets (base64), used in the report cover
├── fulllogo_header.b64
├── fulllogo_white.b64
├── logo.b64
├── icon.b64
└── docs/
    ├── Architecture.docx            System architecture and design decisions
    ├── User_Guide.docx              How to use the app day to day
    ├── Setup_Guide.docx             Per-tenant Entra ID app registration steps
    └── architecture-diagram.png
```

## Setup — per client tenant

Before the report builder can pull data for a client, an app registration must be created in **that client's** Microsoft Entra ID tenant with 12 read-only Graph API permissions and admin consent granted. Full step-by-step instructions are in `docs/Setup_Guide.docx`.

Once registered, you need three values per client:

- Tenant ID
- Client ID
- Client Secret

These are entered directly into the app at runtime — nothing is configured in code or environment variables per client.

## Deployment

No build step is required. Two deployment options:

### Option A — Netlify CLI (recommended)

```bash
npm install -g netlify-cli
netlify login
npm install
netlify deploy --prod
```

### Option B — Netlify UI drag-and-drop

Drag the project folder onto the **Deploys** tab in the Netlify dashboard.

> **Important:** the drag-and-drop method does **not** run `npm install` automatically. If deploying this way, run `npm install` locally first and include the resulting `node_modules` folder in the zip you drop — otherwise the `export-docx` function will fail at runtime with `Cannot find module 'docx'`.

## Local development

```bash
npm install -g netlify-cli
npm install
netlify dev
```

This serves `index.html` and both functions locally with hot reload.

## Documentation

| Document | Covers |
|---|---|
| `docs/Architecture.docx` | System design, data flow, security model, all 52 metrics, key implementation decisions |
| `docs/User_Guide.docx` | Day-to-day usage — generating a report, understanding each section, troubleshooting |
| `docs/Setup_Guide.docx` | Per-tenant Microsoft Entra ID app registration and Graph API permission setup |
| `DEPLOY.md` | Headless GitHub + Netlify deploy runbook, written for a Claude Code session with local git/netlify CLI access |

## Known limitations

- Rate limiting is in-memory and resets on Netlify cold starts — not a durable limit
- Netlify free tier has a 10-second function timeout; very large tenants may see the sign-in log query fall back to a shorter window
- Antivirus detection is Windows Defender-aware only; third-party AV is inferred, not identified by name
- macOS devices are excluded from the encryption count — Intune does not reliably report FileVault state
- Netlify's built-in site password protection blocks serverless function calls — use an in-app password gate or Netlify Identity instead if access control is needed

## License

Proprietary — for internal use by Integricity Technology and authorised developers only.
