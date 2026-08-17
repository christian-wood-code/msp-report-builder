# Deploy runbook: headless GitHub + Netlify setup

How to get this project from local code to a live Netlify site with as
few interactive browser steps as possible — written for a future Claude
Code session doing this deploy (or a redeploy), not for a human reader.
Adapted from the general pattern documented in
`headless_github_netlify_deploy_pattern.md`, applied specifically to this
project.

This project has **no scheduled functions** and **no build step** — it is
simpler than the pattern this runbook is adapted from. Read the whole
thing before running anything; the one step that can't be automated is
called out explicitly near the end.

## Preconditions this runbook assumes

- `netlify` CLI is installed and already authenticated (`netlify status`
  shows a logged-in user). If not, that one login step needs the human —
  `netlify login` opens a browser. Everything after login is scriptable.
- `git` is configured with a working credential helper that already has
  a cached GitHub token from prior use on this machine (Windows: Git
  Credential Manager). No `gh` CLI required.
- Node.js and `npm` are available (`node --check` and `npm install` are
  used below).

## Project facts (fill these in / confirm before running commands)

| Item | Value |
|---|---|
| Project name | Integricity MSP Report Builder |
| Existing live site | `integricitymsp.netlify.app` |
| Netlify Site ID | *(look up with `netlify status` if a `.netlify` folder already exists, or `netlify sites:list`)* |
| Functions directory | `netlify/functions` |
| Publish directory | `.` (repo root — `index.html` sits at the top level) |
| Build command | none — no build step |
| Runtime dependency | `docx` (declared in `package.json`) |
| Scheduled functions | none in this project |
| Env vars required | none — credentials are entered per-session by the user in the browser, not stored as Netlify env vars |

## Step 1 — Get a GitHub token without `gh`

```bash
TOKEN=$(printf 'protocol=https\nhost=github.com\n\n' | git credential fill | grep '^password=' | cut -d= -f2)
echo ${TOKEN:0:4}   # should print gho_
```

If this doesn't print `gho_`, stop and tell the user — it means there is
no cached GitHub credential on this machine and they need to run one
interactive `git push` (or `gh auth login`) first, once, before this
runbook can proceed headlessly.

## Step 2 — Create the GitHub repo via API (skip if it already exists)

```bash
curl -s -X POST https://api.github.com/user/repos \
  -H "Authorization: token $TOKEN" \
  -H "Accept: application/vnd.github+json" \
  -d '{"name":"msp-report-builder","private":true,"description":"Integricity MSP Report Builder — monthly Intune/M365 client reports"}' \
  | grep -E '"full_name"|"html_url"|"message"'
```

Grep for `"message"` too — a name collision or scope problem shows up
there. If it 403s, the cached token may have narrower scope than
expected; surface that to the user rather than guessing around it.

## Step 3 — `git init`, commit, push

```bash
cd /path/to/unzipped/msp-report-builder-github

# .gitignore already excludes node_modules, .netlify, .env — verify, don't assume
cat .gitignore

git init -q
git add -A
git status --short   # eyeball this — node_modules should NOT appear
git commit -q -m "Initial commit"
git branch -M main
git remote add origin https://github.com/OWNER/msp-report-builder.git
git push -u origin main
```

The push authenticates via the same cached credential helper from
Step 1 — no need to embed the token in the remote URL.

## Step 4 — Link to the existing Netlify site (not a new one)

This project already has a live site at `integricitymsp.netlify.app`.
Don't create a new site with `netlify sites:create` — that gives a new
random URL and abandons the existing one. Instead, link the local folder
to the **existing** site:

```bash
netlify link
# When prompted, choose "Use current git remote" or search by name for
# "integricitymsp" and select it.
```

Verify it linked to the right site before continuing:

```bash
netlify status
# Confirm the "Current project" line shows integricitymsp,
# not a new/different site.
```

## Step 5 — Install dependencies and do a manual verification deploy first

Before touching the Git link (next step, which is where things can go
sideways), confirm the code itself is deployable with a plain CLI
deploy — this is the same method that has worked for every previous
build of this project:

```bash
npm install
node --check netlify/functions/intune.js
node --check netlify/functions/export-docx.js
netlify deploy --prod
```

Verify the live functions actually respond (expect an auth error, not a
parse error — a parse error like `Unexpected token '<'` means something
is broken, an auth error means the function loaded correctly and is
just waiting for real credentials):

```bash
curl -s -X POST https://integricitymsp.netlify.app/.netlify/functions/intune \
  -H "Content-Type: application/json" \
  -d '{"tenantId":"test","clientId":"test","clientSecret":"test","reportFrom":"2026-01-01","reportTo":"2026-01-31"}'
```

If this step works, the manual-deploy path is healthy and nothing is
broken. Only proceed to Step 6 if the human actually wants push-to-deploy
from GitHub going forward — it is not required for the site to work.

## Step 6 — The one thing that doesn't work headlessly: Git-linked continuous deployment

**Do not spend time trying to script this.** Linking a GitHub repo to a
Netlify site for automatic deploys requires Netlify to install an SSH
deploy key on the GitHub repo itself, and that installation only happens
through the dashboard's interactive "Import from Git" / "Link repository"
flow — there is no API-only way to complete it.

Symptom if you try anyway via `netlify api updateSite` with a `repo`
block: the call succeeds (200, no error) and looks linked, but the next
auto-triggered build fails with something like:

```
Host key verification failed. fatal: Could not read from remote repository.
```

**Tell the human plainly:** this is a one-time manual step they need to
do themselves —

1. Go to **app.netlify.com** → the `integricitymsp` site
2. **Project configuration → Build & deploy → Continuous deployment**
3. Click **Link repository** (wording may shift slightly over time)
4. Choose GitHub, authorise, select the `msp-report-builder` repo
5. Confirm — Netlify should auto-detect `publish = "."` and
   `functions = "netlify/functions"` from `netlify.toml`

After that one click, every `git push` to `main` triggers an automatic
build, and — this is the actual payoff — **that build always runs
`npm install`**, which permanently fixes the recurring
`Cannot find module 'docx'` failure that happens with drag-and-drop zip
deploys that don't include `node_modules`.

If the human doesn't want to do that manual step right now, that's a
completely fine state to leave the project in — just keep using
`netlify deploy --prod` after `npm install` for future changes (Step 5).

### Safety net for the first Git-triggered build

Linking the repo triggers an initial build immediately. This is safe by
default, but worth doing carefully on a site with real traffic:

- A **failed** build never touches the live site — Netlify simply
  doesn't publish it, and the currently live version keeps serving
  exactly as before.
- If you want to inspect a build before it can possibly go live, disable
  auto-publishing first: **Deploys → Lock** (or **Site configuration →
  Build & deploy → Stop builds** to prevent building entirely). Builds
  will still happen on push, but won't replace what's live until you
  click **Publish Deploy** on the one you want.
- If a build succeeds and publishes something wrong, rollback is
  instant: open **Deploys**, find the last known-good deploy, and click
  **Publish Deploy** on it. This republishes an already-built version —
  it does not trigger a new build, and takes effect immediately.

## No scheduled functions in this project

Unlike the sibling project this pattern was adapted from, this project
has no cron/scheduled functions, so the "Netlify Scheduled Functions
only register during a Git-linked build" gotcha does not apply here.
Nothing to work around — no GitHub Actions cron needed either.

## Quick reference — redeploying after a code change

Once everything above is done once, a future code change is just:

```bash
cd /path/to/msp-report-builder-github
npm install        # only needed if package.json changed
node --check netlify/functions/intune.js
node --check netlify/functions/export-docx.js
netlify deploy --prod
```

Or, if Step 6 (Git link) was completed by the human:

```bash
git add -A
git commit -m "Describe the change"
git push
# Netlify builds and deploys automatically — no netlify CLI needed at all
```
