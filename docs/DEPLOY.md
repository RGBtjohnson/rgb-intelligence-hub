# Deploying Google sign-in, night-audit links, DAR approval and the 2027 baseline

Branch: `feature/sso-dar-baseline`. **Plan: do not deploy before the Dec 1 2026 budget deadline.** The current
deployment keeps working untouched until you switch people over.

## What changes for whom

| Who | Before | After |
|---|---|---|
| RGB GMs, regional, corporate | Username + password on the static Hub | Open the Hub URL, already signed in with their @rgbhospitality.com Google account. What they see comes from Google Groups. |
| Night auditors | Shared username + password | One private link per hotel. Submit-only; nothing on file is shown. |
| Franchise GMs (no Google account) | Username + password | Unchanged for now (static Hub + "Anyone" deployment). |

## 1. Move the script into the RGB Workspace (one time)

"Anyone within rgbhospitality.com" only exists for scripts **owned by an @rgbhospitality.com account**, and only
then does Apps Script reveal who is signed in.

1. Signed in as your @rgbhospitality.com account, create a new Apps Script project (script.google.com → New project).
2. Add these files, pasting from the repo:
   - Script files: `Code.gs`, `Auth.gs`, `Dar.gs`, `Baseline.gs` (from `apps-script/`)
   - HTML files (File → New → HTML, name exactly): `App`, `NightAudit` (from `apps-script/App.html`, `NightAudit.html`)
     and `Hub` (paste the whole of `rgb-intelligence-hub-v4.html`)
3. Project Settings → Time zone: **America/Chicago** (the 10am reminder and "today" use it).
4. Share every backing Google Sheet (IDs at the top of `Code.gs`) with this account as **Editor**.
5. Copy the Script Properties from the old project (Project Settings → Script properties), especially
   `FC_FORECAST_ID`, `FC_LOG_ID`, `FC_BUDGET_ID`, `DAILY_ACTUALS_ID`, so the 2027 forecast keeps its data.
6. Optional, faster group lookups: Services → add **Admin SDK API** (Directory). Only works if the owner can read the
   directory; without it the code checks each group with GroupsApp, which is fine for ~30 groups.

## 2. Google Groups (who sees what)

Create in Google Admin (or Groups), all `@rgbhospitality.com`:

| Group | Gives |
|---|---|
| `corporate@` | Whole portfolio + Settings (baseline assumptions, night-audit links) + resolving flagged DARs |
| `region-<market>@` e.g. `region-houston@`, `region-corpus-christi@` | Every hotel in that market |
| `gm-<property_id>@` e.g. `gm-prov-houes@`, `gm-10795674@` | That hotel |

`<market>` / `<property_id>` are the Properties sheet values, lowercased, with spaces and symbols turned into `-`.
If a hotel's group is named differently, add a `gm_group` column to the Properties sheet with the full address.

The script owner must be able to **see each group's members** (set "Who can view members" to organisation
members, or make the owner a member/manager). Membership changes take effect within 5 minutes (cache).

## 3. Two deployments of the same project

Deploy → New deployment → Web app, twice:

| | A: Staff Hub | B: Public (night audit + franchise) |
|---|---|---|
| Execute as | Me | Me |
| Who has access | **Anyone within rgbhospitality.com** | **Anyone** |
| Used for | `…/exec` (Hub), `…/exec?view=app` (DAR review, baseline, settings) | `…/exec?na=<token>` links; legacy password API for the static Hub |

The code enforces the domain itself on every signed-in page and call, so deployment B can't be used to see Hub
data even though it is public. It only serves the submit-only link form and the old password endpoints.

After deploying:
0. In `rgb-intelligence-hub-v4.html` (the static copy franchise GMs use), set `APPS_SCRIPT_URL` to deployment B's URL.
1. Project Settings → Script properties → add `HUB_EXEC_URL` = deployment A's `/exec` URL (used in emails and links).
2. Open A's URL → **Settings** → *Night-audit links* → paste deployment B's `/exec` URL → Save.
3. Click **Create link** per hotel. The full link is shown **once**; send it to that hotel's night audit.
   Lost or leaked → **Rotate** (old link dies immediately) or **Revoke**.
4. Optional 10am reminder to GMs with pending DARs: in the editor run `installDarReminderTrigger` once
   (`removeDarReminderTrigger` turns it off). Flag emails to corporate@ can be turned off with script property
   `DAR_FLAG_EMAIL` = `off`.
5. When every RGB person has signed in with Google successfully, set script property `SSO_ENFORCED` = `true`:
   @rgbhospitality.com usernames can then no longer use the password login. Franchise logins keep working.

## 4. Cloudflare Access (if the Hub has a Cloudflare-hosted address)

The data boundary is Apps Script (sign-in + server-side filtering). Cloudflare Access is an extra front door for
any page you host on Cloudflare (e.g. a `hub.rgbhospitality.com` that links/redirects to deployment A):

Zero Trust → Access → Applications → Add → Self-hosted → your hostname →
Policy **Allow**, Include **Emails ending in** `@rgbhospitality.com`, identity provider **Google Workspace**.
If the static password Hub stays on Cloudflare for franchise GMs, give it its own application with an
Include **Emails** list of those GMs (One-time PIN).

## 5. Test before switching anyone (real accounts)

As a **GM test account** (in `gm-<id>@` only):
- [ ] Deployment A opens the Hub without a password; only your hotel appears on Portfolio, Budget, DAR Review.
- [ ] No Settings tab. `…?view=app&tab=settings` shows DAR Review instead.
- [ ] Submit a DAR through your hotel's link (incognito window, no Google sign-in) → it appears as *Pending GM review*.
- [ ] Flag it without a note → refused. Flag with a note → corporate@ gets the email.
- [ ] 2027 Baseline shows the *System baseline — not GM-submitted* label and the waterfall for your hotel only.

As a **corporate test account** (in `corporate@`):
- [ ] Every hotel appears; Settings tab present.
- [ ] Reopen the flagged DAR → the link can resubmit → GM approves → the link now says it's locked.
- [ ] Approved DARs tab counts it; pending ones are listed but not totalled.
- [ ] Change an assumption (e.g. Dallas market ADR) → that hotel's baseline changes; GM test account still can't open Settings.
- [ ] Rotate a link → the old one says "not active".

As a **non-RGB Google account** and as an RGB account **in no group**: deployment A refuses.

## Offline tests (no Google needed)

`tests/harness.html` runs the real `.gs` files against mock Apps Script services (108 checks covering GM,
regional, corporate, outsider and anonymous auditor). `tests/ui.html` previews the served pages as each of those
users. Serve the repo with `powershell -File tests/serve.ps1` and open `http://localhost:8765/tests/harness.html`.

## Known limits

- Apps Script does not expose a caller's IP address. The audit log records time, action, the link's last 4
  characters, the name typed on the form, the browser string and a per-session key.
- Rate Shop rows are not filtered per hotel: the cluster → hotel mapping lives in the Hub page, not on the server.
  It is competitor/public rate data, but if GMs should see only their own cluster, that mapping has to move server-side.
- "Rolls to portfolio" means the *Approved DARs* rollup. The Hub's older *DAR Business Summary* screen still shows
  its hard-coded demo data and Portfolio Flash still reads the Flash sheet.
- The `Hub` HTML file in Apps Script is a pasted copy of `rgb-intelligence-hub-v4.html`; re-paste after Hub changes.
