# rgb-intelligence-hub
RGB Hospitality

- `rgb-intelligence-hub-v4.html` — the Hub (static page with password login, or served by Apps Script with Google sign-in)
- `rgb-property-portal.html` — property portal mock-up
- `apps-script/` — the Apps Script backend: `Code.gs` (data API, DAR, 2027 forecast), `Auth.gs` (Google sign-in +
  Google Groups scopes), `Dar.gs` (night-audit links, GM approval), `Baseline.gs` (2027 system baseline), and the
  served pages `App.html` / `NightAudit.html`
- `docs/DEPLOY.md` — setup for Google sign-in, the two deployments, Google Groups, and a GM/corporate test checklist
- `tests/` — offline tests that run the real `.gs` files against mock Apps Script services
