// Runs the real apps-script/*.gs code against the mocks, as a GM, a regional VP, corporate,
// an outsider and an anonymous night auditor. Results land in window.TEST_RESULTS and the page.
const RESULTS = [];
let SECTION = '';
function section(s) { SECTION = s; }
function check(name, cond, detail) { RESULTS.push({ section: SECTION, name, pass: !!cond, detail: cond ? '' : (detail === undefined ? '' : JSON.stringify(detail).slice(0, 400)) }); }
function as(user, fn) { const prev = MOCK.user; MOCK.user = user; try { return fn(); } finally { MOCK.user = prev; } }
function api(user, action, params) { return as(user, () => hubApi(Object.assign({ action }, params || {}))); }
function near(a, b, tol) { return Math.abs(a - b) <= (tol == null ? 1e-6 : tol) * Math.max(1, Math.abs(b)); }
function throws(fn) { try { fn(); return false; } catch (e) { return true; } }

function runAll() {
  loadFixtures();
  const today = todayKey_(), d1 = fcShiftDate_(today, -1), d2 = fcShiftDate_(today, -2), d3 = fcShiftDate_(today, -3);

  // ═══ 1. AUTH ═══
  section('Auth');
  check('no Google session -> no_identity', api('', 'whoami').error === 'no_identity');
  check('non-RGB Google account -> no_identity', api(FX.OUTSIDER, 'whoami').error === 'no_identity');
  check('RGB account in no Hub group -> no_scope', api(FX.NOBODY, 'whoami').error === 'no_scope');
  const gmWho = api(FX.GM, 'whoami');
  check('GM sees only their hotel', gmWho.properties && gmWho.properties.length === 1 && gmWho.properties[0].property_id === '10795674', gmWho);
  check('GM role is GM, not corporate', gmWho.hub_role === 'GM' && gmWho.is_corporate === false);
  const rvpWho = api(FX.RVP, 'whoami');
  check('Regional (region-dallas@) sees both Dallas hotels only', rvpWho.properties.map(p => p.property_id).sort().join() === 'DALLQ,PROV-DALES', rvpWho.properties.map(p => p.property_id));
  const corpWho = api(FX.CORP, 'whoami');
  check('Corporate sees all 5 hotels', corpWho.properties.length === 5 && corpWho.is_corporate === true);
  const gmFlash = api(FX.GM, 'portfolio_flash');
  check('GM Portfolio Flash: only their rows reach the browser',
    gmFlash.flash.length === 12 && gmFlash.flash.every(r => String(r.property_id) === '10795674') &&
    gmFlash.budget.every(r => String(r.property_id) === '10795674') && gmFlash.historical.every(r => String(r.property_id) === '10795674'),
    { flash: gmFlash.flash.length });
  const spoof = api(FX.GM, 'portfolio_flash', { username: FX.CORP, password: 'x' });
  check('username/password sent by the browser is ignored (no escalation)', spoof.properties.length === 1);
  check('identity is cleared after each call', REQUEST_SSO_ === null);
  check('GM cannot open another hotel\'s baseline', api(FX.GM, 'baseline_property', { property_id: 'HOUHX' }).error === 'not_authorized');
  check('GM cannot open settings', api(FX.GM, 'baseline_settings_get').error === 'not_authorized');
  check('GM cannot list night-audit links', api(FX.GM, 'na_tokens').error === 'not_authorized');
  check('GM cannot see another hotel\'s 2027 forecast', !!api(FX.GM, 'forecast_get', { property_id: 'HOUHX' }).error);
  check('maintenance function blocked for a signed-in GM', as(FX.GM, () => throws(() => refreshHubCache())));
  check('maintenance function blocked for anonymous', as('', () => throws(() => setupForecast2027())));
  check('maintenance function allowed for the owner in the editor', as(FX.OWNER, () => !throws(() => refreshHubCache())));
  check('served Hub page refuses an outsider', as(FX.OUTSIDER, () => doGet({ parameter: {} }).kind === 'html'));
  check('served Hub page renders for the GM', as(FX.GM, () => { const o = doGet({ parameter: {} }); return o.kind === 'template' && o.v.file === 'Hub'; }));
  check('served App page renders for corporate', as(FX.CORP, () => doGet({ parameter: { view: 'app', tab: 'settings' } }).v.file === 'App'));
  // legacy franchise login still works; RGB staff can be forced onto Google
  const fr = JSON.parse(doGet({ parameter: { action: 'whoami', username: 'franchise.gm@example.com', password: 'fpw' } }).v);
  check('legacy franchise password login still works', fr.role === 'READ' && fr.properties.length === 1, fr);
  MOCK.props.SSO_ENFORCED = 'true';
  const old = JSON.parse(doGet({ parameter: { action: 'whoami', username: FX.GM, password: 'oldpw' } }).v);
  check('SSO_ENFORCED blocks @rgbhospitality.com password logins', old.error === 'not_authorized', old);
  delete MOCK.props.SSO_ENFORCED;

  // ═══ 2. NIGHT-AUDIT LINKS ═══
  section('Night audit links');
  check('GM cannot create a link', api(FX.GM, 'na_rotate', { property_id: '10795674' }).error === 'not_authorized');
  check('link URL rejects a non-exec URL', !!api(FX.CORP, 'na_set_url', { url: 'https://evil.example.com/exec' }).error);
  check('corporate sets the public URL', api(FX.CORP, 'na_set_url', { url: 'https://script.google.com/macros/s/PUBLICDEPLOY/exec' }).ok === true);
  const rot = api(FX.CORP, 'na_rotate', { property_id: '10795674' });
  const token = (rot.url || '').split('?na=')[1] || '';
  check('rotate returns a link on the public deployment with a 64-hex token', /^https:\/\/script\.google\.com\/macros\/s\/PUBLICDEPLOY\/exec\?na=[0-9a-f]{64}$/.test(rot.url), rot);
  const tokSheet = MOCK.sheets[MOCK.props.NA_TOKENS_ID].sheets[0];
  check('only a hash of the token is stored', JSON.stringify(tokSheet.rows).indexOf(token) < 0 && tokSheet.rows[1][1] === sha256Hex_(token));
  const page = doGet({ parameter: { na: token } });
  const boot = JSON.parse(page.v.boot);
  check('link page knows the hotel from the token (no hotel picker)', boot.property_name === 'Hampton Inn Portland' && !('property_id' in boot));
  check('link page carries no existing data', !('row' in boot) && !('values' in boot));
  check('bad token gets the generic "not active" page', doGet({ parameter: { na: 'a'.repeat(64) } }).kind === 'html');
  const dar = { date: d1, submitter_name: 'Night Auditor', user_agent: 'TestBrowser/1', rooms_rented: 60, room_rental_revenue: 7200, parking: 100, bank_deposits: 500 };
  const r1 = as('', () => naSubmit(token, dar));
  check('submit returns only received/date/time', r1.received === true && Object.keys(r1).sort().join() === 'date,received,received_at', r1);
  let T = darTable_();
  let i = darFindRow_(T, '10795674', d1);
  check('submission lands as Pending GM review', i > 0 && T.data[i][T.col.status] === 'SUBMITTED');
  check('submission source + link hint recorded', T.data[i][T.col.submission_source] === 'night-audit-link' && T.data[i][T.col.token_hint] === token.slice(-4));
  const r2 = as('', () => naSubmit(token, Object.assign({}, dar, { rooms_rented: 61 })));
  T = darTable_();
  check('resubmitting a pending date replaces it (still one row)', r2.received && T.data.filter(r => r[T.col.property_id] === '10795674' && r[T.col.date] === d1).length === 1 && T.data[darFindRow_(T, '10795674', d1)][T.col.rooms_rented] === 61);
  const log = MOCK.sheets[DAR_LOG_SHEET_ID].sheets[0];
  const lh = log.rows[0].map(String);
  const logRows = log.rows.slice(1).map(r => Object.fromEntries(lh.map((h, k) => [h, r[k]])));
  check('audit log: SUBMIT then REPLACE, with time, name, browser, link hint',
    logRows.length === 2 && logRows[0].action === 'SUBMIT' && logRows[1].action === 'REPLACE' && logRows[1].submitter_name === 'Night Auditor' &&
    logRows[1].user_agent === 'TestBrowser/1' && logRows[1].token_hint === token.slice(-4) && logRows[1].logged_at instanceof Date, logRows);
  check('future business date rejected', as('', () => naSubmit(token, Object.assign({}, dar, { date: fcShiftDate_(today, 1) }))).error === 'bad_request');
  check('date older than 14 days rejected', as('', () => naSubmit(token, Object.assign({}, dar, { date: fcShiftDate_(today, -15) }))).error === 'bad_request');
  check('rooms over the room count rejected', as('', () => naSubmit(token, Object.assign({}, dar, { date: d2, rooms_rented: 81 }))).error === 'bad_request');
  check('name required', as('', () => naSubmit(token, Object.assign({}, dar, { date: d2, submitter_name: '' }))).error === 'bad_request');
  // a second hotel's link, for scoping tests below
  const rotD = api(FX.CORP, 'na_rotate', { property_id: 'PROV-DALES' });
  const tokD = rotD.url.split('?na=')[1];
  as('', () => naSubmit(tokD, { date: d1, submitter_name: 'Dallas Audit', rooms_rented: 100, room_rental_revenue: 17000 }));
  const rot2 = api(FX.CORP, 'na_rotate', { property_id: '10795674' });
  const token2 = rot2.url.split('?na=')[1];
  check('after rotation the old link stops working', as('', () => naSubmit(token, Object.assign({}, dar, { date: d3 }))).error === 'invalid_link');
  check('and the new link works', as('', () => naSubmit(token2, Object.assign({}, dar, { date: d3 }))).received === true);
  const tl = api(FX.CORP, 'na_tokens');
  const pl = tl.properties.find(p => p.property_id === '10795674');
  check('link list shows hint + revoked count, never a token', pl.active && pl.hint === token2.slice(-4) && pl.revoked_count === 1 && JSON.stringify(tl).indexOf(token2) < 0);
  // legacy auditor login is write-only too
  const lg = JSON.parse(doGet({ parameter: { action: 'dar_get', username: 'crppo.audit', password: 'auditpw', date: d1 } }).v);
  check('legacy auditor dar_get returns status only, no numbers', lg.exists === true && lg.row === null && lg.status === 'SUBMITTED', lg);

  // ═══ 3. GM APPROVAL ═══
  section('DAR approval');
  const q = api(FX.GM, 'dar_queue');
  check('GM queue shows their pending DARs only', q.pending.length === 2 && q.pending.every(r => r.property_id === '10795674'), q.pending.map(r => r.property_id + ' ' + r.date));
  check('GM cannot decide another hotel\'s DAR', api(FX.GM, 'dar_decide', { property_id: 'PROV-DALES', date: d1, decision: 'APPROVE' }).error === 'not_authorized');
  check('flag without a note is refused', api(FX.GM, 'dar_decide', { property_id: '10795674', date: d1, decision: 'FLAG', note: '' }).error === 'bad_request');
  const fl = api(FX.GM, 'dar_decide', { property_id: '10795674', date: d1, decision: 'FLAG', note: 'Parking looks doubled' });
  check('flag with note -> FLAGGED', fl.status === 'FLAGGED', fl);
  check('flag is emailed to corporate@ with the note', MOCK.mail.some(m => m.to === 'corporate@rgbhospitality.com' && /Parking looks doubled/.test(m.body)));
  check('auditor cannot resubmit a flagged date', as('', () => naSubmit(token2, dar)).error === 'locked');
  check('GM cannot resolve a flag (corporate only)', api(FX.GM, 'dar_corporate', { property_id: '10795674', date: d1, decision: 'APPROVE' }).error === 'not_authorized');
  const cq = api(FX.CORP, 'dar_queue');
  check('corporate queue shows the flag with the GM note', cq.flagged.length === 1 && cq.flagged[0].flag_note === 'Parking looks doubled' && cq.is_corporate);
  check('corporate reopen needs a note', api(FX.CORP, 'dar_corporate', { property_id: '10795674', date: d1, decision: 'REOPEN', note: '' }).error === 'bad_request');
  check('corporate reopens -> REOPENED', api(FX.CORP, 'dar_corporate', { property_id: '10795674', date: d1, decision: 'REOPEN', note: 'Fix parking' }).status === 'REOPENED');
  check('auditor resubmits a reopened date', as('', () => naSubmit(token2, Object.assign({}, dar, { parking: 50 }))).received === true);
  check('...and it is back to Pending GM review', (() => { const t = darTable_(); return t.data[darFindRow_(t, '10795674', d1)][t.col.status] === 'SUBMITTED'; })());
  check('GM approves -> VALIDATED', api(FX.GM, 'dar_decide', { property_id: '10795674', date: d1, decision: 'APPROVE' }).status === 'VALIDATED');
  check('approved date is locked for the auditor link', as('', () => naSubmit(token2, dar)).error === 'locked');
  const legacy = as('', () => handleDarSubmission_({ username: 'crppo.audit', password: 'auditpw', date: d1, rooms_rented: 1 }));
  check('approved date is locked for the legacy auditor login too', legacy.error === 'already_validated');
  check('approved date cannot be decided again', api(FX.GM, 'dar_decide', { property_id: '10795674', date: d1, decision: 'FLAG', note: 'late' }).error === 'bad_state');
  const grid = api(FX.GM, 'dar_status', { days: 21 });
  check('status grid per date (GM: 1 hotel)', grid.properties.length === 1 && grid.properties[0].cells[d1] === 'VALIDATED' && grid.properties[0].cells[d3] === 'SUBMITTED' && !grid.properties[0].cells[d2], grid.properties[0].cells);
  const roll = api(FX.CORP, 'dar_rollup', { month: d1.slice(0, 7) });
  const rp = roll.rows.find(r => r.property_id === '10795674'), rd = roll.rows.find(r => r.property_id === 'PROV-DALES');
  check('rollup counts approved dates only', d1.slice(0, 7) !== d3.slice(0, 7) || (rp.approved_days === 1 && rp.rooms_sold === 60 && rp.pending === 1), rp);
  check('pending Dallas DAR is not in the rollup totals', rd.approved_days === 0 && rd.room_rev === 0 && rd.pending === 1, rd);
  check('GM rollup only has their hotel', api(FX.GM, 'dar_rollup', {}).rows.length === 1);
  MOCK.mail = [];
  MOCK.hour = 7;
  as('', () => sendPendingDarReminders());
  check('reminder does nothing before 10am (an early call can\'t use up the day)', MOCK.mail.length === 0 && !MOCK.props.DAR_REMINDER_LAST);
  MOCK.hour = 10;
  as('', () => sendPendingDarReminders());
  check('10am reminder emails the GM group of each hotel with pending DARs', MOCK.mail.length === 2 && MOCK.mail.some(m => /gm-10795674@/.test(m.to)) && MOCK.mail.some(m => /gm-prov-dales@/.test(m.to)), MOCK.mail.map(m => m.to));
  as('', () => sendPendingDarReminders());
  check('reminder sends at most once a day', MOCK.mail.length === 2);
  as(FX.OWNER, () => installDarReminderTrigger());
  check('reminder trigger installs at 10am', MOCK.triggers.length === 1 && MOCK.triggers[0].hour === 10);
  check('reminder trigger install is owner-only', as(FX.CORP, () => throws(() => installDarReminderTrigger())));

  // ═══ 4. 2027 BASELINE ═══
  section('2027 baseline');
  const port = api(FX.GM, 'baseline_property', { property_id: '10795674' });
  const pp = FX.props.find(p => p[0] === '10795674');
  check('label says "System baseline — not GM-submitted"', port.label === 'System baseline — not GM-submitted');
  check('a) Jan–Oct 2026 = actuals', port.stages.raw.slice(0, 10).every((x, m) => x.src === 'actual' && near(x.rev, FX.y26(pp, m).rev)));
  const nov25 = FX.y25(pp, 10);
  check('a) Nov 2026 = budget revenue, split by 2025 Nov ADR', near(port.stages.raw[10].rev, Math.round(FX.y26(pp, 10).rev)) && near(port.stages.raw[10].rn, Math.round(FX.y26(pp, 10).rev) / (nov25.rev / nov25.rn)), port.stages.raw[10]);
  check('b) non-host market: no event adjustment', !port.assumptions.event && near(port.totals.event.rev, port.totals.raw.rev));
  check('c) Easter: mixed hotel moves 2% of March room nights to April', port.assumptions.mix === 'mixed' && port.assumptions.easter_pct === 2 &&
    near(port.stages.easter[2].rn, port.stages.event[2].rn * 0.98) && near(port.stages.easter[3].rn, port.stages.event[3].rn + port.stages.event[2].rn * 0.02));
  check('c) Easter shift leaves annual revenue unchanged', near(port.totals.easter.rev, port.totals.event.rev));
  check('d) growth: every month × 1.018 ADR × 1.009 occ (RevPAR +2.7%)', port.stages.final.every((x, m) => near(x.rev, port.stages.easter[m].rev * 1.018 * 1.009, 1e-9)));
  check('d) growth source is the default for upper midscale', port.assumptions.growth_adr === 1.8 && port.assumptions.growth_occ === 0.9 && port.assumptions.growth_adr_source === 'default');
  const wf = port.waterfall;
  check('waterfall: raw → event → Easter → growth → 2027 adds up', wf.length === 5 && near(wf[0].rev + wf[1].rev + wf[2].rev + wf[3].rev, wf[4].rev));

  const hou = api(FX.CORP, 'baseline_property', { property_id: 'HOUHX' });
  const hp = FX.props.find(p => p[0] === 'HOUHX');
  check('b) Houston is a World Cup host market', hou.assumptions.event === true && hou.event.applied);
  check('b) trend measured from the other actual months (+2% occ, +3% ADR)', near(hou.event.occ_trend, 0.02, 0.003) && near(hou.event.adr_trend, 0.03, 0.003), hou.event);
  check('b) Jun 2026 replaced with 2025 Jun × (1 + trend)', near(hou.stages.event[5].rn, Math.min(hp[4] * 30, FX.y25(hp, 5).rn * (1 + hou.event.occ_trend))) && hou.stages.event[5].rev < hou.stages.raw[5].rev * 0.8, { ev: hou.stages.event[5], raw: hou.stages.raw[5] });
  check('b) other months untouched by the event step', [0, 1, 2, 3, 4, 7, 8, 9, 10, 11].every(m => near(hou.stages.event[m].rev, hou.stages.raw[m].rev)));
  check('c) corporate-leaning hotel (Galleria) shifts 5%', hou.assumptions.mix === 'corporate' && hou.assumptions.easter_pct === 5);
  check('d) no extra Jun/Jul premium: Jun growth factor = other months', near(hou.stages.final[5].rev / hou.stages.easter[5].rev, hou.stages.final[1].rev / hou.stages.easter[1].rev, 1e-9));

  const es = api(FX.CORP, 'baseline_property', { property_id: 'PROV-DALES' });
  check('b) Embassy Suites Dallas: no 2025 → Manual review required', es.manual_review === true && es.flags.some(f => /Manual review required/.test(f)), es.flags);
  check('b) ...Jun/Jul derived from shoulder months × peer seasonality', /shoulder/.test(es.event.method) && es.event.peers.indexOf('La Quinta Dallas North') >= 0, es.event);
  const sh = [4, 7].reduce((a, m) => ({ rn: a.rn + es.stages.raw[m].rn, av: a.av + es.stages.raw[m].avail }), { rn: 0, av: 0 });
  check('b) ...Jun occupancy = shoulder occupancy × peer ratio', near(es.stages.event[5].rn / es.stages.event[5].avail, Math.min(1, sh.rn / sh.av * es.event.ratios[5].occ), 1e-9));
  check('b) ...and well below the World Cup–inflated raw June', es.stages.event[5].rev < es.stages.raw[5].rev * 0.8);

  const tru = api(FX.CORP, 'baseline_property', { property_id: 'CRPTRU' });
  check('e) TRU inferred economy/midscale → ADR +1.0%, Occ +0.5%', tru.assumptions.tier === 'economy_midscale' && tru.assumptions.growth_adr === 1 && tru.assumptions.growth_occ === 0.5);
  check('c) leisure hotel (Rockport) → 0% Easter shift', tru.assumptions.mix === 'leisure' && tru.easter_moved_rn === 0);
  const lq = api(FX.CORP, 'baseline_property', { property_id: 'DALLQ' });
  check('e) La Quinta inferred economy/midscale', lq.assumptions.tier === 'economy_midscale');

  // settings
  check('GM cannot save settings', api(FX.GM, 'baseline_settings_save', { settings: {} }).error === 'not_authorized');
  const sg = api(FX.CORP, 'baseline_settings_get');
  const s = JSON.parse(JSON.stringify(sg.settings));
  s.growth.markets = { Dallas: { adr: 0.5, occ: '' } };
  s.properties = { '10795674': { tier: 'upscale', mix: 'corporate', event: 'auto', easter_pct: '' } };
  s.gm_threshold_pct = 5;
  const bad = JSON.parse(JSON.stringify(s)); bad.growth.default.adr = 99;
  check('out-of-range assumption refused', api(FX.CORP, 'baseline_settings_save', { settings: bad }).error === 'bad_request');
  const saved = api(FX.CORP, 'baseline_settings_save', { settings: s });
  check('corporate saves settings (stamped with who/when)', saved.ok && saved.settings.updated_by === FX.CORP);
  const lq2 = api(FX.CORP, 'baseline_property', { property_id: 'DALLQ' });
  check('e) market override beats tier for ADR; occ still from tier', lq2.assumptions.growth_adr === 0.5 && /market/.test(lq2.assumptions.growth_adr_source) && lq2.assumptions.growth_occ === 0.5);
  const port2 = api(FX.GM, 'baseline_property', { property_id: '10795674' });
  check('e) per-hotel override: Portland now corporate mix → 5% Easter', port2.assumptions.mix === 'corporate' && port2.assumptions.easter_pct === 5 && !port2.assumptions.mix_inferred);

  // GM entry + threshold + rollup counts
  section('GM 2027 entry vs baseline');
  as(FX.OWNER, () => setupForecast2027());
  let pf = api(FX.CORP, 'baseline_portfolio');
  check('portfolio: 0 GM-submitted, 5 on baseline', pf.counts.gm_submitted === 0 && pf.counts.baseline === 5 && pf.properties.every(r => r.source_label === 'System baseline — not GM-submitted'), pf.counts);
  check('portfolio counts manual-review hotels', pf.counts.manual_review === 1);
  check('GM portfolio view has only their hotel', api(FX.GM, 'baseline_portfolio').properties.length === 1);
  const fg = api(FX.GM, 'forecast_get', { property_id: '10795674' });
  check('budget editor receives the system baseline beside the GM numbers', fg.system_baseline && fg.system_baseline.label === 'System baseline — not GM-submitted' && fg.system_baseline.months.length === 12 && fg.system_baseline.threshold_pct === 5);
  // GM enters every day 15% above baseline RevPAR
  const blm = fg.system_baseline.months;
  const days = fcYearDays_(2027).map(d => { const m = Number(d.slice(5, 7)) - 1; const occ = blm[m].rn / blm[m].avail; const adr = blm[m].rev / blm[m].rn * 1.15; return { date: d, rooms_sold: Math.round(80 * occ), adr: Math.round(adr * 100) / 100, notes: '' }; });
  check('GM saves all 365 days', api(FX.GM, 'forecast_save', { property_id: '10795674', days: days }).changed === 365);
  const sub1 = api(FX.GM, 'forecast_submit', { property_id: '10795674' });
  check('submit >5% from baseline without a reason → justification_required', sub1.error === 'justification_required' && sub1.diff_pct > 5, sub1);
  check('a too-short reason is refused', api(FX.GM, 'forecast_submit', { property_id: '10795674', justification: 'ok' }).error === 'justification_required');
  const sub2 = api(FX.GM, 'forecast_submit', { property_id: '10795674', justification: 'New hospital wing opens Feb 2027, signed contract' });
  check('submit with a justification goes through', sub2.ok && sub2.status === 'SUBMITTED', sub2);
  const st = fcReadStatus_().map['10795674'];
  check('justification + both RevPARs stored on the Status row', /hospital/.test(st.baseline_justification) && Number(st.baseline_diff_pct) > 5 && Number(st.gm_revpar) > Number(st.baseline_revpar), st);
  pf = api(FX.CORP, 'baseline_portfolio');
  const pr = pf.properties.find(r => r.property_id === '10795674');
  check('portfolio: 1 GM-submitted, 4 baseline', pf.counts.gm_submitted === 1 && pf.counts.baseline === 4, pf.counts);
  check('GM entry overrides the baseline in the rollup', pr.source === 'gm' && near(pf.totals.rev, pf.properties.reduce((a, r) => a + (r.source === 'gm' ? r.gm.rev : r.baseline.rev), 0)));
  check('GM vs baseline shown with the threshold breach + justification', pr.over_threshold && /hospital/.test(pr.justification) && near(pr.diff_pct, 15, 0.05), pr);
  const subFlat = (() => { // a GM within 5% needs no justification
    const d2s = fcYearDays_(2027).map(d => { const m = Number(d.slice(5, 7)) - 1; const p = FX.props.find(x => x[0] === 'CRPTRU'); return { date: d, rooms_sold: Math.round(p[4] * 0.6), adr: 0, notes: '' }; });
    const b = api(FX.CORP, 'forecast_get', { property_id: 'CRPTRU' }).system_baseline;
    d2s.forEach(x => { const m = Number(x.date.slice(5, 7)) - 1; x.rooms_sold = Math.round(b.months[m].rn / FX.dim(2027, m)); x.adr = Math.round(b.months[m].rev / b.months[m].rn * 100) / 100; });
    api(FX.CORP, 'forecast_save', { property_id: 'CRPTRU', days: d2s });
    return api(FX.CORP, 'forecast_submit', { property_id: 'CRPTRU' });
  })();
  check('GM within ±5% submits without a justification', subFlat.ok === true, subFlat);

  render();
}

function render() {
  window.TEST_RESULTS = RESULTS;
  const pass = RESULTS.filter(r => r.pass).length;
  const out = document.getElementById('out');
  let html = '<h1>' + pass + ' / ' + RESULTS.length + ' passed</h1>';
  let sec = '';
  RESULTS.forEach(r => {
    if (r.section !== sec) { sec = r.section; html += '<h2>' + sec + '</h2>'; }
    html += '<div class="' + (r.pass ? 'p' : 'f') + '">' + (r.pass ? '✓' : '✗') + ' ' + r.name.replace(/</g, '&lt;') + (r.detail ? '<pre>' + r.detail.replace(/</g, '&lt;') + '</pre>' : '') + '</div>';
  });
  out.innerHTML = html;
}

window.addEventListener('load', () => {
  try { runAll(); } catch (e) { RESULTS.push({ section: 'CRASH', name: String(e && e.stack || e), pass: false, detail: '' }); render(); }
});
