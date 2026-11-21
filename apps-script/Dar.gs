/**
 * RGB Intelligence Hub — Night-audit submission links + GM approval of DARs
 * -------------------------------------------------------------------------
 * SUBMISSION LINKS (write-only)
 *   Every hotel gets one unguessable link: <public exec URL>?na=<64 hex chars, ~244 random bits>.
 *   Only a SHA-256 hash of the token is stored (Night Audit Links sheet), so the sheet alone
 *   can't be used to submit. The token decides the hotel on the server; the form has no hotel
 *   picker. The page and naSubmit() never send back anything already on file: the reply is only
 *   "received", the date and the time.
 *   Corporate rotates (new link, old one stops working at once) or revokes a link from
 *   App > Settings > Night-audit links. The full link is shown once, when it is made.
 *
 * ONE DAR PER HOTEL PER BUSINESS DATE
 *   status (DAR Data sheet)   meaning                              auditor may resubmit?
 *   SUBMITTED                 Pending GM review                    yes, replaces it
 *   VALIDATED                 Approved by GM: locked, in rollups   no
 *   FLAGGED                   Flagged for correction, with corp.   no (corporate decides)
 *   REOPENED                  Corporate sent it back to the hotel  yes, goes back to Pending
 *   Every submission, replacement and decision is appended to the DAR Submission Log with the
 *   time, the action, the link's last 4 characters, the name typed on the form and the browser.
 *   (Apps Script does not expose the caller's IP address, so it cannot be logged.)
 *
 * GM REVIEW
 *   App > DAR Review: the GM approves, or flags with a required note. A flag goes to corporate
 *   (their queue, plus an email to corporate@ unless the DAR_FLAG_EMAIL script property is "off").
 *   Corporate can reopen a flagged date for the hotel to resubmit, or approve it as is.
 *   Optional 10am email to GMs with anything still pending: run installDarReminderTrigger once.
 */

const DAR_STATUS_LABEL = {
  '': 'Not submitted', SUBMITTED: 'Pending GM review', VALIDATED: 'Approved', FLAGGED: 'Flagged — with corporate', REOPENED: 'Reopened — resubmit'
};
const DAR_ROW_BASE_KEYS = ['property_id', 'date', 'submitted_by', 'submitted_at', 'status', 'validated_by', 'validated_at', 'flag_note'];
const DAR_ROW_AUDIT_KEYS = ['submission_source', 'token_hint', 'submitter_name', 'user_agent', 'flag_routed_at', 'corporate_by', 'corporate_at', 'corporate_note'];
const DAR_LOG_AUDIT_KEYS = ['action', 'submission_source', 'token_hint', 'submitter_name', 'user_agent', 'session_key', 'note'];
const NA_MAX_BACKDATE_DAYS = 14;   // a link can submit for today back to this many days ago
const NA_TOKEN_HEADERS = ['property_id', 'token_hash', 'token_hint', 'status', 'created_by', 'created_at', 'revoked_by', 'revoked_at'];

// ── small helpers ──
function sha256Hex_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}
function todayKey_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function darErr_(code, message, status) { return { error: code, message: message, _status: status || 400 }; }
function darIso_(v) { return v instanceof Date ? v.toISOString() : String(v == null ? '' : v); }

/** The DAR Data sheet with every column this file needs, plus a header -> column index map. */
function darTable_() {
  const sheet = firstDataSheet_(DAR_SHEET_ID);
  darEnsureColumns_(sheet, DAR_ROW_BASE_KEYS.concat(darAllKeys_(), DAR_ROW_AUDIT_KEYS));
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());
  const col = {};
  header.forEach((h, i) => { if (!(h in col)) col[h] = i; });
  const tz = sheet.getParent().getSpreadsheetTimeZone();
  for (let i = 1; i < data.length; i++) {          // dates as plain yyyy-MM-dd, even if Sheets made them Dates
    data[i][col.date] = fcDateKey_(data[i][col.date], tz);
    data[i][col.property_id] = String(data[i][col.property_id]);
  }
  return { sheet: sheet, data: data, header: header, col: col };
}
function darFindRow_(T, propertyId, date) {
  for (let i = 1; i < T.data.length; i++) {
    if (T.data[i][T.col.property_id] === String(propertyId) && T.data[i][T.col.date] === date) return i;
  }
  return -1;
}
function darSetCells_(T, i, patch) {
  Object.keys(patch).forEach(k => {
    T.sheet.getRange(i + 1, T.col[k] + 1).setValue(patch[k]);
    T.data[i][T.col[k]] = patch[k];
  });
}
function darValidDate_(date) {
  return /^\d{4}-\d{2}-\d{2}$/.test(date) && fcShiftDate_(date, 0) === date;
}
function darRowValues_(T, r) {
  const v = {};
  darAllKeys_().forEach(k => { v[k] = r[T.col[k]]; });
  return v;
}

// ═══ SUBMISSION LINKS ═══

function naTokenSpec_() {
  if (!FC_SPREADSHEETS.NA_TOKENS_ID) {
    FC_SPREADSHEETS.NA_TOKENS_ID = { title: 'RGB Intel Hub — Night Audit Links (hashed, corporate only)', tabs: [
      { name: 'Tokens', header: NA_TOKEN_HEADERS, text: [1, 2, 3, 4, 5, 6, 7, 8] }
    ] };
  }
}
/** create=false never makes the sheet (so an anonymous visitor can't cause one to be created). */
function naTokenSheet_(create) {
  naTokenSpec_();
  const id = PropertiesService.getScriptProperties().getProperty('NA_TOKENS_ID');
  if (!id && !create) return null;
  const ss = create ? fcEnsureSpreadsheet_('NA_TOKENS_ID') : SpreadsheetApp.openById(id);
  return ss.getSheetByName('Tokens');
}
function naNewToken_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').toLowerCase(); }

/** {property_id, hint} for an ACTIVE token, else null. */
function naLookup_(token) {
  token = String(token || '').trim().toLowerCase();
  if (!/^[0-9a-f]{64}$/.test(token)) return null;
  const sh = naTokenSheet_(false);
  if (!sh) return null;
  const hash = sha256Hex_(token);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === hash && String(data[i][3]) === 'ACTIVE') return { property_id: String(data[i][0]), hint: String(data[i][2]) };
  }
  return null;
}

function naPublicBase_() {
  return PropertiesService.getScriptProperties().getProperty('NA_PUBLIC_URL') || ScriptApp.getService().getUrl();
}

/** The form page behind a submission link. */
function naPage_(token) {
  const t = naLookup_(token);
  if (!t) {
    return HtmlService.createHtmlOutput('<div style="font-family:Arial,sans-serif;max-width:420px;margin:60px auto;color:#222"><h2 style="font-size:18px">Night Audit Submission</h2>' +
      '<p>This link is not active. Ask RGB corporate for your hotel\'s current link.</p></div>').setTitle('Night Audit Submission');
  }
  const prop = getPropertyById_(t.property_id) || {};
  const tmpl = HtmlService.createTemplateFromFile('NightAudit');
  tmpl.boot = safeJsonForHtml_({
    token: String(token).trim().toLowerCase(),
    property_name: String(prop.name || t.property_id),
    rooms: Number(prop.rooms) || 0,
    profile: darProfileForProperty_(t.property_id),
    fields: darFieldsForProperty_(t.property_id).map(f => ({ key: f.key, label: f.label, section: f.section })),
    today: todayKey_(),
    max_backdate_days: NA_MAX_BACKDATE_DAYS
  });
  return tmpl.evaluate().setTitle('Night Audit Submission').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Called from NightAudit.html. Write-only: returns "received", never anything on file. */
function naSubmit(token, payload) {
  payload = payload || {};
  const t = naLookup_(token);
  if (!t) return darErr_('invalid_link', 'This link is no longer active. Ask RGB corporate for your hotel\'s current link.', 403);
  const pid = t.property_id;

  const date = String(payload.date || '').trim();
  if (!darValidDate_(date)) return darErr_('bad_request', 'Pick the business date this DAR is for.');
  const today = todayKey_();
  if (date > today) return darErr_('bad_request', 'The business date can\'t be in the future.');
  if (date < fcShiftDate_(today, -NA_MAX_BACKDATE_DAYS)) return darErr_('bad_request', 'That date is more than ' + NA_MAX_BACKDATE_DAYS + ' days ago. Contact corporate to submit it.');

  const name = String(payload.submitter_name || '').trim().slice(0, 80);
  if (!name) return darErr_('bad_request', 'Enter your name.');

  const prop = getPropertyById_(pid) || {};
  const values = {};
  DAR_FIELDS.forEach(f => { values[f.key] = 0; });
  const fields = darFieldsForProperty_(pid);
  for (let i = 0; i < fields.length; i++) {
    const f = fields[i], raw = payload[f.key];
    const n = raw === '' || raw == null ? 0 : Number(raw);
    if (!isFinite(n)) return darErr_('bad_request', f.label + ' must be a number.');
    values[f.key] = Math.round(n * 100) / 100;
  }
  const rooms = Number(prop.rooms) || 0;
  if (values.rooms_rented < 0 || Math.floor(values.rooms_rented) !== values.rooms_rented || (rooms && values.rooms_rented > rooms)) {
    return darErr_('bad_request', 'Rooms Rented must be a whole number from 0 to ' + (rooms || 'the hotel\'s room count') + '.');
  }
  const totals = computeDarTotals_(values, darProfileForProperty_(pid));

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const T = darTable_();
    const i = darFindRow_(T, pid, date);
    const cur = i >= 0 ? String(T.data[i][T.col.status]) : '';
    if (cur === 'VALIDATED') return darErr_('locked', 'The DAR for ' + date + ' is approved by your GM and locked. Contact your GM if something is wrong.', 409);
    if (cur === 'FLAGGED') return darErr_('locked', 'Your GM flagged the DAR for ' + date + ' for correction and it is with corporate. Corporate will reopen it if it needs resubmitting.', 409);

    const now = new Date();
    const ua = String(payload.user_agent || '').slice(0, 200);
    const audit = { submission_source: 'night-audit-link', token_hint: t.hint, submitter_name: name, user_agent: ua };
    const row = T.header.map(h => {
      if (h === 'property_id') return pid;
      if (h === 'date') return date;
      if (h in values) return values[h];
      if (h in totals) return totals[h];
      if (h in audit) return audit[h];
      if (h === 'submitted_by') return name + ' (link …' + t.hint + ')';
      if (h === 'submitted_at') return now;
      if (h === 'status') return 'SUBMITTED';
      return '';   // validated_*, flag_*, corporate_* start empty; the log keeps the history
    });
    if (i >= 0) T.sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
    else T.sheet.appendRow(row);

    let sessionKey = '';
    try { sessionKey = Session.getTemporaryActiveUserKey() || ''; } catch (e) { /* anonymous */ }
    logDarSubmission_(pid, date, name, values, totals, Object.assign({ action: i >= 0 ? 'REPLACE' : 'SUBMIT', session_key: sessionKey, note: cur ? 'replaced ' + cur : '' }, audit));
    return { received: true, date: date, received_at: now.toISOString() };
  } finally {
    lock.releaseLock();
  }
}

// ── corporate: list / rotate / revoke links ──

function naRequireCorporate_(me) { return me && me.isCorporate ? null : darErr_('not_authorized', 'Only corporate can manage night-audit links.', 403); }

function naTokenList_(me) {
  const deny = naRequireCorporate_(me);
  if (deny) return deny;
  const sh = naTokenSheet_(false);
  const rows = sh ? sh.getDataRange().getValues().slice(1) : [];
  const props = getAuthorizedProperties_(me.grant);
  const list = props.map(p => {
    const pid = String(p.property_id);
    const mine = rows.filter(r => String(r[0]) === pid);
    const active = mine.find(r => String(r[3]) === 'ACTIVE');
    return {
      property_id: pid, name: p.name, market: p.market,
      active: !!active, hint: active ? String(active[2]) : '',
      created_by: active ? String(active[4]) : '', created_at: active ? darIso_(active[5]) : '',
      revoked_count: mine.filter(r => String(r[3]) === 'REVOKED').length
    };
  });
  return { properties: list, public_url: naPublicBase_(), public_url_set: !!PropertiesService.getScriptProperties().getProperty('NA_PUBLIC_URL') };
}

/** Revokes any active link for the hotel and makes a new one. The full link is returned only here. */
function naRotateToken_(me, p) {
  const deny = naRequireCorporate_(me);
  if (deny) return deny;
  const pid = String(p.property_id || '').trim();
  if (!getPropertyById_(pid)) return darErr_('not_found', 'Unknown property.', 404);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = naTokenSheet_(true);
    const now = new Date().toISOString();
    naRevokeRows_(sh, pid, me.username, now);
    const token = naNewToken_();
    fcAppendRows_(sh, [[pid, sha256Hex_(token), token.slice(-4), 'ACTIVE', me.username, now, '', '']], NA_TOKEN_HEADERS.length);
    return { ok: true, property_id: pid, hint: token.slice(-4), url: naPublicBase_() + '?na=' + token, created_at: now };
  } finally {
    lock.releaseLock();
  }
}

function naRevokeToken_(me, p) {
  const deny = naRequireCorporate_(me);
  if (deny) return deny;
  const pid = String(p.property_id || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = naTokenSheet_(false);
    const n = sh ? naRevokeRows_(sh, pid, me.username, new Date().toISOString()) : 0;
    return { ok: true, property_id: pid, revoked: n };
  } finally {
    lock.releaseLock();
  }
}

function naRevokeRows_(sh, pid, who, when) {
  const data = sh.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === pid && String(data[i][3]) === 'ACTIVE') {
      sh.getRange(i + 1, 4, 1, 1).setValue('REVOKED');
      sh.getRange(i + 1, 7, 1, 2).setValues([[who, when]]);
      n++;
    }
  }
  return n;
}

/** Corporate sets the "Anyone" deployment's /exec URL that links should point at (docs/DEPLOY.md). */
function naSetPublicUrl_(me, p) {
  const deny = naRequireCorporate_(me);
  if (deny) return deny;
  const url = String(p.url || '').trim();
  if (!/^https:\/\/script\.google\.com\/(a\/[^/]+\/)?macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(url)) {
    return darErr_('bad_request', 'Paste the public deployment\'s /exec URL (https://script.google.com/macros/s/…/exec).');
  }
  PropertiesService.getScriptProperties().setProperty('NA_PUBLIC_URL', url);
  return { ok: true, public_url: url };
}

// ═══ GM REVIEW ═══

function darPublicRow_(T, r, prop) {
  const pid = r[T.col.property_id];
  const status = String(r[T.col.status] || '');
  const values = {};
  darFieldsForProperty_(pid).forEach(f => { values[f.key] = Number(r[T.col[f.key]]) || 0; });
  const totals = {};
  DAR_TOTAL_KEYS.forEach(k => { const v = r[T.col[k]]; if (v !== '' && v != null) totals[k] = Number(v); });
  return {
    property_id: pid, property_name: prop ? prop.name : pid, rooms: prop ? Number(prop.rooms) || 0 : 0,
    date: r[T.col.date], status: status, status_label: DAR_STATUS_LABEL[status] || status,
    values: values, totals: totals,
    submitted_by: String(r[T.col.submitted_by] || ''), submitted_at: darIso_(r[T.col.submitted_at]),
    submission_source: String(r[T.col.submission_source] || ''),
    validated_by: String(r[T.col.validated_by] || ''), validated_at: darIso_(r[T.col.validated_at]),
    flag_note: String(r[T.col.flag_note] || ''), corporate_note: String(r[T.col.corporate_note] || '')
  };
}

/** Pending DARs for the caller's hotels, plus flagged ones (corporate acts on those). */
function darQueue_(me) {
  const props = getAuthorizedProperties_(me.grant);
  const byId = {};
  props.forEach(p => { byId[String(p.property_id)] = p; });
  const T = darTable_();
  const pending = [], flagged = [];
  T.data.slice(1).forEach(r => {
    const prop = byId[r[T.col.property_id]];
    if (!prop) return;
    const st = String(r[T.col.status]);
    if (st === 'SUBMITTED') pending.push(darPublicRow_(T, r, prop));
    else if (st === 'FLAGGED') flagged.push(darPublicRow_(T, r, prop));
  });
  const byDate = (a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : String(a.property_name).localeCompare(String(b.property_name)));
  pending.sort(byDate); flagged.sort(byDate);
  return {
    pending: pending, flagged: flagged, is_corporate: me.isCorporate,
    labels: DAR_FIELDS.reduce((o, f) => { o[f.key] = f.label; return o; }, {})
  };
}

/** Status per hotel per business date for the last `days` days (default 21). */
function darStatusGrid_(me, p) {
  const n = Math.max(7, Math.min(62, Number(p.days) || 21));
  const today = todayKey_();
  const dates = [];
  for (let i = 0; i < n; i++) dates.push(fcShiftDate_(today, -i));
  const props = getAuthorizedProperties_(me.grant);
  const ids = new Set(props.map(x => String(x.property_id)));
  const T = darTable_();
  const cells = {};
  T.data.slice(1).forEach(r => {
    const pid = r[T.col.property_id];
    if (!ids.has(pid)) return;
    (cells[pid] = cells[pid] || {})[r[T.col.date]] = String(r[T.col.status] || '');
  });
  return {
    dates: dates, labels: DAR_STATUS_LABEL,
    properties: props.map(x => ({ property_id: String(x.property_id), name: x.name, market: x.market, cells: cells[String(x.property_id)] || {} }))
  };
}

/** GM (or anyone with the hotel in scope) approves or flags a pending DAR. A flag needs a note. */
function darDecide_(me, p) {
  const pid = String(p.property_id || '').trim();
  if (!pid || !isAuthorizedForProperty_(me.grant, pid)) return darErr_('not_authorized', 'Not authorized for this property.', 403);
  const date = String(p.date || '').trim();
  if (!darValidDate_(date)) return darErr_('bad_request', 'Bad date.');
  const decision = String(p.decision || '').toUpperCase();
  if (decision !== 'APPROVE' && decision !== 'FLAG') return darErr_('bad_request', 'Decision must be APPROVE or FLAG.');
  const note = String(p.note || '').trim().slice(0, 1000);
  if (decision === 'FLAG' && note.length < 3) return darErr_('bad_request', 'Add a note saying what needs correcting.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let row, prop;
  try {
    const T = darTable_();
    const i = darFindRow_(T, pid, date);
    if (i < 0) return darErr_('not_found', 'No DAR submitted for that hotel and date.', 404);
    if (String(T.data[i][T.col.status]) !== 'SUBMITTED') return darErr_('bad_state', 'This DAR is not pending GM review (it is ' + (DAR_STATUS_LABEL[T.data[i][T.col.status]] || T.data[i][T.col.status]) + ').', 409);
    const now = new Date();
    const newStatus = decision === 'APPROVE' ? 'VALIDATED' : 'FLAGGED';
    darSetCells_(T, i, {
      status: newStatus, validated_by: me.username, validated_at: now,
      flag_note: decision === 'FLAG' ? note : '', flag_routed_at: decision === 'FLAG' ? now : ''
    });
    logDarSubmission_(pid, date, me.username, null, null, { action: decision === 'APPROVE' ? 'APPROVE' : 'FLAG', submission_source: 'review:' + me.via, note: note });
    row = T.data[i];
    prop = getPropertyById_(pid);
    if (decision === 'FLAG') darNotifyCorporate_(prop, date, note, me.username);
    return { ok: true, property_id: pid, date: date, status: newStatus, status_label: DAR_STATUS_LABEL[newStatus] };
  } finally {
    lock.releaseLock();
  }
}

/** Corporate only, on a FLAGGED DAR: REOPEN (hotel resubmits) or APPROVE as is. */
function darCorporateAction_(me, p) {
  if (!me.isCorporate) return darErr_('not_authorized', 'Only corporate can resolve flagged DARs.', 403);
  const pid = String(p.property_id || '').trim();
  const date = String(p.date || '').trim();
  const decision = String(p.decision || '').toUpperCase();
  if (decision !== 'REOPEN' && decision !== 'APPROVE') return darErr_('bad_request', 'Decision must be REOPEN or APPROVE.');
  const note = String(p.note || '').trim().slice(0, 1000);
  if (decision === 'REOPEN' && note.length < 3) return darErr_('bad_request', 'Add a note telling the hotel what to correct.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const T = darTable_();
    const i = darFindRow_(T, pid, date);
    if (i < 0) return darErr_('not_found', 'No DAR for that hotel and date.', 404);
    if (String(T.data[i][T.col.status]) !== 'FLAGGED') return darErr_('bad_state', 'Only flagged DARs can be resolved here.', 409);
    const now = new Date();
    const newStatus = decision === 'APPROVE' ? 'VALIDATED' : 'REOPENED';
    const patch = { status: newStatus, corporate_by: me.username, corporate_at: now, corporate_note: note };
    if (decision === 'APPROVE') { patch.validated_by = me.username; patch.validated_at = now; }
    darSetCells_(T, i, patch);
    logDarSubmission_(pid, date, me.username, null, null, { action: 'CORPORATE_' + decision, submission_source: 'review:' + me.via, note: note });
    return { ok: true, property_id: pid, date: date, status: newStatus, status_label: DAR_STATUS_LABEL[newStatus] };
  } finally {
    lock.releaseLock();
  }
}

/** Approved DARs only, rolled up per hotel for one month ("yyyy-MM"), within the caller's scope. */
function darRollup_(me, p) {
  const month = /^\d{4}-\d{2}$/.test(String(p.month || '')) ? String(p.month) : todayKey_().slice(0, 7);
  const props = getAuthorizedProperties_(me.grant);
  const T = darTable_();
  const agg = {};
  props.forEach(x => { agg[String(x.property_id)] = { days: 0, rooms_sold: 0, room_rev: 0, total_rev: 0, pending: 0, flagged: 0 }; });
  T.data.slice(1).forEach(r => {
    const g = agg[r[T.col.property_id]];
    if (!g || String(r[T.col.date]).slice(0, 7) !== month) return;
    const st = String(r[T.col.status]);
    if (st === 'VALIDATED') {
      g.days++;
      g.rooms_sold += Number(r[T.col.rooms_rented]) || 0;
      g.room_rev += Number(r[T.col.room_rental_revenue]) || 0;
      g.total_rev += Number(r[T.col.total_revenue]) || 0;
    } else if (st === 'SUBMITTED' || st === 'REOPENED') g.pending++;
    else if (st === 'FLAGGED') g.flagged++;
  });
  const rows = props.map(x => {
    const g = agg[String(x.property_id)], rooms = Number(x.rooms) || 0, avail = rooms * g.days;
    return {
      property_id: String(x.property_id), name: x.name, market: x.market, rooms: rooms,
      approved_days: g.days, pending: g.pending, flagged: g.flagged,
      rooms_sold: g.rooms_sold, room_rev: Math.round(g.room_rev * 100) / 100, total_rev: Math.round(g.total_rev * 100) / 100,
      occ: avail ? g.rooms_sold / avail : null, adr: g.rooms_sold ? g.room_rev / g.rooms_sold : null, revpar: avail ? g.room_rev / avail : null
    };
  });
  return { month: month, rows: rows };
}

// ═══ EMAIL ═══

function darAppLink_() { return hubAppUrl_() + '?view=app&tab=dar'; }

function darNotifyCorporate_(prop, date, note, by) {
  if (String(PropertiesService.getScriptProperties().getProperty('DAR_FLAG_EMAIL') || '').toLowerCase() === 'off') return;
  try {
    MailApp.sendEmail({
      to: CORPORATE_GROUP,
      subject: 'DAR flagged: ' + (prop ? prop.name : '') + ' — ' + date,
      body: (prop ? prop.name : '') + '\nBusiness date: ' + date + '\nFlagged by: ' + by + '\n\nNote:\n' + note + '\n\nResolve it: ' + darAppLink_()
    });
  } catch (e) { Logger.log('Flag email not sent: ' + e); }
}

/** Time-driven (10am). Emails each GM group whose hotel has DARs pending review. Sends at most once a day. */
function sendPendingDarReminders() {
  const sp = PropertiesService.getScriptProperties();
  const today = todayKey_();
  // Any served page can call a top-level function, so: never before 10am (an early call can't use up
  // the day's send), and at most once a day.
  if (Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'H')) < 10) return;
  if (sp.getProperty('DAR_REMINDER_LAST') === today) return;
  sp.setProperty('DAR_REMINDER_LAST', today);
  const T = darTable_();
  const byPid = {};
  T.data.slice(1).forEach(r => {
    if (String(r[T.col.status]) === 'SUBMITTED') (byPid[r[T.col.property_id]] = byPid[r[T.col.property_id]] || []).push(r[T.col.date]);
  });
  let sent = 0;
  Object.keys(byPid).forEach(pid => {
    const prop = getPropertyById_(pid);
    if (!prop) return;
    const to = [gmGroupFor_(prop)].concat(prop.gm_email ? [String(prop.gm_email).trim()] : []).join(',');
    const dates = byPid[pid].sort();
    try {
      MailApp.sendEmail({
        to: to,
        subject: dates.length + ' DAR' + (dates.length === 1 ? '' : 's') + ' waiting for your review — ' + prop.name,
        body: prop.name + ' has ' + dates.length + ' night-audit DAR' + (dates.length === 1 ? '' : 's') + ' pending your review:\n\n' +
          dates.map(d => '  • ' + d).join('\n') + '\n\nReview them: ' + darAppLink_()
      });
      sent++;
    } catch (e) { Logger.log('Reminder not sent for ' + pid + ': ' + e); }
  });
  Logger.log('DAR reminders sent: ' + sent);
}

/** Run once from the editor to turn on the 10am reminder (script time zone). Safe to re-run. */
function installDarReminderTrigger() {
  requireOwner_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sendPendingDarReminders').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendPendingDarReminders').timeBased().everyDays(1).atHour(10).nearMinute(0).create();
  Logger.log('10am DAR reminder installed (time zone: ' + Session.getScriptTimeZone() + ').');
}
/** Turns the reminder off. */
function removeDarReminderTrigger() {
  requireOwner_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sendPendingDarReminders').forEach(t => ScriptApp.deleteTrigger(t));
}
