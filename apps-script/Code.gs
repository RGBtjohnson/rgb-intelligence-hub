/**
 * RGB Intelligence Hub — Permissions-filtered data API
 * ------------------------------------------------------
 * Pilot scope: Portfolio Flash module (read) + night-audit data entry (write).
 *
 * IDENTITY: username/password only, for everyone, checked against the
 * Credentials sheet — corporate, GMs (RGB and franchise), and night
 * auditors alike. (An earlier version of this tried to auto-detect RGB
 * Workspace users via Session.getActiveUser() and skip the login step for
 * them. That doesn't work: Apps Script only exposes the caller's identity
 * when the deployment forces a Google sign-in — i.e. "Who has access" is
 * domain-restricted. Since night auditors and franchise GMs have no Google
 * account to sign in with, this deployment must allow "Anyone," and
 * "Anyone" means Google never captures identity at all, for anyone,
 * Workspace or not. So: one login system for everyone, entirely under our
 * own control in the Credentials sheet.)
 *
 * Two roles in Credentials:
 *   READ  — username is the person's email; their actual scope (single
 *           property / market cluster / full portfolio) comes from looking
 *           that email up in the Permissions sheet, same as before.
 *   ENTRY — night auditors. property_id lives directly on the Credentials
 *           row; they can only ever submit data for that one property.
 *
 * SETUP
 * 1. The four backing sheets already exist (IDs below). Their columns:
 *      Properties      property_id, inn_code, name, brand, market, rooms,
 *                      gm_email, relationship, notes
 *      Permissions     email, scope_type, scope_value, role, notes
 *                      scope_type: property | market | portfolio
 *      PortfolioFlash  property_id, name, brand, market, rooms,
 *                      month_index (0-11), room_rev, occupied_rooms, total_rev
 *      Credentials     username, new_password, password_hash, salt,
 *                      property_id, role (ENTRY | READ), display_name,
 *                      active, last_synced
 * 2. To add/update/revoke a night-auditor or franchise-GM login: edit a row
 *    in the Credentials sheet (username, new_password as plaintext, property_id,
 *    role, active), then in the Apps Script editor pick "syncCredentials_" from
 *    the function dropdown (top toolbar) and click Run. That hashes any
 *    plaintext new_password cells, clears them, and stamps last_synced.
 *    To revoke: set active to FALSE and Run syncCredentials_ again.
 * 3. Deploy > New deployment > Web app.
 *      - Execute as: Me
 *        (The script must run with YOUR Drive permissions so it can open
 *        the backing sheets regardless of who is calling.)
 *      - Who has access: Anyone
 *        (Has to be public — auditors/franchise GMs have no Google account.
 *        There is no Google-login gate at all; access control is entirely
 *        the username/password + Permissions/Credentials sheet lookups
 *        below.)
 * 4. Give the frontend the deployed /exec URL (APPS_SCRIPT_URL in the HTML)
 *    — the dashboard will prompt for username/password on first load.
 *    Give night auditors and franchise GMs: `<url>?view=entry`.
 */

const PROPERTIES_SHEET_ID  = '1EZZel1s-F_cOLJ2mJbQ0G4VODahvmk1yVhSniwlqfaY'; // RGB Intel Hub — Properties (Canonical)
const PERMISSIONS_SHEET_ID = '1z0BxwYdXm9tIPypUYY2PJYzDc607C_t663ZFRLUKvbs'; // RGB Intel Hub — Permissions
const FLASH_SHEET_ID       = '126maQ0FS6WFH2xSptj71C8u_kOLnnkFFGB_vF5h2W24'; // RGB Intel Hub — Portfolio Flash Data
const CREDENTIALS_SHEET_ID = '1j9CJbbUw4gEKJoKdECzGaYj9ww1KDKu7pdqEXqqx9Sc'; // RGB Intel Hub — Credentials

// ═══ ENTRY POINTS ═══

function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.view === 'entry') {
    return HtmlService.createHtmlOutput(renderEntryFormHtml_())
      .setTitle('RGB Intelligence Hub — Data Entry')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  try {
    const identity = resolveIdentity_(params);
    if (!identity) {
      return jsonOut_({ error: 'no_identity', message: 'Could not verify identity. Google sign-in not detected and no valid username/password provided.' }, 401);
    }

    const module = params.module || 'portfolio-flash';
    if (module !== 'portfolio-flash') {
      return jsonOut_({ error: 'unknown_module', message: `Module "${module}" is not wired to the API yet.` }, 400);
    }

    const properties = getAuthorizedProperties_(identity.grant);
    const flashRows = getFlashRows_(properties.map(p => p.property_id));

    return jsonOut_({
      caller: identity.label,
      scope: identity.grant,
      properties: properties,
      flash: flashRows
    }, 200);

  } catch (err) {
    return jsonOut_({ error: 'server_error', message: String(err) }, 500);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = body.action || 'submit_flash';

    if (action === 'submit_flash') {
      return jsonOut_(handleFlashSubmission_(body), 200);
    }
    return jsonOut_({ error: 'unknown_action', message: `Action "${action}" is not supported.` }, 400);

  } catch (err) {
    return jsonOut_({ error: 'server_error', message: String(err) }, 500);
  }
}

/** Temporary wrapper — if syncCredentials_ won't show in the function
 *  dropdown, run this instead (does the exact same thing). Safe to leave
 *  in place, or delete once syncCredentials_ shows up normally. */
function runSync(){ syncCredentials_(); }

// ═══ IDENTITY ═══

/** Username/password only — see the file header for why Google auto-identity doesn't work here. */
function resolveIdentity_(params) {
  if (!params.username || !params.password) return null;

  const cred = checkCredentials_(params.username, params.password);
  if (!cred || cred.role !== 'READ') return null;

  // Credentials only proves who they are. Their actual scope (single property,
  // market cluster, or full portfolio) is still looked up from Permissions,
  // keyed by the same email — one source of truth for "what can this person see."
  const grant = lookupPermission_(cred.username);
  if (!grant) return null;
  return { label: cred.username, grant: grant };
}

function lookupPermission_(email) {
  const sheet = SpreadsheetApp.openById(PERMISSIONS_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());
  const emailCol = header.indexOf('email');
  const typeCol = header.indexOf('scope_type');
  const valCol = header.indexOf('scope_value');
  const roleCol = header.indexOf('role');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[emailCol]).toLowerCase().trim() === email) {
      return {
        scope_type: String(row[typeCol]).trim(),
        scope_value: String(row[valCol]).trim(),
        role: String(row[roleCol]).trim()
      };
    }
  }
  return null;
}

// ═══ CREDENTIALS (username/password path — auditors & franchise GMs) ═══

/** Returns {username, property_id, role} if the login is valid and active, else null. */
function checkCredentials_(username, password) {
  const sheet = SpreadsheetApp.openById(CREDENTIALS_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());
  const userCol = header.indexOf('username');
  const hashCol = header.indexOf('password_hash');
  const saltCol = header.indexOf('salt');
  const propCol = header.indexOf('property_id');
  const roleCol = header.indexOf('role');
  const activeCol = header.indexOf('active');

  const uname = String(username).toLowerCase().trim();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[userCol]).toLowerCase().trim() !== uname) continue;
    if (String(row[activeCol]).toUpperCase().trim() !== 'TRUE') return null;
    const salt = String(row[saltCol]);
    const hash = String(row[hashCol]);
    if (!salt || !hash) return null; // not synced yet
    if (hashPassword_(password, salt) !== hash) return null;
    return {
      username: String(row[userCol]),
      property_id: String(row[propCol]),
      role: String(row[roleCol]).trim().toUpperCase()
    };
  }
  return null;
}

function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + ':' + salt);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function randomSalt_() {
  return Utilities.getUuid();
}

/** Run manually from the Apps Script editor (function dropdown > syncCredentials_ > Run)
 *  whenever you add, change, or revoke a login. Hashes any plaintext new_password
 *  cells and clears them so plaintext never sits in the sheet for long. */
function syncCredentials_() {
  const sheet = SpreadsheetApp.openById(CREDENTIALS_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());
  const userCol = header.indexOf('username');
  const pwCol = header.indexOf('new_password');
  const hashCol = header.indexOf('password_hash');
  const saltCol = header.indexOf('salt');
  const activeCol = header.indexOf('active');
  const syncCol = header.indexOf('last_synced');

  let synced = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const plain = row[pwCol];
    if (!plain) continue;
    const salt = randomSalt_();
    const hash = hashPassword_(String(plain), salt);
    sheet.getRange(i + 1, saltCol + 1).setValue(salt);
    sheet.getRange(i + 1, hashCol + 1).setValue(hash);
    sheet.getRange(i + 1, pwCol + 1).setValue(''); // wipe plaintext
    if (String(row[activeCol]).toUpperCase().trim() !== 'FALSE') {
      sheet.getRange(i + 1, activeCol + 1).setValue('TRUE');
    }
    sheet.getRange(i + 1, syncCol + 1).setValue(new Date());
    synced++;
  }
  // Run from the Apps Script editor (no Sheets UI context available there),
  // so report via the execution log rather than a UI alert.
  Logger.log(synced + ' credential(s) synced. Plaintext passwords cleared from the sheet.');
}

// ═══ DATA ACCESS ═══

function getAuthorizedProperties_(grant) {
  const sheet = SpreadsheetApp.openById(PROPERTIES_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());

  const rows = data.slice(1).map(r => {
    const obj = {};
    header.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });

  if (grant.scope_type === 'portfolio') return rows;
  if (grant.scope_type === 'market') {
    return rows.filter(r => String(r.market).toLowerCase() === grant.scope_value.toLowerCase());
  }
  return rows.filter(r => String(r.property_id) === grant.scope_value);
}

function getFlashRows_(propertyIds) {
  if (!propertyIds.length) return [];
  const idSet = new Set(propertyIds.map(String));
  const sheet = SpreadsheetApp.openById(FLASH_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());
  const idCol = header.indexOf('property_id');

  return data.slice(1)
    .filter(r => idSet.has(String(r[idCol])))
    .map(r => {
      const obj = {};
      header.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
}

/** Night-audit submission. Property comes from the caller's credential, never from the payload. */
function handleFlashSubmission_(body) {
  const cred = checkCredentials_(body.username, body.password);
  if (!cred) return { error: 'not_authorized', message: 'Invalid or inactive credentials.', _status: 401 };
  if (cred.role !== 'ENTRY') return { error: 'not_authorized', message: 'This login is not permitted to submit data.', _status: 403 };

  const monthIndex = Number(body.month_index);
  if (!(monthIndex >= 0 && monthIndex <= 11)) {
    return { error: 'bad_request', message: 'month_index must be 0-11.', _status: 400 };
  }
  const roomRev = Number(body.room_rev) || 0;
  const occupiedRooms = Number(body.occupied_rooms) || 0;
  const totalRev = Number(body.total_rev) || 0;

  const sheet = SpreadsheetApp.openById(FLASH_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());
  const idCol = header.indexOf('property_id');
  const monthCol = header.indexOf('month_index');
  const rvCol = header.indexOf('room_rev');
  const ocCol = header.indexOf('occupied_rooms');
  const trCol = header.indexOf('total_rev');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === cred.property_id && Number(data[i][monthCol]) === monthIndex) {
      sheet.getRange(i + 1, rvCol + 1).setValue(roomRev);
      sheet.getRange(i + 1, ocCol + 1).setValue(occupiedRooms);
      sheet.getRange(i + 1, trCol + 1).setValue(totalRev);
      return { ok: true, property_id: cred.property_id, month_index: monthIndex, updated: true };
    }
  }
  return { error: 'not_found', message: 'No matching property/month row in Portfolio Flash Data to update.', _status: 404 };
}

// ═══ ENTRY FORM (served HTML, no separate hosting needed) ═══

function renderEntryFormHtml_() {
  return `<!DOCTYPE html><html><head><base target="_top">
<style>
body{font-family:Arial,sans-serif;max-width:420px;margin:40px auto;padding:0 16px;color:#222}
h2{font-size:18px}
label{display:block;margin-top:12px;font-size:13px;font-weight:bold}
input{width:100%;padding:8px;font-size:14px;box-sizing:border-box;margin-top:4px}
button{margin-top:20px;padding:10px 16px;font-size:14px;background:#1a5;color:#fff;border:none;cursor:pointer}
#msg{margin-top:16px;font-size:13px}
.err{color:#c0392b} .ok{color:#1a5}
</style></head><body>
<h2>RGB Intelligence Hub — Night Audit Entry</h2>
<div id="loginBox">
<label>Username <input id="u"></label>
<label>Password <input id="p" type="password"></label>
</div>
<div id="entryBox" style="display:none">
<label>Month (0=Jan ... 11=Dec) <input id="m" type="number" min="0" max="11"></label>
<label>Room Revenue ($) <input id="rv" type="number" step="0.01"></label>
<label>Occupied Room-Nights <input id="oc" type="number"></label>
<label>Total Revenue ($) <input id="tr" type="number" step="0.01"></label>
</div>
<button onclick="submitForm()">Submit</button>
<div id="msg"></div>
<script>
function submitForm(){
  var body={
    action:'submit_flash',
    username:document.getElementById('u').value,
    password:document.getElementById('p').value,
    month_index:document.getElementById('m').value,
    room_rev:document.getElementById('rv').value,
    occupied_rooms:document.getElementById('oc').value,
    total_rev:document.getElementById('tr').value
  };
  document.getElementById('entryBox').style.display='block';
  document.getElementById('msg').textContent='Submitting...';
  document.getElementById('msg').className='';
  fetch(window.location.href.split('?')[0], {method:'POST', body:JSON.stringify(body)})
    .then(function(r){return r.json();})
    .then(function(j){
      if(j.error){ document.getElementById('msg').textContent='Error: '+j.message; document.getElementById('msg').className='err'; }
      else { document.getElementById('msg').textContent='Saved.'; document.getElementById('msg').className='ok'; }
    })
    .catch(function(err){ document.getElementById('msg').textContent='Network error: '+err; document.getElementById('msg').className='err'; });
}
</script>
</body></html>`;
}

function jsonOut_(obj, statusCode) {
  obj._status = statusCode;
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
