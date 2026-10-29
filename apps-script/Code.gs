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
 *      Budget 2026     property_id, name, month_index (0-11), budget_room_rev
 *      Historical      property_id, name, year (2024|2025), month_index (0-11),
 *                      room_rev, occupied_rooms — feeds the Portfolio Flash
 *                      screen's Year-over-Year / vs-Budget comparison toggle
 *      Rate Shop Data  property_name, date, rate_display, rate_num — a flat
 *                      rate lookup table, populated by pasting the OTA
 *                      Insight "Rates" export into that spreadsheet's
 *                      "Raw Paste" tab and running runParseRateShop (see the
 *                      "RATE SHOP" section below)
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
const DAR_SHEET_ID         = '1298lXdy5cfNNo3HMISjI2wMn2Zu0dnUX5_a0a0qbafQ'; // RGB Intel Hub — DAR Data
const DAR_LOG_SHEET_ID     = '1ORM876qEYclj3jFlSIJe_UW2gsxluEC4Sedi17yV4JY'; // RGB Intel Hub — DAR Submission Log (append-only, never overwritten)
const AVAILABILITY_SHEET_ID = '1I61_RQ7NLmjGJHyRtXESfIzZUHNqlKVVkH3mF92J8m4'; // RGB Intel Hub — Availability
const AVAILABILITY_WINDOW_DAYS = 21; // auditors enter/update this many days forward each night
const BUDGET_SHEET_ID       = '1uj3wiFtnyWOoWnJpwyGhzrRDujp4ir9l1gfZb3pYOFY'; // RGB Intel Hub — Budget 2026
                                                                               //   property_id, name, month_index, budget_room_rev
const HISTORICAL_SHEET_ID   = '1pDYuQkJWvDBCt9rk_XWXfaLgSVF2pRkN8sYM-vuNVXE'; // RGB Intel Hub — Historical Actuals
                                                                               //   property_id, name, year, month_index, room_rev, occupied_rooms
const RATE_SHOP_SHEET_ID    = '1JD2yi77dgUhFpWPkwCxRiYKYJ1RDEVDoSiHLq3eIHZc'; // RGB Intel Hub — Rate Shop Data
                                                                               //   property_name, date, rate_display, rate_num — a flat
                                                                               //   rate lookup table. Cluster/RGB-vs-comp assignment lives
                                                                               //   in CLUSTER_DEFS on the frontend, not here — see the
                                                                               //   "RATE SHOP" section below for why and how this is
                                                                               //   populated each morning.

// Never assume the data lives on getSheets()[0] — a Read Me tab (or anything
// else) inserted ahead of it silently breaks that. This picks the first tab
// that isn't a known non-data tab by name. If you add a personal/reference
// tab to any of these spreadsheets, name it something in NON_DATA_TAB_NAMES
// (or add your name here) so it's automatically skipped.
const NON_DATA_TAB_NAMES = new Set(['read me', 'readme', 'notes', 'raw paste']);
function firstDataSheet_(spreadsheetId) {
  const sheets = SpreadsheetApp.openById(spreadsheetId).getSheets();
  const dataSheet = sheets.find(s => !NON_DATA_TAB_NAMES.has(s.getName().trim().toLowerCase()));
  if (!dataSheet) throw new Error(`No data tab found in spreadsheet ${spreadsheetId} — every tab matched a non-data name.`);
  return dataSheet;
}

// Properties whose DAR includes actual restaurant revenue (vs. just a small
// self-serve market). Hardcoded for now since the Properties sheet doesn't
// have a column for it yet — move this to a real "has_restaurant" column on
// Properties if the list grows or needs editing without a code change.
const RESTAURANT_PROPERTY_IDS = new Set(['PROV-AUSES', 'PROV-HOUES', 'PROV-DALES']);
function propertyHasRestaurant_(propertyId) { return RESTAURANT_PROPERTY_IDS.has(String(propertyId)); }

// Every field an auditor can enter on the DAR, grouped into form sections.
// restaurantOnly fields are hidden entirely for properties without a restaurant.
const DAR_FIELDS = [
  { key: 'rooms_rented', label: 'Rooms Rented', section: 'Room Revenue' },
  { key: 'room_rental_revenue', label: 'Room Rental Revenue', section: 'Room Revenue' },
  { key: 'dry_cleaning_laundry', label: 'Dry Cleaning / Laundry', section: 'Room Revenue' },
  { key: 'parking', label: 'Parking', section: 'Room Revenue' },
  { key: 'early_late_fee', label: 'Early/Late Fee', section: 'Room Revenue' },
  { key: 'hotel_misc', label: 'Hotel Misc.', section: 'Room Revenue' },
  { key: 'data_internet_charge', label: 'Data Service / Internet Charge', section: 'Room Revenue' },
  { key: 'market_food', label: 'Market Food/Sundry', section: 'Market' },
  { key: 'market_beverage', label: 'Market Beverage', section: 'Market' },
  { key: 'market_beer', label: 'Market Beer', section: 'Market' },
  { key: 'market_wine', label: 'Market Wine', section: 'Market' },
  { key: 'restaurant_food', label: 'Restaurant Food', section: 'Restaurant', restaurantOnly: true },
  { key: 'restaurant_liquor', label: 'Restaurant Liquor', section: 'Restaurant', restaurantOnly: true },
  { key: 'restaurant_beer', label: 'Restaurant Beer', section: 'Restaurant', restaurantOnly: true },
  { key: 'restaurant_wine', label: 'Restaurant Wine', section: 'Restaurant', restaurantOnly: true },
  { key: 'restaurant_food_tax', label: 'Restaurant Food Tax', section: 'Restaurant', restaurantOnly: true },
  { key: 'meeting_room_rental', label: 'Meeting Room Rental', section: 'Meeting & Banquet' },
  { key: 'banquet_food', label: 'Banquet Food', section: 'Meeting & Banquet' },
  { key: 'banquet_misc', label: 'Banquet Misc.', section: 'Meeting & Banquet' },
  { key: 'banquet_service_charge', label: 'Banquet Service Charge', section: 'Meeting & Banquet' },
  { key: 'state_meeting_room_tax', label: 'State Meeting Room Tax', section: 'Meeting & Banquet' },
  { key: 'meeting_room_tax_exempt', label: 'Meeting Room Tax Exempt', section: 'Meeting & Banquet' },
  { key: 'state_hotel_tax', label: 'State Hotel Tax', section: 'Taxes' },
  { key: 'city_hotel_tax', label: 'City Hotel Tax', section: 'Taxes' },
  { key: 'county_tax', label: 'County Tax', section: 'Taxes' },
  { key: 'sales_tax', label: 'Sales Tax', section: 'Taxes' },
  { key: 'tabc_sales_tax', label: 'TABC Sales Tax', section: 'Taxes' },
  { key: 'state_cost_recovery_tax', label: 'State Cost Recovery Tax', section: 'Taxes', restaurantOnly: true }
];

// DEFERRED — not on the entry form yet (accounting reconciliation, not
// performance data, and the settlement/ledger structure genuinely differs
// by brand: Hilton uses City/Tray Ledger + Closed Folio; IHG uses Guest
// Ledger + Accounts Receivable + Advance Deposit + a separate POS terminal
// breakdown; Marriott/Wyndham/Hyatt not yet seen at all). The DAR Data sheet
// already has columns reserved for all of these — enabling this section
// later is just uncommenting entries here (per-brand if the structures
// don't reconcile into one shared list) plus a UI pass, no schema change.
// const DAR_FIELDS_DEFERRED = [
//   { key: 'employee_accrual', label: 'Employee Accrual', section: 'Other' },
//   { key: 'telephone_allowance', label: 'Telephone Allowance', section: 'Other' },
//   { key: 'gtd_allowance_100pct', label: '100% GTD Allowance', section: 'Other' },
//   { key: 'advance_purchase_allowance', label: 'Advance Purchase Allowance', section: 'Other' },
//   { key: 'writeoffs', label: 'Writeoffs', section: 'Other' },
//   { key: 'tips_payable', label: 'Tips Payable', section: 'Other', restaurantOnly: true },
//   { key: 'paid_outs', label: 'Paid Outs', section: 'Other' },
//   { key: 'bank_deposits', label: 'Bank Deposits', section: 'Other' },
//   { key: 'loyalty_advance_purchase', label: 'Loyalty / Advance Purchase Settlement (brand-specific label)', section: 'Settlement & Ledger' },
//   { key: 'amex', label: 'Amex', section: 'Settlement & Ledger' },
//   { key: 'visa_mc', label: 'Visa / MC', section: 'Settlement & Ledger' },
//   { key: 'discover', label: 'Discover', section: 'Settlement & Ledger' },
//   { key: 'other_card_allowance', label: 'Other Card Allowance', section: 'Settlement & Ledger' },
//   { key: 'over_short', label: 'Over / Short', section: 'Settlement & Ledger' },
//   { key: 'city_ledger', label: 'City Ledger Balance', section: 'Settlement & Ledger' },
//   { key: 'tray_guest_ledger', label: 'Guest (Tray) Ledger Balance', section: 'Settlement & Ledger' },
//   { key: 'closed_folio_group_master', label: 'Closed Folio / Group Master', section: 'Settlement & Ledger' }
// ];

// ═══ ENTRY POINTS ═══

function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.view === 'entry') {
    return HtmlService.createHtmlOutput(renderEntryFormHtml_())
      .setTitle('RGB Intelligence Hub — Data Entry')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (params.view === 'dar-entry') {
    return HtmlService.createHtmlOutput(renderDarEntryFormHtml_())
      .setTitle('RGB Intelligence Hub — DAR Entry')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (params.view === 'dar-review') {
    return HtmlService.createHtmlOutput(renderDarReviewFormHtml_())
      .setTitle('RGB Intelligence Hub — DAR Review')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  try {
    if (params.action === 'whoami') return jsonOut_(handleWhoami_(params), 200);
    if (params.action === 'dar_get') return jsonOut_(handleDarGet_(params), 200);
    if (params.action === 'dar_pending') return jsonOut_(handleDarPending_(params), 200);
    if (params.action === 'availability_get') return jsonOut_(handleAvailabilityGet_(params), 200);
    if (params.action === 'forecast_summary') return jsonOut_(handleForecastSummary_(params), 200);
    if (params.action === 'forecast_get') return jsonOut_(handleForecastGet_(params), 200);

    const identity = resolveIdentity_(params);
    if (!identity) {
      return jsonOut_({ error: 'no_identity', message: 'Could not verify identity. Google sign-in not detected and no valid username/password provided.' }, 401);
    }

    const module = params.module || 'portfolio-flash';

    if (module === 'rate-shop') {
      return jsonOut_({ caller: identity.label, scope: identity.grant, rateShop: getRateShopRows_() }, 200);
    }
    if (module !== 'portfolio-flash') {
      return jsonOut_({ error: 'unknown_module', message: `Module "${module}" is not wired to the API yet.` }, 400);
    }

    const properties = getAuthorizedProperties_(identity.grant);
    const idList = properties.map(p => p.property_id);
    const flashRows = getFlashRows_(idList);
    const budgetRows = getBudgetRows_(idList);
    const historicalRows = getHistoricalRows_(idList);

    return jsonOut_({
      caller: identity.label,
      scope: identity.grant,
      properties: properties,
      flash: flashRows,
      budget: budgetRows,
      historical: historicalRows
    }, 200);

  } catch (err) {
    return jsonOut_({ error: 'server_error', message: String(err) }, 500);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = body.action || 'submit_flash';

    if (action === 'submit_flash') return jsonOut_(handleFlashSubmission_(body), 200);
    if (action === 'submit_dar') return jsonOut_(handleDarSubmission_(body), 200);
    if (action === 'validate_dar') return jsonOut_(handleDarValidation_(body, 'VALIDATED'), 200);
    if (action === 'flag_dar') return jsonOut_(handleDarValidation_(body, 'FLAGGED'), 200);
    if (action === 'submit_availability') return jsonOut_(handleAvailabilitySubmission_(body), 200);
    if (action === 'forecast_save') return jsonOut_(handleForecastSave_(body), 200);
    if (action === 'forecast_submit') return jsonOut_(handleForecastSubmit_(body), 200);
    if (action === 'forecast_review') return jsonOut_(handleForecastReview_(body), 200);
    return jsonOut_({ error: 'unknown_action', message: `Action "${action}" is not supported.` }, 400);

  } catch (err) {
    return jsonOut_({ error: 'server_error', message: String(err) }, 500);
  }
}

/** Temporary wrapper — if syncCredentials_ won't show in the function
 *  dropdown, run this instead (does the exact same thing). Safe to leave
 *  in place, or delete once syncCredentials_ shows up normally. */
function runSync(){ syncCredentials_(); }

/** Wrapper for addReadmeTab_ — the Run dropdown hides any function whose name
 *  ends in "_" (that's the "private helper" naming convention), so this
 *  plain-named wrapper is what actually shows up to click Run on. */
function runAddReadmeTab(){ addReadmeTab_(); }

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
  const sheet = firstDataSheet_(PERMISSIONS_SHEET_ID);
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
  const sheet = firstDataSheet_(CREDENTIALS_SHEET_ID);
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
  const sheet = firstDataSheet_(CREDENTIALS_SHEET_ID);
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
  const sheet = firstDataSheet_(PROPERTIES_SHEET_ID);
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
  const sheet = firstDataSheet_(FLASH_SHEET_ID);
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

/** Budget 2026 rows (property_id, name, month_index, budget_room_rev) for the given property IDs only. */
function getBudgetRows_(propertyIds) {
  if (!propertyIds.length) return [];
  const idSet = new Set(propertyIds.map(String));
  const sheet = firstDataSheet_(BUDGET_SHEET_ID);
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

/** 2024/2025 actuals (property_id, name, year, month_index, room_rev, occupied_rooms) for the given property IDs only. */
function getHistoricalRows_(propertyIds) {
  if (!propertyIds.length) return [];
  const idSet = new Set(propertyIds.map(String));
  const sheet = firstDataSheet_(HISTORICAL_SHEET_ID);
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

  const sheet = firstDataSheet_(FLASH_SHEET_ID);
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

function getPropertyById_(propertyId) {
  const sheet = firstDataSheet_(PROPERTIES_SHEET_ID);
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());
  const idCol = header.indexOf('property_id');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(propertyId)) {
      const obj = {};
      header.forEach((h, j) => obj[h] = data[i][j]);
      return obj;
    }
  }
  return null;
}

function isAuthorizedForProperty_(grant, propertyId) {
  if (grant.scope_type === 'portfolio') return true;
  if (grant.scope_type === 'property') return grant.scope_value === String(propertyId);
  if (grant.scope_type === 'market') {
    const prop = getPropertyById_(propertyId);
    return !!prop && String(prop.market).toLowerCase() === grant.scope_value.toLowerCase();
  }
  return false;
}

// ═══ DAR (Daily Audit Report) ═══

/** Identity check used by the DAR entry/review pages before rendering their form fields. */
function handleWhoami_(params) {
  const cred = checkCredentials_(params.username, params.password);
  if (!cred) return { error: 'not_authorized', message: 'Invalid or inactive credentials.', _status: 401 };

  if (cred.role === 'ENTRY') {
    const prop = getPropertyById_(cred.property_id);
    return {
      role: 'ENTRY',
      property_id: cred.property_id,
      property_name: prop ? prop.name : cred.property_id,
      has_restaurant: propertyHasRestaurant_(cred.property_id)
    };
  }

  const grant = lookupPermission_(cred.username);
  if (!grant) return { error: 'not_authorized', message: 'Not on the Permissions sheet.', _status: 403 };
  const properties = getAuthorizedProperties_(grant);
  return { role: 'READ', username: cred.username, scope: grant, properties: properties };
}

function readDarSheet_() {
  const sheet = firstDataSheet_(DAR_SHEET_ID);
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());
  return { sheet, data, header };
}

function darRowToObject_(header, row) {
  const obj = {};
  header.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

function findDarRowIndex_(data, header, propertyId, date) {
  const idCol = header.indexOf('property_id');
  const dateCol = header.indexOf('date');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(propertyId) && String(data[i][dateCol]) === date) return i;
  }
  return -1;
}

/** Prefill for the entry form, or read-only detail for the review page. */
function handleDarGet_(params) {
  const cred = checkCredentials_(params.username, params.password);
  if (!cred) return { error: 'not_authorized', message: 'Invalid or inactive credentials.', _status: 401 };

  let propertyId;
  if (cred.role === 'ENTRY') {
    propertyId = cred.property_id;
  } else {
    const grant = lookupPermission_(cred.username);
    if (!grant) return { error: 'not_authorized', message: 'Not on the Permissions sheet.', _status: 403 };
    propertyId = params.property_id;
    if (!propertyId || !isAuthorizedForProperty_(grant, propertyId)) {
      return { error: 'not_authorized', message: 'Not authorized for this property.', _status: 403 };
    }
  }

  const date = String(params.date || '').trim();
  const { data, header } = readDarSheet_();
  const rowIndex = findDarRowIndex_(data, header, propertyId, date);
  return {
    property_id: propertyId,
    date: date,
    exists: rowIndex >= 0,
    row: rowIndex >= 0 ? darRowToObject_(header, data[rowIndex]) : null
  };
}

/** List of SUBMITTED (pending-review) DAR rows across whatever properties this login can see. */
function handleDarPending_(params) {
  const cred = checkCredentials_(params.username, params.password);
  if (!cred) return { error: 'not_authorized', message: 'Invalid or inactive credentials.', _status: 401 };
  if (cred.role !== 'READ') return { error: 'not_authorized', message: 'This login cannot review DARs.', _status: 403 };

  const grant = lookupPermission_(cred.username);
  if (!grant) return { error: 'not_authorized', message: 'Not on the Permissions sheet.', _status: 403 };

  const authorizedIds = new Set(getAuthorizedProperties_(grant).map(p => String(p.property_id)));
  const { data, header } = readDarSheet_();
  const idCol = header.indexOf('property_id');
  const statusCol = header.indexOf('status');

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (!authorizedIds.has(String(data[i][idCol]))) continue;
    if (String(data[i][statusCol]) !== 'SUBMITTED') continue;
    rows.push(darRowToObject_(header, data[i]));
  }
  return { rows };
}

/** Revenue-generating lines only — the totals below intentionally reference every
 *  possible field (including ones not yet on the form) so turning on a deferred
 *  field later doesn't require touching this formula. */
function computeDarTotals_(v) {
  const n = k => Number(v[k]) || 0;
  const total_revenue = n('room_rental_revenue') + n('dry_cleaning_laundry') + n('parking') + n('early_late_fee') +
    n('hotel_misc') + n('data_internet_charge') + n('market_food') + n('market_beverage') + n('market_beer') +
    n('market_wine') + n('restaurant_food') + n('restaurant_liquor') + n('restaurant_beer') + n('restaurant_wine') +
    n('meeting_room_rental') + n('banquet_food') + n('banquet_misc');
  const total_non_revenue = n('state_hotel_tax') + n('city_hotel_tax') + n('county_tax') + n('sales_tax') +
    n('tabc_sales_tax') + n('state_meeting_room_tax') + n('state_cost_recovery_tax') + n('restaurant_food_tax');
  const total_allowances = n('telephone_allowance') + n('gtd_allowance_100pct') + n('advance_purchase_allowance');
  const total_deposits = n('loyalty_advance_purchase') + n('amex') + n('visa_mc') + n('discover') +
    n('other_card_allowance') + n('over_short') + n('bank_deposits');
  return { total_revenue, total_non_revenue, total_allowances, total_deposits };
}

function handleDarSubmission_(body) {
  const cred = checkCredentials_(body.username, body.password);
  if (!cred) return { error: 'not_authorized', message: 'Invalid or inactive credentials.', _status: 401 };
  if (cred.role !== 'ENTRY') return { error: 'not_authorized', message: 'This login is not permitted to submit DAR data.', _status: 403 };

  const date = String(body.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return { error: 'bad_request', message: 'date must be YYYY-MM-DD.', _status: 400 };
  }

  const { sheet, data, header } = readDarSheet_();
  const rowIndex = findDarRowIndex_(data, header, cred.property_id, date);
  if (rowIndex >= 0 && String(data[rowIndex][header.indexOf('status')]) === 'VALIDATED') {
    return { error: 'already_validated', message: 'This date was already approved by the GM and can no longer be resubmitted. Contact your GM if it needs correction.', _status: 409 };
  }

  const values = {};
  DAR_FIELDS.forEach(f => { values[f.key] = Number(body[f.key]) || 0; });
  const totals = computeDarTotals_(values);
  const now = new Date();

  const rowData = header.map(h => {
    if (h === 'property_id') return cred.property_id;
    if (h === 'date') return date;
    if (h in values) return values[h];
    if (h in totals) return totals[h];
    if (h === 'submitted_by') return cred.username;
    if (h === 'submitted_at') return now;
    if (h === 'status') return 'SUBMITTED';
    return ''; // validated_by, validated_at, flag_note, and any deferred field not in this submission
  });

  if (rowIndex >= 0) {
    sheet.getRange(rowIndex + 1, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  // Permanent, append-only record of this exact submission — kept even if a
  // later resubmission overwrites the "current" row above. Never updated or
  // deleted by any code path; this is the answer to "what did the auditor
  // actually submit, and when" regardless of what happened afterward.
  logDarSubmission_(cred.property_id, date, cred.username, values, totals);

  return { ok: true, property_id: cred.property_id, date: date, status: 'SUBMITTED', totals: totals };
}

function logDarSubmission_(propertyId, date, submittedBy, values, totals) {
  const sheet = firstDataSheet_(DAR_LOG_SHEET_ID);
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
  const row = header.map(h => {
    if (h === 'logged_at') return new Date();
    if (h === 'property_id') return propertyId;
    if (h === 'date') return date;
    if (h === 'submitted_by') return submittedBy;
    if (h in values) return values[h];
    if (h in totals) return totals[h];
    return '';
  });
  sheet.appendRow(row);
}

/** GM (or corporate) approves or flags a SUBMITTED DAR. Read-only review — no direct edits. */
function handleDarValidation_(body, newStatus) {
  const cred = checkCredentials_(body.username, body.password);
  if (!cred) return { error: 'not_authorized', message: 'Invalid or inactive credentials.', _status: 401 };
  if (cred.role !== 'READ') return { error: 'not_authorized', message: 'This login cannot validate DARs.', _status: 403 };

  const grant = lookupPermission_(cred.username);
  if (!grant || !isAuthorizedForProperty_(grant, body.property_id)) {
    return { error: 'not_authorized', message: 'Not authorized for this property.', _status: 403 };
  }

  const date = String(body.date || '').trim();
  const { sheet, data, header } = readDarSheet_();
  const rowIndex = findDarRowIndex_(data, header, body.property_id, date);
  if (rowIndex < 0) return { error: 'not_found', message: 'No DAR submission found for that property/date.', _status: 404 };
  if (String(data[rowIndex][header.indexOf('status')]) !== 'SUBMITTED') {
    return { error: 'bad_state', message: 'This DAR is not pending review.', _status: 409 };
  }

  const statusCol = header.indexOf('status');
  const validatedByCol = header.indexOf('validated_by');
  const validatedAtCol = header.indexOf('validated_at');
  const flagNoteCol = header.indexOf('flag_note');

  sheet.getRange(rowIndex + 1, statusCol + 1).setValue(newStatus);
  sheet.getRange(rowIndex + 1, validatedByCol + 1).setValue(cred.username);
  sheet.getRange(rowIndex + 1, validatedAtCol + 1).setValue(new Date());
  sheet.getRange(rowIndex + 1, flagNoteCol + 1).setValue(newStatus === 'FLAGGED' ? String(body.note || '') : '');

  return { ok: true, property_id: body.property_id, date: date, status: newStatus };
}

// ═══ AVAILABILITY (rooms to sell, rolling window — feeds Rate Shop) ═══
// Replaces someone manually logging into 20 PMS systems each morning: the
// auditor enters/updates this alongside their DAR, one number per day for
// the next AVAILABILITY_WINDOW_DAYS days. Comp-set rates themselves still
// come from the OTA Insight upload — this only covers the inventory half.

function nextNDates_(n) {
  const dates = [];
  const tz = Session.getScriptTimeZone();
  for (let i = 0; i < n; i++) {
    const d = new Date();
    d.setDate(d.getDate() + i);
    dates.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
  }
  return dates;
}

/** Prefill for the auditor's availability grid: today + next N-1 days, with whatever's already on file. */
function handleAvailabilityGet_(params) {
  const cred = checkCredentials_(params.username, params.password);
  if (!cred) return { error: 'not_authorized', message: 'Invalid or inactive credentials.', _status: 401 };
  if (cred.role !== 'ENTRY') return { error: 'not_authorized', message: 'This login cannot view availability.', _status: 403 };

  const sheet = firstDataSheet_(AVAILABILITY_SHEET_ID);
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());
  const idCol = header.indexOf('property_id');
  const dateCol = header.indexOf('date');
  const roomsCol = header.indexOf('rooms_available');

  const existing = {};
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === cred.property_id) {
      existing[String(data[i][dateCol])] = data[i][roomsCol];
    }
  }

  const dates = nextNDates_(AVAILABILITY_WINDOW_DAYS);
  return {
    property_id: cred.property_id,
    days: dates.map(d => ({ date: d, rooms_available: existing[d] != null ? existing[d] : '' }))
  };
}

/** Upserts a batch of {date, rooms_available} entries for the auditor's own property in one call. */
function handleAvailabilitySubmission_(body) {
  const cred = checkCredentials_(body.username, body.password);
  if (!cred) return { error: 'not_authorized', message: 'Invalid or inactive credentials.', _status: 401 };
  if (cred.role !== 'ENTRY') return { error: 'not_authorized', message: 'This login is not permitted to submit availability.', _status: 403 };

  const days = Array.isArray(body.days) ? body.days : [];
  if (!days.length) return { error: 'bad_request', message: 'No days submitted.', _status: 400 };

  const sheet = firstDataSheet_(AVAILABILITY_SHEET_ID);
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());
  const idCol = header.indexOf('property_id');
  const dateCol = header.indexOf('date');
  const roomsCol = header.indexOf('rooms_available');
  const byCol = header.indexOf('submitted_by');
  const atCol = header.indexOf('submitted_at');

  const now = new Date();
  let updated = 0, inserted = 0;

  days.forEach(d => {
    const date = String(d.date || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return; // skip malformed entries silently rather than fail the whole batch
    const rooms = Number(d.rooms_available);
    if (isNaN(rooms)) return;

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === cred.property_id && String(data[i][dateCol]) === date) { rowIndex = i; break; }
    }
    if (rowIndex >= 0) {
      sheet.getRange(rowIndex + 1, roomsCol + 1).setValue(rooms);
      sheet.getRange(rowIndex + 1, byCol + 1).setValue(cred.username);
      sheet.getRange(rowIndex + 1, atCol + 1).setValue(now);
      updated++;
    } else {
      const row = header.map(h => {
        if (h === 'property_id') return cred.property_id;
        if (h === 'date') return date;
        if (h === 'rooms_available') return rooms;
        if (h === 'submitted_by') return cred.username;
        if (h === 'submitted_at') return now;
        return '';
      });
      sheet.appendRow(row);
      data.push(row); // keep our in-memory copy consistent for subsequent iterations of this same batch
      inserted++;
    }
  });

  return { ok: true, property_id: cred.property_id, updated: updated, inserted: inserted };
}

// ═══ RATE SHOP (comp-set rates, OTA Insight "Rates" import) ═══
// The cluster list, property display names, and RGB/Compset table format on
// the Rate Shop screen are the original design, unchanged — see CLUSTER_DEFS
// in rgb-intelligence-hub-v4.html. That's ALSO where each property's raw
// OTA Insight name and RGB-vs-comp classification live, not here: the export
// itself flags some of your own properties as "comp" (its "My Hotels" list is
// incomplete — 5 confirmed short as of the export this was built from:
// Residence Inn Downtown, Hampton Navigation, TRU Corpus, Staybridge Portland,
// La Quinta Portland), so this pipeline never trusts the file's own RGB/comp
// signal — it's a flat name → rate lookup table, nothing more. Whether a
// property is RGB is decided once, by hand, in CLUSTER_DEFS.
//
// The "Rate" data is genuinely external (competitor pricing OTA Insight
// scrapes) — pulling that export each morning isn't going away. This
// pipeline only replaces hand-editing a JS object every time it changes.
//
// MORNING WORKFLOW
// 1. Open the "RGB Intel Hub — Rate Shop Data" spreadsheet. If there's no
//    "Raw Paste" tab yet, add one (right-click any tab > Insert sheet >
//    rename it "Raw Paste") — only needs doing once, ever.
// 2. Export OTA Insight's "Rates" report (Best flexible rate / Brand.com,
//    whatever LOS/guest count you use) and paste it into Raw Paste starting
//    at cell A1, values only.
// 3. In the Apps Script editor: function dropdown > runParseRateShop > Run.
// 4. Done — the Hub's Rate Shop screen reads the parsed result immediately,
//    no redeploy needed.

function firstOrCreateSheet_(spreadsheetId, name) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** Wrapper — see firstDataSheet_'s header comment for why this exists (the
 *  Run dropdown hides any function whose name ends in "_"). */
function runParseRateShop(){ parseRateShopImport_(); }

/** Recognizes a date in an OTA Insight "Rates" header cell and normalizes it
 *  to "yyyy-MM-dd", or returns null if the cell isn't a date at all. Handles
 *  two shapes because Excel-to-Sheets paste can produce either: the literal
 *  text "Thu 27 Aug 2026" (a weekday name, day, month name, year), OR a real
 *  Date value if Sheets auto-detected that text as a date during paste (a
 *  common paste-from-Excel quirk — not a locale-dependent guess, since a
 *  genuine Date object here unambiguously came from that recognition, not
 *  from parsing arbitrary user text). */
function parseRatesDateCell_(cell) {
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const m = /^[A-Za-z]{3}\s+(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})$/.exec(String(cell || '').trim());
  if (!m) return null;
  const months = { Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11 };
  const monthIdx = months[m[2]];
  if (monthIdx === undefined) return null;
  const d = new Date(Number(m[3]), monthIdx, Number(m[1]));
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

const RATE_SHOP_SKIP_NAMES = new Set(['holidays', 'events']); // annotation rows in the export, not properties

/** Parses the flat OTA Insight "Rates" export (Raw Paste tab) — one row per
 *  property, a rate column per date (RGB-flagged properties also carry a
 *  % change column next to it, which this ignores — see this section's
 *  header comment for why) — into normalized rows on the Rate Shop Data tab.
 *  Re-run each morning after pasting a fresh export; this overwrites the
 *  prior snapshot rather than appending, since Rate Shop is "what does the
 *  market look like right now," not a history to preserve.
 *
 *  Finds date COLUMNS by scanning header-candidate rows for cells that are
 *  actually dates, rather than assuming a fixed column/stride — an Excel
 *  paste can shift, merge, or duplicate header cells in ways position-based
 *  guessing doesn't survive, but the date values themselves are reliable to
 *  search for directly. If a date appears in two adjacent columns (a merged
 *  header cell can read back as duplicated), only the first is kept — the
 *  rate for that date reliably comes first, with its % change column right
 *  after. */
function parseRateShopImport_() {
  const raw = firstOrCreateSheet_(RATE_SHOP_SHEET_ID, 'Raw Paste');
  const data = raw.getDataRange().getValues();

  let headerRowIdx = -1, dateCols = [];
  for (let r = 0; r < Math.min(data.length, 10); r++) {
    const row = data[r];
    const cols = [];
    const seenDates = new Set();
    for (let c = 1; c < row.length; c++) {
      const date = parseRatesDateCell_(row[c]);
      if (date && !seenDates.has(date)) { cols.push(c); seenDates.add(date); }
    }
    if (cols.length >= 2) { headerRowIdx = r; dateCols = cols; break; }
  }
  if (headerRowIdx < 0) {
    throw new Error('Could not find a header row with recognizable dates in Raw Paste. Check the paste matches the OTA Insight "Rates" export format, starting at cell A1.');
  }
  const headerRow = data[headerRowIdx];
  const dates = dateCols.map(c => parseRatesDateCell_(headerRow[c]));

  const out = [];
  for (let r = headerRowIdx + 1; r < data.length; r++) {
    const row = data[r];
    const name = String(row[0] || '').trim();
    if (!name || RATE_SHOP_SKIP_NAMES.has(name.toLowerCase())) continue;

    dateCols.forEach((c, i) => {
      const cellVal = row[c];
      const rateDisplay = String(cellVal == null ? '' : cellVal).trim();
      if (!rateDisplay) return; // no rate for this property/date — just omit the row
      const rateNum = /^-?[\d,.]+$/.test(rateDisplay) ? Number(rateDisplay.replace(/,/g, '')) : '';
      out.push([name, dates[i], rateDisplay, rateNum]);
    });
  }

  // Full wipe, not just the first 4 columns — this sheet's schema changed
  // more than once while this pipeline was being built, and a partial clear
  // left stale columns (a duplicated rate_display/rate_num from an older
  // version) sitting past column D. Those silently corrupted every row: the
  // frontend keys each row by header name, so a repeated blank column
  // overwrote the real value with nothing. A clean sheet each run avoids
  // that regardless of what schema used to be here.
  const outSheet = firstDataSheet_(RATE_SHOP_SHEET_ID);
  outSheet.clear();
  outSheet.getRange(1, 1, 1, 4).setValues([['property_name', 'date', 'rate_display', 'rate_num']]);
  if (out.length) {
    // Force column B to plain text before writing — otherwise Sheets'
    // auto-formatting re-recognizes "2026-08-27" as an actual date and
    // silently converts the cell, the same auto-conversion this whole
    // section already works around on the Raw Paste side.
    outSheet.getRange(2, 2, out.length, 1).setNumberFormat('@');
    outSheet.getRange(2, 1, out.length, 4).setValues(out);
  }

  Logger.log(`Rate Shop: parsed ${out.length} property/date rate(s) across ${new Set(out.map(r => r[0])).size} propert(y/ies).`);
}

/** Every parsed Rate Shop row: {property_name, date, rate_display, rate_num}.
 *  Cluster assignment and RGB-vs-comp classification happen entirely on the
 *  frontend (CLUSTER_DEFS) — this is just the flat rate lookup table. */
function getRateShopRows_() {
  const sheet = firstDataSheet_(RATE_SHOP_SHEET_ID);
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(h => String(h).trim().toLowerCase());
  return data.slice(1)
    .filter(r => r.some(v => v !== '' && v !== null)) // skip fully-blank trailing rows
    .map(r => {
      const obj = {};
      header.forEach((h, i) => {
        let v = r[i];
        // Defensive, in addition to writing the date column as plain text
        // above — if a cell ever comes back as a real Date object anyway
        // (a manual edit to the sheet, a format that slipped through),
        // JSON.stringify would otherwise turn it into a full ISO timestamp
        // and break the frontend's date parsing rather than just the
        // display looking odd.
        if (h === 'date' && v instanceof Date) {
          v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        obj[h] = v;
      });
      return obj;
    });
}

// ═══ 2027 BUDGET — DAILY FORECAST (GMs enter, corporate reviews, budget locks on approval) ═══
// The forecast is the living, day-by-day number (rooms sold + ADR + a note per
// day) and is never locked. Only the BUDGET is locked: when corporate approves
// a submitted forecast, a frozen copy of those 365 days is written to the
// Budget 2027 sheet, and no code path in this file edits or deletes a row there.
//
// STATE (per hotel, on the Status tab):
//   (none) -> DRAFT -> SUBMITTED -> CHANGES_REQUESTED -> DRAFT ... -> APPROVED
//   SUBMITTED freezes the GM's own edits until corporate responds. APPROVED
//   only means the budget snapshot exists; the forecast stays editable.
//
// SETUP (once): run setupForecast2027 from the Apps Script editor. It creates
// three private spreadsheets under the account that owns this script and
// remembers their IDs in Script Properties (no IDs to paste):
//   Forecast 2027       tabs: Forecast (property_id,date,rooms_sold,adr,notes,...) and Status
//   Forecast Log 2027   append-only: every changed day (old/new) and every submit/review/approve
//   Budget 2027         locked snapshots
// Also set DAILY_ACTUALS_SHEET_ID below (last year's actuals, used for the
// "last year same weekday" columns). Until it's set the forecast still works,
// it just shows no last-year comparison.

const FORECAST_YEAR = 2027;
const DAILY_ACTUALS_SHEET_ID = ''; // RGB Intel Hub — Daily Actuals (property_id, date, room_rev, rooms_occ, total_rev), 2022-2026

const FC_HEADERS = {
  forecast: ['property_id', 'date', 'rooms_sold', 'adr', 'notes', 'updated_by', 'updated_at'],
  status: ['property_id', 'status', 'submitted_by', 'submitted_at', 'reviewed_by', 'reviewed_at', 'review_note', 'locked_at', 'locked_by'],
  log: ['logged_at', 'property_id', 'date', 'action', 'old_rooms_sold', 'old_adr', 'new_rooms_sold', 'new_adr', 'notes', 'user'],
  budget: ['property_id', 'date', 'rooms_sold', 'adr', 'notes', 'locked_by', 'locked_at']
};
// Columns kept as plain text so Sheets never re-interprets "2027-01-05" as a
// real date (the same auto-conversion that broke the Rate Shop headers).
const FC_SPREADSHEETS = {
  FC_FORECAST_ID: { title: 'RGB Intel Hub — Forecast 2027', tabs: [
    { name: 'Forecast', header: FC_HEADERS.forecast, text: [1, 2, 5, 6, 7] },
    { name: 'Status', header: FC_HEADERS.status, text: [1, 2, 3, 4, 5, 6, 7, 8, 9] }
  ] },
  FC_LOG_ID: { title: 'RGB Intel Hub — Forecast Log 2027 (append-only)', tabs: [
    { name: 'Log', header: FC_HEADERS.log, text: [1, 2, 3, 4, 9, 10] }
  ] },
  FC_BUDGET_ID: { title: 'RGB Intel Hub — Budget 2027 (locked)', tabs: [
    { name: 'Budget', header: FC_HEADERS.budget, text: [1, 2, 5, 6, 7] }
  ] }
};

/** Run once from the Apps Script editor. Safe to re-run: it only creates what's missing. */
function setupForecast2027() {
  Object.keys(FC_SPREADSHEETS).forEach(key => {
    const ss = fcEnsureSpreadsheet_(key);
    Logger.log(FC_SPREADSHEETS[key].title + ' -> ' + ss.getUrl());
  });
  if (!DAILY_ACTUALS_SHEET_ID) Logger.log('NOTE: DAILY_ACTUALS_SHEET_ID is still empty — set it at the top of the 2027 BUDGET section to enable last-year comparisons.');
}

function fcEnsureSpreadsheet_(key) {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(key);
  if (existingId) {
    try { return SpreadsheetApp.openById(existingId); } catch (e) { /* deleted or no longer accessible — recreate below */ }
  }
  const spec = FC_SPREADSHEETS[key];
  const ss = SpreadsheetApp.create(spec.title);
  spec.tabs.forEach((t, i) => {
    const sh = i === 0 ? ss.getSheets()[0].setName(t.name) : ss.insertSheet(t.name);
    t.text.forEach(c => sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'));
    sh.getRange(1, 1, 1, t.header.length).setValues([t.header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  props.setProperty(key, ss.getId());
  return ss;
}

function fcSheet_(key, tabName) {
  const id = PropertiesService.getScriptProperties().getProperty(key);
  if (!id) throw new Error('The 2027 budget sheets are not set up yet — run setupForecast2027 once from the Apps Script editor.');
  const sh = SpreadsheetApp.openById(id).getSheetByName(tabName);
  if (!sh) throw new Error('Missing tab "' + tabName + '" in the 2027 budget spreadsheet.');
  return sh;
}

function fcErr_(code, message) { return { error: code, message: message }; }

/** Budget screens are for logged-in READ users (GM / regional / corporate) with a row on the Permissions sheet. */
function fcAuth_(source) {
  const cred = checkCredentials_(source.username, source.password);
  if (!cred) return { err: fcErr_('not_authorized', 'Invalid or inactive credentials.') };
  if (cred.role !== 'READ') return { err: fcErr_('not_authorized', 'This login cannot access budgets.') };
  const grant = lookupPermission_(cred.username);
  if (!grant) return { err: fcErr_('not_authorized', 'Not on the Permissions sheet.') };
  return { cred: cred, grant: grant, isCorporate: grant.scope_type === 'portfolio' };
}

// ── dates ──
// Everything is a plain 'yyyy-MM-dd' string; Dates only appear when Sheets
// hands one back, and are converted with the SPREADSHEET's timezone (not the
// script's) so a midnight date can't slip to the neighbouring day.
function fcDateKey_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim().slice(0, 10);
}
function fcShiftDate_(key, deltaDays) {
  const p = key.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2] + deltaDays)).toISOString().slice(0, 10);
}
function fcYearDays_(year) {
  const out = [];
  for (let d = new Date(Date.UTC(year, 0, 1)); d.getUTCFullYear() === year; d.setUTCDate(d.getUTCDate() + 1)) {
    out.push(d.toISOString().slice(0, 10));
  }
  return out;
}

// ── last year's actuals → "same weekday last year" baseline ──
/** {propertyId: {'yyyy-MM-dd': [roomsSold, adr]}} for 2025 onward, for the given properties only. */
function fcLoadActuals_(propertyIds) {
  const out = {};
  if (!DAILY_ACTUALS_SHEET_ID) return out;
  const ss = SpreadsheetApp.openById(DAILY_ACTUALS_SHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const sheet = ss.getSheets().find(s => !NON_DATA_TAB_NAMES.has(s.getName().trim().toLowerCase())) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const h = data[0].map(x => String(x).trim().toLowerCase());
  const iId = h.indexOf('property_id'), iDate = h.indexOf('date'), iRev = h.indexOf('room_rev'), iOcc = h.indexOf('rooms_occ');
  if (iId < 0 || iDate < 0 || iRev < 0 || iOcc < 0) throw new Error('Daily Actuals sheet needs columns: property_id, date, room_rev, rooms_occ.');
  const want = new Set(propertyIds.map(String));
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const pid = String(row[iId]);
    if (!want.has(pid)) continue;
    const key = fcDateKey_(row[iDate], tz);
    if (key < '2025-01-01') continue;
    const sold = Number(row[iOcc]) || 0, rev = Number(row[iRev]) || 0;
    if (sold <= 0) continue;
    (out[pid] = out[pid] || {})[key] = [sold, rev / sold];
  }
  return out;
}

/** Baseline for a 2027 date: the same WEEKDAY one year back (364 days — keeps Fri on Fri,
 *  and lands Thanksgiving on Thanksgiving), falling back to two years back (728 days) where
 *  last year isn't available yet (Nov-Dec 2026). Returns [roomsSold, adr, sourceYear] or null. */
function fcBase_(actualsForProperty, key) {
  const a = actualsForProperty[fcShiftDate_(key, -364)];
  if (a) return [a[0], a[1], FORECAST_YEAR - 1];
  const b = actualsForProperty[fcShiftDate_(key, -728)];
  if (b) return [b[0], b[1], FORECAST_YEAR - 2];
  return null;
}

// ── sheet access ──
function fcReadForecast_() {
  const sheet = fcSheet_('FC_FORECAST_ID', 'Forecast');
  const tz = sheet.getParent().getSpreadsheetTimeZone();
  const width = FC_HEADERS.forecast.length;
  // Exactly `width` columns per row (a stray extra column someone adds must not break the write-back).
  const values = sheet.getDataRange().getValues().map(r => {
    const row = r.slice(0, width);
    while (row.length < width) row.push('');
    return row;
  });
  for (let i = 1; i < values.length; i++) {          // normalize to text keys so round-trips are stable
    values[i][0] = String(values[i][0]);
    values[i][1] = fcDateKey_(values[i][1], tz);
  }
  return { sheet: sheet, tz: tz, values: values };
}

function fcReadStatus_() {
  const sheet = fcSheet_('FC_FORECAST_ID', 'Status');
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === '' || data[i][0] == null) continue;
    const o = { _row: i + 1 };
    FC_HEADERS.status.forEach((k, j) => { o[k] = data[i][j] == null ? '' : data[i][j]; });
    o.property_id = String(o.property_id);
    map[o.property_id] = o;
  }
  return { sheet: sheet, map: map };
}

function fcSetStatus_(propertyId, patch) {
  const s = fcReadStatus_();
  const cur = s.map[propertyId];
  const vals = FC_HEADERS.status.map(k => patch[k] !== undefined ? patch[k] : (cur ? cur[k] : ''));
  vals[0] = propertyId;
  if (cur) s.sheet.getRange(cur._row, 1, 1, vals.length).setValues([vals]);
  else s.sheet.appendRow(vals);
}

function fcPublicStatus_(st) {
  st = st || {};
  return {
    status: st.status || '', submitted_by: st.submitted_by || '', submitted_at: st.submitted_at || '',
    reviewed_by: st.reviewed_by || '', reviewed_at: st.reviewed_at || '', review_note: st.review_note || '',
    locked_at: st.locked_at || '', locked_by: st.locked_by || ''
  };
}

function fcAppendRows_(sheet, rows, width) {
  if (!rows.length) return;
  const start = sheet.getLastRow() + 1;
  const need = start + rows.length - 1;
  if (sheet.getMaxRows() < need) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());
  sheet.getRange(start, 1, rows.length, width).setValues(rows);
}

function fcLog_(rows) {
  fcAppendRows_(fcSheet_('FC_LOG_ID', 'Log'), rows, FC_HEADERS.log.length);
}

function fcNowIso_() { return new Date().toISOString(); }

/** Days (this year only, both numbers present) a property has entered: {date: [rooms, adr, notes]}. */
function fcPropertyDays_(values, propertyId) {
  const out = {};
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (String(r[0]) !== propertyId) continue;
    if (String(r[1]).slice(0, 4) !== String(FORECAST_YEAR)) continue;
    if (r[2] === '' || r[3] === '' || !isFinite(Number(r[2])) || !isFinite(Number(r[3]))) continue;
    out[r[1]] = [Number(r[2]), Number(r[3]), String(r[4] || '')];
  }
  return out;
}

function fcValidateDay_(d, roomCount) {
  const date = String((d && d.date) || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date.slice(0, 4) !== String(FORECAST_YEAR) || fcShiftDate_(date, 0) !== date) {
    return 'Not a valid ' + FORECAST_YEAR + ' date: ' + date;
  }
  const sold = Number(d.rooms_sold), adr = Number(d.adr);
  if (d.rooms_sold === '' || d.rooms_sold == null || !isFinite(sold) || sold < 0 || sold > roomCount || Math.floor(sold) !== sold) {
    return date + ': rooms sold must be a whole number from 0 to ' + roomCount + ' (the hotel\'s room count).';
  }
  if (d.adr === '' || d.adr == null || !isFinite(adr) || adr < 0 || adr > 2000) {
    return date + ': ADR must be between $0 and $2,000.';
  }
  return null;
}

// ── handlers ──

/** One row per hotel the caller can see: status, how many of 365 days are entered, the forecast
 *  total, and that total against last year over the SAME days (so a half-finished forecast isn't
 *  compared with a full year). */
function handleForecastSummary_(params) {
  const a = fcAuth_(params);
  if (a.err) return a.err;
  const props = getAuthorizedProperties_(a.grant);
  const ids = props.map(p => String(p.property_id));
  const idSet = new Set(ids);

  let actuals = {}, actualsLoaded = !!DAILY_ACTUALS_SHEET_ID;
  if (actualsLoaded) { try { actuals = fcLoadActuals_(ids); } catch (e) { actualsLoaded = false; } }

  const fc = fcReadForecast_();
  const statusMap = fcReadStatus_().map;
  const yearDays = fcYearDays_(FORECAST_YEAR);

  const agg = {};
  fc.values.slice(1).forEach(r => {
    const pid = String(r[0]);
    if (!idSet.has(pid) || String(r[1]).slice(0, 4) !== String(FORECAST_YEAR)) return;
    if (r[2] === '' || r[3] === '' || !isFinite(Number(r[2])) || !isFinite(Number(r[3]))) return;
    const g = agg[pid] = agg[pid] || { days: 0, rooms: 0, rev: 0, revMatched: 0, baseRev: 0, updated: '' };
    const sold = Number(r[2]), adr = Number(r[3]);
    g.days++; g.rooms += sold; g.rev += sold * adr;
    // "vs last year" only compares days that have BOTH a forecast and a baseline,
    // so a day with no last-year number can't inflate the variance.
    const b = fcBase_(actuals[pid] || {}, r[1]);
    if (b) { g.baseRev += b[0] * b[1]; g.revMatched += sold * adr; }
    const u = String(r[6] || '');
    if (u > g.updated) g.updated = u;
  });

  const list = props.map(p => {
    const pid = String(p.property_id);
    const g = agg[pid] || { days: 0, rooms: 0, rev: 0, revMatched: 0, baseRev: 0, updated: '' };
    let fullBaseRev = 0;
    if (actuals[pid]) yearDays.forEach(d => { const b = fcBase_(actuals[pid], d); if (b) fullBaseRev += b[0] * b[1]; });
    const st = fcPublicStatus_(statusMap[pid]);
    return {
      property_id: pid, name: p.name, brand: p.brand, market: p.market, rooms: Number(p.rooms) || 0, gm_email: p.gm_email || '',
      status: st.status, days_entered: g.days, fc_rooms: g.rooms, fc_rev: Math.round(g.rev),
      fc_rev_matched: Math.round(g.revMatched), base_rev_matched: Math.round(g.baseRev), base_rev_year: Math.round(fullBaseRev),
      last_updated: g.updated, review_note: st.review_note, locked_at: st.locked_at
    };
  });

  return { year: FORECAST_YEAR, scope: a.grant, can_review: a.isCorporate, daily_actuals_loaded: actualsLoaded, properties: list };
}

/** Everything the editor needs for one hotel: its saved forecast, the last-year baseline for every 2027 day, status. */
function handleForecastGet_(params) {
  const a = fcAuth_(params);
  if (a.err) return a.err;
  const pid = String(params.property_id || '').trim();
  if (!pid || !isAuthorizedForProperty_(a.grant, pid)) return fcErr_('not_authorized', 'Not authorized for this property.');
  const prop = getPropertyById_(pid);
  if (!prop) return fcErr_('not_found', 'Unknown property.');

  let actuals = {}, baseError = '';
  if (DAILY_ACTUALS_SHEET_ID) {
    try { actuals = fcLoadActuals_([pid])[pid] || {}; } catch (e) { baseError = String(e.message || e); }
  } else {
    baseError = 'Last-year comparison is not loaded yet.';
  }
  const base = {};
  fcYearDays_(FORECAST_YEAR).forEach(d => {
    const b = fcBase_(actuals, d);
    if (b) base[d] = [b[0], Math.round(b[1] * 100) / 100, b[2]];
  });

  const days = fcPropertyDays_(fcReadForecast_().values, pid);
  const st = fcPublicStatus_(fcReadStatus_().map[pid]);
  return {
    year: FORECAST_YEAR,
    property: { property_id: pid, name: prop.name, brand: prop.brand, market: prop.market, rooms: Number(prop.rooms) || 0 },
    status: st, forecast: days, base: base, base_error: baseError,
    can_review: a.isCorporate,
    can_edit: a.isCorporate || st.status !== 'SUBMITTED'
  };
}

/** Upserts the days the caller changed. Only changed days are written, and each one is logged old -> new. */
function handleForecastSave_(body) {
  const a = fcAuth_(body);
  if (a.err) return a.err;
  const pid = String(body.property_id || '').trim();
  if (!pid || !isAuthorizedForProperty_(a.grant, pid)) return fcErr_('not_authorized', 'Not authorized for this property.');
  const prop = getPropertyById_(pid);
  if (!prop) return fcErr_('not_found', 'Unknown property.');
  const roomCount = Number(prop.rooms) || 0;

  const days = Array.isArray(body.days) ? body.days : [];
  if (!days.length) return fcErr_('bad_request', 'No days to save.');
  if (days.length > 400) return fcErr_('bad_request', 'Too many days in one save.');
  const clean = [], seen = new Set();
  for (let i = 0; i < days.length; i++) {
    const err = fcValidateDay_(days[i], roomCount);
    if (err) return fcErr_('bad_request', err);
    const date = String(days[i].date).trim();
    if (seen.has(date)) continue;
    seen.add(date);
    clean.push({ date: date, rooms_sold: Number(days[i].rooms_sold), adr: Math.round(Number(days[i].adr) * 100) / 100, notes: String(days[i].notes || '').slice(0, 300) });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const stMap = fcReadStatus_().map;
    const curStatus = stMap[pid] ? stMap[pid].status : '';
    if (curStatus === 'SUBMITTED' && !a.isCorporate) {
      return fcErr_('locked_for_review', 'This forecast is submitted and waiting for review. Ask corporate to send it back if you need to change it.');
    }

    const fc = fcReadForecast_();
    const values = fc.values;
    const index = {};
    for (let i = 1; i < values.length; i++) if (values[i][0] === pid) index[values[i][1]] = i;

    const now = fcNowIso_(), who = a.cred.username;
    const logRows = [];
    clean.forEach(d => {
      const i = index[d.date];
      if (i !== undefined) {
        const r = values[i];
        const same = r[2] !== '' && Number(r[2]) === d.rooms_sold && r[3] !== '' && Number(r[3]) === d.adr && String(r[4] || '') === d.notes;
        if (same) return;
        logRows.push([now, pid, d.date, 'SAVE_DAY', r[2], r[3], d.rooms_sold, d.adr, d.notes, who]);
        values[i] = [pid, d.date, d.rooms_sold, d.adr, d.notes, who, now];
      } else {
        logRows.push([now, pid, d.date, 'SAVE_DAY', '', '', d.rooms_sold, d.adr, d.notes, who]);
        values.push([pid, d.date, d.rooms_sold, d.adr, d.notes, who, now]);
        index[d.date] = values.length - 1;
      }
    });

    if (!logRows.length) return { ok: true, changed: 0, status: curStatus };

    if (fc.sheet.getMaxRows() < values.length) fc.sheet.insertRowsAfter(fc.sheet.getMaxRows(), values.length - fc.sheet.getMaxRows());
    fc.sheet.getRange(1, 1, values.length, FC_HEADERS.forecast.length).setValues(values);
    fcLog_(logRows);

    let newStatus = curStatus;
    if (curStatus === '' || curStatus === 'CHANGES_REQUESTED') { newStatus = 'DRAFT'; fcSetStatus_(pid, { status: 'DRAFT' }); }
    return { ok: true, changed: logRows.length, status: newStatus };
  } finally {
    lock.releaseLock();
  }
}

/** GM hands the forecast to corporate. Requires every day of the year to be filled in. */
function handleForecastSubmit_(body) {
  const a = fcAuth_(body);
  if (a.err) return a.err;
  const pid = String(body.property_id || '').trim();
  if (!pid || !isAuthorizedForProperty_(a.grant, pid)) return fcErr_('not_authorized', 'Not authorized for this property.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const cur = fcReadStatus_().map[pid];
    if (cur && cur.status === 'APPROVED') return fcErr_('bad_state', 'This budget is already approved and locked.');
    if (cur && cur.status === 'SUBMITTED') return fcErr_('bad_state', 'Already submitted — waiting for corporate review.');

    const days = fcPropertyDays_(fcReadForecast_().values, pid);
    const missingByMonth = [];
    const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const counts = new Array(12).fill(0);
    fcYearDays_(FORECAST_YEAR).forEach(d => { if (!days[d]) counts[Number(d.slice(5, 7)) - 1]++; });
    counts.forEach((n, m) => { if (n) missingByMonth.push(MONTHS[m] + ' (' + n + ')'); });
    if (missingByMonth.length) {
      return fcErr_('incomplete', 'Every day needs rooms sold and ADR before submitting. Still missing: ' + missingByMonth.join(', ') + '.');
    }

    const now = fcNowIso_();
    fcSetStatus_(pid, { status: 'SUBMITTED', submitted_by: a.cred.username, submitted_at: now });
    fcLog_([[now, pid, '', 'SUBMIT', '', '', '', '', '', a.cred.username]]);
    return { ok: true, status: 'SUBMITTED' };
  } finally {
    lock.releaseLock();
  }
}

/** Corporate only: COMMENT (note, no state change), REQUEST_CHANGES (back to the GM), or APPROVE (lock the budget). */
function handleForecastReview_(body) {
  const a = fcAuth_(body);
  if (a.err) return a.err;
  if (!a.isCorporate) return fcErr_('not_authorized', 'Only corporate can review budgets.');
  const pid = String(body.property_id || '').trim();
  if (!pid || !getPropertyById_(pid)) return fcErr_('not_found', 'Unknown property.');
  const decision = String(body.decision || '').toUpperCase();
  const note = String(body.note || '').slice(0, 1000);
  if (['COMMENT', 'REQUEST_CHANGES', 'APPROVE'].indexOf(decision) < 0) return fcErr_('bad_request', 'Unknown decision.');
  if (decision !== 'APPROVE' && !note.trim()) return fcErr_('bad_request', 'Add a note so the GM knows what to change.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const cur = fcReadStatus_().map[pid];
    const curStatus = cur ? cur.status : '';
    const now = fcNowIso_(), who = a.cred.username;

    if (decision === 'COMMENT') {
      fcSetStatus_(pid, { status: curStatus || 'DRAFT', reviewed_by: who, reviewed_at: now, review_note: note });
      fcLog_([[now, pid, '', 'COMMENT', '', '', '', '', note, who]]);
      return { ok: true, status: curStatus || 'DRAFT' };
    }

    if (decision === 'REQUEST_CHANGES') {
      if (curStatus === 'APPROVED') return fcErr_('bad_state', 'This budget is already locked.');
      fcSetStatus_(pid, { status: 'CHANGES_REQUESTED', reviewed_by: who, reviewed_at: now, review_note: note });
      fcLog_([[now, pid, '', 'REQUEST_CHANGES', '', '', '', '', note, who]]);
      return { ok: true, status: 'CHANGES_REQUESTED' };
    }

    // APPROVE: freeze a copy of the 365 days into Budget 2027.
    if (curStatus === 'APPROVED') return fcErr_('bad_state', 'This budget is already approved and locked.');
    if (curStatus !== 'SUBMITTED') return fcErr_('bad_state', 'Only a submitted forecast can be approved.');
    const budgetSheet = fcSheet_('FC_BUDGET_ID', 'Budget');
    const existing = budgetSheet.getDataRange().getValues();
    for (let i = 1; i < existing.length; i++) {
      if (String(existing[i][0]) === pid) return fcErr_('bad_state', 'A locked budget already exists for this hotel.');
    }
    const days = fcPropertyDays_(fcReadForecast_().values, pid);
    const yearDays = fcYearDays_(FORECAST_YEAR);
    if (yearDays.some(d => !days[d])) return fcErr_('incomplete', 'Forecast is missing days — send it back instead of approving.');
    const snapshot = yearDays.map(d => [pid, d, days[d][0], days[d][1], days[d][2], who, now]);
    fcAppendRows_(budgetSheet, snapshot, FC_HEADERS.budget.length);
    fcSetStatus_(pid, { status: 'APPROVED', reviewed_by: who, reviewed_at: now, review_note: note, locked_at: now, locked_by: who });
    fcLog_([[now, pid, '', 'APPROVE', '', '', '', '', note, who]]);
    return { ok: true, status: 'APPROVED', locked_at: now };
  } finally {
    lock.releaseLock();
  }
}

// ═══ DAR ENTRY FORM (auditor-facing) ═══

function renderDarEntryFormHtml_() {
  const fieldsJson = JSON.stringify(DAR_FIELDS);
  const scriptUrl = ScriptApp.getService().getUrl();
  return `<!DOCTYPE html><html><head><base target="_top">
<meta charset="utf-8">
<style>
body{font-family:Arial,sans-serif;max-width:520px;margin:24px auto;padding:0 16px;color:#222}
h2{font-size:18px;margin-bottom:4px}
h3{font-size:13px;text-transform:uppercase;letter-spacing:.03em;color:#555;margin:20px 0 6px;border-bottom:1px solid #ddd;padding-bottom:4px}
label{display:block;margin-top:8px;font-size:13px;font-weight:bold}
input{width:100%;padding:7px;font-size:14px;box-sizing:border-box;margin-top:3px}
button{margin-top:16px;padding:10px 16px;font-size:14px;background:#1a5;color:#fff;border:none;cursor:pointer;border-radius:3px}
button.secondary{background:#666}
#msg{margin-top:14px;font-size:13px}
.err{color:#c0392b} .ok{color:#1a5}
#propInfo{font-size:13px;color:#555;margin-top:6px}
#formBody{display:none}
</style></head><body>
<h2>RGB Intelligence Hub — DAR Entry</h2>
<div id="loginBox">
<label>Username <input id="u"></label>
<label>Password <input id="p" type="password"></label>
<label>Date (the day this DAR is for) <input id="d" type="date"></label>
<button onclick="loadForm()">Continue</button>
<div id="propInfo"></div>
</div>
<div id="formBody"></div>
<button id="submitBtn" style="display:none" onclick="submitDar()">Submit DAR</button>
<div id="msg"></div>
<div id="availSection" style="display:none">
  <h3>21-Day Availability (Rooms to Sell)</h3>
  <div style="font-size:12px;color:#666;margin-bottom:6px">Update as many days as you have visibility into — this feeds Rate Shop instead of someone pulling it from your PMS each morning.</div>
  <div id="availBody"></div>
  <button onclick="submitAvailability()">Save Availability</button>
  <div id="availMsg" style="margin-top:10px;font-size:13px"></div>
</div>
<script>
const FIELDS = ${fieldsJson};
const API_URL = '${scriptUrl}'; // the real /exec URL, injected server-side — window.location.href
                                 // inside a served Apps Script page points at an internal
                                 // googleusercontent.com content URL, not this one, so it can't
                                 // be used to build fetch() targets.
let CTX = null; // {property_id, has_restaurant, username, password, date}

function yesterday(){ const d=new Date(); d.setDate(d.getDate()-1); return d.toISOString().slice(0,10); }
document.getElementById('d').value = yesterday();
// If arriving from the main Hub's splash login, the username is carried over
// in the URL (never the password) so the auditor only has to type it once.
(function prefillUsername(){
  const params = new URLSearchParams(window.location.search);
  const u = params.get('username');
  if(u){ document.getElementById('u').value = u; document.getElementById('p').focus(); }
})();

function loadForm(){
  const u=document.getElementById('u').value.trim();
  const p=document.getElementById('p').value;
  const d=document.getElementById('d').value;
  if(!u||!p||!d){ setMsg('Fill in username, password, and date.', true); return; }
  setMsg('Loading...', false);
  fetch(base()+'?action=whoami&username='+encodeURIComponent(u)+'&password='+encodeURIComponent(p))
    .then(r=>r.json())
    .then(who=>{
      if(who.error){ setMsg(who.message||'Login failed.', true); return; }
      if(who.role!=='ENTRY'){ setMsg('This login is not a night-audit login.', true); return; }
      CTX = {property_id: who.property_id, has_restaurant: who.has_restaurant, username:u, password:p, date:d};
      document.getElementById('propInfo').textContent = 'Property: ' + who.property_name;
      return fetch(base()+'?action=dar_get&username='+encodeURIComponent(u)+'&password='+encodeURIComponent(p)+'&date='+d);
    })
    .then(r=>r && r.json())
    .then(existing=>{
      renderFields(existing && existing.exists ? existing.row : {});
      document.getElementById('submitBtn').style.display='block';
      setMsg(existing && existing.exists ? 'Editing existing submission for this date.' : 'New entry for this date.', false);
      return fetch(base()+'?action=availability_get&username='+encodeURIComponent(CTX.username)+'&password='+encodeURIComponent(CTX.password));
    })
    .then(r=>r && r.json())
    .then(avail=>{
      if(avail && avail.days) renderAvailability(avail.days);
    })
    .catch(err=>setMsg('Network error: '+err, true));
}

function renderAvailability(days){
  const el=document.getElementById('availBody');
  el.innerHTML = days.map(d=>{
    const dow = new Date(d.date+'T12:00:00').toLocaleDateString('en-US',{weekday:'short',month:'short',day:'numeric'});
    return '<label style="display:flex;align-items:center;gap:10px;font-weight:normal;margin-top:6px">'+
      '<span style="width:110px;flex-shrink:0">'+dow+'</span>'+
      '<input type="number" step="1" id="avail_'+d.date+'" value="'+d.rooms_available+'" style="margin-top:0">'+
    '</label>';
  }).join('');
  document.getElementById('availSection').style.display='block';
}

function submitAvailability(){
  const el=document.getElementById('availBody');
  const inputs = el.querySelectorAll('input[id^="avail_"]');
  const days = [];
  inputs.forEach(inp=>{
    if(inp.value === '') return; // leave blank days untouched rather than writing 0
    days.push({ date: inp.id.replace('avail_',''), rooms_available: inp.value });
  });
  if(!days.length){ setAvailMsg('Nothing entered to save.', true); return; }
  setAvailMsg('Saving...', false);
  fetch(base(), {method:'POST', body:JSON.stringify({action:'submit_availability', username:CTX.username, password:CTX.password, days:days})})
    .then(r=>r.json())
    .then(j=>{
      if(j.error){ setAvailMsg(j.message||'Save failed.', true); }
      else { setAvailMsg('Saved '+(j.updated+j.inserted)+' day(s).', false); }
    })
    .catch(err=>setAvailMsg('Network error: '+err, true));
}

function setAvailMsg(text, isErr){ const m=document.getElementById('availMsg'); m.textContent=text; m.style.color = isErr?'#c0392b':'#1a5'; }

function renderFields(existing){
  const body=document.getElementById('formBody');
  body.style.display='block';
  const sections={};
  FIELDS.forEach(f=>{
    if(f.restaurantOnly && !CTX.has_restaurant) return;
    (sections[f.section]=sections[f.section]||[]).push(f);
  });
  let html='';
  Object.keys(sections).forEach(sec=>{
    html+='<h3>'+sec+'</h3>';
    sections[sec].forEach(f=>{
      const val = existing[f.key]!=null ? existing[f.key] : '';
      html+='<label>'+f.label+' <input type="number" step="0.01" id="f_'+f.key+'" value="'+val+'"></label>';
    });
  });
  body.innerHTML=html;
}

function submitDar(){
  const payload={action:'submit_dar', username:CTX.username, password:CTX.password, date:CTX.date};
  FIELDS.forEach(f=>{
    const el=document.getElementById('f_'+f.key);
    payload[f.key] = el ? el.value : 0;
  });
  setMsg('Submitting...', false);
  fetch(base(), {method:'POST', body:JSON.stringify(payload)})
    .then(r=>r.json())
    .then(j=>{
      if(j.error){ setMsg(j.message||'Submit failed.', true); }
      else { setMsg('Saved. Total Revenue: $'+j.totals.total_revenue.toFixed(2), false); }
    })
    .catch(err=>setMsg('Network error: '+err, true));
}

function base(){ return API_URL; }
function setMsg(text, isErr){ const m=document.getElementById('msg'); m.textContent=text; m.className=isErr?'err':'ok'; }
</script>
</body></html>`;
}

// ═══ DAR REVIEW FORM (GM-facing) ═══

function renderDarReviewFormHtml_() {
  const fieldsJson = JSON.stringify(DAR_FIELDS);
  const scriptUrl = ScriptApp.getService().getUrl();
  return `<!DOCTYPE html><html><head><base target="_top">
<meta charset="utf-8">
<style>
body{font-family:Arial,sans-serif;max-width:640px;margin:24px auto;padding:0 16px;color:#222}
h2{font-size:18px}
label{display:block;margin-top:8px;font-size:13px;font-weight:bold}
input{padding:7px;font-size:14px;margin-top:3px}
button{padding:8px 14px;font-size:13px;border:none;cursor:pointer;border-radius:3px;margin-right:6px}
.approve{background:#1a5;color:#fff} .flag{background:#c0392b;color:#fff}
#msg{margin-top:14px;font-size:13px}
.err{color:#c0392b} .ok{color:#1a5}
.card{border:1px solid #ddd;border-radius:4px;padding:12px;margin-top:14px}
.card h4{margin:0 0 8px 0}
table{width:100%;border-collapse:collapse;font-size:12px}
td{padding:2px 4px;border-bottom:1px solid #eee}
</style></head><body>
<h2>RGB Intelligence Hub — DAR Review</h2>
<label>Username <input id="u"></label>
<label>Password <input id="p" type="password"></label>
<button onclick="loadPending()">Load Pending DARs</button>
<div id="list"></div>
<div id="msg"></div>
<script>
const FIELDS = ${fieldsJson};
const API_URL = '${scriptUrl}'; // the real /exec URL, injected server-side — window.location.href
                                 // inside a served Apps Script page points at an internal
                                 // googleusercontent.com content URL, not this one, so it can't
                                 // be used to build fetch() targets.
let CREDS = null;

function loadPending(){
  CREDS = {username:document.getElementById('u').value.trim(), password:document.getElementById('p').value};
  setMsg('Loading...', false);
  fetch(base()+'?action=dar_pending&username='+encodeURIComponent(CREDS.username)+'&password='+encodeURIComponent(CREDS.password))
    .then(r=>r.json())
    .then(j=>{
      if(j.error){ setMsg(j.message||'Load failed.', true); return; }
      renderList(j.rows||[]);
      setMsg(j.rows && j.rows.length ? '' : 'Nothing pending review.', false);
    })
    .catch(err=>setMsg('Network error: '+err, true));
}

function renderList(rows){
  const list=document.getElementById('list');
  if(!rows.length){ list.innerHTML=''; return; }
  list.innerHTML = rows.map((r,idx)=>{
    const fieldRows = FIELDS.filter(f=>Number(r[f.key])).map(f=>
      '<tr><td>'+f.label+'</td><td style="text-align:right">$'+Number(r[f.key]).toFixed(2)+'</td></tr>'
    ).join('');
    return '<div class="card">'+
      '<h4>'+r.property_id+' — '+r.date+'</h4>'+
      '<div>Total Revenue: $'+Number(r.total_revenue||0).toFixed(2)+' | Total Non-Revenue (taxes): $'+Number(r.total_non_revenue||0).toFixed(2)+'</div>'+
      '<table>'+fieldRows+'</table>'+
      '<div style="margin-top:8px">'+
        '<button class="approve" onclick="act('+idx+',\\'validate_dar\\')">Approve</button>'+
        '<button class="flag" onclick="flagWithNote('+idx+')">Flag Issue</button>'+
      '</div>'+
      '<div id="rowmsg_'+idx+'" style="margin-top:6px;font-size:12px"></div>'+
    '</div>';
  }).join('');
  window._pendingRows = rows;
}

function flagWithNote(idx){
  const note = prompt('What is wrong with this submission?');
  if(note===null) return;
  act(idx, 'flag_dar', note);
}

function act(idx, action, note){
  const r = window._pendingRows[idx];
  const payload = {action:action, username:CREDS.username, password:CREDS.password, property_id:r.property_id, date:r.date};
  if(note) payload.note = note;
  fetch(base(), {method:'POST', body:JSON.stringify(payload)})
    .then(resp=>resp.json())
    .then(j=>{
      const m=document.getElementById('rowmsg_'+idx);
      if(j.error){ m.textContent=j.message||'Failed.'; m.className='err'; }
      else { m.textContent='Marked '+j.status+'.'; m.className='ok'; }
    });
}

function base(){ return API_URL; }
function setMsg(text, isErr){ const m=document.getElementById('msg'); m.textContent=text; m.className=isErr?'err':'ok'; }
</script>
</body></html>`;
}

// ═══ ENTRY FORM (served HTML, no separate hosting needed) ═══

function renderEntryFormHtml_() {
  const scriptUrl = ScriptApp.getService().getUrl();
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
  fetch('${scriptUrl}', {method:'POST', body:JSON.stringify(body)})
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

// ═══ NIGHT AUDIT ARCHIVE ═══
// Unrelated to the web app above — this runs on its own time-driven trigger,
// not through doGet/doPost, so the deployment's Execute-as/access settings
// don't affect it. Archives every attachment sent to Audit@rgbhospitality.com
// into Drive, one folder per calendar day (by the message's date, not the day
// the trigger happens to run — so a late-arriving report still lands right).
//
// SETUP (one-time):
// 1. In the Apps Script editor: Triggers (clock icon, left sidebar) > Add Trigger
// 2. Function: archiveNightAuditAttachments | Event source: Time-driven
//    Type: Day timer | Time of day: 7am to 8am (or your preferred window)
// 3. Before relying on the trigger, run it once manually (function dropdown >
//    archiveNightAuditAttachments > Run) and check the destination folder.

const NIGHT_AUDIT_FOLDER_ID = '1yHhh_1mMY98DNNeQS2bwsQuYceJGUbVU'; // Corporate > Reports > Night Audits
const NIGHT_AUDIT_LABEL = 'NightAuditArchived';
const NIGHT_AUDIT_LOOKBACK = 'newer_than:7d'; // safety net if a run is ever missed

function archiveNightAuditAttachments() {
  let label = GmailApp.getUserLabelByName(NIGHT_AUDIT_LABEL);
  if (!label) label = GmailApp.createLabel(NIGHT_AUDIT_LABEL);

  const rootFolder = DriveApp.getFolderById(NIGHT_AUDIT_FOLDER_ID);
  const threads = GmailApp.search(`to:audit@rgbhospitality.com ${NIGHT_AUDIT_LOOKBACK}`, 0, 100);

  let threadsProcessed = 0, filesSaved = 0;
  threads.forEach(thread => {
    const alreadyDone = thread.getLabels().some(l => l.getName() === NIGHT_AUDIT_LABEL);
    if (alreadyDone) return;

    thread.getMessages().forEach(msg => {
      const attachments = msg.getAttachments({ includeInlineImages: false });
      if (!attachments.length) return;
      const dateFolder = getOrCreateDateFolder_(rootFolder, msg.getDate());
      attachments.forEach(att => {
        dateFolder.createFile(att).setName(att.getName());
        filesSaved++;
      });
    });
    thread.addLabel(label);
    threadsProcessed++;
  });

  Logger.log(`Night audit archive: processed ${threadsProcessed} new thread(s), saved ${filesSaved} attachment(s).`);
}

function getOrCreateDateFolder_(rootFolder, date) {
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const existing = rootFolder.getFoldersByName(dateStr);
  return existing.hasNext() ? existing.next() : rootFolder.createFolder(dateStr);
}

// ═══ ONE-TIME SETUP HELPERS ═══

/** Run once from the Apps Script editor (function dropdown > addReadmeTab_ > Run).
 *  Adds a "Read Me" tab to the Credentials sheet with the add/remove/reassign
 *  instructions, so they live next to the data instead of only in chat history.
 *  Safe to run again later — it just overwrites the existing Read Me tab rather
 *  than creating a duplicate. */
function addReadmeTab_() {
  const ss = SpreadsheetApp.openById(CREDENTIALS_SHEET_ID);
  let sheet = ss.getSheetByName('Read Me');
  if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet('Read Me', 0); }

  const rows = [
    ['RGB Intelligence Hub — Credentials Sheet: How To', ''],
    ['', ''],
    ['ADD A USER', ''],
    ['1. Add a row below with: username, new_password (temporary, plaintext), property_id, role (ENTRY or READ), active = TRUE.', ''],
    ['2. If role = READ (GM, regional, or corporate), also add a matching row in the Permissions sheet (email, scope_type, scope_value, role) — that\'s what determines what they can actually see.', ''],
    ['3. In the Apps Script editor, pick syncCredentials_ from the function dropdown (top toolbar) and click Run. This hashes the password and wipes the plaintext cell.', ''],
    ['', ''],
    ['REMOVE A USER', ''],
    ['1. Set their "active" cell to FALSE.', ''],
    ['2. Run syncCredentials_ again. This blocks their login immediately.', ''],
    ['Prefer this over deleting the row — it keeps a record and you can restore access anytime by setting active back to TRUE.', ''],
    ['', ''],
    ['REASSIGN A USER TO A DIFFERENT PROPERTY', ''],
    ['1. Edit their property_id cell.', ''],
    ['2. Run syncCredentials_ again.', ''],
    ['', ''],
    ['ROLES', ''],
    ['ENTRY = night auditor. Can only ever submit DAR data for the one property_id on their row.', ''],
    ['READ = anyone who views dashboards (GM, regional, corporate). Their actual scope (one property / a market cluster / the full portfolio) comes from the Permissions sheet, not from this sheet.', ''],
    ['', ''],
    ['COLUMNS ON THE CREDENTIALS TAB', ''],
    ['username | new_password | password_hash | salt | property_id | role | display_name | active | last_synced', ''],
    ['Only ever fill in: username, new_password, property_id, role, display_name, active. Leave password_hash, salt, and last_synced alone — syncCredentials_ fills those in automatically.', '']
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  ['A3', 'A8', 'A13', 'A17', 'A21'].forEach(cell => sheet.getRange(cell).setFontWeight('bold'));
  sheet.setColumnWidth(1, 700);
  sheet.getRange(1, 1, rows.length, 1).setWrap(true);

  Logger.log('Read Me tab added to the Credentials sheet.');
}
