/**
 * RGB Intelligence Hub — Google sign-in (SSO) and Google Groups scopes
 * --------------------------------------------------------------------
 * WHO GETS IN
 *   Only @rgbhospitality.com Google accounts. The deployment that serves the Hub is set to
 *   "Who has access: Anyone within rgbhospitality.com" (see docs/DEPLOY.md), and on top of that
 *   every call checks the signed-in email's domain here, so a mis-set deployment still can't
 *   leak data.
 *
 * WHAT THEY SEE (no username sheet: Google Groups decide)
 *   corporate@rgbhospitality.com          -> the whole portfolio, plus the corporate settings
 *   region-<market>@rgbhospitality.com     -> every hotel in that market (e.g. region-houston@)
 *   gm-<property>@rgbhospitality.com       -> that one hotel (e.g. gm-prov-houes@)
 *   <market> and <property> are the Properties sheet's market / property_id, lowercased, with
 *   anything that isn't a letter or digit turned into "-". A hotel whose group is named
 *   differently can set it in an optional gm_group column on the Properties sheet.
 *   Someone in several groups gets all of them combined.
 *
 * SERVER-SIDE ONLY
 *   The browser calls hubApi() through google.script.run. hubApi works out the caller from the
 *   Google session, never from anything the browser sends (any username/password in the request
 *   is thrown away), and every handler filters to the caller's hotels before returning.
 *
 * Group lookups take a moment, so a person's scope is cached for SCOPE_CACHE_SECONDS. Adding or
 * removing someone from a group takes effect within that time.
 */

const SSO_DOMAIN = 'rgbhospitality.com';
const CORPORATE_GROUP = 'corporate@' + SSO_DOMAIN;
const SCOPE_CACHE_SECONDS = 300;

// Set by hubApi for the length of one call. Apps Script runs every request in its own
// execution, so this can never carry one person's identity into another person's request.
let REQUEST_SSO_ = null;

function hubSlug_(s) {
  return String(s == null ? '' : s).trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
}
function gmGroupFor_(prop) {
  const custom = String((prop && prop.gm_group) || '').trim().toLowerCase();
  return custom || ('gm-' + hubSlug_(prop.property_id) + '@' + SSO_DOMAIN);
}
function regionGroupFor_(market) { return 'region-' + hubSlug_(market) + '@' + SSO_DOMAIN; }

/** The signed-in Google account, only if it is an @rgbhospitality.com account; otherwise ''. */
function ssoEmail_() {
  let email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { /* not signed in */ }
  return email && email.slice(-(SSO_DOMAIN.length + 1)) === '@' + SSO_DOMAIN ? email : '';
}

/** Which of `candidates` (group emails) this person belongs to. Uses the Admin SDK Directory
 *  service when it is enabled (one call), otherwise checks each group with GroupsApp. */
function userGroups_(email, candidates) {
  const found = new Set();
  if (typeof AdminDirectory !== 'undefined') {
    try {
      let pageToken;
      do {
        const r = AdminDirectory.Groups.list({ userKey: email, maxResults: 200, pageToken: pageToken });
        (r.groups || []).forEach(g => found.add(String(g.email).toLowerCase()));
        pageToken = r.nextPageToken;
      } while (pageToken);
      return found;
    } catch (e) { /* no directory permission: fall back to GroupsApp below */ }
  }
  candidates.forEach(g => {
    try {
      const group = GroupsApp.getGroupByEmail(g);
      if (group && group.hasUser(email)) found.add(g);
    } catch (e) { /* group doesn't exist or isn't visible to the script owner */ }
  });
  return found;
}

/** {scope_type, property_ids, markets, role, groups} for a signed-in person, or null if they are in no Hub group. */
function ssoGrant_(email) {
  const cacheKey = 'hub:scope:' + email;
  const hit = cacheGet_(cacheKey);
  if (hit) return hit.none ? null : hit;

  const props = cachedSheetObjects_('hub:props', PROPERTIES_SHEET_ID);
  const markets = Array.from(new Set(props.map(p => String(p.market || '').trim()).filter(Boolean)));
  const candidates = Array.from(new Set([CORPORATE_GROUP].concat(markets.map(regionGroupFor_), props.map(gmGroupFor_))));
  const groups = userGroups_(email, candidates);

  let grant = null;
  if (groups.has(CORPORATE_GROUP)) {
    grant = { scope_type: 'portfolio', scope_value: '', role: 'CORPORATE', property_ids: [], markets: [], groups: [CORPORATE_GROUP] };
  } else {
    const myMarkets = markets.filter(m => groups.has(regionGroupFor_(m)));
    const myProps = props.filter(p => groups.has(gmGroupFor_(p))).map(p => String(p.property_id));
    if (myMarkets.length || myProps.length) {
      grant = {
        scope_type: 'set', scope_value: '', role: myMarkets.length ? 'REGIONAL' : 'GM',
        property_ids: myProps, markets: myMarkets,
        groups: candidates.filter(g => groups.has(g))
      };
    }
  }
  cachePut_(cacheKey, grant || { none: true }, SCOPE_CACHE_SECONDS);
  return grant;
}

/** The caller of a READ-type endpoint: the Google-signed-in person when called through hubApi,
 *  otherwise the legacy username/password (franchise GMs, until they have Google accounts).
 *  Returns {username, cred, grant, isCorporate, via} or {err}. */
function readAuth_(source) {
  if (REQUEST_SSO_) return REQUEST_SSO_;
  source = source || {};
  const cred = checkCredentials_(source.username, source.password);
  if (!cred) return { err: { error: 'not_authorized', message: 'Invalid or inactive credentials.', _status: 401 } };
  if (cred.role !== 'READ') return { err: { error: 'not_authorized', message: 'This login cannot view this.', _status: 403 }, cred: cred };
  const grant = lookupPermission_(String(cred.username).toLowerCase().trim());
  if (!grant) return { err: { error: 'not_authorized', message: 'Not on the Permissions sheet.', _status: 403 } };
  return { username: cred.username, cred: cred, grant: grant, isCorporate: grant.scope_type === 'portfolio', via: 'password' };
}

/** When the SSO_ENFORCED script property is "true", @rgbhospitality.com people can no longer use
 *  a username/password READ login — they must use Google sign-in. Franchise logins are unaffected. */
function ssoEnforcedFor_(username) {
  const on = String(PropertiesService.getScriptProperties().getProperty('SSO_ENFORCED') || '').toLowerCase() === 'true';
  return on && String(username || '').toLowerCase().trim().slice(-(SSO_DOMAIN.length + 1)) === '@' + SSO_DOMAIN;
}

/** Maintenance functions (setup, re-import, cache refresh...) are top-level functions, and a served
 *  page can call any top-level function through google.script.run. This stops anyone but the
 *  script owner, running it from the Apps Script editor, from doing so. */
function requireOwner_() {
  let active = '', effective = '';
  try { active = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { /* none */ }
  try { effective = String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) { /* none */ }
  if (!active || active !== effective) throw new Error('Run this from the Apps Script editor, signed in as the script owner.');
}

function hubAppUrl_() {
  const stored = PropertiesService.getScriptProperties().getProperty('HUB_EXEC_URL');
  return stored || ScriptApp.getService().getUrl();
}

// ═══ hubApi: the one function the signed-in pages call ═══

/** Called from the served Hub / App pages with google.script.run.hubApi({action, ...}). */
function hubApi(req) {
  req = req || {};
  const email = ssoEmail_();
  if (!email) return { error: 'no_identity', message: 'Sign in with your @' + SSO_DOMAIN + ' Google account.', _status: 401 };
  let grant;
  try { grant = ssoGrant_(email); } catch (e) { return { error: 'server_error', message: 'Could not check your groups: ' + (e.message || e) }; }
  if (!grant) {
    return { error: 'no_scope', _status: 403, email: email,
      message: email + ' is not in any Hub group yet (corporate@, region-<market>@ or gm-<property>@). Ask corporate to add you.' };
  }
  REQUEST_SSO_ = { username: email, cred: { username: email, role: 'READ' }, grant: grant, isCorporate: grant.scope_type === 'portfolio', via: 'google' };
  try {
    const params = Object.assign({}, req);
    delete params.username; delete params.password;   // identity never comes from the browser
    const out = hubDispatch_(String(params.action || ''), params);
    return JSON.parse(JSON.stringify(out));             // Dates -> strings (google.script.run can't return Dates)
  } catch (err) {
    return { error: 'server_error', message: String((err && err.message) || err) };
  } finally {
    REQUEST_SSO_ = null;
  }
}

function hubDispatch_(action, p) {
  const me = REQUEST_SSO_;
  switch (action) {
    case 'whoami': {
      const props = getAuthorizedProperties_(me.grant);
      return {
        role: 'READ', sso: true, username: me.username, scope: me.grant, hub_role: me.grant.role,
        is_corporate: me.isCorporate, properties: props, app_url: hubAppUrl_()
      };
    }
    case 'portfolio_flash': return portfolioFlashPayload_(me);
    case 'rate_shop': return { caller: me.username, scope: me.grant, rateShop: getRateShopRows_() };

    case 'forecast_summary': return handleForecastSummary_(p);
    case 'forecast_get': return handleForecastGet_(p);
    case 'forecast_save': return handleForecastSave_(p);
    case 'forecast_submit': return handleForecastSubmit_(p);
    case 'forecast_review': return handleForecastReview_(p);

    case 'dar_queue': return darQueue_(me);
    case 'dar_status': return darStatusGrid_(me, p);
    case 'dar_decide': return darDecide_(me, p);
    case 'dar_corporate': return darCorporateAction_(me, p);
    case 'dar_rollup': return darRollup_(me, p);

    case 'na_tokens': return naTokenList_(me);
    case 'na_rotate': return naRotateToken_(me, p);
    case 'na_revoke': return naRevokeToken_(me, p);
    case 'na_set_url': return naSetPublicUrl_(me, p);

    case 'baseline_portfolio': return blPortfolio_(me);
    case 'baseline_property': return blPropertyDetail_(me, p);
    case 'baseline_settings_get': return blSettingsGet_(me);
    case 'baseline_settings_save': return blSettingsSave_(me, p);
  }
  return { error: 'unknown_action', message: 'Action "' + action + '" is not supported.', _status: 400 };
}
