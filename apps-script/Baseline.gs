/**
 * RGB Intelligence Hub — 2027 top-line system baseline
 * ----------------------------------------------------
 * For a hotel whose GM hasn't submitted 2027, a monthly baseline is built in this order:
 *
 *  a) 2026 base   Portfolio Flash actuals for Jan..(actuals_through_month, default Oct), then the
 *                 Budget 2026 room revenue for the remaining months. Budget 2026 has no occupancy,
 *                 so budget months are split into room nights with the hotel's 2025 ADR for that
 *                 month (2026 YTD ADR if there is no 2025 month).
 *  b) Event       Hotels in a World Cup 2026 host market (default Houston, Dallas): Jun and Jul 2026
 *                 are replaced with 2025 Jun/Jul x (1 + the hotel's own 2026-vs-2025 trend over the
 *                 other actual months), separately for occupancy and ADR, capped at +/-trend_cap_pct.
 *                 No 2025 Jun/Jul (e.g. Embassy Suites Dallas): Jun/Jul come from the hotel's 2026
 *                 May+Aug shoulder months x the Jun/Jul-vs-May/Aug seasonality of its 2025 peers
 *                 (same market, else same brand tier, else the portfolio), and the hotel is flagged
 *                 "Manual review required".
 *  c) Easter      Easter moves from Apr 5 (2026) to Mar 28 (2027). A % of March room nights moves to
 *                 April at March's ADR, by demand mix (corporate / mixed / leisure) or a per-hotel %.
 *  d) Growth      Occupancy and ADR growth applied uniformly to every month of the normalized base:
 *                 market override, else brand-tier override, else the default (ADR +1.8%, Occ +0.9%,
 *                 = RevPAR +2.7%). No extra Jun/Jul premium is layered on top of normalized data.
 *  e) Settings    All assumptions live in the BASELINE_SETTINGS script property and are edited on the
 *                 corporate-only App > Settings screen.
 *
 * A GM's submitted (or approved) 2027 forecast overrides the baseline. On submit, a GM RevPAR more
 * than gm_threshold_pct (default 5%) away from the baseline needs a written justification.
 */

const BL_BASE_YEAR = 2026;
const BL_YEAR = 2027;
const BL_LABEL = 'System baseline — not GM-submitted';
const BL_TIERS = ['economy_midscale', 'upper_midscale', 'upscale'];
const BL_TIER_LABEL = { economy_midscale: 'Economy / midscale', upper_midscale: 'Upper midscale', upscale: 'Upscale' };
const BL_MIXES = ['corporate', 'mixed', 'leisure'];
const BL_MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const BL_SHOULDER = [4, 7];   // May, Aug

function blDefaults_() {
  return {
    actuals_through_month: 9,
    event: { markets: ['Houston', 'Dallas'], months: [5, 6], trend_cap_pct: 10 },
    easter: { corporate: 5, mixed: 2, leisure: 0 },
    growth: {
      default: { adr: 1.8, occ: 0.9 },
      tiers: { economy_midscale: { adr: 1.0, occ: 0.5 }, upper_midscale: {}, upscale: {} },
      markets: {}
    },
    gm_threshold_pct: 5,
    properties: {},
    updated_by: '', updated_at: ''
  };
}

function blSettings_() {
  const d = blDefaults_();
  let s = {};
  try { s = JSON.parse(PropertiesService.getScriptProperties().getProperty('BASELINE_SETTINGS') || '{}'); } catch (e) { s = {}; }
  return {
    actuals_through_month: s.actuals_through_month != null ? s.actuals_through_month : d.actuals_through_month,
    event: Object.assign({}, d.event, s.event || {}),
    easter: Object.assign({}, d.easter, s.easter || {}),
    growth: {
      default: Object.assign({}, d.growth.default, (s.growth && s.growth.default) || {}),
      tiers: Object.assign({}, d.growth.tiers, (s.growth && s.growth.tiers) || {}),
      markets: Object.assign({}, (s.growth && s.growth.markets) || {})
    },
    gm_threshold_pct: s.gm_threshold_pct != null ? s.gm_threshold_pct : d.gm_threshold_pct,
    properties: s.properties || {},
    updated_by: s.updated_by || '', updated_at: s.updated_at || ''
  };
}

// ── per-hotel classification (inferred from the name; corporate can override each one) ──
function blInferTier_(p) {
  const s = String(p.name || '').toLowerCase();
  if (/la quinta|\btru\b|days inn|super 8|microtel|baymont|red roof|motel 6|sleep inn|quality inn|wingate|ramada|travelodge|econo lodge|americas best/.test(s)) return 'economy_midscale';
  if (/embassy|residence inn|staybridge|hyatt|courtyard|hilton garden|homewood|springhill|cambria|doubletree|aloft|element/.test(s)) return 'upscale';
  return 'upper_midscale';
}
function blInferMix_(p) {
  const s = (String(p.name || '') + ' ' + String(p.market || '')).toLowerCase();
  if (/beach|island|rockport|port aransas|galveston|padre|resort|waterpark/.test(s)) return 'leisure';
  if (/downtown|airport|medical|galleria|energy corridor|convention|westchase|cbd|financial/.test(s)) return 'corporate';
  return 'mixed';
}
function blNum_(v) { return v === '' || v == null || !isFinite(Number(v)) ? null : Number(v); }
function blDim_(y, m) { return new Date(Date.UTC(y, m + 1, 0)).getUTCDate(); }

function blResolve_(prop, S) {
  const pid = String(prop.property_id);
  const ps = S.properties[pid] || {};
  const tier = BL_TIERS.indexOf(ps.tier) >= 0 ? ps.tier : blInferTier_(prop);
  const mix = BL_MIXES.indexOf(ps.mix) >= 0 ? ps.mix : blInferMix_(prop);
  const market = String(prop.market || '').trim();
  const mo = S.growth.markets[market] || {}, to = S.growth.tiers[tier] || {}, d = S.growth.default;
  const pick = k => blNum_(mo[k]) !== null ? [blNum_(mo[k]), 'market: ' + market] : blNum_(to[k]) !== null ? [blNum_(to[k]), 'tier: ' + BL_TIER_LABEL[tier]] : [Number(d[k]) || 0, 'default'];
  const adr = pick('adr'), occ = pick('occ');
  const autoEvent = (S.event.markets || []).some(m => m && market.toLowerCase().indexOf(String(m).toLowerCase()) >= 0);
  const isEvent = ps.event === 'yes' || (ps.event !== 'no' && autoEvent);
  const easterOverride = blNum_(ps.easter_pct);
  return {
    tier: tier, tier_inferred: BL_TIERS.indexOf(ps.tier) < 0, mix: mix, mix_inferred: BL_MIXES.indexOf(ps.mix) < 0,
    growth_adr: adr[0], growth_adr_source: adr[1], growth_occ: occ[0], growth_occ_source: occ[1],
    event: isEvent, event_source: ps.event === 'yes' || ps.event === 'no' ? 'override' : 'market',
    easter_pct: easterOverride !== null ? easterOverride : Number(S.easter[mix]) || 0, easter_source: easterOverride !== null ? 'override' : 'mix: ' + mix
  };
}

// ── data (all hotels: peers are needed for seasonality; only outputs for the caller's hotels leave the server) ──
function blLoadData_() {
  const all = cachedSheetObjects_('hub:props', PROPERTIES_SHEET_ID);
  const ids = all.map(p => String(p.property_id));
  const flash = {}, h25 = {}, budget = {};
  getFlashRows_(ids).forEach(r => {
    const pid = String(r.property_id), m = Number(r.month_index);
    if (m >= 0 && m <= 11) (flash[pid] = flash[pid] || {})[m] = { rn: Number(r.occupied_rooms) || 0, rev: Number(r.room_rev) || 0 };
  });
  getHistoricalRows_(ids).forEach(r => {
    if (Number(r.year) !== BL_BASE_YEAR - 1) return;
    const pid = String(r.property_id), m = Number(r.month_index);
    if (m >= 0 && m <= 11) (h25[pid] = h25[pid] || {})[m] = { rn: Number(r.occupied_rooms) || 0, rev: Number(r.room_rev) || 0 };
  });
  getBudgetRows_(ids).forEach(r => {
    const pid = String(r.property_id), m = Number(r.month_index);
    if (m >= 0 && m <= 11) (budget[pid] = budget[pid] || {})[m] = Number(r.budget_room_rev) || 0;
  });
  return { all: all, flash: flash, h25: h25, budget: budget };
}

/** Jun/Jul vs May+Aug seasonality from 2025 peers: {months: {m: {occ, adr}}, basis, peers}. */
function blPeerSeasonality_(prop, tier, D, S, months) {
  const ok = pid => {
    const h = D.h25[pid] || {};
    return BL_SHOULDER.concat(months).every(m => h[m] && h[m].rn > 0 && h[m].rev > 0);
  };
  const others = D.all.filter(p => String(p.property_id) !== String(prop.property_id) && Number(p.rooms) > 0 && ok(String(p.property_id)));
  const sameMarket = others.filter(p => String(p.market || '').toLowerCase() === String(prop.market || '').toLowerCase());
  const sameTier = others.filter(p => blResolve_(p, S).tier === tier);
  const peers = sameMarket.length ? sameMarket : sameTier.length ? sameTier : others;
  const basis = sameMarket.length ? 'same market' : sameTier.length ? 'same brand tier' : others.length ? 'portfolio' : 'none (ratio 1.0)';
  const out = {};
  months.forEach(m => {
    let occR = 0, adrR = 0;
    peers.forEach(p => {
      const h = D.h25[String(p.property_id)], rooms = Number(p.rooms);
      const shRn = BL_SHOULDER.reduce((a, s) => a + h[s].rn, 0), shRev = BL_SHOULDER.reduce((a, s) => a + h[s].rev, 0);
      const shAvail = BL_SHOULDER.reduce((a, s) => a + rooms * blDim_(BL_BASE_YEAR - 1, s), 0);
      occR += (h[m].rn / (rooms * blDim_(BL_BASE_YEAR - 1, m))) / (shRn / shAvail);
      adrR += (h[m].rev / h[m].rn) / (shRev / shRn);
    });
    out[m] = peers.length ? { occ: occR / peers.length, adr: adrR / peers.length } : { occ: 1, adr: 1 };
  });
  return { months: out, basis: basis, peers: peers.map(p => p.name) };
}

function blClone_(arr) { return arr.map(x => ({ rn: x.rn, rev: x.rev, avail: x.avail, src: x.src })); }
function blTotals_(arr) {
  const rn = arr.reduce((a, x) => a + x.rn, 0), rev = arr.reduce((a, x) => a + x.rev, 0), avail = arr.reduce((a, x) => a + x.avail, 0);
  return { rn: rn, rev: rev, avail: avail, occ: avail ? rn / avail : 0, adr: rn ? rev / rn : 0, revpar: avail ? rev / avail : 0 };
}

/** The full build for one hotel: every stage month by month, the waterfall, flags. */
function blCompute_(prop, D, S) {
  const pid = String(prop.property_id), rooms = Number(prop.rooms) || 0;
  const R = blResolve_(prop, S);
  const through = Math.max(-1, Math.min(11, Math.floor(Number(S.actuals_through_month))));
  const f26 = D.flash[pid] || {}, h25 = D.h25[pid] || {}, bud = D.budget[pid] || {};
  const flags = [];
  let manual = false;

  // 2026 YTD ADR, the fallback for splitting a budget month that has no 2025 ADR
  let ytdRn = 0, ytdRev = 0;
  for (let m = 0; m <= through; m++) if (f26[m] && f26[m].rn > 0) { ytdRn += f26[m].rn; ytdRev += f26[m].rev; }
  const ytdAdr = ytdRn ? ytdRev / ytdRn : 0;

  // a) 2026 base
  const raw = [];
  for (let m = 0; m < 12; m++) {
    const avail = rooms * blDim_(BL_BASE_YEAR, m);
    const a = f26[m], b = bud[m], l = h25[m];
    if (m <= through && a && a.rn > 0) { raw.push({ rn: a.rn, rev: a.rev, avail: avail, src: 'actual' }); continue; }
    if (b > 0) {
      const adr = l && l.rn > 0 ? l.rev / l.rn : ytdAdr;
      const rn = adr > 0 ? Math.min(avail, b / adr) : 0;
      raw.push({ rn: rn, rev: b, avail: avail, src: l && l.rn > 0 ? 'budget (2025 ADR split)' : 'budget (2026 YTD ADR split)' });
      if (m <= through) flags.push('No ' + BL_MON[m] + ' ' + BL_BASE_YEAR + ' actual on file — used budget.');
      continue;
    }
    if (l && l.rn > 0) {
      raw.push({ rn: l.rn, rev: l.rev, avail: avail, src: BL_BASE_YEAR - 1 + ' actual (no ' + BL_BASE_YEAR + ' data)' });
      flags.push(BL_MON[m] + ': no ' + BL_BASE_YEAR + ' actual or budget — used ' + (BL_BASE_YEAR - 1) + '.');
    } else {
      raw.push({ rn: 0, rev: 0, avail: avail, src: 'no data' });
      flags.push(BL_MON[m] + ': no data at all.');
      manual = true;
    }
  }

  // b) event normalization
  const ev = blClone_(raw);
  let event = { applied: false };
  if (R.event) {
    const months = (S.event.months || []).filter(m => m >= 0 && m <= 11);
    const cap = (Number(S.event.trend_cap_pct) || 0) / 100;
    const have25 = months.length && months.every(m => h25[m] && h25[m].rn > 0 && h25[m].rev > 0);
    if (have25) {
      let rn26 = 0, rev26 = 0, rn25 = 0, rev25 = 0, n = 0;
      for (let m = 0; m <= through; m++) {
        if (months.indexOf(m) >= 0 || raw[m].src !== 'actual' || !(h25[m] && h25[m].rn > 0)) continue;
        rn26 += raw[m].rn; rev26 += raw[m].rev; rn25 += h25[m].rn; rev25 += h25[m].rev; n++;
      }
      const clamp = x => Math.max(-cap, Math.min(cap, x));
      const occT = rn25 && rn26 ? clamp(rn26 / rn25 - 1) : 0;
      const adrT = rn25 && rn26 ? clamp((rev26 / rn26) / (rev25 / rn25) - 1) : 0;
      months.forEach(m => {
        const rn = Math.min(ev[m].avail, h25[m].rn * (1 + occT));
        const adr = (h25[m].rev / h25[m].rn) * (1 + adrT);
        ev[m] = { rn: rn, rev: rn * adr, avail: ev[m].avail, src: (BL_BASE_YEAR - 1) + ' × trend' };
      });
      event = { applied: true, method: (BL_BASE_YEAR - 1) + ' ' + months.map(m => BL_MON[m]).join('/') + ' × trend', occ_trend: occT, adr_trend: adrT, trend_months: n };
      if (!n) flags.push('Event months: no overlapping ' + BL_BASE_YEAR + '/' + (BL_BASE_YEAR - 1) + ' actual months to measure trend — used ' + (BL_BASE_YEAR - 1) + ' flat.');
    } else {
      manual = true;
      const sh = BL_SHOULDER.filter(m => raw[m].rn > 0);
      if (sh.length) {
        const shRn = sh.reduce((a, m) => a + raw[m].rn, 0), shRev = sh.reduce((a, m) => a + raw[m].rev, 0), shAvail = sh.reduce((a, m) => a + raw[m].avail, 0);
        const seas = blPeerSeasonality_(prop, R.tier, D, S, months);
        months.forEach(m => {
          const r = seas.months[m];
          const rn = Math.min(ev[m].avail, (shRn / shAvail) * r.occ * ev[m].avail);
          const adr = (shRev / shRn) * r.adr;
          ev[m] = { rn: rn, rev: rn * adr, avail: ev[m].avail, src: 'shoulder × peer seasonality' };
        });
        event = { applied: true, method: 'May/Aug shoulder × peer seasonality (' + seas.basis + ')', peers: seas.peers, ratios: seas.months };
      } else {
        event = { applied: false, method: 'no shoulder-month data — left as ' + BL_BASE_YEAR + ' raw' };
      }
      flags.push('Manual review required: no ' + (BL_BASE_YEAR - 1) + ' Jun/Jul history, so Jun/Jul were derived from shoulder-month seasonality.');
    }
  }

  // c) Easter shift (March -> April, at March ADR)
  const ea = blClone_(ev);
  let moved = 0;
  if (R.easter_pct > 0 && ea[2].rn > 0) {
    const adrMar = ea[2].rev / ea[2].rn;
    moved = Math.min(ea[2].rn * R.easter_pct / 100, Math.max(0, ea[3].avail - ea[3].rn));
    ea[2].rn -= moved; ea[2].rev -= moved * adrMar;
    ea[3].rn += moved; ea[3].rev += moved * adrMar;
  }

  // d) growth, uniform across months
  const fin = ea.map((x, m) => {
    const avail = rooms * blDim_(BL_YEAR, m);
    const scale = x.avail ? avail / x.avail : 1;
    const rn = Math.min(avail, x.rn * scale * (1 + R.growth_occ / 100));
    const adr = x.rn > 0 ? (x.rev / x.rn) * (1 + R.growth_adr / 100) : 0;
    return { rn: rn, rev: rn * adr, avail: avail, src: 'baseline' };
  });

  const t = { raw: blTotals_(raw), event: blTotals_(ev), easter: blTotals_(ea), final: blTotals_(fin) };
  return {
    property_id: pid, name: prop.name, brand: prop.brand, market: prop.market, rooms: rooms,
    label: BL_LABEL, assumptions: R, event: event, easter_moved_rn: moved,
    stages: { raw: raw, event: ev, easter: ea, final: fin }, totals: t,
    waterfall: [
      { step: BL_BASE_YEAR + ' raw', kind: 'total', rev: t.raw.rev, revpar: t.raw.revpar },
      { step: 'Event adj', kind: 'delta', rev: t.event.rev - t.raw.rev, revpar: t.event.revpar - t.raw.revpar },
      { step: 'Easter shift', kind: 'delta', rev: t.easter.rev - t.event.rev, revpar: t.easter.revpar - t.event.revpar },
      { step: 'Growth', kind: 'delta', rev: t.final.rev - t.easter.rev, revpar: t.final.revpar - t.easter.revpar },
      { step: BL_YEAR + ' baseline', kind: 'total', rev: t.final.rev, revpar: t.final.revpar }
    ],
    manual_review: manual, flags: flags
  };
}

// ── GM 2027 forecast, rolled to months ──
function blGmData_(pids) {
  const out = {};
  pids.forEach(pid => { out[pid] = { status: '', months: BL_MON.map(() => ({ rn: 0, rev: 0, days: 0 })), days: 0 }; });
  try {
    const st = fcReadStatus_().map;
    const values = fcReadForecast_().values;
    for (let i = 1; i < values.length; i++) {
      const r = values[i], g = out[String(r[0])];
      if (!g || String(r[1]).slice(0, 4) !== String(BL_YEAR)) continue;
      if (r[2] === '' || r[3] === '' || !isFinite(Number(r[2])) || !isFinite(Number(r[3]))) continue;
      const m = Number(String(r[1]).slice(5, 7)) - 1;
      g.months[m].rn += Number(r[2]); g.months[m].rev += Number(r[2]) * Number(r[3]); g.months[m].days++; g.days++;
    }
    pids.forEach(pid => { out[pid].status = st[pid] ? String(st[pid].status || '') : ''; out[pid].justification = st[pid] ? String(st[pid].baseline_justification || '') : ''; });
  } catch (e) { /* 2027 forecast sheets not set up yet — every hotel is on the baseline */ }
  return out;
}
function blGmSubmitted_(status) { return status === 'SUBMITTED' || status === 'APPROVED'; }

// ── handlers ──

function blPortfolio_(me) {
  const S = blSettings_(), D = blLoadData_();
  const props = getAuthorizedProperties_(me.grant);
  const gm = blGmData_(props.map(p => String(p.property_id)));
  const thr = Number(S.gm_threshold_pct) || 0;
  const tot = { rn: 0, rev: 0, avail: 0 };
  let nGm = 0, nBase = 0, nManual = 0;
  const rows = props.map(p => {
    const b = blCompute_(p, D, S);
    const g = gm[b.property_id];
    const avail27 = b.totals.final.avail;
    const gmRev = g.months.reduce((a, x) => a + x.rev, 0), gmRn = g.months.reduce((a, x) => a + x.rn, 0);
    const submitted = blGmSubmitted_(g.status);
    const gmRevpar = g.days === 365 || submitted ? (avail27 ? gmRev / avail27 : 0) : null;
    const diff = gmRevpar !== null && b.totals.final.revpar ? (gmRevpar / b.totals.final.revpar - 1) * 100 : null;
    const eff = submitted ? { rn: gmRn, rev: gmRev } : { rn: b.totals.final.rn, rev: b.totals.final.rev };
    tot.rn += eff.rn; tot.rev += eff.rev; tot.avail += avail27;
    if (submitted) nGm++; else nBase++;
    if (b.manual_review) nManual++;
    return {
      property_id: b.property_id, name: b.name, brand: b.brand, market: b.market, rooms: b.rooms,
      source: submitted ? 'gm' : 'baseline', source_label: submitted ? 'GM-submitted' : BL_LABEL,
      gm_status: g.status, gm_days: g.days,
      baseline: { rev: b.totals.final.rev, occ: b.totals.final.occ, adr: b.totals.final.adr, revpar: b.totals.final.revpar },
      raw2026: { rev: b.totals.raw.rev, revpar: b.totals.raw.revpar },
      gm: gmRevpar !== null ? { rev: gmRev, occ: avail27 ? gmRn / avail27 : 0, adr: gmRn ? gmRev / gmRn : 0, revpar: gmRevpar } : null,
      diff_pct: diff, over_threshold: diff !== null && Math.abs(diff) > thr, justification: g.justification || '',
      manual_review: b.manual_review, event: b.assumptions.event, tier: b.assumptions.tier, mix: b.assumptions.mix
    };
  });
  return {
    year: BL_YEAR, label: BL_LABEL, threshold_pct: thr, is_corporate: me.isCorporate,
    counts: { gm_submitted: nGm, baseline: nBase, manual_review: nManual, total: rows.length },
    totals: { rn: tot.rn, rev: tot.rev, avail: tot.avail, occ: tot.avail ? tot.rn / tot.avail : 0, adr: tot.rn ? tot.rev / tot.rn : 0, revpar: tot.avail ? tot.rev / tot.avail : 0 },
    properties: rows, settings_updated_at: S.updated_at, settings_updated_by: S.updated_by
  };
}

function blPropertyDetail_(me, p) {
  const pid = String(p.property_id || '').trim();
  if (!pid || !isAuthorizedForProperty_(me.grant, pid)) return { error: 'not_authorized', message: 'Not authorized for this property.', _status: 403 };
  const prop = getPropertyById_(pid);
  if (!prop) return { error: 'not_found', message: 'Unknown property.', _status: 404 };
  const S = blSettings_();
  const b = blCompute_(prop, blLoadData_(), S);
  const g = blGmData_([pid])[pid];
  b.gm = { status: g.status, submitted: blGmSubmitted_(g.status), days: g.days, months: g.months, justification: g.justification || '' };
  b.threshold_pct = Number(S.gm_threshold_pct) || 0;
  b.year = BL_YEAR; b.base_year = BL_BASE_YEAR; b.months = BL_MON;
  b.actuals_through_month = S.actuals_through_month;
  return b;
}

/** Small version for the Hub's budget editor (forecast_get). Never throws. */
function blCompactForProperty_(pid) {
  try {
    const prop = getPropertyById_(pid);
    if (!prop) return null;
    const S = blSettings_();
    const b = blCompute_(prop, blLoadData_(), S);
    return {
      label: BL_LABEL, threshold_pct: Number(S.gm_threshold_pct) || 0, manual_review: b.manual_review, flags: b.flags,
      months: b.stages.final.map(x => ({ rn: Math.round(x.rn), rev: Math.round(x.rev), avail: x.avail })),
      totals: b.totals.final
    };
  } catch (e) {
    return null;
  }
}

/** forecast_submit gate: {status: {...columns}} or {error: 'justification_required', ...}. */
function blSubmitGate_(pid, days, justification) {
  const bl = blCompactForProperty_(pid);
  if (!bl || !bl.totals.revpar) return { status: {} };
  let rev = 0;
  Object.keys(days).forEach(d => { rev += days[d][0] * days[d][1]; });
  const gmRevpar = bl.totals.avail ? rev / bl.totals.avail : 0;
  const diff = (gmRevpar / bl.totals.revpar - 1) * 100;
  const note = String(justification || '').trim().slice(0, 1000);
  const r2 = x => Math.round(x * 100) / 100;
  if (Math.abs(diff) > bl.threshold_pct && note.length < 10) {
    return {
      error: 'justification_required', _status: 400,
      message: 'Your 2027 RevPAR ($' + gmRevpar.toFixed(2) + ') is ' + (diff > 0 ? '+' : '') + diff.toFixed(1) + '% vs the system baseline ($' +
        bl.totals.revpar.toFixed(2) + '), beyond the ±' + bl.threshold_pct + '% threshold. Add a short justification (at least 10 characters) to submit.',
      gm_revpar: r2(gmRevpar), baseline_revpar: r2(bl.totals.revpar), diff_pct: r2(diff), threshold_pct: bl.threshold_pct
    };
  }
  return { status: { baseline_revpar: r2(bl.totals.revpar), gm_revpar: r2(gmRevpar), baseline_diff_pct: r2(diff), baseline_justification: note } };
}

// ── settings (corporate only) ──

function blSettingsGet_(me) {
  if (!me.isCorporate) return { error: 'not_authorized', message: 'Only corporate can see baseline settings.', _status: 403 };
  const S = blSettings_();
  const props = getAuthorizedProperties_(me.grant);
  return {
    settings: S, defaults: blDefaults_(), tiers: BL_TIERS, tier_labels: BL_TIER_LABEL, mixes: BL_MIXES,
    markets: Array.from(new Set(props.map(p => String(p.market || '').trim()).filter(Boolean))).sort(),
    properties: props.map(p => ({ property_id: String(p.property_id), name: p.name, market: p.market, inferred_tier: blInferTier_(p), inferred_mix: blInferMix_(p) }))
  };
}

function blSettingsSave_(me, p) {
  if (!me.isCorporate) return { error: 'not_authorized', message: 'Only corporate can change baseline settings.', _status: 403 };
  const s = p.settings || {};
  const err = [];
  const num = (v, lo, hi, label, blankOk) => {
    if ((v === '' || v == null) && blankOk) return '';
    const n = Number(v);
    if (v === '' || v == null || !isFinite(n) || n < lo || n > hi) { err.push(label + ' must be between ' + lo + ' and ' + hi + '.'); return 0; }
    return Math.round(n * 1000) / 1000;
  };
  const g = s.growth || {}, e = s.event || {}, ea = s.easter || {};
  const out = {
    actuals_through_month: Math.floor(num(s.actuals_through_month, 0, 11, 'Actuals through month')),
    event: {
      markets: (Array.isArray(e.markets) ? e.markets : []).map(x => String(x).trim().slice(0, 40)).filter(Boolean).slice(0, 12),
      months: (Array.isArray(e.months) ? e.months : [5, 6]).map(Number).filter(m => m >= 0 && m <= 11 && Math.floor(m) === m).slice(0, 4),
      trend_cap_pct: num(e.trend_cap_pct, 0, 50, 'Event trend cap')
    },
    easter: {}, growth: { default: {}, tiers: {}, markets: {} },
    gm_threshold_pct: num(s.gm_threshold_pct, 0, 100, 'GM threshold'),
    properties: {}, updated_by: me.username, updated_at: new Date().toISOString()
  };
  BL_MIXES.forEach(m => { out.easter[m] = num(ea[m], 0, 50, 'Easter % (' + m + ')'); });
  out.growth.default = { adr: num((g.default || {}).adr, -20, 20, 'Default ADR growth'), occ: num((g.default || {}).occ, -20, 20, 'Default occupancy growth') };
  BL_TIERS.forEach(t => {
    const x = (g.tiers || {})[t] || {};
    out.growth.tiers[t] = { adr: num(x.adr, -20, 20, BL_TIER_LABEL[t] + ' ADR', true), occ: num(x.occ, -20, 20, BL_TIER_LABEL[t] + ' occupancy', true) };
  });
  const knownMarkets = new Set(getAuthorizedProperties_(me.grant).map(x => String(x.market || '').trim()));
  Object.keys(g.markets || {}).forEach(mk => {
    if (!knownMarkets.has(mk)) return;
    const x = g.markets[mk] || {};
    const o = { adr: num(x.adr, -20, 20, mk + ' ADR', true), occ: num(x.occ, -20, 20, mk + ' occupancy', true) };
    if (o.adr !== '' || o.occ !== '') out.growth.markets[mk] = o;
  });
  Object.keys(s.properties || {}).forEach(pid => {
    if (!getPropertyById_(pid)) return;
    const x = s.properties[pid] || {};
    const o = {
      tier: BL_TIERS.indexOf(x.tier) >= 0 ? x.tier : '',
      mix: BL_MIXES.indexOf(x.mix) >= 0 ? x.mix : '',
      event: x.event === 'yes' || x.event === 'no' ? x.event : 'auto',
      easter_pct: num(x.easter_pct, 0, 50, 'Easter % for ' + pid, true)
    };
    if (o.tier || o.mix || o.event !== 'auto' || o.easter_pct !== '') out.properties[pid] = o;
  });
  if (err.length) return { error: 'bad_request', message: err.join(' '), _status: 400 };

  const json = JSON.stringify(out);
  if (json.length > 8500) return { error: 'bad_request', message: 'Too many overrides to store — remove some.', _status: 400 };
  PropertiesService.getScriptProperties().setProperty('BASELINE_SETTINGS', json);
  try { fcLog_([[out.updated_at, '', '', 'BASELINE_SETTINGS', '', '', '', '', json.slice(0, 45000), me.username]]); } catch (x) { /* forecast log not set up */ }
  return { ok: true, settings: blSettings_() };
}
