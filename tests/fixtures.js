// Test data: 5 hotels, groups for a GM, a regional VP and corporate. Loaded AFTER the .gs files
// (it uses their sheet-ID constants). All numbers are synthetic.
const FX = {
  OWNER: 'tjohnson@rgbhospitality.com',
  CORP: 'corp.analyst@rgbhospitality.com',
  GM: 'gm.portland@rgbhospitality.com',
  RVP: 'rvp.dallas@rgbhospitality.com',
  NOBODY: 'new.hire@rgbhospitality.com',
  OUTSIDER: 'someone@gmail.com',
  props: [
    // id, name, brand, market, rooms, base ADR, has 2025 history
    ['HOUHX', 'Hampton Inn Houston Galleria', 'HILTON', 'Houston', 120, 130, true],
    ['PROV-DALES', 'Embassy Suites Dallas', 'HILTON', 'Dallas', 150, 170, false],
    ['DALLQ', 'La Quinta Dallas North', 'WYNDHAM', 'Dallas', 100, 85, true],
    ['10795674', 'Hampton Inn Portland', 'HILTON', 'Corpus Christi', 80, 120, true],
    ['CRPTRU', 'TRU by Hilton Rockport', 'HILTON', 'Corpus Christi', 90, 95, true]
  ],
  OCC: [.58, .64, .74, .72, .70, .76, .75, .68, .63, .66, .57, .50],
  ADR: [.95, .98, 1.04, 1.02, 1.0, 1.03, 1.02, .99, .98, 1.0, .96, .93]
};
FX.dim = (y, m) => new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
FX.y25 = (p, m) => { const rn = Math.round(p[4] * FX.dim(2025, m) * FX.OCC[m]); return { rn, rev: Math.round(rn * p[5] * FX.ADR[m] * 100) / 100 }; };
FX.y26 = (p, m) => {
  const host = p[3] === 'Houston' || p[3] === 'Dallas';
  const wc = host && (m === 5 || m === 6);
  const occ = wc ? Math.min(0.97, FX.OCC[m] + 0.15) : FX.OCC[m] * 1.02;
  const rn = Math.round(p[4] * FX.dim(2026, m) * occ);
  const adr = p[5] * FX.ADR[m] * (wc ? 1.6 : 1.03);
  return { rn, rev: Math.round(rn * adr * 100) / 100 };
};

function loadFixtures() {
  Object.assign(MOCK, { user: '', mail: [], sheets: {}, props: {}, cache: {}, triggers: [], groups: {} });
  MOCK.groups['corporate@rgbhospitality.com'] = [FX.CORP, FX.OWNER];
  MOCK.groups['gm-10795674@rgbhospitality.com'] = [FX.GM];
  MOCK.groups['region-dallas@rgbhospitality.com'] = [FX.RVP];
  MOCK.groups['gm-prov-dales@rgbhospitality.com'] = [];

  mockAddSpreadsheet(PROPERTIES_SHEET_ID, [{ name: 'Read Me', rows: [['notes']] }, { name: 'Properties', rows: [['property_id', 'inn_code', 'name', 'brand', 'market', 'rooms', 'gm_email', 'relationship', 'notes']]
    .concat(FX.props.map(p => [p[0], '', p[1], p[2], p[3], p[4], '', 'RGB', ''])) }]);
  mockAddSpreadsheet(PERMISSIONS_SHEET_ID, [{ name: 'Permissions', rows: [['email', 'scope_type', 'scope_value', 'role', 'notes'], ['franchise.gm@example.com', 'property', 'CRPTRU', 'GM', '']] }]);
  const flash = [['property_id', 'name', 'brand', 'market', 'rooms', 'month_index', 'room_rev', 'occupied_rooms', 'total_rev']];
  const hist = [['property_id', 'name', 'year', 'month_index', 'room_rev', 'occupied_rooms']];
  const budget = [['property_id', 'name', 'month_index', 'budget_room_rev']];
  FX.props.forEach(p => {
    for (let m = 0; m < 12; m++) {
      const a = FX.y26(p, m), h = FX.y25(p, m);
      flash.push([p[0], p[1], p[2], p[3], p[4], m, m <= 9 ? a.rev : 0, m <= 9 ? a.rn : 0, m <= 9 ? a.rev * 1.1 : 0]);
      if (p[6]) hist.push([p[0], p[1], 2025, m, h.rev, h.rn]);
      budget.push([p[0], p[1], m, Math.round(a.rev)]);
    }
  });
  mockAddSpreadsheet(FLASH_SHEET_ID, [{ name: 'Flash', rows: flash }]);
  mockAddSpreadsheet(HISTORICAL_SHEET_ID, [{ name: 'Historical', rows: hist }]);
  mockAddSpreadsheet(BUDGET_SHEET_ID, [{ name: 'Budget', rows: budget }]);
  mockAddSpreadsheet(CREDENTIALS_SHEET_ID, [{ name: 'Credentials', rows: [
    ['username', 'new_password', 'password_hash', 'salt', 'property_id', 'role', 'display_name', 'active', 'last_synced'],
    ['crppo.audit', '', hashPassword_('auditpw', 's1'), 's1', '10795674', 'ENTRY', 'Portland audit', 'TRUE', ''],
    ['franchise.gm@example.com', '', hashPassword_('fpw', 's2'), 's2', '', 'READ', 'Franchise GM', 'TRUE', ''],
    ['gm.portland@rgbhospitality.com', '', hashPassword_('oldpw', 's3'), 's3', '', 'READ', 'GM (old pw)', 'TRUE', '']
  ] }]);
  MOCK.sheets[PERMISSIONS_SHEET_ID].sheets[0].rows.push(['gm.portland@rgbhospitality.com', 'property', '10795674', 'GM', '']);
  mockAddSpreadsheet(DAR_SHEET_ID, [{ name: 'DAR', rows: [['property_id', 'date', 'submitted_by', 'submitted_at', 'status', 'validated_by', 'validated_at', 'flag_note']] }]);
  mockAddSpreadsheet(DAR_LOG_SHEET_ID, [{ name: 'Log', rows: [['logged_at', 'property_id', 'date', 'submitted_by']] }]);
  mockAddSpreadsheet(RATE_SHOP_SHEET_ID, [{ name: 'Rates', rows: [['property_name', 'date', 'rate_display', 'rate_num']] }]);
}
