// In-memory stand-ins for the Apps Script services the Hub uses, so apps-script/*.gs can run in a
// browser. Only what the code calls is implemented. MOCK.user is "who is signed in" for a call.
const MOCK = {
  owner: 'tjohnson@rgbhospitality.com',
  user: '',
  groups: {},          // group email -> [member emails]
  mail: [],
  sheets: {},          // spreadsheet id -> FakeSpreadsheet
  props: {},
  cache: {},
  triggers: [],
  nextId: 1
};

// ── SHA-256 (sync; Apps Script's computeDigest is sync) ──
function mockSha256Bytes(str) {
  const bytes = new TextEncoder().encode(str);
  const K = [0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
  let H = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19];
  const l = bytes.length, withPad = ((l + 9 + 63) >> 6) << 6;
  const m = new Uint8Array(withPad); m.set(bytes); m[l] = 0x80;
  const bits = l * 8; const dv = new DataView(m.buffer);
  dv.setUint32(withPad - 4, bits >>> 0); dv.setUint32(withPad - 8, Math.floor(bits / 2 ** 32));
  const w = new Uint32Array(64), rot = (x, n) => (x >>> n) | (x << (32 - n));
  for (let o = 0; o < withPad; o += 64) {
    for (let i = 0; i < 16; i++) w[i] = dv.getUint32(o + i * 4);
    for (let i = 16; i < 64; i++) {
      const s0 = rot(w[i-15],7) ^ rot(w[i-15],18) ^ (w[i-15] >>> 3), s1 = rot(w[i-2],17) ^ rot(w[i-2],19) ^ (w[i-2] >>> 10);
      w[i] = (w[i-16] + s0 + w[i-7] + s1) >>> 0;
    }
    let [a,b,c,d,e,f,g,h] = H;
    for (let i = 0; i < 64; i++) {
      const t1 = (h + (rot(e,6) ^ rot(e,11) ^ rot(e,25)) + ((e & f) ^ (~e & g)) + K[i] + w[i]) >>> 0;
      const t2 = ((rot(a,2) ^ rot(a,13) ^ rot(a,22)) + ((a & b) ^ (a & c) ^ (b & c))) >>> 0;
      h = g; g = f; f = e; e = (d + t1) >>> 0; d = c; c = b; b = a; a = (t1 + t2) >>> 0;
    }
    H = H.map((x, i) => (x + [a,b,c,d,e,f,g,h][i]) >>> 0);
  }
  const out = [];
  H.forEach(x => { for (let s = 24; s >= 0; s -= 8) { const b = (x >>> s) & 0xff; out.push(b > 127 ? b - 256 : b); } });
  return out;   // signed bytes, like Apps Script
}

// ── Sheets ──
class FakeRange {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const v = (this.sheet.rows[this.r - 1 + i] || [])[this.c - 1 + j]; row.push(v === undefined ? '' : v); } out.push(row); } return out; }
  setValues(vals) { vals.forEach((row, i) => row.forEach((v, j) => this.sheet.put(this.r + i, this.c + j, v))); return this; }
  setValue(v) { this.sheet.put(this.r, this.c, v); return this; }
  setNumberFormat() { return this; } setFontWeight() { return this; } setFontSize() { return this; } setWrap() { return this; }
}
class FakeSheet {
  constructor(ss, name, rows) { this.ss = ss; this.name = name; this.rows = rows || []; this.maxRows = 1000; this.maxCols = 26; }
  width() { return this.rows.reduce((a, r) => Math.max(a, r.length), 0); }
  put(r, c, v) { while (this.rows.length < r) this.rows.push([]); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; this.maxCols = Math.max(this.maxCols, c); }
  getName() { return this.name; } setName(n) { this.name = n; return this; }
  getParent() { return this.ss; }
  getLastRow() { let n = this.rows.length; while (n && this.rows[n - 1].every(v => v === '' || v == null)) n--; return n; }
  getLastColumn() { return this.width(); }
  getMaxRows() { return Math.max(this.maxRows, this.rows.length); } getMaxColumns() { return Math.max(this.maxCols, this.width()); }
  insertRowsAfter(after, n) { this.maxRows += n; } insertColumnsAfter(after, n) { this.maxCols += n; }
  getDataRange() { const w = this.width(), h = this.getLastRow(); return new FakeRange(this, 1, 1, Math.max(h, 1), Math.max(w, 1)); }
  getRange(r, c, nr, nc) { if (typeof r === 'string') return new FakeRange(this, 1, 1, 1, 1); return new FakeRange(this, r, c, nr || 1, nc || 1); }
  appendRow(row) { this.rows.splice(this.getLastRow(), 0, row.slice()); this.rows.length = this.getLastRow(); return this; }
  clear() { this.rows = []; return this; } setFrozenRows() { return this; } setColumnWidth() { return this; }
}
class FakeSpreadsheet {
  constructor(id, title) { this.id = id; this.title = title; this.sheets = []; this.tz = 'America/Chicago'; }
  getId() { return this.id; } getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id; }
  getSheets() { return this.sheets; } getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  insertSheet(n, idx) { const s = new FakeSheet(this, n); if (idx === 0) this.sheets.unshift(s); else this.sheets.push(s); return s; }
  getSpreadsheetTimeZone() { return this.tz; }
}
function mockAddSpreadsheet(id, tabs) {
  const ss = new FakeSpreadsheet(id, id);
  tabs.forEach(t => ss.sheets.push(new FakeSheet(ss, t.name, t.rows.map(r => r.slice()))));
  MOCK.sheets[id] = ss; return ss;
}
const SpreadsheetApp = {
  openById(id) { const s = MOCK.sheets[id]; if (!s) throw new Error('No spreadsheet ' + id); return s; },
  create(title) { const id = 'created-' + (MOCK.nextId++); const ss = new FakeSpreadsheet(id, title); ss.sheets.push(new FakeSheet(ss, 'Sheet1')); MOCK.sheets[id] = ss; return ss; }
};

// ── everything else ──
const PropertiesService = { getScriptProperties() { return {
  getProperty: k => (k in MOCK.props ? MOCK.props[k] : null), setProperty: (k, v) => { MOCK.props[k] = String(v); }, deleteProperty: k => { delete MOCK.props[k]; }
}; } };
const CacheService = { getScriptCache() { return {
  get: k => MOCK.cache[k] || null, put: (k, v) => { MOCK.cache[k] = v; },
  getAll: keys => { const o = {}; keys.forEach(k => { if (MOCK.cache[k]) o[k] = MOCK.cache[k]; }); return o; },
  removeAll: keys => keys.forEach(k => delete MOCK.cache[k])
}; } };
const Session = {
  getActiveUser: () => ({ getEmail: () => MOCK.user }),
  getEffectiveUser: () => ({ getEmail: () => MOCK.owner }),
  getScriptTimeZone: () => 'America/Chicago',
  getTemporaryActiveUserKey: () => (MOCK.user ? 'tmpkey-' + MOCK.user.length : '')
};
const Utilities = {
  DigestAlgorithm: { SHA_256: 'SHA_256' }, Charset: { UTF_8: 'UTF_8' },
  computeDigest: (alg, s) => mockSha256Bytes(String(s)),
  getUuid: () => crypto.randomUUID(),
  formatDate(d, tz, fmt) {
    const p = new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
    if (fmt === 'yyyy-MM-dd') return p;
    if (fmt === 'H') return MOCK.hour != null ? String(MOCK.hour) : String(Number(new Intl.DateTimeFormat('en-US', { timeZone: tz, hour: 'numeric', hourCycle: 'h23' }).format(d)));
    throw new Error('mock formatDate: unsupported ' + fmt);
  },
  parseCsv: () => []
};
const GroupsApp = { getGroupByEmail(g) {
  if (!(g in MOCK.groups)) throw new Error('Group not found: ' + g);
  return { hasUser: e => MOCK.groups[g].indexOf(e) >= 0 };
} };
const LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
const MailApp = { sendEmail(o) { MOCK.mail.push(o); } };
class FakeOutput { constructor(kind, v) { this.kind = kind; this.v = v; } setTitle(t) { this.title = t; return this; } addMetaTag() { return this; } setMimeType() { return this; } }
const HtmlService = {
  createHtmlOutput: html => new FakeOutput('html', html),
  createTemplateFromFile: file => ({ file, evaluate() { return new FakeOutput('template', { file: this.file, boot: this.boot }); } })
};
const ContentService = { MimeType: { JSON: 'json' }, createTextOutput: s => new FakeOutput('text', s) };
const ScriptApp = {
  getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/DOMAINDEPLOY/exec' }),
  getProjectTriggers: () => MOCK.triggers, deleteTrigger: t => { MOCK.triggers = MOCK.triggers.filter(x => x !== t); },
  newTrigger(fn) { const t = { getHandlerFunction: () => fn }; const chain = { timeBased: () => chain, everyDays: () => chain, atHour: h => { t.hour = h; return chain; }, nearMinute: () => chain, create: () => { MOCK.triggers.push(t); return t; } }; return chain; }
};
const DriveApp = { getFilesByName: () => ({ hasNext: () => false }) };
const MimeType = { GOOGLE_SHEETS: 'sheets' };
const GmailApp = {};
const Logger = { log: (...a) => console.log('[Logger]', ...a) };
