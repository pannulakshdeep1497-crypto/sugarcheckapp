/**
 * Blood sugar log – Google Sheet backend
 * Paste this into Extensions > Apps Script of your Google Sheet.
 * 1) Change PASSCODE below to your own (letters and numbers, 12+ characters).
 * 2) Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 * 3) Copy the web app address (ends in /exec) into the app, with the same passcode.
 */
const PASSCODE = 'CHANGE-ME-to-something-long';

const DAY_SHEET = 'Readings';
const PROFILE_SHEET = 'Profile';
const DAY_HEADER = [
  'Date', 'Before breakfast', 'After breakfast', 'Breakfast check',
  'After lunch', 'Lunch check', 'After supper', 'Supper check',
  'Breakfast med dose', 'Lunch med dose', 'Supper med dose', 'Bedtime med dose',
  'Notes', 'Breakfast start (ms)', 'Lunch start (ms)', 'Supper start (ms)', 'Updated (ms)'
];
const PROFILE_FIELDS = ['name', 'dob', 'meds', 'medsChanged', 'u'];
const PROFILE_LABELS = ['Full name', 'Date of birth', 'Medications', 'Medication last changed', 'Updated (ms)'];

function doGet() {
  return ContentService.createTextOutput('Blood sugar log backend is running.');
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  if (String(p.key || '') !== PASSCODE || PASSCODE.indexOf('CHANGE-ME') === 0) {
    return out_({ ok: false, error: 'bad_key' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (p.action === 'all') return out_({ ok: true, profile: readProfile_(), days: readDays_() });
    if (p.action === 'saveDay') return out_({ ok: true, day: saveDay_(String(p.date), p.day || {}) });
    if (p.action === 'saveProfile') return out_({ ok: true, profile: saveProfile_(p.profile || {}) });
    return out_({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- sheets ---------- */
function daySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DAY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DAY_SHEET);
    sh.getRange(1, 1, 1, DAY_HEADER.length).setValues([DAY_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    sh.hideColumns(14, 4);
  }
  return sh;
}
function profileSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PROFILE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PROFILE_SHEET);
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange(1, 1, PROFILE_LABELS.length, 1).setValues(PROFILE_LABELS.map(l => [l])).setFontWeight('bold');
  }
  return sh;
}

/* ---------- conversion ---------- */
function dateText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim();
}
function num_(v) { return (v === '' || v === null || v === undefined || isNaN(Number(v))) ? null : Number(v); }
function safe_(v) {
  v = v == null ? '' : String(v);
  return /^[=+@\-]/.test(v) ? "'" + v : v;
}
function unsafe_(v) { v = v == null ? '' : String(v); return /^'[=+@\-]/.test(v) ? v.slice(1) : v; }
function reading_(r) { r = r || {}; return { v: num_(r.v), k: r.k === '2h' ? '2h' : '1h', start: num_(r.start) }; }

function dayToRow_(date, d) {
  const f = reading_(d.fast), bf = reading_(d.bf), lu = reading_(d.lu), su = reading_(d.su);
  const kind = r => (r.v == null ? '' : (r.k === '2h' ? '2 hr' : '1 hr'));
  const n = x => (x == null ? '' : x);
  return [
    date, n(f.v), n(bf.v), kind(bf), n(lu.v), kind(lu), n(su.v), kind(su),
    safe_(d.mB), safe_(d.mL), safe_(d.mS), safe_(d.mBed), safe_(d.notes),
    n(bf.start), n(lu.start), n(su.start), Number(d.u) || Date.now()
  ];
}
function rowToDay_(r) {
  const k = x => (String(x).indexOf('2') === 0 ? '2h' : '1h');
  return {
    fast: { v: num_(r[1]) },
    bf: { v: num_(r[2]), k: k(r[3]), start: num_(r[13]) },
    lu: { v: num_(r[4]), k: k(r[5]), start: num_(r[14]) },
    su: { v: num_(r[6]), k: k(r[7]), start: num_(r[15]) },
    mB: unsafe_(r[8]), mL: unsafe_(r[9]), mS: unsafe_(r[10]), mBed: unsafe_(r[11]), notes: unsafe_(r[12]),
    u: Number(r[16]) || 0
  };
}

/* ---------- read / write ---------- */
function readDays_() {
  const sh = daySheet_(), last = sh.getLastRow(), days = {};
  if (last < 2) return days;
  sh.getRange(2, 1, last - 1, DAY_HEADER.length).getValues().forEach(r => {
    const date = dateText_(r[0]);
    if (/^\d{4}-\d{2}-\d{2}$/.test(date)) days[date] = rowToDay_(r);
  });
  return days;
}
function saveDay_(date, day) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('bad date');
  const sh = daySheet_(), last = sh.getLastRow();
  let rowIndex = -1;
  if (last >= 2) {
    const dates = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) if (dateText_(dates[i][0]) === date) { rowIndex = i + 2; break; }
  }
  if (rowIndex > 0) {
    const current = rowToDay_(sh.getRange(rowIndex, 1, 1, DAY_HEADER.length).getValues()[0]);
    if ((current.u || 0) > (Number(day.u) || 0)) return current; // the other phone saved something newer
    sh.getRange(rowIndex, 1, 1, DAY_HEADER.length).setValues([dayToRow_(date, day)]);
  } else {
    sh.appendRow(dayToRow_(date, day));
    const n = sh.getLastRow();
    if (n > 2) sh.getRange(2, 1, n - 1, DAY_HEADER.length).sort({ column: 1, ascending: true });
  }
  return rowToDay_(dayToRow_(date, day));
}
function readProfile_() {
  const vals = profileSheet_().getRange(1, 2, PROFILE_FIELDS.length, 1).getValues();
  const p = {};
  PROFILE_FIELDS.forEach((f, i) => { p[f] = f === 'u' ? (Number(vals[i][0]) || 0) : (f === 'dob' || f === 'medsChanged' ? dateText_(vals[i][0]) : unsafe_(vals[i][0])); });
  return p;
}
function saveProfile_(p) {
  const current = readProfile_();
  if ((current.u || 0) > (Number(p.u) || 0)) return current;
  profileSheet_().getRange(1, 2, PROFILE_FIELDS.length, 1)
    .setValues(PROFILE_FIELDS.map(f => [f === 'u' ? (Number(p.u) || Date.now()) : safe_(p[f])]));
  return readProfile_();
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
