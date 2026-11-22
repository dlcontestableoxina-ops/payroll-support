/* FILE 1 of 4 · Code.gs · server code. Replace everything in Code.gs with this. */
/**
 * PAYROLL SUPPORT · STATIONS DEPARTMENT (Rephil / CAPAREL)
 * Modules: Station Records (List of Stations · Approved Manpower) · Attendance Records (Daily Attendance Checking · HR Attendance Processed: Weekly Processed-Validation)
 * Google Sheet = database. Run "setup" once (or just sign in) and every tab is created and filled.
 */

var APP_NAME = 'Payroll Support';
var APP_VERSION = 'v26 web';
var MP_SEED_VERSION = '2026-11-05'; // latest APPROVED MANPOWER file built into the app
var SHEET_USERS = 'Users';
var SHEET_COMPANIES = 'Companies';
var SHEET_BATCHES = 'Add-back Cut-offs';
var SHEET_STATIONS = 'Stations';
var SHEET_MANPOWER = 'Approved Manpower';
var SHEET_WEEKS = 'Attendance Weeks';
var ATT_PREFIX = 'ATT ';
var SHEET_ACTIVE = 'Active Manpower Lists';
var ACT_PREFIX = 'ACT ';
var SHEET_CHECKS = 'Attendance Checks';
var CHK_PREFIX = 'CHK ';
var SHEET_UPX = 'Uploadtrix Cut-offs';
var UPX_PREFIX = 'UPX ';
var UPX_HEADERS = ['Cut-off ID', 'Tab', 'Start', 'End', 'Stations', 'Done', 'Manual', 'Created By', 'Created At'];
var UPX_ROW_HEADERS = ['No.', 'Station', 'Company', 'Status', 'Remarks', 'Updated By', 'Updated At'];
var CHECK_HEADERS = ['Check ID', 'Tab', 'Period', 'Start', 'End', 'Rows', 'Stations', 'Imported By', 'Imported At'];
var SESSION_HOURS = 6;

var BATCH_HEADERS = ['ID', 'Data Tab', 'Cut-off Start', 'Cut-off End', 'Reflect Start', 'Reflect End',
  'Report Date', 'Last Cut-off Label', 'Last Cut-off Count', 'Prepared By', 'Position',
  'Notice Text', 'Challenges (JSON)', 'Actions (JSON)', 'Updated'];

// Code (station suffix) | SC Code (HR add-back file) | Old Company | New Company
var COMPANY_HEADERS = ['Code', 'SC Code', 'Old Company', 'New Company'];
var DEFAULT_COMPANIES = [
  ['FGI', 'SC01', 'Fortune Global Inc.', 'Fortune Global Inc.'],
  ['RBI', 'SC02', 'Rephil Basak Inc.', 'Fast and Faster Co., Inc.'],
  ['RAI', 'SC03', 'Rephil Antipolo Inc.', 'Quickcharge Express Co., Inc.'],
  ['RMI', 'SC04', 'Rephil Marikina Inc.', 'Swiftie United Co., Inc.'],
  ['ASF', 'SC05', 'Rephil AS Fortuna Inc.', 'Accelerate Filling Co., Inc.'],
  ['RGI', 'SC06', 'Rephil Guerilla Inc.', 'Turbo Charge Co., Inc.'],
  ['RASI', 'SC07', 'Rephil Abad Santos Inc.', 'JustDrive Station Co., Inc.'],
  ['REPI', 'SC08', 'Rephil Edsa Pasay Inc.', 'Too Fast Co., Inc.'],
  ['FFC', 'SC09', 'Flash Fill Co., Inc.', 'Flash Fill Co., Inc.'],
  ['FFFCI', 'SC10', 'Fast Forward Fuels Co., Inc.', 'Fast Forward Fuels Co., Inc.']
];
var STATION_HEADERS = ['No.', 'Station', 'Code', 'Area / Branch', 'HDI Code', 'Group', 'Area Supervisor', 'Active'];
var GROUPS = ['NORTH', 'SOUTH', 'VISMIN'];
// Group from the Area / Branch: Visayas–Mindanao, North Luzon, or South (NCR, Rizal, Cavite, South Luzon)
function groupOf_(area) {
  var a = String(area || '').toLowerCase();
  if (/cebu|visayas|mindanao|vismin|iloilo|bacolod|davao|cagayan de oro|leyte|bohol|negros|samar|panay|zamboanga/.test(a)) return 'VISMIN';
  if (/bulacan|pampanga|central luzon|tarlac|nueva|pangasinan|bataan|zambales|north|ilocos|isabela|baguio|la union/.test(a)) return 'NORTH';
  return 'SOUTH';
}
function groupFix_(g, area) { g = String(g || '').trim().toUpperCase(); return GROUPS.indexOf(g) >= 0 ? g : groupOf_(area); }
var MANPOWER_HEADERS = ['Station', 'CA (Required)', 'PA (Required)', 'REL (Required)', 'CA (Approved)', 'PA (Approved)',
  'Additional Hrs (Non-24)', 'Remarks', 'Updated', 'Required Total', 'Approved Total', 'Approved Hrs / Day', 'Approved Weekly Hrs', 'Supervisor'];
var WEEK_HEADERS = ['Week ID', 'Tab', 'Start', 'End', 'Holidays', 'Rows', 'Imported By', 'Imported At'];
var ACTIVE_HEADERS = ['List ID', 'Tab', 'As of', 'Rows', 'CAPAREL', 'SS & CC', 'Imported By', 'Imported At'];

/* ===================== Web app entry ===================== */
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Payroll Support · Stations Department')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ===================== Sheet menu & setup ===================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu(APP_NAME)
    .addItem('Set up / repair all tabs', 'setup')
    .addToUi();
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

var USER_HEADERS = ['Username', 'Password', 'Display Name', 'Position', 'Active'];

function setup() {
  setupSheets_();
  try {
    SpreadsheetApp.getUi().alert('Done. Add or edit logins in the "Users" tab: type a Username and Password on a new row. Default login: admin / ChangeMe123');
  } catch (e) { /* run from editor */ }
}

function setupSheets_() {
  var ss = ss_();
  var u = ss.getSheetByName(SHEET_USERS) || ss.insertSheet(SHEET_USERS);
  usersSheet_(u);
  var c = ss.getSheetByName(SHEET_COMPANIES) || ss.insertSheet(SHEET_COMPANIES);
  if (c.getLastRow() === 0 || String(c.getRange(1, 2).getValue()) !== 'SC Code') {
    c.clear();
    c.getRange(1, 1, 1, COMPANY_HEADERS.length).setValues([COMPANY_HEADERS]);
    c.getRange(2, 1, DEFAULT_COMPANIES.length, 4).setNumberFormat('@').setValues(DEFAULT_COMPANIES);
    head_(c, COMPANY_HEADERS.length);
  }
  var st = ss.getSheetByName(SHEET_STATIONS) || ss.insertSheet(SHEET_STATIONS);
  if (st.getLastRow() === 0) {
    st.getRange(1, 1, 1, STATION_HEADERS.length).setValues([STATION_HEADERS]);
    var srows = SEED.stations.map(function (r) { return [r[0], r[1], codeOf_(r[1]), r[3] || '', r[4] || '', groupOf_(r[3]), '', 'TRUE']; });
    st.getRange(2, 1, srows.length, STATION_HEADERS.length).setNumberFormat('@').setValues(srows.map(function (r) { return r.map(String); }));
    head_(st, STATION_HEADERS.length);
  }
  var mp = ss.getSheetByName(SHEET_MANPOWER) || ss.insertSheet(SHEET_MANPOWER);
  if (mp.getLastRow() === 0) {
    mp.getRange(1, 1, 1, MANPOWER_HEADERS.length).setValues([MANPOWER_HEADERS]);
    var mrows = SEED.manpower.map(function (r, i) { return manpowerRow_(i + 2, { station: r[0], caR: r[1], paR: r[2], relR: r[3], caA: r[4], paA: r[5], addl: r[6], remarks: 'As of Nov 5, 2026', updated: MP_SEED_VERSION }); });
    mp.getRange(2, 1, mrows.length, MANPOWER_HEADERS.length).setValues(mrows);
    head_(mp, MANPOWER_HEADERS.length);
  } else if (String(mp.getRange(1, 14).getValue()) !== 'Supervisor') { // v8 sheet: add the Supervisor column
    mp.getRange(1, 14).setValue('Supervisor');
    head_(mp, MANPOWER_HEADERS.length);
  }
  var ux = ss.getSheetByName(SHEET_UPX) || ss.insertSheet(SHEET_UPX);
  if (ux.getLastRow() === 0) { ux.getRange(1, 1, 1, UPX_HEADERS.length).setValues([UPX_HEADERS]); head_(ux, UPX_HEADERS.length); }
  var ck = ss.getSheetByName(SHEET_CHECKS) || ss.insertSheet(SHEET_CHECKS);
  if (ck.getLastRow() === 0) { ck.getRange(1, 1, 1, CHECK_HEADERS.length).setValues([CHECK_HEADERS]); head_(ck, CHECK_HEADERS.length); }
  var ac = ss.getSheetByName(SHEET_ACTIVE) || ss.insertSheet(SHEET_ACTIVE);
  if (ac.getLastRow() === 0) { ac.getRange(1, 1, 1, ACTIVE_HEADERS.length).setValues([ACTIVE_HEADERS]); head_(ac, ACTIVE_HEADERS.length); }
  var w = ss.getSheetByName(SHEET_WEEKS) || ss.insertSheet(SHEET_WEEKS);
  if (w.getLastRow() === 0) { w.getRange(1, 1, 1, WEEK_HEADERS.length).setValues([WEEK_HEADERS]); head_(w, WEEK_HEADERS.length); }
  var b = ss.getSheetByName(SHEET_BATCHES) || ss.getSheetByName('Cut-offs');
  if (b && b.getName() === 'Cut-offs') b.setName(SHEET_BATCHES); // older version's name
  if (!b) b = ss.insertSheet(SHEET_BATCHES);
  if (b.getLastRow() === 0) { b.getRange(1, 1, 1, BATCH_HEADERS.length).setValues([BATCH_HEADERS]); head_(b, BATCH_HEADERS.length); }
  var s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

function head_(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#0F1F5B').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
}
function codeOf_(station) { var m = String(station || '').match(/-\s*([A-Za-z]+)\s*$/); return m ? m[1].toUpperCase() : ''; }
function manpowerRow_(r, m) {
  return [m.station, +m.caR || 0, +m.paR || 0, +m.relR || 0, +m.caA || 0, +m.paA || 0, +m.addl || 0, m.remarks || '', m.updated || '',
    '=SUM(B' + r + ':D' + r + ')', '=SUM(E' + r + ':F' + r + ')', '=K' + r + '*8+G' + r, '=L' + r + '*7', m.supervisor || ''];
}

/**
 * Makes sure the Users tab has the simple layout:
 * Username | Password | Display Name | Position | Active
 * An older layout (Password Hash / Salt) is converted; old passwords can't be read back,
 * so admin gets ChangeMe123 and other users get a blank password to fill in.
 */
function usersSheet_(u) {
  u = u || ss_().getSheetByName(SHEET_USERS) || ss_().insertSheet(SHEET_USERS);
  var lr = u.getLastRow(), lc = Math.max(u.getLastColumn(), 1);
  var head = lr ? u.getRange(1, 1, 1, lc).getValues()[0].map(function (h) { return String(h).trim(); }) : [];
  if (head[0] === 'Username' && head[1] === 'Password') return u;
  var rows = [];
  if (lr > 1) {
    var vals = u.getRange(2, 1, lr - 1, lc).getValues();
    var col = function (name) { return head.indexOf(name); };
    vals.forEach(function (r) {
      var name = String(r[0] || '').trim();
      if (!name) return;
      var dn = col('Display Name') >= 0 ? r[col('Display Name')] : '';
      var pos = col('Position') >= 0 ? r[col('Position')] : '';
      var act = col('Active') >= 0 ? r[col('Active')] : true;
      rows.push([name, name.toLowerCase() === 'admin' ? 'ChangeMe123' : '', dn, pos || 'Payroll Support Assistant', act === '' ? true : act]);
    });
  }
  if (!rows.length) rows.push(['admin', 'ChangeMe123', 'Den Dwight Laurence Oxina', 'Payroll Support Assistant', true]);
  u.clear();
  if (lc > 1) u.showColumns(1, Math.max(lc, 5));
  u.getRange(1, 1, 1, 5).setValues([USER_HEADERS]).setFontWeight('bold').setBackground('#0D0788').setFontColor('#FFFFFF');
  u.getRange(2, 1, Math.max(rows.length, 50), 2).setNumberFormat('@'); // keep passwords like 001234 as typed
  u.getRange(2, 1, rows.length, 5).setValues(rows);
  u.setFrozenRows(1);
  u.setColumnWidths(1, 2, 160); u.setColumnWidth(3, 240); u.setColumnWidth(4, 200); u.setColumnWidth(5, 80);
  u.getRange(2, 5, rows.length, 1).insertCheckboxes()
    .setValues(rows.map(function (r) { return [r[4] !== false && String(r[4]).toUpperCase() !== 'FALSE']; }));
  u.getRange(1, 7).setValue('To add a login, type a Username and Password on a new row. After the first sign-in the password is stored scrambled (sha256:…); to reset it, type a new password over it. To block someone, type FALSE (or untick) in Active.')
    .setFontColor('#595959').setFontStyle('italic');
  return u;
}

/* ===================== Auth ===================== */
function users_() {
  var sh = usersSheet_();
  var lr = sh.getLastRow();
  if (lr < 2) return { sh: sh, rows: [] };
  return { sh: sh, rows: sh.getRange(2, 1, lr - 1, 5).getDisplayValues() };
}

function apiLogin(username, password) {
  username = String(username || '').trim().toLowerCase();
  password = String(password || '');
  var cache = CacheService.getScriptCache();
  var failKey = 'fail_' + username, fails = +(cache.get(failKey) || 0);
  if (fails >= 5) throw new Error('Too many attempts. Try again in 10 minutes.');
  setupSheets_();
  var U = users_();
  for (var i = 0; i < U.rows.length; i++) {
    var r = U.rows[i];
    if (String(r[0]).trim().toLowerCase() !== username) continue;
    if (String(r[4]).toUpperCase() === 'FALSE') throw new Error('This account is switched off (Active is unticked in the Users tab).');
    if (!String(r[1])) throw new Error('No password set for this username yet. Type one in the Users tab.');
    var stored = String(r[1]).trim(), ok = stored.indexOf('sha256:') === 0 ? stored === hash_(username, password.trim()) : stored === password.trim();
    if (ok) {
      // passwords typed in the Users tab are stored scrambled after the first sign-in
      if (stored.indexOf('sha256:') !== 0) { try { U.sh.getRange(i + 2, 2).setNumberFormat('@').setValue(hash_(username, password.trim())); } catch (e) {} }
      logSignIn_(username, true);
      var token = Utilities.getUuid();
      var user = { username: username, name: r[2] || r[0], position: r[3] || 'Payroll Support Assistant' };
      cache.put('tok_' + token, JSON.stringify(user), SESSION_HOURS * 3600);
      cache.remove(failKey);
      return { token: token, user: user };
    }
  }
  cache.put(failKey, String(fails + 1), 600);
  logSignIn_(username, false);
  throw new Error('Wrong username or password. Check the Users tab in the Google Sheet.');
}

function apiLogout(token) { CacheService.getScriptCache().remove('tok_' + token); return true; }

function auth_(token) {
  var raw = token && CacheService.getScriptCache().get('tok_' + token);
  if (!raw) throw new Error('SESSION_EXPIRED');
  return JSON.parse(raw);
}

function apiChangePassword(token, oldPw, newPw) {
  var user = auth_(token);
  if (!newPw) throw new Error('Type a new password.');
  var U = users_();
  for (var i = 0; i < U.rows.length; i++) {
    if (String(U.rows[i][0]).trim().toLowerCase() === user.username) {
      var st = String(U.rows[i][1]).trim(), o = String(oldPw || '').trim();
      if (st.indexOf('sha256:') === 0 ? st !== hash_(user.username, o) : st !== o) throw new Error('Current password is wrong.');
      U.sh.getRange(i + 2, 2).setNumberFormat('@').setValue(hash_(user.username, String(newPw).trim()));
      return true;
    }
  }
  throw new Error('User not found.');
}

/* ===================== Helpers ===================== */
var SHEET_SETTINGS = 'Settings', SHEET_HOLIDAYS = 'Holidays', SHEET_ENDORSE = 'Endorsements', SHEET_LOG = 'Sign-in Log';
var SYSTEM_SHEETS = [SHEET_USERS, SHEET_COMPANIES, SHEET_BATCHES, SHEET_STATIONS, SHEET_MANPOWER, SHEET_WEEKS, SHEET_ACTIVE, SHEET_CHECKS, SHEET_UPX, 'Cut-offs', SHEET_SETTINGS, SHEET_HOLIDAYS, SHEET_ENDORSE, SHEET_LOG];

/* ===================== v21: settings, holidays, endorsements, sign-in log ===================== */
var SETTINGS_DEFAULT = [
  ['paidComplete', 10, 'Rest day / holiday: a paid duty is complete when the logs are this many hours apart or less'],
  ['breakHrs', 1, 'Break taken off a long rest day / holiday duty (hours)'],
  ['maxFile', 16, 'Most hours that can be filed for a rest day or holiday duty'],
  ['lateWrong', 2, 'Late (hours) that means a wrong schedule'],
  ['otMin', 1, 'Hours outside the shift that count as OT to confirm']
];
function sheetOf_(name, head) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, head.length).setValues([head]); head_(sh, head.length); sh.setFrozenRows(1); }
  return sh;
}
function settings_() {
  var sh = sheetOf_(SHEET_SETTINGS, ['Key', 'Value', 'What it means']), out = {};
  if (sh.getLastRow() < 2) sh.getRange(2, 1, SETTINGS_DEFAULT.length, 3).setValues(SETTINGS_DEFAULT);
  SETTINGS_DEFAULT.forEach(function (d) { out[d[0]] = d[1]; });
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { var k = String(r[0]).trim(); if (k && r[1] !== '' && !isNaN(+r[1])) out[k] = +r[1]; });
  return out;
}
function apiSaveSettings(token, cfg) {
  var user = auth_(token);
  if (user.username !== 'admin' && user.role !== 'admin') { var first = users_().rows[0]; if (!first || String(first[0]).toLowerCase() !== user.username) throw new Error('Only the first user in the Users tab can change Settings.'); }
  var sh = sheetOf_(SHEET_SETTINGS, ['Key', 'Value', 'What it means']);
  var rows = SETTINGS_DEFAULT.map(function (d) { var v = cfg && cfg[d[0]] !== undefined ? +cfg[d[0]] : d[1]; if (isNaN(v) || v < 0 || v > 24) throw new Error('“' + d[2] + '” must be a number of hours from 0 to 24.'); return [d[0], v, d[2]]; });
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, rows.length), 3).clearContent();
  sh.getRange(2, 1, rows.length, 3).setValues(rows);
  return settings_();
}
function holidays_() {
  var sh = sheetOf_(SHEET_HOLIDAYS, ['Date', 'Holiday', 'Type']);
  if (sh.getLastRow() < 2) {
    var seed = [['2026-11-01', 'All Saints\u2019 Day', 'Special'], ['2026-11-02', 'All Souls\u2019 Day', 'Special'], ['2026-11-30', 'Bonifacio Day', 'Regular'], ['2026-12-08', 'Immaculate Conception', 'Special'],
      ['2026-12-24', 'Christmas Eve', 'Special'], ['2026-12-25', 'Christmas Day', 'Regular'], ['2026-12-30', 'Rizal Day', 'Regular'], ['2026-12-31', 'Last Day of the Year', 'Special']];
    sh.getRange(2, 1, seed.length, 1).setNumberFormat('@'); sh.getRange(2, 1, seed.length, 3).setValues(seed);
  }
  return sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).getDisplayValues().filter(function (r) { return /^\d{4}-\d{2}-\d{2}$/.test(String(r[0]).trim()); })
    .map(function (r) { return { date: String(r[0]).trim(), name: String(r[1]).trim(), type: String(r[2]).trim() || 'Special' }; }).sort(function (a, b) { return a.date < b.date ? -1 : 1; });
}
function apiSaveHolidays(token, list) {
  auth_(token);
  var sh = sheetOf_(SHEET_HOLIDAYS, ['Date', 'Holiday', 'Type']);
  var rows = (list || []).filter(function (h) { return /^\d{4}-\d{2}-\d{2}$/.test(h.date) && h.name; }).map(function (h) { return [h.date, String(h.name).slice(0, 80), h.type === 'Regular' ? 'Regular' : 'Special']; });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  if (rows.length) { sh.getRange(2, 1, rows.length, 1).setNumberFormat('@'); sh.getRange(2, 1, rows.length, 3).setValues(rows); }
  else sh.getRange(2, 1, 1, 3).setValues([['', '(no holiday)', '']]);
  return holidays_();
}
// Endorsed for Add-back: one row per concern, the whole entry kept as JSON
var ENDORSE_HEAD = ['ID', 'Cut-off Start', 'Status', 'Name', 'Station', 'Dates', 'Description', 'Updated By', 'Updated At', 'Data (JSON)'];
function endorse_() {
  var sh = sheetOf_(SHEET_ENDORSE, ENDORSE_HEAD);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ENDORSE_HEAD.length).getValues().map(function (r) { try { return JSON.parse(r[9]); } catch (e) { return null; } }).filter(Boolean);
}
function apiEndorseSave(token, list) {
  var user = auth_(token), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheetOf_(SHEET_ENDORSE, ENDORSE_HEAD), now = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm');
    var rows = (list || []).map(function (e) { var j = JSON.stringify(e); if (j.length > 45000) throw new Error('One concern is too long to save: ' + e.name); return [e.id, e.cutStart || '', e.status || '', e.name || '', e.station || '', e.dates || '', e.desc || '', e.validatedBy || e.by || user.name, now, j]; });
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, ENDORSE_HEAD.length).clearContent();
    if (rows.length) { sh.getRange(2, 1, rows.length, ENDORSE_HEAD.length).setNumberFormat('@'); sh.getRange(2, 1, rows.length, ENDORSE_HEAD.length).setValues(rows); }
    return { saved: rows.length, at: now };
  } finally { lock.releaseLock(); }
}
function hash_(username, password) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'ps21:' + String(username).toLowerCase() + ':' + password, Utilities.Charset.UTF_8);
  return 'sha256:' + b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}
function logSignIn_(username, ok) {
  try { var sh = sheetOf_(SHEET_LOG, ['When', 'Username', 'Result']); sh.appendRow([Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'), username, ok ? 'Signed in' : 'Wrong password']); } catch (e) {}
}

function tz_() { return ss_().getSpreadsheetTimeZone() || 'Asia/Manila'; }

function ymd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return v ? String(v) : '';
}

// Dates can't cross google.script.run, and "Oct 4" pasted in Sheets becomes a date → show it as text again.
function clean_(values) {
  var tz = tz_();
  return values.map(function (r) {
    return r.map(function (v) { return v instanceof Date ? Utilities.formatDate(v, tz, 'MMM d') : v; });
  });
}

function companyList_() {
  var sh = ss_().getSheetByName(SHEET_COMPANIES);
  var rows = (!sh || sh.getLastRow() < 2) ? DEFAULT_COMPANIES : sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues();
  return rows.filter(function (r) { return r[0]; }).map(function (r, i) {
    return { row: i + 2, code: String(r[0]).trim().toUpperCase(), sc: String(r[1]).trim(), oldCompany: String(r[2]).trim(), newCompany: String(r[3]).trim() };
  });
}
// shape used by the Add-back engine: {code: SC01, name: new company, abbr: FGI}
function companies_() {
  return companyList_().filter(function (c) { return c.sc; }).map(function (c) { return { code: c.sc, name: c.newCompany || c.oldCompany, abbr: c.code }; })
    .sort(function (a, b) { return a.code < b.code ? -1 : 1; });
}

// Google Sheets may turn a tab name like "Oct 29" into a date. Map it back to the real tab.
function resolveTab_(v, names) {
  if (!(v instanceof Date)) {
    var str = String(v || '').trim();
    if (!str || names.indexOf(str) >= 0) return str;
    // text like "Thu Oct 29 2026 15:00:00 GMT+0800 (China Standard Time)" saved by an older version
    var m = str.match(/^[A-Za-z]{3}\s+([A-Za-z]{3})\s+(\d{1,2})\s+\d{4}\b/);
    if (m) {
      var cands = [m[1] + ' ' + (+m[2]), m[1] + ' ' + (m[2].length < 2 ? '0' + m[2] : m[2])];
      for (var c = 0; c < cands.length; c++) if (names.indexOf(cands[c]) >= 0) return cands[c];
      var parsed = new Date(str);
      if (!isNaN(parsed)) v = parsed; else return str;
    } else return str;
  }
  var tzs = [tz_(), Session.getScriptTimeZone()];
  var fmts = ['MMM d', 'MMM dd', 'MMMM d', 'MMMM dd', 'M/d', 'M/d/yyyy', 'MM/dd/yyyy', 'yyyy-MM-dd', 'MMM d, yyyy', 'd MMM'];
  for (var t = 0; t < tzs.length; t++) for (var f = 0; f < fmts.length; f++) {
    var n = Utilities.formatDate(v, tzs[t], fmts[f]);
    if (names.indexOf(n) >= 0) return n;
  }
  return Utilities.formatDate(v, tz_(), 'MMM d');
}

// Payroll weeks run Sunday–Saturday: snap any stored start date to its Sunday (fixes time-zone shifts).
function weekStart_(iso) {
  if (!iso) return '';
  var p = String(iso).split('-'); if (p.length < 3) return String(iso);
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
  var dow = d.getUTCDay();
  if (dow !== 0) d.setUTCDate(d.getUTCDate() + (dow === 6 ? 1 : -dow));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}
function plus6_(iso) {
  if (!iso) return '';
  var p = iso.split('-'); var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + 6));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function batches_() {
  var sh = ss_().getSheetByName(SHEET_BATCHES);
  if (!sh || sh.getLastRow() < 2) return [];
  var names = ss_().getSheets().map(function (x) { return x.getName(); });
  return sh.getRange(2, 1, sh.getLastRow() - 1, BATCH_HEADERS.length).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r, i) {
      var st = weekStart_(ymd_(r[2])), rs = weekStart_(ymd_(r[4]));
      return {
        row: i + 2, id: String(r[0]), tab: resolveTab_(r[1], names), tabIsDate: r[1] instanceof Date,
        start: st, end: plus6_(st), reflectStart: rs, reflectEnd: plus6_(rs),
        reportDate: ymd_(r[6]), lastLabel: String(r[7] || ''), lastCount: r[8] === '' ? '' : +r[8],
        preparedBy: String(r[9] || ''), position: String(r[10] || ''), notice: String(r[11] || ''),
        challenges: String(r[12] || ''), actions: String(r[13] || ''), updated: ymd_(r[14])
      };
    })
    .sort(function (a, b) { return a.start < b.start ? 1 : -1; });
}

function headerRowIndex_(values) {
  for (var i = 0; i < Math.min(values.length, 15); i++) {
    var line = values[i].map(function (x) { return String(x || '').toLowerCase(); }).join('|');
    if (/name/.test(line) && /(amount|add-?back)/.test(line)) return i;
  }
  return -1;
}

function tabValues_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Tab "' + name + '" not found.');
  var lr = sh.getLastRow(), lc = Math.min(sh.getLastColumn(), 20);
  if (lr < 1) return [];
  return clean_(sh.getRange(1, 1, lr, lc).getValues());
}

/* ===================== API ===================== */
function apiBootstrap(token) {
  var user = auth_(token);
  var notice = '';
  try { notice = upgradeManpower_(); } catch (e) { notice = ''; }
  try { upgradeGroups_(); } catch (e) {}
  var tabs = ss_().getSheets()
    .filter(function (s) { return SYSTEM_SHEETS.indexOf(s.getName()) < 0 && s.getName().indexOf(ATT_PREFIX) !== 0 && s.getName().indexOf(ACT_PREFIX) !== 0 && s.getName().indexOf(CHK_PREFIX) !== 0 && s.getName().indexOf(UPX_PREFIX) !== 0; })
    .map(function (s) {
      var lr = s.getLastRow(), head = lr ? s.getRange(1, 1, Math.min(lr, 15), Math.min(Math.max(s.getLastColumn(), 1), 20)).getValues() : [];
      var hi = headerRowIndex_(head);
      return { name: s.getName(), rows: hi >= 0 ? Math.max(0, lr - hi - 1) : 0, valid: hi >= 0 };
    });
  var bs = batches_(), bsh = ss_().getSheetByName(SHEET_BATCHES);
  if (bsh) {
    var raw = bsh.getRange(2, 2, Math.max(1, bsh.getLastRow() - 1), 1).getValues();
    // repair tab names that were turned into dates / date text
    bs.forEach(function (b) {
      if (String(raw[b.row - 2][0]) !== b.tab) bsh.getRange(b.row, 2).setNumberFormat('@').setValue(b.tab);
    });
    // one row per tab: keep the newest row for each tab, remove older duplicates
    var keep = {}, drop = [];
    bs.slice().sort(function (a, b) { return b.row - a.row; }).forEach(function (b) {
      if (keep[b.tab]) drop.push(b.row); else keep[b.tab] = 1;
    });
    drop.sort(function (a, b) { return b - a; }).forEach(function (r) { bsh.deleteRow(r); });
    if (drop.length) bs = batches_();
  }
  return { user: user, batches: bs, tabs: tabs, companies: companies_(), companyList: companyList_(),
    stations: stations_(), manpower: manpower_(), weeks: weeks_(), actives: actives_(), checks: checks_(), upx: upxList_(), settings: settings_(), holidays: holidays_(), endorse: endorse_(), sheetUrl: ss_().getUrl(), version: APP_VERSION, notice: notice };
}

function apiGetBatch(token, id) {
  auth_(token);
  var all = batches_();
  var idx = -1;
  for (var i = 0; i < all.length; i++) if (all[i].id === id) idx = i;
  if (idx < 0) throw new Error('Cut-off not found.');
  var b = all[idx], prev = null; // previous = latest cut-off that starts before this one
  for (var j = 0; j < all.length; j++) { if (all[j].start < b.start && (!prev || all[j].start > prev.start)) prev = all[j]; }
  var out = { batch: b, values: tabValues_(b.tab), prev: null };
  if (prev) {
    try { out.prev = { batch: prev, values: tabValues_(prev.tab) }; } catch (e) { out.prev = null; }
  }
  return out;
}

// every add-back data tab (used to detect concerns that were already added back)
function apiAddbackAll(token) {
  auth_(token);
  return batches_().map(function (b) { var v = []; try { v = tabValues_(b.tab); } catch (e) {} return { batch: b, values: v }; });
}

function apiSaveBatch(token, b) {
  var user = auth_(token);
  var sh = ss_().getSheetByName(SHEET_BATCHES);
  if (!sh) { setupSheets_(); sh = ss_().getSheetByName(SHEET_BATCHES); }
  if (!b.tab || !ss_().getSheetByName(b.tab)) throw new Error('Pick the tab that holds this cut-off\'s data.');
  if (!b.start || !b.end) throw new Error('Cut-off start and end dates are required.');
  // everything is stored as plain text so Google Sheets can't turn "Oct 29" or a date into something else
  function d(s) { return s ? String(s).slice(0, 10) : ''; }
  var row = [b.id || ('c' + Utilities.getUuid().slice(0, 7)), b.tab, d(b.start), d(b.end), d(b.reflectStart), d(b.reflectEnd),
    d(b.reportDate), b.lastLabel || '', b.lastCount === '' || b.lastCount === null || b.lastCount === undefined ? '' : String(b.lastCount),
    b.preparedBy || user.name, b.position || user.position, b.notice || '', b.challenges || '', b.actions || '',
    Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm')];
  var list = batches_(), target = null;
  list.forEach(function (x) { if (b.id ? x.id === b.id : x.tab === b.tab) target = x; });
  if (target && !b.id) row[0] = target.id;
  var at = target ? target.row : sh.getLastRow() + 1;
  sh.getRange(at, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  return row[0];
}

function apiDeleteBatch(token, id, deleteTab) {
  auth_(token);
  var ss = ss_(), sh = ss.getSheetByName(SHEET_BATCHES);
  var t = batches_().filter(function (x) { return x.id === id; })[0];
  if (t) {
    sh.deleteRow(t.row);
    if (deleteTab) { var tab = ss.getSheetByName(t.tab); if (tab && ss.getSheets().length > 1) ss.deleteSheet(tab); }
  }
  return true;
}

/**
 * Paste import: values = 2D array (header row optional). mode = 'new' | 'replace' | 'append'
 */
function apiImportPaste(token, tabName, values, mode, meta) {
  var user = auth_(token);
  tabName = String(tabName || '').trim();
  if (!tabName) throw new Error('Tab name is required.');
  if (SYSTEM_SHEETS.indexOf(tabName) >= 0) throw new Error('That name is reserved.');
  if (!values || !values.length) throw new Error('Nothing to import.');
  var HEAD = ['No.', 'Code', 'Company', 'Name', 'Station', 'Date(s) of Concern', 'Description of Concern',
    'Payroll Cut-Off', 'HR Remarks/Type of Error', 'Amount for Add-Back'];
  var hasHeader = headerRowIndex_(values.slice(0, 1)) === 0;
  var ss = ss_(), sh = ss.getSheetByName(tabName);
  if (sh && mode === 'new') throw new Error('Tab "' + tabName + '" already exists. Choose Replace or Append.');
  if (!sh) sh = ss.insertSheet(tabName);
  if (mode === 'replace') sh.clear();
  var width = Math.max.apply(null, values.map(function (r) { return r.length; }));
  width = Math.max(width, HEAD.length);
  function padRow(r) { r = r.slice(); while (r.length < width) r.push(''); return r; }
  var rows = values.map(padRow);
  if (sh.getLastRow() === 0) {
    var head = hasHeader ? rows.shift() : padRow(HEAD);
    sh.getRange(1, 1, 1, width).setValues([head]).setFontWeight('bold').setBackground('#0D0788').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  } else if (hasHeader) {
    rows.shift();
  }
  if (!rows.length) return { tab: tabName, added: 0, id: meta ? saveMeta_(token, tabName, meta) : '' };
  var start = sh.getLastRow() + 1;
  var range = sh.getRange(start, 1, rows.length, width);
  sh.getRange(start, 1, rows.length, 9).setNumberFormat('@'); // keep "Oct 4", "Oct04 Oct10" as text
  sh.getRange(start, 10, rows.length, 1).setNumberFormat('#,##0.00');
  rows = rows.map(function (r) {
    var a = String(r[9] || '').replace(/[₱,\s]/g, '');
    if (a !== '' && !isNaN(+a)) r[9] = +a; else if (/^=/.test(a)) r[9] = a;
    return r;
  });
  range.setValues(rows);
  SpreadsheetApp.flush();
  return { tab: tabName, added: rows.length, id: meta ? saveMeta_(token, tabName, meta) : '' };
}

// registers (or updates) the cut-off settings for a data tab
function saveMeta_(token, tabName, meta) {
  var existing = batches_().filter(function (x) { return x.tab === tabName; })[0] || {};
  var b = {};
  Object.keys(existing).forEach(function (k) { b[k] = existing[k]; });
  b.tab = tabName;
  ['start', 'end', 'reflectStart', 'reflectEnd'].forEach(function (k) { if (meta[k]) b[k] = meta[k]; });
  if (!b.reportDate) b.reportDate = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  return apiSaveBatch(token, b);
}


/* ===================== Check (run from the Apps Script editor) ===================== */
/**
 * Select "diagnose" in the function dropdown and click Run.
 * The Execution log lists every tab, the Cut-offs rows and whether each data tab can be read.
 */
function diagnose() {
  var out = [];
  try { setupSheets_(); out.push('✓ System tabs OK'); } catch (e) { out.push('✗ Setup error: ' + e.message); }
  ss_().getSheets().forEach(function (s) {
    var lr = s.getLastRow(), lc = Math.min(Math.max(s.getLastColumn(), 1), 10);
    var head = lr ? s.getRange(1, 1, 1, lc).getDisplayValues()[0] : [];
    out.push('Tab "' + s.getName() + '": ' + lr + ' rows · header: ' + head.join(' | '));
  });
  var bs = [];
  try { bs = batches_(); out.push('Cut-offs registered: ' + bs.length); } catch (e) { out.push('✗ Cut-offs read error: ' + e.message); }
  bs.forEach(function (b) {
    out.push(' • tab "' + b.tab + '" · concerns for ' + b.start + ' to ' + b.end + ' · reflected ' + (b.reflectStart || '(not set)'));
    try { var v = tabValues_(b.tab); JSON.stringify(v); out.push('   ✓ read ' + v.length + ' rows'); }
    catch (e) { out.push('   ✗ could not read: ' + e.message); }
  });
  try { var boot = { batches: batches_(), companies: companies_() }; JSON.stringify(boot); out.push('✓ Dashboard data OK'); }
  catch (e) { out.push('✗ Dashboard data error: ' + e.message); }
  Logger.log(out.join('\n'));
  return out.join('\n');
}


/* =====================================================================
   STATIONS · APPROVED MANPOWER · ATTENDANCE WEEKS
   ===================================================================== */
function stations_() {
  var sh = ss_().getSheetByName(SHEET_STATIONS);
  if (!sh || sh.getLastRow() < 2) return [];
  var comp = {}; companyList_().forEach(function (c) { comp[c.code] = c; });
  return sh.getRange(2, 1, sh.getLastRow() - 1, STATION_HEADERS.length).getDisplayValues()
    .map(function (r, i) {
      var code = String(r[2] || codeOf_(r[1])).toUpperCase(), c = comp[code] || {};
      return { row: i + 2, no: r[0] === '' ? '' : +r[0], station: String(r[1]).trim(), code: code, area: r[3], hdi: r[4], group: groupFix_(r[5], r[3]),
        supervisor: r[6], active: String(r[7]).toUpperCase() !== 'FALSE', oldCompany: c.oldCompany || '', newCompany: c.newCompany || '' };
    })
    .filter(function (s) { return s.station; });
}

function manpower_() {
  var sh = ss_().getSheetByName(SHEET_MANPOWER);
  if (!sh || sh.getLastRow() < 2) return [];
  var w = Math.min(Math.max(sh.getLastColumn(), 9), 14);
  return sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues()
    .map(function (r, i) {
      return { row: i + 2, station: String(r[0]).trim(), caR: +r[1] || 0, paR: +r[2] || 0, relR: +r[3] || 0, caA: +r[4] || 0, paA: +r[5] || 0,
        addl: +r[6] || 0, remarks: String(r[7] || ''), updated: r[8] instanceof Date ? Utilities.formatDate(r[8], tz_(), 'yyyy-MM-dd') : String(r[8] || ''),
        supervisor: String(r[13] === undefined ? '' : r[13]).trim() };
    })
    .filter(function (m) { return m.station; });
}

function weeks_() {
  var sh = ss_().getSheetByName(SHEET_WEEKS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, WEEK_HEADERS.length).getDisplayValues()
    .map(function (r, i) { return { row: i + 2, id: r[0], tab: r[1], start: r[2], end: r[3], holidays: r[4] ? r[4].split(/\s*,\s*/).filter(String) : [], rows: +r[5] || 0, by: r[6], at: r[7] }; })
    .filter(function (w) { return w.id; })
    .sort(function (a, b) { return a.start < b.start ? 1 : -1; });
}

function stationNo_(s) { var m = String(s || '').match(/^\s*(\d+)/); return m ? +m[1] : null; }

function apiSaveStation(token, s, originalName) {
  auth_(token);
  var name = String(s.station || '').trim();
  if (!name) throw new Error('Type the station name, for example "123-Sample - RGI".');
  var sh = ss_().getSheetByName(SHEET_STATIONS); if (!sh) { setupSheets_(); sh = ss_().getSheetByName(SHEET_STATIONS); }
  var list = stations_(), target = null;
  list.forEach(function (x) {
    if (originalName && x.station === originalName) target = x;
    else if (x.station.toLowerCase() === name.toLowerCase()) throw new Error('"' + name + '" is already in the list.');
    else if (s.no !== '' && s.no !== undefined && String(x.no) === String(s.no) && x.station !== originalName) throw new Error('Station number ' + s.no + ' is already used by "' + x.station + '".');
  });
  var no = s.no !== '' && s.no !== undefined ? s.no : stationNo_(name);
  var row = [no === null ? '' : String(no), name, String(s.code || codeOf_(name)).toUpperCase(), s.area || '', s.hdi || '', groupFix_(s.group, s.area), s.supervisor || '', s.active === false ? 'FALSE' : 'TRUE'];
  var at = target ? target.row : sh.getLastRow() + 1;
  sh.getRange(at, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  if (target && originalName && originalName !== name) { // keep Approved Manpower linked to the renamed station
    var mp = ss_().getSheetByName(SHEET_MANPOWER);
    manpower_().forEach(function (m) { if (m.station === originalName) mp.getRange(m.row, 1).setValue(name); });
  }
  return stations_();
}

function apiDeleteStation(token, name) {
  auth_(token);
  var sh = ss_().getSheetByName(SHEET_STATIONS);
  var t = stations_().filter(function (x) { return x.station === name; })[0];
  if (t) sh.deleteRow(t.row);
  return stations_();
}

function apiSaveCompany(token, c, originalCode) {
  auth_(token);
  var code = String(c.code || '').trim().toUpperCase();
  if (!code) throw new Error('Type the company code, for example "RGI".');
  var sh = ss_().getSheetByName(SHEET_COMPANIES);
  var list = companyList_(), target = null;
  list.forEach(function (x) { if (originalCode && x.code === originalCode) target = x; else if (x.code === code) throw new Error('Code "' + code + '" already exists.'); });
  var row = [code, String(c.sc || '').trim().toUpperCase(), String(c.oldCompany || '').trim(), String(c.newCompany || '').trim()];
  var at = target ? target.row : sh.getLastRow() + 1;
  sh.getRange(at, 1, 1, 4).setNumberFormat('@').setValues([row]);
  return { companyList: companyList_(), companies: companies_(), stations: stations_() };
}

function apiSaveManpower(token, m, originalName) {
  auth_(token);
  var name = String(m.station || '').trim();
  if (!name) throw new Error('Pick a station.');
  var sh = ss_().getSheetByName(SHEET_MANPOWER); if (!sh) { setupSheets_(); sh = ss_().getSheetByName(SHEET_MANPOWER); }
  var list = manpower_(), target = null;
  list.forEach(function (x) {
    if (originalName && x.station === originalName) target = x;
    else if (x.station === name) throw new Error('"' + name + '" already has approved manpower. Edit that row instead.');
  });
  var at = target ? target.row : sh.getLastRow() + 1;
  m.updated = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  sh.getRange(at, 1, 1, MANPOWER_HEADERS.length).setValues([manpowerRow_(at, m)]);
  return manpower_();
}

function apiDeleteManpower(token, name) {
  auth_(token);
  var sh = ss_().getSheetByName(SHEET_MANPOWER);
  var t = manpower_().filter(function (x) { return x.station === name; })[0];
  if (t) {
    sh.deleteRow(t.row);
    // re-write formulas below the deleted row so they point at the right row numbers
    var rest = manpower_().filter(function (x) { return x.row >= t.row; });
    rest.forEach(function (x) { sh.getRange(x.row, 10, 1, 4).setFormulas([manpowerRow_(x.row, x).slice(9, 13)]); });
  }
  return manpower_();
}

/* ---------- Attendance (chunked import so 10,000+ rows go through) ---------- */
function weekTab_(start) { return ATT_PREFIX + start; }

function apiAttendanceBegin(token, meta, header) {
  auth_(token);
  if (!meta || !meta.start) throw new Error('Could not detect the week from the data.');
  var ss = ss_(), name = weekTab_(meta.start), sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name, ss.getSheets().length); else sh.clear();
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  head_(sh, header.length);
  return { tab: name };
}

function apiWriteChunk(token, tab, rows, startRow) { return apiAttendanceChunk(token, tab, rows, startRow); }
function apiAttendanceChunk(token, tab, rows, startRow) {
  auth_(token);
  if (!rows || !rows.length) return 0;
  var sh = ss_().getSheetByName(tab);
  if (!sh) throw new Error('Tab "' + tab + '" not found.');
  var rg = sh.getRange(startRow, 1, rows.length, rows[0].length);
  rg.setNumberFormat('@');
  rg.setValues(rows.map(function (r) { return r.map(function (v) { return v === null || v === undefined ? '' : String(v); }); }));
  return rows.length;
}

function apiAttendanceFinish(token, meta, count) {
  var user = auth_(token);
  var sh = ss_().getSheetByName(SHEET_WEEKS); if (!sh) { setupSheets_(); sh = ss_().getSheetByName(SHEET_WEEKS); }
  var id = 'W' + meta.start.replace(/-/g, '');
  var row = [id, weekTab_(meta.start), meta.start, meta.end, (meta.holidays || []).join(', '), String(count), user.name, Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm')];
  var t = weeks_().filter(function (w) { return w.id === id; })[0];
  var at = t ? t.row : sh.getLastRow() + 1;
  sh.getRange(at, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  return weeks_();
}

function apiGetWeek(token, id) {
  auth_(token);
  var w = weeks_().filter(function (x) { return x.id === id; })[0];
  if (!w) throw new Error('Week not found.');
  var sh = ss_().getSheetByName(w.tab);
  if (!sh) throw new Error('Tab "' + w.tab + '" not found.');
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  return { week: w, values: lr ? sh.getRange(1, 1, lr, lc).getDisplayValues() : [] };
}

function apiSaveWeekHolidays(token, id, holidays) {
  auth_(token);
  var w = weeks_().filter(function (x) { return x.id === id; })[0];
  if (!w) throw new Error('Week not found.');
  ss_().getSheetByName(SHEET_WEEKS).getRange(w.row, 5).setNumberFormat('@').setValue((holidays || []).join(', '));
  return weeks_();
}

function apiDeleteWeek(token, id) {
  auth_(token);
  var w = weeks_().filter(function (x) { return x.id === id; })[0];
  if (!w) return weeks_();
  ss_().getSheetByName(SHEET_WEEKS).deleteRow(w.row);
  var tab = ss_().getSheetByName(w.tab); if (tab) ss_().deleteSheet(tab);
  return weeks_();
}

/* ---------- Starting data: LIST OF STATIONS + Approved Manpower (as of Sep 30, 2026) ---------- */

/* =====================================================================
   v9 · APPROVED MANPOWER IMPORT · ACTIVE MANPOWER · DATA TABS
   ===================================================================== */
function key_(name) { var n = stationNo_(name); return n !== null ? 'N' + n : 'S' + String(name || '').toLowerCase().replace(/\s+/g, ''); }

/**
 * items: [{station, no, company, area, hdi, depot, caR, paR, relR, caA, paA, addl, supervisor, remarks}]
 * mode: 'merge' (update matching stations, add new ones) | 'replace' (the file becomes the whole list)
 * Stations that are not yet in "Stations" are added there too.
 */
function apiImportManpower(token, items, mode, asOf) {
  auth_(token);
  return importManpower_(items, mode, asOf);
}
// one time per new built-in file: update Approved Manpower to the latest numbers (supervisors and extra stations are kept)
// v15: the Depot column of Stations becomes Group (NORTH / SOUTH / VISMIN), filled from the Area / Branch
function upgradeGroups_() {
  var sh = ss_().getSheetByName(SHEET_STATIONS); if (!sh) return;
  if (String(sh.getRange(1, 6).getValue()).trim().toLowerCase() === 'group') return;
  sh.getRange(1, 6).setValue('Group');
  var n = sh.getLastRow() - 1; if (n < 1) return;
  var areas = sh.getRange(2, 4, n, 1).getDisplayValues();
  sh.getRange(2, 6, n, 1).setNumberFormat('@').setValues(areas.map(function (a) { return [groupOf_(a[0])]; }));
}

function upgradeManpower_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('mpSeed') === MP_SEED_VERSION) return '';
  props.setProperty('mpSeed', MP_SEED_VERSION);
  var sh = ss_().getSheetByName(SHEET_MANPOWER);
  if (!sh || sh.getLastRow() < 2) return '';
  var cur = {}; manpower_().forEach(function (m) { cur[key_(m.station)] = m; });
  var items = SEED.manpower.map(function (r) { return { station: r[0], caR: r[1], paR: r[2], relR: r[3], caA: r[4], paA: r[5], addl: r[6] }; });
  var changed = items.filter(function (it) { var c = cur[key_(it.station)]; return !c || ['caR', 'paR', 'relR', 'caA', 'paA', 'addl'].some(function (f) { return +c[f] !== +it[f]; }); });
  if (!changed.length) return '';
  importManpower_(changed, 'merge', MP_SEED_VERSION);
  return 'Approved Manpower updated to the latest file: ' + changed.length + ' station(s) changed.';
}
function importManpower_(items, mode, asOf) {
  if (!items || !items.length) throw new Error('No stations found in the file.');
  var ss = ss_(), sh = ss.getSheetByName(SHEET_MANPOWER); if (!sh) { setupSheets_(); sh = ss.getSheetByName(SHEET_MANPOWER); }
  var today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  var cur = manpower_(), byKey = {}, order = [];
  cur.forEach(function (m) { byKey[key_(m.station)] = m; order.push(key_(m.station)); });
  var added = 0, updated = 0, seen = {};
  var finalList = mode === 'replace' ? [] : cur.map(function (m) { return m; });
  var idx = {}; finalList.forEach(function (m, i) { idx[key_(m.station)] = i; });
  items.forEach(function (it) {
    var k = key_(it.station); if (!it.station || seen[k]) return; seen[k] = 1;
    var old = byKey[k];
    var m = { station: old && mode !== 'replace' ? old.station : String(it.station).trim(), caR: it.caR, paR: it.paR, relR: it.relR, caA: it.caA, paA: it.paA, addl: it.addl,
      remarks: it.remarks || (asOf ? 'As of ' + Utilities.formatDate(new Date(asOf + 'T12:00:00Z'), 'UTC', 'MMM d, yyyy') : (old ? old.remarks : '')), updated: asOf || today,
      supervisor: it.supervisor || (old ? old.supervisor : '') };
    if (old) updated++; else added++;
    if (idx[k] !== undefined) finalList[idx[k]] = m; else { idx[k] = finalList.length; finalList.push(m); }
  });
  finalList.sort(function (a, b) { return (stationNo_(a.station) || 9999) - (stationNo_(b.station) || 9999); });
  var lr = sh.getLastRow(); if (lr > 1) sh.getRange(2, 1, lr - 1, MANPOWER_HEADERS.length).clearContent();
  var rows = finalList.map(function (m, i) { return manpowerRow_(i + 2, m); });
  if (rows.length) sh.getRange(2, 1, rows.length, MANPOWER_HEADERS.length).setValues(rows);
  // add stations that are missing from the Stations list
  var st = ss.getSheetByName(SHEET_STATIONS), have = {}, newSt = 0;
  stations_().forEach(function (x) { have[key_(x.station)] = 1; });
  var comp = {}; companyList_().forEach(function (c) { comp[c.oldCompany.toLowerCase()] = c.code; comp[c.newCompany.toLowerCase()] = c.code; });
  var srows = [];
  items.forEach(function (it) {
    var k = key_(it.station); if (!it.station || have[k]) return; have[k] = 1;
    var no = it.no !== '' && it.no !== undefined && it.no !== null ? it.no : stationNo_(it.station);
    srows.push([no === null ? '' : String(no), String(it.station).trim(), comp[String(it.company || '').toLowerCase()] || codeOf_(it.station), it.area || '', it.hdi || '', groupFix_(it.group, it.area), '', 'TRUE']);
  });
  if (srows.length) { st.getRange(st.getLastRow() + 1, 1, srows.length, 8).setNumberFormat('@').setValues(srows); newSt = srows.length; }
  return { manpower: manpower_(), stations: stations_(), added: added, updated: updated, newStations: newSt, total: rows.length };
}

/* ---------- Active Manpower (one tab per weekly list) ---------- */
function actives_() {
  var sh = ss_().getSheetByName(SHEET_ACTIVE);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ACTIVE_HEADERS.length).getDisplayValues()
    .map(function (r, i) { return { row: i + 2, id: r[0], tab: r[1], asOf: r[2], rows: +r[3] || 0, caparel: +r[4] || 0, sscc: +r[5] || 0, by: r[6], at: r[7] }; })
    .filter(function (a) { return a.id; })
    .sort(function (a, b) { return a.asOf < b.asOf ? 1 : -1; });
}

function apiActiveBegin(token, meta, header) {
  auth_(token);
  if (!meta || !/^\d{4}-\d{2}-\d{2}$/.test(meta.asOf || '')) throw new Error('Pick the "as of" date of this list.');
  var ss = ss_(); if (!ss.getSheetByName(SHEET_ACTIVE)) setupSheets_();
  var name = ACT_PREFIX + meta.asOf, sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name, ss.getSheets().length); else sh.clear();
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  head_(sh, header.length);
  return { tab: name };
}

// supervisors: {stationName: 'Supervisor name'} found in the list → fills blank Supervisor cells in Approved Manpower
function apiActiveFinish(token, meta, counts, supervisors, overwrite) {
  var user = auth_(token);
  var sh = ss_().getSheetByName(SHEET_ACTIVE);
  var id = 'A' + meta.asOf.replace(/-/g, '');
  var row = [id, ACT_PREFIX + meta.asOf, meta.asOf, String(counts.rows || 0), String(counts.caparel || 0), String(counts.sscc || 0), user.name, Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm')];
  var t = actives_().filter(function (a) { return a.id === id; })[0];
  var at = t ? t.row : sh.getLastRow() + 1;
  sh.getRange(at, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  var filled = 0;
  if (supervisors) {
    var by = {}; Object.keys(supervisors).forEach(function (n) { by[key_(n)] = supervisors[n]; });
    var mp = ss_().getSheetByName(SHEET_MANPOWER);
    manpower_().forEach(function (m) {
      var name = by[key_(m.station)];
      if (name && (overwrite || !m.supervisor) && name !== m.supervisor) { mp.getRange(m.row, 14).setValue(name); filled++; }
    });
  }
  return { actives: actives_(), manpower: manpower_(), filled: filled };
}

function apiGetActive(token, id) {
  auth_(token);
  var a = actives_().filter(function (x) { return x.id === id; })[0];
  if (!a) throw new Error('Active list not found.');
  var sh = ss_().getSheetByName(a.tab);
  if (!sh) throw new Error('Tab "' + a.tab + '" not found.');
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  return { list: a, values: lr ? sh.getRange(1, 1, lr, lc).getDisplayValues() : [] };
}

// supervisors: {stationName: 'Name'} → Supervisor column of Approved Manpower
function apiFillSupervisors(token, supervisors, overwrite) {
  auth_(token);
  var by = {}; Object.keys(supervisors || {}).forEach(function (n) { by[key_(n)] = supervisors[n]; });
  var mp = ss_().getSheetByName(SHEET_MANPOWER), filled = 0;
  manpower_().forEach(function (m) {
    var name = by[key_(m.station)];
    if (name && (overwrite || !m.supervisor) && name !== m.supervisor) { mp.getRange(m.row, 14).setValue(name); filled++; }
  });
  return { manpower: manpower_(), filled: filled };
}


/* ---------- Daily Attendance Checking (one tab per import: "CHK yyyy-mm-dd to yyyy-mm-dd") ---------- */
function checks_() {
  var sh = ss_().getSheetByName(SHEET_CHECKS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, CHECK_HEADERS.length).getDisplayValues()
    .map(function (r, i) { return { row: i + 2, id: r[0], tab: r[1], period: r[2], start: r[3], end: r[4], rows: +r[5] || 0, stations: +r[6] || 0, by: r[7], at: r[8] }; })
    .filter(function (c) { return c.id; })
    .sort(function (a, b) { return a.end === b.end ? (a.at < b.at ? 1 : -1) : (a.end < b.end ? 1 : -1); });
}
function checkTab_(meta) { return CHK_PREFIX + meta.start + (meta.end && meta.end !== meta.start ? ' to ' + meta.end : ''); }
function apiCheckBegin(token, meta, header) {
  auth_(token);
  if (!meta || !/^\d{4}-\d{2}-\d{2}$/.test(meta.start || '')) throw new Error('Set the period of this checking.');
  var ss = ss_(); if (!ss.getSheetByName(SHEET_CHECKS)) setupSheets_();
  var name = checkTab_(meta), sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name, ss.getSheets().length); else sh.clear();
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  head_(sh, header.length);
  return { tab: name };
}
function apiCheckFinish(token, meta, counts) {
  var user = auth_(token);
  var sh = ss_().getSheetByName(SHEET_CHECKS), tab = checkTab_(meta);
  var id = 'K' + meta.start.replace(/-/g, '') + (meta.end && meta.end !== meta.start ? meta.end.replace(/-/g, '').slice(4) : '');
  var row = [id, tab, meta.period || '', meta.start, meta.end || meta.start, String(counts.rows || 0), String(counts.stations || 0), user.name, Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm')];
  var t = checks_().filter(function (c) { return c.id === id; })[0];
  sh.getRange(t ? t.row : sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  return { checks: checks_() };
}
function apiGetCheck(token, id) {
  auth_(token);
  var c = checks_().filter(function (x) { return x.id === id; })[0];
  if (!c) throw new Error('Checking not found.');
  var sh = ss_().getSheetByName(c.tab);
  if (!sh) throw new Error('Tab "' + c.tab + '" not found.');
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  return { check: c, values: lr ? sh.getRange(1, 1, lr, lc).getDisplayValues() : [] };
}


/* ---------- Uploadtrix confirmation per cut-off (DONE / MANUAL) ---------- */
function upxList_() {
  var sh = ss_().getSheetByName(SHEET_UPX);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, UPX_HEADERS.length).getDisplayValues()
    .map(function (r, i) { return { row: i + 2, id: r[0], tab: r[1], start: r[2], end: r[3], stations: +r[4] || 0, done: +r[5] || 0, manual: +r[6] || 0, by: r[7], at: r[8] }; })
    .filter(function (u) { return u.id; })
    .sort(function (a, b) { return a.start < b.start ? 1 : -1; });
}
function upxRows_(tab) {
  var sh = ss_().getSheetByName(tab); if (!sh) throw new Error('Tab "' + tab + '" not found.');
  var lr = sh.getLastRow(); if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, UPX_ROW_HEADERS.length).getDisplayValues().map(function (r, i) {
    return { row: i + 2, no: r[0], station: r[1], company: r[2], status: String(r[3] || '').toUpperCase(), remarks: r[4], by: r[5], at: r[6] };
  }).filter(function (x) { return x.station; });
}
function upxCount_(id) {
  var u = upxList_().filter(function (x) { return x.id === id; })[0]; if (!u) return;
  var rows = upxRows_(u.tab), d = 0, m = 0;
  rows.forEach(function (r) { if (r.status === 'DONE') d++; else if (r.status === 'MANUAL') m++; });
  ss_().getSheetByName(SHEET_UPX).getRange(u.row, 5, 1, 3).setNumberFormat('@').setValues([[String(rows.length), String(d), String(m)]]);
}
// meta: {start, end}. A row for every active station of the List of Stations.
function apiUpxCreate(token, meta) {
  var user = auth_(token);
  if (!meta || !/^\d{4}-\d{2}-\d{2}$/.test(meta.start || '') || !/^\d{4}-\d{2}-\d{2}$/.test(meta.end || '')) throw new Error('Set the cut-off dates.');
  var ss = ss_(); if (!ss.getSheetByName(SHEET_UPX)) setupSheets_();
  var id = 'U' + meta.start.replace(/-/g, '');
  if (upxList_().filter(function (x) { return x.id === id; })[0]) throw new Error('This cut-off already exists. Pick it from the list.');
  var tab = UPX_PREFIX + meta.start + ' to ' + meta.end;
  var sh = ss.getSheetByName(tab) || ss.insertSheet(tab, ss.getSheets().length); sh.clear();
  sh.getRange(1, 1, 1, UPX_ROW_HEADERS.length).setValues([UPX_ROW_HEADERS]); head_(sh, UPX_ROW_HEADERS.length);
  var comp = {}; companyList_().forEach(function (c) { comp[c.code] = c.newCompany || c.oldCompany; });
  var rows = stations_().filter(function (s) { return s.active; }).map(function (s) {
    return [String(s.no || stationNo_(s.station) || ''), s.station, s.newCompany || s.oldCompany || '', '', '', '', ''];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, UPX_ROW_HEADERS.length).setNumberFormat('@').setValues(rows);
  var idx = ss.getSheetByName(SHEET_UPX);
  idx.getRange(idx.getLastRow() + 1, 1, 1, UPX_HEADERS.length).setNumberFormat('@')
    .setValues([[id, tab, meta.start, meta.end, String(rows.length), '0', '0', user.name, Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm')]]);
  return { upx: upxList_(), id: id };
}
function apiUpxGet(token, id) {
  auth_(token);
  var u = upxList_().filter(function (x) { return x.id === id; })[0]; if (!u) throw new Error('Cut-off not found.');
  return { cutoff: u, rows: upxRows_(u.tab) };
}
// updates: [{station, status: 'DONE'|'MANUAL'|'', remarks}]
function apiUpxSet(token, id, updates) {
  var user = auth_(token);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var u = upxList_().filter(function (x) { return x.id === id; })[0]; if (!u) throw new Error('Cut-off not found.');
    var sh = ss_().getSheetByName(u.tab), rows = upxRows_(u.tab), by = {};
    rows.forEach(function (r) { by[key_(r.station)] = r; });
    var now = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm');
    (updates || []).forEach(function (x) {
      var r = by[key_(x.station)];
      if (!r) { // a station not in the list yet
        var nr = sh.getLastRow() + 1;
        sh.getRange(nr, 1, 1, UPX_ROW_HEADERS.length).setNumberFormat('@').setValues([[String(stationNo_(x.station) || ''), x.station, x.company || '', '', '', '', '']]);
        r = { row: nr, remarks: '' }; by[key_(x.station)] = r;
      }
      var st = String(x.status === undefined ? r.status || '' : x.status).toUpperCase();
      var rem = x.remarks === undefined ? r.remarks : x.remarks;
      sh.getRange(r.row, 4, 1, 4).setNumberFormat('@').setValues([[st, rem || '', user.name, now]]);
    });
    upxCount_(id);
    return { upx: upxList_(), rows: upxRows_(u.tab) };
  } finally { lock.releaseLock(); }
}
/* ---------- Data tabs: list every tab by module, delete the ones you don't need ---------- */
function apiListTabs(token) {
  auth_(token);
  var reg = {}; batches_().forEach(function (b) { reg[b.tab] = 'Add-back cut-off ' + b.start + ' to ' + b.end; });
  var wk = {}; weeks_().forEach(function (w) { wk[w.tab] = 'Week ' + w.start + ' to ' + w.end; });
  var ac = {}; actives_().forEach(function (a) { ac[a.tab] = 'Active list as of ' + a.asOf; });
  var ux = {}; upxList_().forEach(function (u) { ux[u.tab] = 'Uploadtrix cut-off ' + u.start + ' to ' + u.end + ' · ' + u.done + ' done, ' + u.manual + ' manual'; });
  var ck = {}; checks_().forEach(function (c) { ck[c.tab] = 'Attendance checking · ' + c.period + ' · ' + c.rows + ' findings'; });
  return ss_().getSheets().map(function (s) {
    var n = s.getName(), lr = s.getLastRow(), lc = s.getLastColumn();
    var t = { name: n, rows: Math.max(0, lr - 1), cols: lc, cells: lr * lc, system: SYSTEM_SHEETS.indexOf(n) >= 0, module: 'other', note: '' };
    if (t.system) { t.module = 'system'; t.note = 'Built-in tab (cannot be deleted here)'; }
    else if (n.indexOf(ATT_PREFIX) === 0) { t.module = 'attendance'; t.note = wk[n] || 'Not linked to a week'; }
    else if (n.indexOf(ACT_PREFIX) === 0) { t.module = 'active'; t.note = ac[n] || 'Not linked to a list'; }
    else if (n.indexOf(CHK_PREFIX) === 0) { t.module = 'checking'; t.note = ck[n] || 'Not linked to a checking'; }
    else if (n.indexOf(UPX_PREFIX) === 0) { t.module = 'uploadtrix'; t.note = ux[n] || 'Not linked to a cut-off'; }
    else if (reg[n]) { t.module = 'addback'; t.note = reg[n]; }
    else {
      var head = lr ? s.getRange(1, 1, Math.min(lr, 15), Math.min(Math.max(lc, 1), 20)).getValues() : [];
      if (headerRowIndex_(head) >= 0) { t.module = 'addback'; t.note = 'Add-back data (no cut-off set yet)'; }
    }
    return t;
  });
}

function apiDeleteTabs(token, names) {
  auth_(token);
  var ss = ss_();
  (names || []).forEach(function (n) {
    if (SYSTEM_SHEETS.indexOf(n) >= 0) throw new Error('"' + n + '" is a built-in tab and cannot be deleted.');
  });
  (names || []).forEach(function (n) {
    var w = weeks_().filter(function (x) { return x.tab === n; })[0]; if (w) ss.getSheetByName(SHEET_WEEKS).deleteRow(w.row);
    var a = actives_().filter(function (x) { return x.tab === n; })[0]; if (a) ss.getSheetByName(SHEET_ACTIVE).deleteRow(a.row);
    var k = checks_().filter(function (x) { return x.tab === n; })[0]; if (k) ss.getSheetByName(SHEET_CHECKS).deleteRow(k.row);
    var u = upxList_().filter(function (x) { return x.tab === n; })[0]; if (u) ss.getSheetByName(SHEET_UPX).deleteRow(u.row);
    var b = batches_().filter(function (x) { return x.tab === n; })[0]; if (b) ss.getSheetByName(SHEET_BATCHES).deleteRow(b.row);
    var sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  return apiBootstrap(token);
}

var SEED = {"stations":[[1,"01-Marikina - RGI","Rephil Guerilla Inc.","GMA-Rizal","RGI","Sariaya"],[2,"02-Sampaloc - RAI","Rephil Antipolo Inc.","NCR","RAI","Sariaya"],[3,"03-Taytay - RMI","Rephil Marikina Inc.","GMA-Rizal","RMI","Sariaya"],[4,"04-Lubao - RAI","Rephil Antipolo Inc.","Pampanga 2","RAI","Subic"],[5,"05-Batangas - RMI","Rephil Marikina Inc.","South Luzon 1","RMI","Sariaya"],[6,"06-Balagtas - FGI","Fortune Global Inc.","South Bulacan","FGI","Subic"],[7,"07-San Ildefonso - FGI","Fortune Global Inc.","North Bulacan","FGI","Subic"],[8,"08-Basak - RBI","Rephil Basak Inc.","Cebu","RBI","Cebu"],[9,"09-Naga - RASI","Rephil Abad Santos Inc.","Cebu","RASI","Cebu"],[10,"10-JP Rizal - RMI","Rephil Marikina Inc.","NCR","RMI","Sariaya"],[11,"11-Masinag - RAI","Rephil Antipolo Inc.","GMA-Rizal","RAI","Sariaya"],[12,"12-Juan Luna - RMI","Rephil Marikina Inc.","NCR","RMI","Sariaya"],[13,"13-Cainta - RAI","Rephil Antipolo Inc.","GMA-Rizal","RAI","Sariaya"],[14,"14-Carcar - RASI","Rephil Abad Santos Inc.","Cebu","RASI","Cebu"],[15,"15-AS Fortuna - ASF","Rephil AS Fortuna Inc.","Cebu","ASF","Cebu"],[16,"16-Capitangan - RGI","Rephil Guerilla Inc.","Pampanga 2","RGI","Subic"],[17,"17-Molino - RGI","Rephil Guerilla Inc.","Cavite","RGI","Sariaya"],[18,"18-Las Pinas - ASF","Rephil AS Fortuna Inc.","GMA-Rizal","ASF","Sariaya"],[19,"19-Balibago - FFFCI","Fast Forward Fuels Co., Inc.","South","FFCI","Sariaya"],[20,"20-Edsa Pasay - REPI","Rephil Edsa Pasay Inc.","GMA-Rizal","REPI","Sariaya"],[21,"21-Nasugbu - RMI","Rephil Marikina Inc.","South Luzon 2","RMI","Sariaya"],[22,"22-Camarin - RBI","Rephil Basak Inc.","NCR","RBI","Sariaya"],[23,"23-Sabang - REPI","Rephil Edsa Pasay Inc.","South Luzon 1","REPI","Sariaya"],[24,"24-Cordova - FFC","Flash Fill Co., Inc.","Cebu","FFC","Cebu"],[25,"25-Gun-ob - FFC","Flash Fill Co., Inc.","Cebu","FFC","Cebu"],[26,"26-Marilao - FFC","Flash Fill Co., Inc.","South Bulacan","FFC","Sariaya"],[27,"27-Tayud - RASI","Rephil Abad Santos Inc.","Cebu","RASI","Cebu"],[28,"28-Biga - ASF","Rephil AS Fortuna Inc.","Cavite","ASF","Sariaya"],[29,"29-Pardo - RASI","Rephil Abad Santos Inc.","Cebu","RASI","Cebu"],[30,"30-Calumpit - FFC","Flash Fill Co., Inc.","South Bulacan","FFC","Subic"],[31,"31-San Jose - REPI","Rephil Edsa Pasay Inc.","North Bulacan","REPI","Subic"],[32,"32-Cavite - FFC","Flash Fill Co., Inc.","Cavite","FFC","Sariaya"],[33,"33-Malolos - FFC","Flash Fill Co., Inc.","North Bulacan","FFC","Subic"],[34,"34-Mexico - REPI","Rephil Edsa Pasay Inc.","Pampanga 1","REPI","Subic"],[35,"35-Baesa - FFC","Flash Fill Co., Inc.","NCR","FFC","Sariaya"],[36,"36-Silang - REPI","Rephil Edsa Pasay Inc.","Cavite","REPI","Sariaya"],[37,"37-Caloocan - REPI","Rephil Edsa Pasay Inc.","NCR","REPI","Sariaya"],[38,"38-Baliwag - REPI","Rephil Edsa Pasay Inc.","North Bulacan","REPI","Subic"],[39,"39- Bocaue - REPI","Rephil Edsa Pasay Inc.","South Bulacan","REPI","Sariaya"],[40,"40-Cabuyao - REPI","Rephil Edsa Pasay Inc.","South Luzon 2","REPI","Sariaya"],[41,"41- San Juan - RASI","Rephil Abad Santos Inc.","GMA-Rizal","RASI","Sariaya"],[42,"42-Sto. Tomas - RASI","Rephil Abad Santos Inc.","Pampanga 1","RASI","Subic"],[43,"43-Cabiao - RGI","Rephil Guerilla Inc.","Central Luzon","RGI","Subic"],[44,"44-Arayat - RGI","Rephil Guerilla Inc.","Central Luzon","RGI","Subic"],[45,"45- Del Pilar - RGI","Rephil Guerilla Inc.","Pampanga 1","RGI","Subic"],[46,"46-San Matias - RMI","Rephil Marikina Inc.","Pampanga 1","RMI","Subic"],[47,"47-Bacoor - ASF","Rephil AS Fortuna Inc.","GMA-Rizal","ASF","Sariaya"],[48,"48-San Fernando - RGI","Rephil Guerilla Inc.","Pampanga 1","RGI","Subic"],[49,"49- Sta. Maria - RAI","Rephil Antipolo Inc.","South Bulacan","RAI","Subic"],[50,"50- San Agustin - RBI","Rephil Basak Inc.","Pampanga 1","RBI","Subic"],[51,"51-Carmona - RGI","Rephil Guerilla Inc.","Cavite","RGI","Sariaya"],[52,"52-Lazatin - RBI","Rephil Basak Inc.","Pampanga 1","RBI","Subic"],[53,"53-Pulilan - RBI","Rephil Basak Inc.","North Bulacan","RBI","Subic"],[54,"54-Susano - FFC","Flash Fill Co., Inc.","NCR","FFC","Sariaya"],[55,"55-Sta. Rosa - RBI","Rephil Basak Inc.","North Bulacan","RBI","Subic"],[56,"56-Rosario - ASF","Rephil AS Fortuna Inc.","Cavite","ASF","Sariaya"],[57,"57-Pradera - RBI","Rephil Basak Inc.","Pampanga 2","RBI","Subic"],[58,"58-San Leonardo - FGI","Fortune Global Inc.","North Bulacan","FGI","Subic"],[59,"59-San Miguel - FGI","Fortune Global Inc.","North Bulacan","FGI","Subic"],[60,"60-Bagbaguin - FGI","Fortune Global Inc.","South Bulacan","FGI","Sariaya"],[61,"61-Tuy - FGI","Fortune Global Inc.","South Luzon 2","FGI","Sariaya"],[62,"62-Tagaytay - FGI","Fortune Global Inc.","South Luzon 2","FGI","Sariaya"],[63,"63-Sta. Ana - RBI","Rephil Basak Inc.","Central Luzon","RBI","Subic"],[64,"64-Magalang - FGI","Fortune Global Inc.","Central Luzon","FGI","Subic"],[65,"65-Alaminos - FGI","Fortune Global Inc.","South Luzon 1","FGI","Sariaya"],[66,"66-Rodriguez - FGI","Fortune Global Inc.","NCR","FGI","Sariaya"],[67,"67-Mabalacat - FGI","Fortune Global Inc.","Central Luzon","FGI","Subic"],[68,"68-Pulong Buhangin - FGI","Fortune Global Inc.","South Bulacan","FGI","Subic"],[69,"69-Sta. Clara - FGI","Fortune Global Inc.","South Bulacan","FGI","Sariaya"],[70,"70-Hermosa - FGI","Fortune Global Inc.","Pampanga 2","FGI","Subic"],[71,"71-Bacolor - FGI","Fortune Global Inc.","Pampanga 2","FGI","Subic"],[72,"72-Muzon - FGI","Fortune Global Inc.","South Bulacan","FGI","Subic"],[73,"73-Urdaneta - FFC","Flash Fill Co., Inc.","Central Luzon","FFC","Subic"],[74,"74-Malvar - FGI","Fortune Global Inc.","South Luzon 1","FGI","Sariaya"],[75,"75-Tarlac City - FFC","Flash Fill Co., Inc.","Central Luzon","FFC","Subic"],[76,"76-Quezon - FFC","Flash Fill Co., Inc.","South Luzon 1","FFC","Sariaya"],[77,"77-Calamba - ASF","Rephil AS Fortuna Inc.","South Luzon 2","ASF","Sariaya"],[78,"78-Guagua - ASF","Rephil AS Fortuna Inc.","Pampanga 2","ASF","Subic"],[79,"79-Jasa - ASF","Rephil AS Fortuna Inc.","Central Luzon","ASF","Subic"],[80,"80-Trece Martirez - ASF","Rephil AS Fortuna Inc.","Cavite","ASF","Sariaya"],[81,"81-General Trias - ASF","Rephil AS Fortuna Inc.","Cavite","ASF","Sariaya"],[82,"82-Paciano - ASF","Rephil AS Fortuna Inc.","South Luzon 2","ASF","Sariaya"],[83,"83-San Roque - RGI","Rephil Guerilla Inc.","GMA-Rizal","RGI","Sariaya"],[84,"84-Candelaria - FFC","Flash Fill Co., Inc.","South Luzon 1","FFC","Sariaya"],[85,"85-Palico - ASF","Rephil AS Fortuna Inc.","South Luzon 2","ASF","Sariaya"],[86,"86-Sta. Barbara - RBI","Rephil Basak Inc.","Pampanga 2","RBI","Subic"],[87,"87-Alaminos 2 - REPI","Rephil Edsa Pasay Inc.","South Luzon 1","REPI","Sariaya"],[88,"88-Ibaan - REPI","Rephil Edsa Pasay Inc.","South Luzon 1","REPI","Sariaya"],[89,"89-San Rafael - REPI","Rephil Edsa Pasay Inc.","North Bulacan","REPI","Subic"],[90,"90-Dolores - RGI","Rephil Guerilla Inc.","Pampanga 1","RGI","Subic"],[91,"91-San Pablo - REPI","Rephil Edsa Pasay Inc.","South Luzon 1","REPI","Sariaya"],[92,"92-San Pascual - RMI","Rephil Marikina Inc.","South Luzon 1","RMI","Sariaya"],[93,"93-DRT - RBI","Rephil Basak Inc.","North Bulacan","RBI","Subic"],[94,"94-Balanga - RMI","Rephil Marikina Inc.","Pampanga 2","RMI","Subic"],[95,"95-Lemery - RMI","Rephil Marikina Inc.","South Luzon 2","RMI","Sariaya"],[96,"96-Toledo - ASF","Rephil AS Fortuna Inc.","Cebu","ASF","Cebu"],[97,"97-Tanza - FFFCI","Fast Forward Fuels Co., Inc.","Cavite","FFFCI","Sariaya"],[98,"98-Guiguinto - RAI","Rephil Antipolo Inc.","South Bulacan","RAI","Subic"],[99,"99-Commonwealth - REPI","Rephil Edsa Pasay Inc.","NCR","REPI","Sariaya"],[100,"100-Belfast - FFC","Flash Fill Co., Inc.","NCR","FFC","Sariaya"],[101,"101-Gen. Luis - ASF","Rephil AS Fortuna Inc.","NCR","ASF","Sariaya"],[102,"102-Balamban - RBI","Rephil Basak Inc.","Cebu","RBI","Cebu"],[103,"103-Regalado - RBI","Rephil Basak Inc.","NCR","RBI","Sariaya"],[104,"104-Arnaldo - ASF","Rephil AS Fortuna Inc.","Cavite","ASF","Sariaya"],[105,"105-Sta. Rita - FFFCI","Fast Forward Fuels Co., Inc.","South Bulacan","FFFCI","Subic"],[106,"106-Tina-an - FFFCI","Fast Forward Fuels Co., Inc.","Cebu","FFFCI","Cebu"],[107,"107-Balayan - FFFCI","Fast Forward Fuels Co., Inc.","South Luzon 2","FFFCI","Sariaya"],[108,"108-Meycauayan - FFFCI","Fast Forward Fuels Co., Inc.","South Bulacan","FFFCI","Sariaya"],[109,"109-Milagrosa - FFFCI","Fast Forward Fuels Co., Inc.","South Luzon 2","FFFCI","Sariaya"],[110,"110-Edsa Monumento - FFFCI","Fast Forward Fuels Co., Inc.","NCR","FFFCI","Sariaya"],[111,"111-Manggahan - FFFCI","Fast Forward Fuels Co., Inc.","Cavite","FFFCI","Sariaya"],[112,"112-Sumulong - FFFCI","Fast Forward Fuels Co., Inc.","GMA-Rizal","FFFCI","Sariaya"],[113,"113-Pandan - FFFCI","Fast Forward Fuels Co., Inc.","Pampanga 2","FFFCI","Subic"],[114,"114-Zuellig - FFFCI","Fast Forward Fuels Co., Inc.","Cebu","FFFCI","Cebu"],[115,"115-Tejero - FFC","Flash Fill Co., Inc.","Cavite","FFC","Sariaya"],[116,"116-Langkaan - RASI","Rephil Abad Santos Inc.","Cavite","RASI","Sariaya"],[117,"117-Padre Pio - RGI","Rephil Guerilla Inc.","South Luzon 2","RGI","Sariaya"],[118,"118-Abellana - FFFCI","Fast Forward Fuels Co., Inc.","Cebu","FFFCI","Cebu"],[119,"119-Malabañas - RBI","Rephil Basak Inc.","Pampanga 1","RBI","Subic"],[120,"120-Antipolo - RAI","Rephil Antipolo Inc.","GMA-Rizal","RAI","Sariaya"],[121,"121-Buaya - RGI","Rephil Guerilla Inc.","Cebu","RGI","Cebu"],[122,"122-Minglanilla - FFFCI","Fast Forward Fuels Co., Inc.","Cebu","FFFCI","Cebu"]],"manpower":[["01-Marikina - RGI",3,10,3,3,10,0],["02-Sampaloc - RAI",2,5,1,2,5,2],["03-Taytay - RMI",3,6,1,3,6,0],["04-Lubao - RAI",3,6,1,3,6,0],["05-Batangas - RMI",2,6,1,2,6,2],["06-Balagtas - FGI",3,7,1,3,7,0],["07-San Ildefonso - FGI",3,9,3,3,9,0],["08-Basak - RBI",3,9,3,3,9,0],["09-Naga - RASI",3,9,3,3,9,0],["10-JP Rizal - RMI",3,9,3,3,9,0],["11-Masinag - RAI",3,9,3,3,9,0],["12-Juan Luna - RMI",3,6,1,3,6,0],["13-Cainta - RAI",3,8,2,3,8,0],["14-Carcar - RASI",2,7,1,2,7,1],["15-AS Fortuna - ASF",3,8,2,3,8,0],["16-Capitangan - RGI",3,9,3,3,9,0],["17-Molino - RGI",3,6,1,3,6,0],["18-Las Pinas - ASF",3,9,3,3,9,0],["19-Balibago - FFFCI",3,9,3,3,9,0],["20-Edsa Pasay - REPI",3,10,3,3,10,0],["21-Nasugbu - RMI",3,6,1,3,6,0],["22-Camarin - RBI",3,9,3,3,9,0],["23-Sabang - REPI",3,8,2,3,8,0],["24-Cordova - FFC",2,5,1,2,5,1],["25-Gun-ob - FFC",2,6,1,2,6,1],["26-Marilao - FFC",3,9,3,3,9,0],["27-Tayud - RASI",3,9,3,3,9,0],["28-Biga - ASF",3,15,5,3,15,0],["29-Pardo - RASI",3,6,1,3,6,0],["30-Calumpit - FFC",3,8,2,3,8,0],["31-San Jose - REPI",2,5,1,2,5,2],["32-Cavite - FFC",3,6,1,3,6,0],["33-Malolos - FFC",3,6,1,3,6,0],["34-Mexico - REPI",3,8,2,3,8,0],["35-Baesa - FFC",3,13,5,3,13,0],["36-Silang - REPI",3,7,1,3,7,0],["37-Caloocan - REPI",2,4,1,2,4,4],["38-Baliwag - REPI",3,6,1,3,6,0],["39- Bocaue - REPI",2,6,1,2,6,2],["40-Cabuyao - REPI",3,9,3,3,9,0],["41- San Juan - RASI",2,6,1,2,6,2],["42-Sto. Tomas - RASI",3,7,1,3,7,0],["43-Cabiao - RGI",2,4,1,2,4,2],["44-Arayat - RGI",2,4,1,2,4,2],["45- Del Pilar - RGI",3,6,1,3,6,0],["46-San Matias - RMI",3,6,1,3,6,0],["47-Bacoor - ASF",3,5,1,3,5,0],["48-San Fernando - RGI",3,9,3,3,9,0],["49- Sta. Maria - RAI",3,15,5,3,15,0],["50- San Agustin - RBI",3,9,3,3,9,0],["51-Carmona - RGI",3,13,5,3,13,0],["52-Lazatin - RBI",3,6,1,3,6,0],["53-Pulilan - RBI",3,9,3,3,9,0],["54-Susano - FFC",3,8,2,3,8,0],["55-Sta. Rosa - RBI",2,4,1,2,4,2],["56-Rosario - ASF",2,4,1,2,4,2],["57-Pradera - RBI",3,9,3,3,9,0],["58-San Leonardo - FGI",3,7,1,3,7,0],["59-San Miguel - FGI",3,7,1,3,7,0],["60-Bagbaguin - FGI",3,10,3,3,10,0],["61-Tuy - FGI",3,8,2,3,8,0],["62-Tagaytay - FGI",3,10,3,3,10,0],["63-Sta. Ana - RBI",2,4,1,2,4,3],["64-Magalang - FGI",3,9,3,3,9,0],["65-Alaminos - FGI",3,6,1,3,6,0],["66-Rodriguez - FGI",3,8,2,3,8,0],["67-Mabalacat - FGI",3,13,5,3,13,0],["68-Pulong Buhangin - FGI",3,6,1,3,6,0],["69-Sta. Clara - FGI",3,9,3,3,9,0],["70-Hermosa - FGI",3,12,4,3,12,0],["71-Bacolor - FGI",3,6,1,3,6,0],["72-Muzon - FGI",3,8,2,3,8,0],["73-Urdaneta - FFC",3,6,1,3,6,0],["74-Malvar - FGI",3,8,2,3,8,0],["75-Tarlac City - FFC",3,8,2,3,8,0],["76-Quezon - FFC",2,6,1,2,6,2],["77-Calamba - ASF",3,9,3,3,9,0],["78-Guagua - ASF",3,9,3,3,9,0],["79-Jasa - ASF",3,8,2,3,8,0],["80-Trece Martirez - ASF",2,6,1,2,6,2],["81-General Trias - ASF",3,7,1,3,7,0],["82-Paciano - ASF",3,12,3,3,12,0],["83-San Roque - RGI",3,6,2,3,6,0],["84-Candelaria - FFC",3,7,1,3,7,0],["85-Palico - ASF",2,4,1,2,4,2],["86-Sta. Barbara - RBI",3,12,3,3,12,0],["87-Alaminos 2 - REPI",3,10,3,3,10,0],["88-Ibaan - REPI",3,12,3,3,12,0],["89-San Rafael - REPI",3,10,3,3,10,0],["90-Dolores - RGI",3,9,3,3,9,0],["91-San Pablo - REPI",3,6,1,3,6,0],["92-San Pascual - RMI",3,9,3,3,9,0],["93-DRT - RBI",3,6,1,3,6,0],["94-Balanga - RMI",3,10,3,3,10,0],["95-Lemery - RMI",3,7,1,3,7,0],["96-Toledo - ASF",2,4,1,2,4,1],["97-Tanza - FFFCI",2,6,1,2,6,3],["98-Guiguinto - RAI",3,12,3,3,12,0],["99-Commonwealth - REPI",3,10,3,3,10,0],["100-Belfast - FFC",3,12,3,3,12,0],["101-Gen. Luis - ASF",3,10,3,3,10,0],["102-Balamban - RBI",2,4,1,2,4,1],["103-Regalado - RBI",3,9,3,3,9,0],["104-Arnaldo - ASF",3,6,1,3,6,0],["105-Sta. Rita - FFFCI",3,9,3,3,9,0],["106-Tina-an - FFFCI",3,10,3,3,10,0],["107-Balayan - FFFCI",3,5,1,3,5,0],["108-Meycauayan - FFFCI",3,9,3,3,9,0],["109-Milagrosa - FFFCI",3,9,3,3,9,0],["110-Edsa Monumento - FFFCI",3,9,3,3,9,0],["111-Manggahan - FFFCI",3,7,1,3,7,0],["112-Sumulong - FFFCI",3,9,3,3,9,0],["113-Pandan - FFFCI",3,9,3,3,9,0],["114-Zuellig - FFFCI",3,8,2,3,8,0],["115-Tejero - FFC",3,6,1,3,6,0],["116-Langkaan - RASI",3,9,3,3,9,0],["117-Padre Pio - RGI",3,8,3,3,8,0],["118-Abellana - FFFCI",3,7,1,3,7,0],["119-Malabañas - RBI",3,6,1,3,6,0],["120-Antipolo - RAI",3,6,1,3,6,0]]};
