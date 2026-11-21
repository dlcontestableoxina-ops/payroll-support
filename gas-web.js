/* Payroll Support · free web version
   Runs the same server code (code.js = Code.gs) in the browser. Every Google Sheet tab is kept in Supabase:
   table ps_sheets (one row per tab: version, size, first rows) and ps_parts (the tab's rows, gzip-compressed).
   Saving checks the version, so two people saving at the same time never overwrite each other: the second one is re-run on the newest data.
   Sign-in is a Supabase account (email + password). Only signed-in users can read or write (Row Level Security). */
(function () {
  'use strict';
  var CFG = window.PS_CONFIG || {};
  var sb = null;
  function client() {
    if (sb) return sb;
    if (!CFG.supabaseUrl || !CFG.supabaseAnonKey || /YOUR-/.test(CFG.supabaseUrl + CFG.supabaseAnonKey)) throw new Error('Setup not finished: open config.js and paste your Supabase Project URL and anon key.');
    var make = window.PS_SUPABASE_FACTORY || (window.supabase && window.supabase.createClient);
    if (!make) throw new Error('Could not load Supabase. Check your internet connection, then reload.');
    sb = make(CFG.supabaseUrl, CFG.supabaseAnonKey, { auth: { persistSession: true, autoRefreshToken: true, storageKey: 'ps-web-auth' } });
    return sb;
  }
  function fail(e) { var m = e && (e.message || e.error_description || e.msg) || String(e); if (/PS_CONFLICT/.test(m)) { var c = new Error('conflict'); c.conflict = true; return c; } if (/JWT|not authenticated|permission denied|row-level security/i.test(m)) return new Error('SESSION_EXPIRED'); return new Error(m); }

  /* ---------- values ↔ compressed text ---------- */
  function enc(rows) { return JSON.stringify(rows, function (k, v) { var o = this[k]; return o instanceof Date ? { $d: o.getTime() } : v; }); }
  function dec(s) { return JSON.parse(s, function (k, v) { return v && typeof v === 'object' && !Array.isArray(v) && typeof v.$d === 'number' && Object.keys(v).length === 1 ? new Date(v.$d) : v; }); }
  function toB64(buf) { var b = new Uint8Array(buf), s = '', i; for (i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000)); return btoa(s); }
  function fromB64(t) { var s = atob(t), b = new Uint8Array(s.length); for (var i = 0; i < s.length; i++) b[i] = s.charCodeAt(i); return b; }
  async function gz(str) { if (typeof CompressionStream === 'undefined') return 'J' + str; var st = new Blob([str]).stream().pipeThrough(new CompressionStream('gzip')); return 'G' + toB64(await new Response(st).arrayBuffer()); }
  async function ungz(t) { if (t[0] === 'J') return t.slice(1); var st = new Blob([fromB64(t.slice(1))]).stream().pipeThrough(new DecompressionStream('gzip')); return await new Response(st).text(); }
  var PART = 700000, HEAD_ROWS = 2;

  /* ---------- what is stored ---------- */
  var META = {}, ORDER = [], DATA = {}, metaAt = 0;
  async function refreshMeta(force) {
    if (!force && Date.now() - metaAt < 800) return;
    var r = await client().from('ps_sheets').select('name,version,parts,rows,cols,head').order('created_at', { ascending: true });
    if (r.error) throw fail(r.error);
    META = {}; ORDER = [];
    r.data.forEach(function (m) { m.headRows = m.head ? dec(m.head) : []; META[m.name] = m; ORDER.push(m.name); });
    Object.keys(DATA).forEach(function (n) { if (!META[n] || META[n].version !== DATA[n].version) delete DATA[n]; });
    metaAt = Date.now();
  }
  async function load(names) {
    names = names.filter(function (n, i) { return META[n] && !DATA[n] && names.indexOf(n) === i; });
    for (var i = 0; i < names.length; i += 15) {
      var chunk = names.slice(i, i + 15), r = await client().from('ps_parts').select('name,version,part,data').in('name', chunk);
      if (r.error) throw fail(r.error);
      for (var j = 0; j < chunk.length; j++) {
        var n = chunk[j], v = META[n].version, ps = r.data.filter(function (p) { return p.name === n && p.version === v; }).sort(function (a, b) { return a.part - b.part; });
        if (ps.length !== (META[n].parts || ps.length) || !ps.length) { metaAt = 0; continue; } // saved by someone this second: read again
        DATA[n] = { version: v, d: dec(await ungz(ps.map(function (p) { return p.data; }).join(''))) };
      }
    }
  }
  function smallOnes() { return ORDER.filter(function (n) { var m = META[n]; return !DATA[n] && m.rows * Math.max(1, m.cols) <= 30000; }); }

  /* ---------- the Sheet that the server code sees during one call ---------- */
  function Need(n) { this.need = n; }
  var W = null;
  function Sheet(name, isNew) { this.name = name; this.orig = isNew ? null : name; this.d = isNew ? [] : null; this.own = !!isNew; this.dirty = !!isNew; }
  Sheet.prototype = {
    _src: function () { if (this.own) return this.d; var c = DATA[this.orig]; if (!c) throw new Need(this.orig); return c.d; },
    _mut: function () { if (!this.own) { this.d = this._src().map(function (r) { return r.slice(); }); this.own = true; } this.dirty = true; return this.d; },
    _loaded: function () { return this.own || !!DATA[this.orig]; },
    getName: function () { return this.name; },
    setName: function (n) { if (n === this.name) return this; if (W.byName[n]) throw new Error('A sheet named "' + n + '" already exists.'); this._mut(); delete W.byName[this.name]; if (this.orig) W.deleted[this.orig] = true; this.orig = null; this.name = n; W.byName[n] = this; return this; },
    getLastRow: function () { return this._loaded() ? this._src().length : META[this.orig].rows; },
    getLastColumn: function () { return this._loaded() ? this._src().reduce(function (m, r) { return Math.max(m, r.length); }, 0) : META[this.orig].cols; },
    getMaxRows: function () { return Math.max(1000, this.getLastRow()); }, getMaxColumns: function () { return Math.max(26, this.getLastColumn()); },
    getRange: function (r, c, nr, nc) {
      var sh = this; nr = nr || 1; nc = nc || 1;
      function rows() { if (!sh._loaded() && r + nr - 1 <= Math.min(HEAD_ROWS, META[sh.orig].headRows.length)) return META[sh.orig].headRows; return sh._src(); }
      var R = {
        getValues: function () { var d = rows(), o = []; for (var i = 0; i < nr; i++) { var row = [], src = d[r - 1 + i] || []; for (var j = 0; j < nc; j++) { var v = src[c - 1 + j]; row.push(v === undefined || v === null ? '' : v); } o.push(row); } return o; },
        getValue: function () { return R.getValues()[0][0]; },
        getDisplayValues: function () { return R.getValues().map(function (x) { return x.map(function (v) { return String(v); }); }); },
        setValues: function (v) { var d = sh._mut(); v.forEach(function (row, i) { while (d.length < r + i) d.push([]); var t = d[r - 1 + i]; row.forEach(function (x, j) { while (t.length < c - 1 + j) t.push(''); t[c - 1 + j] = x; }); }); return R; },
        setFormulas: function (v) { return R.setValues(v); }, setValue: function (x) { return R.setValues([[x]]); },
        clearContent: function () { var d = sh._mut(); for (var i = 0; i < nr; i++) for (var j = 0; j < nc; j++) if (d[r - 1 + i] && d[r - 1 + i].length > c - 1 + j) d[r - 1 + i][c - 1 + j] = ''; return R; },
        setNumberFormat: function () { return R; }, setFontWeight: function () { return R; }, setBackground: function () { return R; }, setFontColor: function () { return R; }, setFontStyle: function () { return R; },
        insertCheckboxes: function () { return R; }, setWrap: function () { return R; }, setHorizontalAlignment: function () { return R; }, setVerticalAlignment: function () { return R; }, setFontSize: function () { return R; }, setBorder: function () { return R; }, setNote: function () { return R; }
      };
      return R;
    },
    clear: function () { this.d = []; this.own = true; this.dirty = true; return this; },
    appendRow: function (row) { this._mut().push(row.slice()); return this; },
    deleteRow: function (r) { this._mut().splice(r - 1, 1); return this; },
    deleteRows: function (r, n) { this._mut().splice(r - 1, n); return this; },
    showColumns: function () {}, hideColumns: function () {}, setFrozenRows: function () {}, setColumnWidths: function () {}, setColumnWidth: function () {}, autoResizeColumns: function () {}, setTabColor: function () {}, activate: function () { return this; }
  };
  function hidden(n) { return n.indexOf('__') === 0; }
  var SS = {
    getSheetByName: function (n) {
      if (W.byName[n]) return W.byName[n];
      if (W.deleted[n] || !META[n]) return null;
      return (W.byName[n] = new Sheet(n, false));
    },
    insertSheet: function (n) { if (SS.getSheetByName(n)) throw new Error('A sheet with the name "' + n + '" already exists.'); return (W.byName[n] = new Sheet(n, true)); },
    deleteSheet: function (s) { delete W.byName[s.name]; if (s.orig) W.deleted[s.orig] = true; },
    getSheets: function () {
      var out = ORDER.filter(function (n) { return !hidden(n) && !W.deleted[n]; }).map(function (n) { return SS.getSheetByName(n); }).filter(Boolean);
      Object.keys(W.byName).forEach(function (n) { var s = W.byName[n]; if (!hidden(n) && out.indexOf(s) < 0) out.push(s); });
      return out;
    },
    getUrl: function () { return CFG.supabaseUrl ? CFG.supabaseUrl.replace(/^https:\/\/([a-z0-9]+)\.supabase\.co.*$/, 'https://supabase.com/dashboard/project/$1/editor') : '#'; },
    getSpreadsheetTimeZone: function () { return 'Asia/Manila'; }, getId: function () { return 'supabase'; }, getName: function () { return 'Payroll Support'; }
  };
  window.SpreadsheetApp = { getActiveSpreadsheet: function () { if (!W) throw new Error('No open call'); return SS; }, openById: function () { return SS; }, flush: function () {}, getUi: function () { throw new Error('no ui'); } };
  window.Session = { getScriptTimeZone: function () { return 'Asia/Manila'; } };
  window.LockService = { getScriptLock: function () { return { waitLock: function () {}, tryLock: function () { return true; }, releaseLock: function () {} }; } };
  window.Logger = { log: function () {} };
  // settings the server code keeps (PropertiesService) live in a hidden tab "__props"
  window.PropertiesService = { getScriptProperties: function () {
    function sh() { return SS.getSheetByName('__props') || SS.insertSheet('__props'); }
    return {
      getProperty: function (k) { var d = sh().getLastRow() ? sh().getRange(1, 1, sh().getLastRow(), 2).getValues() : []; for (var i = 0; i < d.length; i++) if (d[i][0] === k) return String(d[i][1]); return null; },
      setProperty: function (k, v) { var s = sh(), n = s.getLastRow(), d = n ? s.getRange(1, 1, n, 2).getValues() : []; for (var i = 0; i < d.length; i++) if (d[i][0] === k) { s.getRange(i + 1, 2).setValue(String(v)); return; } s.appendRow([k, String(v)]); }
    };
  } };
  var CACHE = {};
  window.CacheService = { getScriptCache: function () { return { get: function (k) { var x = CACHE[k]; return x && (!x.t || x.t > Date.now()) ? x.v : null; }, put: function (k, v, sec) { CACHE[k] = { v: String(v), t: sec ? Date.now() + sec * 1000 : 0 }; }, remove: function (k) { delete CACHE[k]; } }; } };
  if (!window.Utilities) window.Utilities = {};
  Utilities.getUuid = function () { return (crypto.randomUUID ? crypto.randomUUID() : 'u' + Math.random().toString(16).slice(2) + Date.now().toString(16)); };
  Utilities.formatDate = function (d, tz, f) {
    var p = {}; new Intl.DateTimeFormat('en-US', { timeZone: tz || 'Asia/Manila', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).formatToParts(d).forEach(function (x) { p[x.type] = x.value; });
    var MO = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][+p.month - 1], hh = p.hour === '24' ? '00' : p.hour;
    var h12 = (+hh % 12) || 12, ap = +hh < 12 ? 'AM' : 'PM';
    return f.replace(/yyyy|MMMM|MMM|MM|dd|HH|hh|mm|ss|a|EEE|\bd\b|\bM\b|\bh\b/g, function (t) {
      return { yyyy: p.year, MMMM: MO, MMM: MO.slice(0, 3), MM: p.month, dd: p.day, HH: hh, hh: (h12 < 10 ? '0' : '') + h12, mm: p.minute, ss: p.second, a: ap, d: String(+p.day), M: String(+p.month), h: String(h12),
        EEE: new Intl.DateTimeFormat('en-US', { timeZone: tz || 'Asia/Manila', weekday: 'short' }).format(d) }[t];
    });
  };
  // SHA-256 for the old Users-tab passwords (kept so Code.gs runs unchanged)
  (function(){function sha(m){function r(n,x){return(x>>>n)|(x<<(32-n));}var K=[1116352408,1899447441,3049323471,3921009573,961987163,1508970993,2453635748,2870763221,3624381080,310598401,607225278,1426881987,1925078388,2162078206,2614888103,3248222580,3835390401,4022224774,264347078,604807628,770255983,1249150122,1555081692,1996064986,2554220882,2821834349,2952996808,3210313671,3336571891,3584528711,113926993,338241895,666307205,773529912,1294757372,1396182291,1695183700,1986661051,2177026350,2456956037,2730485921,2820302411,3259730800,3345764771,3516065817,3600352804,4094571909,275423344,430227734,506948616,659060556,883997877,958139571,1322822218,1537002063,1747873779,1955562222,2024104815,2227730452,2361852424,2428436474,2756734187,3204031479,3329325298];
var H=[1779033703,3144134277,1013904242,2773480762,1359893119,2600822924,528734635,1541459225];var b=Array.from(new TextEncoder().encode(m)),l=b.length*8;b.push(128);while(b.length%64!==56)b.push(0);for(var i=7;i>=0;i--)b.push((l/Math.pow(2,i*8))&255);
for(var o=0;o<b.length;o+=64){var w=[];for(i=0;i<16;i++)w[i]=(b[o+i*4]<<24)|(b[o+i*4+1]<<16)|(b[o+i*4+2]<<8)|b[o+i*4+3];for(i=16;i<64;i++){var s0=r(7,w[i-15])^r(18,w[i-15])^(w[i-15]>>>3),s1=r(17,w[i-2])^r(19,w[i-2])^(w[i-2]>>>10);w[i]=(w[i-16]+s0+w[i-7]+s1)|0;}
var a=H[0],c=H[1],d=H[2],e=H[3],f=H[4],g=H[5],h=H[6],k=H[7];for(i=0;i<64;i++){var S1=r(6,f)^r(11,f)^r(25,f),ch=(f&g)^(~f&h),t1=(k+S1+ch+K[i]+w[i])|0,S0=r(2,a)^r(13,a)^r(22,a),mj=(a&c)^(a&d)^(c&d),t2=(S0+mj)|0;k=h;h=g;g=f;f=(e+t1)|0;e=d;d=c;c=a;a=(t1+t2)|0;}
H=[H[0]+a|0,H[1]+c|0,H[2]+d|0,H[3]+e|0,H[4]+f|0,H[5]+g|0,H[6]+h|0,H[7]+k|0];}var out=[];H.forEach(function(x){out.push((x>>24)<<24>>24,(x<<8>>24),(x<<16>>24),(x<<24>>24));});return out;}
window.Utilities.computeDigest=function(alg,s){return sha(s);};window.Utilities.DigestAlgorithm={SHA_256:'SHA_256'};window.Utilities.Charset={UTF_8:'UTF_8'};})();

  /* ---------- run one server function and save what it changed ---------- */
  async function commit(w) {
    var names = Object.keys(w.byName).filter(function (n) { return w.byName[n].dirty; });
    for (var i = 0; i < names.length; i++) {
      var s = w.byName[names[i]], n = s.name, d = s.d, old = META[n] ? META[n].version : 0;
      var txt = await gz(enc(d)), parts = [];
      for (var k = 0; k < txt.length; k += PART) parts.push(txt.slice(k, k + PART));
      if (!parts.length) parts.push('');
      var nv = old + 1, cols = d.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
      var r1 = await client().from('ps_parts').delete().eq('name', n).eq('version', nv); if (r1.error) throw fail(r1.error);
      var r2 = await client().from('ps_parts').insert(parts.map(function (p, j) { return { name: n, version: nv, part: j, data: p }; })); if (r2.error) throw fail(r2.error);
      var r3 = await client().rpc('ps_commit', { p_name: n, p_old: old, p_parts: parts.length, p_rows: d.length, p_cols: cols, p_head: enc(d.slice(0, HEAD_ROWS)) });
      if (r3.error) { await client().from('ps_parts').delete().eq('name', n).eq('version', nv); throw fail(r3.error); }
      client().from('ps_parts').delete().eq('name', n).lt('version', nv).then(function () {});
      DATA[n] = { version: nv, d: d };
      if (!META[n]) ORDER.push(n);
      META[n] = { name: n, version: nv, rows: d.length, cols: cols, parts: parts.length, headRows: d.slice(0, HEAD_ROWS) };
    }
    var del = Object.keys(w.deleted).filter(function (n) { return META[n]; });
    for (var j = 0; j < del.length; j++) {
      var r4 = await client().rpc('ps_drop', { p_name: del[j] }); if (r4.error) throw fail(r4.error);
      delete META[del[j]]; delete DATA[del[j]]; ORDER = ORDER.filter(function (x) { return x !== del[j]; });
    }
  }
  var queue = Promise.resolve();
  function run(fn, args) { // one call at a time, like Apps Script with a lock
    var p = queue.then(function () { return runNow(fn, args); });
    queue = p.catch(function () {});
    return p;
  }
  async function runNow(fn, args) {
    if (typeof window[fn] !== 'function') throw new Error('Unknown function ' + fn);
    for (var attempt = 0; attempt < 40; attempt++) {
      await refreshMeta(attempt > 0 && attempt % 5 === 0);
      W = { byName: {}, deleted: {} };
      var res, w;
      try { res = window[fn].apply(null, JSON.parse(JSON.stringify(args))); w = W; }
      catch (e) { W = null; if (e instanceof Need) { await load([e.need].concat(smallOnes())); continue; } throw e; }
      W = null;
      try { await commit(w); }
      catch (e) { if (e.conflict) { metaAt = 0; await new Promise(function (r) { setTimeout(r, 150 + Math.random() * 400); }); continue; } throw e; }
      return res === undefined ? null : JSON.parse(JSON.stringify(res));
    }
    throw new Error('The data is busy (someone else is saving). Try again in a moment.');
  }

  /* ---------- sign-in with a Supabase account ---------- */
  var SESSION_KEY = 'ps_web_session';
  function keep(token, user) { CACHE['tok_' + token] = { v: JSON.stringify(user), t: 0 }; try { localStorage.setItem(SESSION_KEY, JSON.stringify({ token: token, user: user })); } catch (e) {} }
  // the Users tab is the list of people: name, position and Active. The first person to sign in takes the first row.
  window.__webProfile = function (email, metaName) {
    setupSheets_();
    var U = users_(), rows = U.rows;
    for (var i = 0; i < rows.length; i++) if (String(rows[i][0]).trim().toLowerCase() === email) {
      if (String(rows[i][4]).toUpperCase() === 'FALSE') throw new Error('This account is switched off (Active is FALSE in the Users tab).');
      if (String(rows[i][1]) !== '(Supabase login)') U.sh.getRange(i + 2, 2).setValue('(Supabase login)');
      logSignIn_(email, true);
      return { username: email, name: rows[i][2] || metaName || email.split('@')[0], position: rows[i][3] || 'Payroll Support Assistant' };
    }
    // the old Apps Script "admin" row (from the Google Sheet) is taken over by the first email that signs in
    for (var k = 0; k < rows.length; k++) if (String(rows[k][0]).trim().toLowerCase() === 'admin') {
      U.sh.getRange(k + 2, 1, 1, 2).setValues([[email, '(Supabase login)']]); logSignIn_(email, true);
      return { username: email, name: rows[k][2] || metaName || email.split('@')[0], position: rows[k][3] || 'Payroll Support Assistant' };
    }
    var seed = false;
    var name = seed ? rows[0][2] : metaName || email.split('@')[0].replace(/[._]/g, ' ').replace(/\b\w/g, function (c) { return c.toUpperCase(); });
    if (seed) U.sh.getRange(2, 1, 1, 2).setValues([[email, '(Supabase login)']]);
    else U.sh.appendRow([email, '(Supabase login)', name, 'Payroll Support Assistant', true]);
    logSignIn_(email, true);
    return { username: email, name: name, position: seed ? rows[0][3] : 'Payroll Support Assistant' };
  };
  // Data → Settings (web): copy a tab from the Google Sheet download, or give every tab for the backup
  var ISO_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/;
  window.__webRestore = function (token, name, values) {
    auth_(token);
    name = String(name || '').trim(); if (!name || hidden(name)) throw new Error('Tab name missing.');
    var w = values.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    var rows = values.map(function (r) { var o = []; for (var j = 0; j < w; j++) { var v = r[j]; o.push(v === null || v === undefined ? '' : typeof v === 'string' && ISO_RE.test(v) ? new Date(v) : v); } return o; });
    while (rows.length && rows[rows.length - 1].every(function (v) { return v === ''; })) rows.pop();
    var sh = SS.getSheetByName(name) || SS.insertSheet(name); sh.clear();
    if (rows.length && w) sh.getRange(1, 1, rows.length, w).setValues(rows);
    return { name: name, rows: rows.length };
  };
  window.__webTabs = function (token) { auth_(token); return SS.getSheets().map(function (s) { return { name: s.getName(), rows: s.getLastRow() }; }); };
  window.__webTab = function (token, name) { auth_(token); var sh = SS.getSheetByName(name); if (!sh || !sh.getLastRow()) return []; return sh.getRange(1, 1, sh.getLastRow(), Math.max(1, sh.getLastColumn())).getValues(); };
  async function afterAuth(u) {
    var email = String(u.email || '').toLowerCase();
    var user = await run('__webProfile', [email, (u.user_metadata && (u.user_metadata.name || u.user_metadata.full_name)) || '']);
    var token = Utilities.getUuid(); keep(token, user);
    load(smallOnes()).catch(function () {});
    return { token: token, user: user };
  }
  var WEB = {
    apiLogin: async function (email, pw) {
      email = String(email || '').trim().toLowerCase();
      if (email.indexOf('@') < 0) throw new Error('Sign in with your email address (the one added in Supabase → Authentication → Users).');
      var r = await client().auth.signInWithPassword({ email: email, password: String(pw || '') });
      if (r.error) throw new Error(/invalid login/i.test(r.error.message) ? 'Wrong email or password.' : r.error.message);
      await refreshMeta(true);
      return afterAuth(r.data.user);
    },
    apiLogout: async function (token) { delete CACHE['tok_' + token]; try { localStorage.removeItem(SESSION_KEY); } catch (e) {} try { await client().auth.signOut(); } catch (e) {} return true; },
    apiChangePassword: async function (token, oldPw, newPw) {
      var user = auth_(token);
      if (!newPw || String(newPw).length < 6) throw new Error('The new password needs at least 6 characters.');
      var r = await client().auth.signInWithPassword({ email: user.username, password: String(oldPw || '') });
      if (r.error) throw new Error('Current password is wrong.');
      var u = await client().auth.updateUser({ password: String(newPw) });
      if (u.error) throw new Error(u.error.message);
      return true;
    }
  };
  // reopening the page: the Supabase session is still valid → stay signed in
  var ready = (async function () {
    try {
      var s = await client().auth.getSession(), saved = JSON.parse(localStorage.getItem(SESSION_KEY) || 'null');
      if (s.data && s.data.session && saved && saved.token) { CACHE['tok_' + saved.token] = { v: JSON.stringify(saved.user), t: 0 }; try { localStorage.setItem('abh_token', JSON.stringify(saved.token)); } catch (e) {} refreshMeta(true).then(function () { return load(smallOnes()); }).catch(function () {}); }
      else { try { localStorage.removeItem(SESSION_KEY); } catch (e) {} }
    } catch (e) { window.__psSetupError = e.message; }
  })();

  /* ---------- google.script.run, the way the app calls the server ---------- */
  function runner(ok, bad) {
    var r = { withSuccessHandler: function (f) { return runner(f, bad); }, withFailureHandler: function (f) { return runner(ok, f); }, withUserObject: function () { return r; } };
    return new Proxy(r, { get: function (t, k) {
      if (k in t) return t[k];
      return function () {
        var a = [].slice.call(arguments);
        ready.then(function () { return WEB[k] ? WEB[k].apply(null, a) : run(k, a); })
          .then(function (res) { ok && ok(res); }, function (e) { console.error('[' + k + ']', e); bad && bad(e instanceof Error ? e : new Error(String(e))); });
      };
    } });
  }
  window.google = { script: { run: runner() } };
  window.PS_WEB = { run: run, refresh: function () { metaAt = 0; DATA = {}; }, stats: function () { return { sheets: ORDER.length, loaded: Object.keys(DATA).length }; } };
})();
