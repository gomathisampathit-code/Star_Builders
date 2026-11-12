/**
 * STAR BUILDERS – shared database (Google Apps Script, bound to a Google Sheet)
 * ---------------------------------------------------------------------------
 * Every entry saved in the app is stored in the "Entries" tab of this sheet,
 * and every device reads the same data from here.
 *
 * Passwords: the SAME two passwords as the app (stored here only as hashes).
 * To change a password, see the "How to change a password" note in the setup guide.
 */
const SALT = 'starbuilders:';
const HASH_USER  = 'bef79138a21127a8b1631fd45caabd85f5b3a661e593f2ad5a6be52891ef2f53';
const HASH_ADMIN = '997c02d779750e97e0a99cb7535036d3f3374e22781d11ee41222f778ce2e0ea';
const TZ = 'Asia/Kolkata';

const E_HEAD = ['id', 'site', 'date', 'paid_for', 'description', 'nos', 'paid_to', 'amount', 'advance', 'breakup_json', 'updated'];
const S_HEAD = ['site'];
const A_HEAD = ['id', 'site', 'date', 'labour', 'number', 'contractor', 'note', 'updated'];   // attendance
const R_HEAD = ['labour', 'daily_wage'];                                                      // wages

function doGet() {
  return out_({ ok: true, msg: 'Star Builders API is running' });
}

function doPost(e) {
  let lock = null;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const role = roleOf_(b.pw);
    if (!role) return out_({ ok: false, err: 'auth' });

    const ops = Array.isArray(b.ops) ? b.ops : [];
    if (!ops.length) {                       // plain read
      const ver = getVer_();
      if (b.read && Number(b.ver) === ver) return out_({ ok: true, role: role, ver: ver, unchanged: true });
      return out_({ ok: true, role: role, ver: ver, data: dump_(load_()) });
    }

    lock = LockService.getScriptLock();
    lock.waitLock(25000);
    const db = load_();
    const errs = [];
    let changed = false;
    ops.forEach(function (op) {
      try { if (apply_(db, op, role, errs)) changed = true; }
      catch (x) { errs.push(String(x)); }
    });
    if (changed) { store_(db); bumpVer_(); }
    return out_({ ok: true, role: role, ver: getVer_(), data: dump_(db), errs: errs });
  } catch (x) {
    return out_({ ok: false, err: String(x) });
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (_) {} }
  }
}

/* ---------------- operations ---------------- */
function apply_(db, op, role, errs) {
  const adm = role === 'admin';
  const t = String(op && op.t || '');
  const need = function () { if (!adm) { errs.push('Admin only: ' + t); return false; } return true; };

  if (t === 'put') {
    const site = str_(op.site);
    if (!site || !op.row) return false;
    let name = findSite_(db, site);
    if (!name) {
      if (!adm) { errs.push('Site not found: ' + site); return false; }
      db.sites.push(site); name = site;
    }
    return putRow_(db, name, op.row, errs);
  }
  if (t === 'del') {
    if (!need()) return false;
    const n = db.rows.length;
    db.rows = db.rows.filter(function (r) { return r.id !== Number(op.id); });
    return db.rows.length !== n;
  }
  if (t === 'attPut') {                                  // attendance entry (any user)
    const site = str_(op.site);
    if (!site || !op.row) return false;
    let name = findSite_(db, site);
    if (!name) {
      if (!adm) { errs.push('Site not found: ' + site); return false; }
      db.sites.push(site); name = site;
    }
    return putAtt_(db, name, op.row, errs);
  }
  if (t === 'attDel') {
    if (!need()) return false;
    const n = db.att.length;
    db.att = db.att.filter(function (a) { return a.id !== Number(op.id); });
    return db.att.length !== n;
  }
  if (t === 'rate') {                                    // daily wage of a labour type (admin)
    if (!need()) return false;
    const k = str_(op.lab), v = Number(op.amt);
    if (!k) return false;
    if (v > 0) db.rates[k] = v; else delete db.rates[k];
    return true;
  }
  if (t === 'addSite') {
    const n = str_(op.name);
    if (!n || findSite_(db, n)) return false;
    db.sites.push(n); return true;
  }
  if (t === 'renSite') {
    const from = findSite_(db, str_(op.from)), to = str_(op.to);
    if (!from || !to) return false;
    const clash = findSite_(db, to);
    if (clash && clash.toLowerCase() !== from.toLowerCase()) { errs.push('A site named "' + to + '" already exists'); return false; }
    db.sites[db.sites.indexOf(from)] = to;
    db.rows.forEach(function (r) { if (r.site === from) r.site = to; });
    db.att.forEach(function (a) { if (a.site === from) a.site = to; });
    return true;
  }
  if (t === 'delSite') {
    if (!need()) return false;
    const n = findSite_(db, str_(op.name));
    if (!n) return false;
    if (db.sites.length < 2) { errs.push('At least one site must remain'); return false; }
    db.sites.splice(db.sites.indexOf(n), 1);
    db.rows = db.rows.filter(function (r) { return r.site !== n; });
    db.att = db.att.filter(function (a) { return a.site !== n; });
    return true;
  }
  if (t === 'bulk' || t === 'seed') {
    if (t === 'seed' && !need()) return false;   // Excel import ('bulk') is open to normal users; the one-time seed is admin-only
    if (t === 'seed' && (db.rows.length || db.sites.length || db.att.length)) { errs.push('The shared database already has data, so nothing was uploaded.'); return false; }
    let any = false;
    (op.sites || []).forEach(function (s) {
      const n = str_(s.name);
      if (!n) return;
      let name = findSite_(db, n);
      if (!name) { db.sites.push(n); name = n; any = true; }
      (s.rows || []).forEach(function (r) { if (putRow_(db, name, r, errs)) any = true; });
      (s.att || []).forEach(function (a) { if (putAtt_(db, name, a, errs)) any = true; });
    });
    return any;
  }
  errs.push('Unknown operation: ' + t);
  return false;
}

function putAtt_(db, site, r, errs) {
  const id = Number(r.id), n = Number(r.n);
  if (!isFinite(id) || id <= 0) { errs.push('Attendance entry without a valid id was skipped'); return false; }
  if (!isFinite(n) || n <= 0) { errs.push('Attendance entry without a number was skipped'); return false; }
  const row = { id: id, site: site, date: str_(r.date).slice(0, 10), lab: str_(r.lab), n: Math.round(n * 100) / 100, by: str_(r.by), note: str_(r.note), ts: Date.now() };
  if (!row.date || !row.lab) { errs.push('Attendance entry without date or labour type was skipped'); return false; }
  const i = db.att.findIndex(function (x) { return x.id === id; });
  if (i >= 0) db.att[i] = row; else db.att.push(row);
  return true;
}

function putRow_(db, site, r, errs) {
  const id = Number(r.id);
  if (!isFinite(id) || id <= 0) { errs.push('Entry without a valid id was skipped'); return false; }
  const row = {
    id: id, site: site,
    date: str_(r.date).slice(0, 10), cat: str_(r.cat), desc: str_(r.desc),
    nos: str_(r.nos), to: str_(r.to),
    amt: Number(r.amt) || 0, adv: Number(r.adv) || 0,
    br: (Array.isArray(r.br) ? r.br : []).map(function (l) {
      return { nos: str_(l.nos), date: str_(l.date).slice(0, 10), cat: str_(l.cat), desc: str_(l.desc), to: str_(l.to), amt: Number(l.amt) || 0 };
    }),
    ts: Date.now()
  };
  const i = db.rows.findIndex(function (x) { return x.id === id; });
  if (i >= 0) db.rows[i] = row; else db.rows.push(row);
  return true;
}

/* ---------------- sheet <-> memory ---------------- */
function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function load_() {
  const es = sheet_('Entries', E_HEAD), ss = sheet_('Sites', S_HEAD);
  const sites = ss.getDataRange().getValues().slice(1).map(function (r) { return str_(r[0]); }).filter(Boolean);
  const rows = [];
  es.getDataRange().getValues().slice(1).forEach(function (a) {
    const id = Number(a[0]);
    if (!isFinite(id) || id <= 0) return;
    let br = [];
    try { br = JSON.parse(a[9] || '[]'); } catch (_) {}
    const site = str_(a[1]);
    if (site && sites.indexOf(site) < 0) sites.push(site);
    rows.push({
      id: id, site: site, date: str_(a[2]).slice(0, 10), cat: str_(a[3]), desc: str_(a[4]),
      nos: str_(a[5]), to: str_(a[6]), amt: Number(a[7]) || 0, adv: Number(a[8]) || 0,
      br: Array.isArray(br) ? br : [], ts: Number(a[10]) || 0
    });
  });
  const att = [];
  sheet_('Attendance', A_HEAD).getDataRange().getValues().slice(1).forEach(function (a) {
    const id = Number(a[0]);
    if (!isFinite(id) || id <= 0) return;
    const site = str_(a[1]);
    if (site && sites.indexOf(site) < 0) sites.push(site);
    att.push({ id: id, site: site, date: str_(a[2]).slice(0, 10), lab: str_(a[3]), n: Number(a[4]) || 0, by: str_(a[5]), note: str_(a[6]), ts: Number(a[7]) || 0 });
  });
  const rates = {};
  sheet_('Rates', R_HEAD).getDataRange().getValues().slice(1).forEach(function (r) {
    const k = str_(r[0]), v = Number(r[1]);
    if (k && v > 0) rates[k] = v;
  });
  return { sites: sites, rows: rows, att: att, rates: rates };
}

function store_(db) {
  const es = sheet_('Entries', E_HEAD), ss = sheet_('Sites', S_HEAD);
  // sites
  const sMax = Math.max(db.sites.length, 1);
  if (ss.getMaxRows() < sMax + 1) ss.insertRowsAfter(ss.getMaxRows(), sMax + 1 - ss.getMaxRows());
  ss.getRange(2, 1, Math.max(ss.getMaxRows() - 1, 1), 1).clearContent();
  if (db.sites.length) ss.getRange(2, 1, db.sites.length, 1).setNumberFormat('@').setValues(db.sites.map(function (n) { return [n]; }));
  // entries
  const n = db.rows.length;
  if (es.getMaxRows() < n + 1) es.insertRowsAfter(es.getMaxRows(), n + 1 - es.getMaxRows());
  es.getRange(2, 1, Math.max(es.getMaxRows() - 1, 1), E_HEAD.length).clearContent();
  if (n) {
    es.getRange(2, 2, n, 6).setNumberFormat('@');      // site..paid_to as plain text
    es.getRange(2, 10, n, 1).setNumberFormat('@');
    es.getRange(2, 1, n, E_HEAD.length).setValues(db.rows.map(function (r) {
      return [r.id, r.site, r.date, r.cat, r.desc, r.nos, r.to, r.amt, r.adv, JSON.stringify(r.br || []), r.ts];
    }));
  }
  // attendance
  const as = sheet_('Attendance', A_HEAD), na = db.att.length;
  if (as.getMaxRows() < na + 1) as.insertRowsAfter(as.getMaxRows(), na + 1 - as.getMaxRows());
  as.getRange(2, 1, Math.max(as.getMaxRows() - 1, 1), A_HEAD.length).clearContent();
  if (na) {
    as.getRange(2, 2, na, 3).setNumberFormat('@');
    as.getRange(2, 6, na, 2).setNumberFormat('@');
    as.getRange(2, 1, na, A_HEAD.length).setValues(db.att.map(function (a) { return [a.id, a.site, a.date, a.lab, a.n, a.by, a.note, a.ts]; }));
  }
  // wages
  const rs = sheet_('Rates', R_HEAD), keys = Object.keys(db.rates);
  if (rs.getMaxRows() < keys.length + 1) rs.insertRowsAfter(rs.getMaxRows(), keys.length + 1 - rs.getMaxRows());
  rs.getRange(2, 1, Math.max(rs.getMaxRows() - 1, 1), 2).clearContent();
  if (keys.length) {
    rs.getRange(2, 1, keys.length, 1).setNumberFormat('@');
    rs.getRange(2, 1, keys.length, 2).setValues(keys.map(function (k) { return [k, db.rates[k]]; }));
  }
}

function dump_(db) {
  return {
    sites: db.sites.map(function (name) {
      return {
        name: name,
        rows: db.rows.filter(function (r) { return r.site === name; }).map(function (r) {
          return { id: r.id, date: r.date, cat: r.cat, desc: r.desc, nos: r.nos, to: r.to, amt: r.amt, adv: r.adv, br: r.br };
        }),
        att: db.att.filter(function (a) { return a.site === name; }).map(function (a) {
          return { id: a.id, date: a.date, lab: a.lab, n: a.n, by: a.by, note: a.note };
        })
      };
    }),
    rates: db.rates
  };
}

/* ---------------- helpers ---------------- */
function str_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v).trim();
}
function findSite_(db, n) {
  const k = String(n).toLowerCase();
  for (let i = 0; i < db.sites.length; i++) if (db.sites[i].toLowerCase() === k) return db.sites[i];
  return '';
}
function sha_(pw) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SALT + pw, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}
function roleOf_(pw) {
  if (pw === undefined || pw === null || pw === '') return '';
  const h = sha_(String(pw));
  return h === HASH_ADMIN ? 'admin' : h === HASH_USER ? 'user' : '';
}
function getVer_() { return Number(PropertiesService.getScriptProperties().getProperty('ver') || 0); }
function bumpVer_() { PropertiesService.getScriptProperties().setProperty('ver', String(getVer_() + 1)); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
