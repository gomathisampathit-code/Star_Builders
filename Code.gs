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
const SC_HEAD = ['id', 'floor', 'seq', 'activity', 'planned_start', 'planned_end', 'concrete', 'actual_start', 'actual_end', 'status', 'remarks', 'updated'];   // sheet "Schedule"
const F_HEAD = ['floor'];                                                                     // sheet "Floors"
const SCH_ST = ['Not started', 'In progress', 'Completed', 'On hold'];

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
  if (t.indexOf('sch') === 0) return schOp_(db, t, op, adm, errs);
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

/* ---- construction schedule (sheets "Schedule" and "Floors") ---- */
function dstr_(v) { const s = str_(v).slice(0, 10); return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : ''; }
function addDays_(s, n) { const p = s.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2]) + n * 864e5).toISOString().slice(0, 10); }
function stOk_(st, a, e) {
  if (e) return 'Completed';
  if (SCH_ST.indexOf(st) < 0) return a ? 'In progress' : 'Not started';
  if (st === 'Completed') return a ? 'In progress' : 'Not started';
  if (st === 'Not started' && a) return 'In progress';
  return st;
}
function schOp_(db, t, op, adm, errs) {
  const q = db.sch;
  const need = function () { if (!adm) { errs.push('Admin only: ' + t); return false; } return true; };
  if (t === 'schPut') {
    const r = op.row || {}, id = Number(r.id);
    if (!isFinite(id) || id <= 0) { errs.push('Schedule entry without a valid id was skipped'); return false; }
    const a = dstr_(r.as), e = dstr_(r.ae);
    if (a && e && e < a) { errs.push('Actual End is before Actual Start - not saved'); return false; }
    const ex = q.rows.find(function (x) { return x.id === id; });
    if (ex) {                                              // existing activity: anyone may change actual dates, status, remarks
      ex.as = a; ex.ae = e; ex.st = stOk_(str_(r.st), a, e); ex.rm = str_(r.rm); ex.ts = Date.now();
      return true;
    }
    if (!need()) return false;                             // new activity: admin only
    const floor = str_(r.floor), ps = dstr_(r.ps), pe = dstr_(r.pe), name = str_(r.name);
    if (!floor || !name || !ps || !pe || pe < ps) { errs.push('Activity needs a floor, name and valid planned dates'); return false; }
    if (q.floors.indexOf(floor) < 0) q.floors.push(floor);
    q.rows.push({ id: id, floor: floor, seq: Number(r.seq) || 0, name: name, ps: ps, pe: pe, conc: !!r.conc, as: a, ae: e, st: stOk_(str_(r.st), a, e), rm: str_(r.rm), ts: Date.now() });
    return true;
  }
  if (!need()) return false;
  if (t === 'schPlan') {
    const r = op.row || {}, ex = q.rows.find(function (x) { return x.id === Number(r.id); });
    const ps = dstr_(r.ps), pe = dstr_(r.pe), name = str_(r.name);
    if (!ex || !name || !ps || !pe || pe < ps) { errs.push('Planned dates not saved - check the activity name and dates'); return false; }
    ex.name = name; ex.ps = ps; ex.pe = pe; ex.conc = !!r.conc; ex.ts = Date.now();
    return true;
  }
  if (t === 'schDel') { const n = q.rows.length; q.rows = q.rows.filter(function (x) { return x.id !== Number(op.id); }); return q.rows.length !== n; }
  if (t === 'schFloorAdd') { const n = str_(op.name); if (!n || q.floors.indexOf(n) >= 0) return false; q.floors.push(n); return true; }
  if (t === 'schFloorRen') {
    const f = str_(op.from), n = str_(op.to), i = q.floors.indexOf(f);
    if (i < 0 || !n || q.floors.indexOf(n) >= 0) return false;
    q.floors[i] = n; q.rows.forEach(function (x) { if (x.floor === f) x.floor = n; }); return true;
  }
  if (t === 'schFloorDel') {
    const f = str_(op.name), n = q.rows.length, had = q.floors.indexOf(f) >= 0;
    q.floors = q.floors.filter(function (x) { return x !== f; }); q.rows = q.rows.filter(function (x) { return x.floor !== f; });
    return had || q.rows.length !== n;
  }
  if (t === 'schReplan') {
    const f = str_(op.floor), d = Math.round(Number(op.days));
    if (!isFinite(d) || !d) return false;
    let any = false; q.rows.forEach(function (x) { if (x.floor === f) { x.ps = addDays_(x.ps, d); x.pe = addDays_(x.pe, d); any = true; } }); return any;
  }
  if (t === 'schImport') {
    const f = str_(op.floor);
    if (!f) return false;
    if (op.replace) q.rows = q.rows.filter(function (x) { return x.floor !== f; });
    if (q.floors.indexOf(f) < 0) q.floors.push(f);
    (op.rows || []).forEach(function (r) {
      const id = Number(r.id), ps = dstr_(r.ps), pe = dstr_(r.pe), name = str_(r.name);
      if (!isFinite(id) || id <= 0 || !name || !ps || !pe || pe < ps) { errs.push('Skipped an imported row: ' + name); return; }
      const a = dstr_(r.as), e = dstr_(r.ae), i = q.rows.findIndex(function (x) { return x.id === id; });
      const row = { id: id, floor: f, seq: Number(r.seq) || 0, name: name, ps: ps, pe: pe, conc: !!r.conc, as: a, ae: e, st: stOk_(str_(r.st), a, e), rm: str_(r.rm), ts: Date.now() };
      if (i >= 0) q.rows[i] = row; else q.rows.push(row);
    });
    return true;
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
  const floors = [];
  sheet_('Floors', F_HEAD).getDataRange().getValues().slice(1).forEach(function (r) { const f = str_(r[0]); if (f && floors.indexOf(f) < 0) floors.push(f); });
  const srows = [];
  sheet_('Schedule', SC_HEAD).getDataRange().getValues().slice(1).forEach(function (a) {
    const id = Number(a[0]);
    if (!isFinite(id) || id <= 0) return;
    const f = str_(a[1]);
    if (f && floors.indexOf(f) < 0) floors.push(f);
    srows.push({ id: id, floor: f, seq: Number(a[2]) || 0, name: str_(a[3]), ps: dstr_(a[4]), pe: dstr_(a[5]), conc: String(a[6]) === '1' || a[6] === true, as: dstr_(a[7]), ae: dstr_(a[8]), st: str_(a[9]), rm: str_(a[10]), ts: Number(a[11]) || 0 });
  });
  return { sites: sites, rows: rows, att: att, rates: rates, sch: { floors: floors, rows: srows } };
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
  // floors -> sheet "Floors"
  const fs = sheet_('Floors', F_HEAD), nf = db.sch.floors.length;
  if (fs.getMaxRows() < nf + 1) fs.insertRowsAfter(fs.getMaxRows(), nf + 1 - fs.getMaxRows());
  fs.getRange(2, 1, Math.max(fs.getMaxRows() - 1, 1), 1).clearContent();
  if (nf) fs.getRange(2, 1, nf, 1).setNumberFormat('@').setValues(db.sch.floors.map(function (f) { return [f]; }));
  // construction schedule -> sheet "Schedule"
  const cs = sheet_('Schedule', SC_HEAD), nc = db.sch.rows.length;
  if (cs.getMaxRows() < nc + 1) cs.insertRowsAfter(cs.getMaxRows(), nc + 1 - cs.getMaxRows());
  cs.getRange(2, 1, Math.max(cs.getMaxRows() - 1, 1), SC_HEAD.length).clearContent();
  if (nc) {
    cs.getRange(2, 2, nc, 1).setNumberFormat('@');
    cs.getRange(2, 4, nc, 3).setNumberFormat('@');
    cs.getRange(2, 8, nc, 4).setNumberFormat('@');
    cs.getRange(2, 1, nc, SC_HEAD.length).setValues(db.sch.rows.map(function (x) {
      return [x.id, x.floor, x.seq, x.name, x.ps, x.pe, x.conc ? 1 : 0, x.as, x.ae, x.st, x.rm, x.ts];
    }));
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
    rates: db.rates,
    sch: { floors: db.sch.floors, rows: db.sch.rows.map(function (x) { return { id: x.id, floor: x.floor, seq: x.seq, name: x.name, ps: x.ps, pe: x.pe, conc: x.conc, as: x.as, ae: x.ae, st: x.st, rm: x.rm }; }) }
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
