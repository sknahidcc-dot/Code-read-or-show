/**
 * Job Express Attendance - Google Apps Script backend
 * Deploy: Extensions > Apps Script > paste > Deploy > New deployment > Web app
 *   Execute as: Me   |   Who has access: Anyone
 * Use the /exec URL in the QR code (Admin > QR Code).
 */
var TZ = 'Asia/Dhaka';
var RADIUS_M = 300; // allowed distance from the institution (meters)

var CFG = {
  library:  { students: 'Library Students',  att: 'Library Attendance',  idKey: 'lid', third: 'shelf number', heads: ['Lid', 'Name', 'Shelf Number', 'Date', 'Time'] },
  coaching: { students: 'Coaching Students', att: 'Coaching Attendance', idKey: 'cid', third: 'batch name',   heads: ['Cid', 'Name', 'Batch Name', 'Date', 'Time'] }
};

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function norm_(s) { return String(s === undefined || s === null ? '' : s).trim(); }
function low_(s) { return norm_(s).toLowerCase(); }
function digits_(s) { return norm_(s).replace(/\D/g, ''); }

function sheet_(name) {
  var all = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < all.length; i++) {
    if (low_(all[i].getName()) === low_(name)) return all[i];
  }
  return null;
}
function rows_(name) {
  var sh = sheet_(name);
  if (!sh) return [];
  var v = sh.getDataRange().getDisplayValues();
  if (v.length < 2) return [];
  var h = v[0].map(low_);
  var res = [];
  for (var r = 1; r < v.length; r++) {
    var o = {}, any = false;
    for (var c = 0; c < h.length; c++) {
      o[h[c]] = norm_(v[r][c]);
      if (o[h[c]]) any = true;
    }
    if (any) res.push(o);
  }
  return res;
}

function settings_() {
  var sh = sheet_('Setting');
  var res = { batches: [], user: '', pass: '', lat: null, lng: null };
  if (!sh) return res;
  var v = sh.getDataRange().getDisplayValues();
  if (v.length < 1) return res;
  var h = v[0].map(low_);
  var cb = h.indexOf('batch name'), cu = h.indexOf('admin user name'), cp = h.indexOf('admin password');
  var cla = h.indexOf('latitude'), cln = h.indexOf('longitude');
  for (var r = 1; r < v.length; r++) {
    if (cb >= 0 && norm_(v[r][cb])) res.batches.push(norm_(v[r][cb]));
  }
  if (v.length > 1) {
    if (cu >= 0) res.user = norm_(v[1][cu]);
    if (cp >= 0) res.pass = norm_(v[1][cp]);
    if (cla >= 0 && !isNaN(parseFloat(v[1][cla]))) res.lat = parseFloat(v[1][cla]);
    if (cln >= 0 && !isNaN(parseFloat(v[1][cln]))) res.lng = parseFloat(v[1][cln]);
  }
  return res;
}

function isAdmin_(p) {
  var s = settings_();
  var u = s.user || 'Job', w = s.pass || 'express';
  return norm_(p.u) === u && norm_(p.pw) === w;
}

function dist_(la1, ln1, la2, ln2) {
  var R = 6371000, rad = Math.PI / 180;
  var dLa = (la2 - la1) * rad, dLn = (ln2 - ln1) * rad;
  var a = Math.sin(dLa / 2) * Math.sin(dLa / 2) +
          Math.cos(la1 * rad) * Math.cos(la2 * rad) * Math.sin(dLn / 2) * Math.sin(dLn / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function findStudent_(type, id) {
  var cfg = CFG[type];
  var list = rows_(cfg.students);
  for (var i = 0; i < list.length; i++) {
    if (low_(list[i][cfg.idKey]) === low_(id)) return list[i];
  }
  return null;
}
function pub_(type, st) {
  return {
    id: st[CFG[type].idKey], name: st['name'] || '', phone: st['phone number'] || '',
    shelf: st['shelf number'] || '', batch: st['batch name'] || '', number: st['number'] || '',
    blood: st['blood group'] || '', email: st['email'] || '', photo: st['photo link'] || ''
  };
}
function attRows_(type) {
  var cfg = CFG[type];
  return rows_(cfg.att).map(function (o) {
    return {
      t: type, id: o[cfg.idKey] || o['lid'] || o['cid'] || '', name: o['name'] || '',
      g: o[cfg.third] || o['shelf number'] || o['batch name'] || '', d: o['date'] || '', tm: o['time'] || ''
    };
  });
}
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function doGet(e) {
  var p = (e && e.parameter) || {};
  var lock = null;
  try {
    var r;
    if (p.action === 'mark' || p.action === 'addStudent') {
      lock = LockService.getScriptLock();
      lock.waitLock(20000);
    }
    switch (p.action) {
      case 'settings': r = settingsAction_(); break;
      case 'today': r = todayAction_(); break;
      case 'mark': r = mark_(p); break;
      case 'profile': r = profile_(p); break;
      case 'adminStats': r = isAdmin_(p) ? adminStats_() : { ok: true, status: 'auth' }; break;
      case 'records': r = isAdmin_(p) ? records_(p) : { ok: true, status: 'auth' }; break;
      case 'addStudent': r = isAdmin_(p) ? addStudent_(p) : { ok: true, status: 'auth' }; break;
      default: r = { ok: true, message: 'Job Express Attendance API running' };
    }
    return out_(r);
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (x) {} }
  }
}

function settingsAction_() {
  var s = settings_();
  return { ok: true, batches: s.batches, lat: s.lat, lng: s.lng };
}

function countToday_(type) {
  var d = today_(), seen = {};
  attRows_(type).forEach(function (r) { if (r.d === d) seen[low_(r.id)] = 1; });
  return Object.keys(seen).length;
}
function todayAction_() {
  return { ok: true, library: countToday_('library'), coaching: countToday_('coaching') };
}
function adminStats_() {
  return {
    ok: true,
    libraryTotal: rows_(CFG.library.students).length,
    coachingTotal: rows_(CFG.coaching.students).length,
    libraryToday: countToday_('library'),
    coachingToday: countToday_('coaching')
  };
}

function mark_(p) {
  var cfg = CFG[p.type];
  if (!cfg) return { ok: false, error: 'bad_type' };
  var id = norm_(p.id);
  var st = findStudent_(p.type, id);
  if (!st) return { ok: true, status: 'not_member' };
  var student = pub_(p.type, st);

  if (p.type === 'coaching' && norm_(p.batch) && low_(st['batch name']) !== low_(p.batch)) {
    return { ok: true, status: 'wrong_batch', student: student };
  }

  // location check
  var s = settings_();
  var tLat = s.lat !== null ? s.lat : parseFloat(p.qlat);
  var tLng = s.lng !== null ? s.lng : parseFloat(p.qlng);
  if (!isNaN(tLat) && !isNaN(tLng)) {
    var la = parseFloat(p.lat), ln = parseFloat(p.lng);
    if (isNaN(la) || isNaN(ln)) return { ok: true, status: 'no_location' };
    var d = dist_(la, ln, tLat, tLng);
    if (d > RADIUS_M) return { ok: true, status: 'too_far', distance: Math.round(d) };
  }

  // identity: device binding + phone verification on first use
  var props = PropertiesService.getScriptProperties();
  var key = 'dev_' + p.type + '_' + low_(id);
  var bound = props.getProperty(key);
  if (bound && bound !== p.device) return { ok: true, status: 'device_mismatch' };
  if (!bound) {
    var ph = digits_(st['phone number']);
    var q = digits_(p.phone4);
    if (q.length < 4) return { ok: true, status: 'need_verify' };
    if (ph.length < 4 || ph.slice(-4) !== q.slice(-4)) return { ok: true, status: 'bad_phone' };
    props.setProperty(key, p.device);
  }

  var date = today_();
  var time = Utilities.formatDate(new Date(), TZ, 'hh:mm a');
  var dup = attRows_(p.type).some(function (r) { return r.d === date && low_(r.id) === low_(id); });
  if (dup) {
    var prev = attRows_(p.type).filter(function (r) { return r.d === date && low_(r.id) === low_(id); })[0];
    return { ok: true, status: 'already', student: student, date: date, time: prev ? prev.tm : '' };
  }

  var sh = sheet_(cfg.att);
  if (!sh) {
    sh = SpreadsheetApp.getActiveSpreadsheet().insertSheet(cfg.att);
    sh.getRange(1, 1, 1, 5).setValues([cfg.heads]);
  }
  var third = p.type === 'library' ? st['shelf number'] : st['batch name'];
  var row = sh.getLastRow() + 1;
  var rng = sh.getRange(row, 1, 1, 5);
  rng.setNumberFormat('@');
  rng.setValues([[st[cfg.idKey], st['name'], third || '', date, time]]);
  SpreadsheetApp.flush();
  return { ok: true, status: 'ok', student: student, date: date, time: time };
}

function profile_(p) {
  var cfg = CFG[p.type];
  if (!cfg) return { ok: false, error: 'bad_type' };
  var st = findStudent_(p.type, norm_(p.id));
  if (!st) return { ok: true, status: 'not_member' };
  var bound = PropertiesService.getScriptProperties().getProperty('dev_' + p.type + '_' + low_(p.id));
  if (bound && bound !== p.device) return { ok: true, status: 'device_mismatch' };
  var recs = attRows_(p.type).filter(function (r) { return low_(r.id) === low_(p.id); });
  return { ok: true, student: pub_(p.type, st), records: recs.slice(-500) };
}

function records_(p) {
  var days = parseInt(p.days, 10) || 30;
  var start = Utilities.formatDate(new Date(Date.now() - (days - 1) * 86400000), TZ, 'yyyy-MM-dd');
  var types = p.type === 'all' ? ['library', 'coaching'] : [p.type];
  var res = [];
  types.forEach(function (t) {
    if (!CFG[t]) return;
    attRows_(t).forEach(function (r) {
      if (r.d >= start && (!norm_(p.id) || low_(r.id) === low_(p.id))) res.push(r);
    });
  });
  return { ok: true, records: res };
}

function addStudent_(p) {
  var cfg = CFG[p.type];
  if (!cfg) return { ok: false, error: 'bad_type' };
  if (findStudent_(p.type, norm_(p.id))) return { ok: true, status: 'exists' };
  var sh = sheet_(cfg.students);
  if (!sh) return { ok: false, error: 'sheet_missing' };
  var heads = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0].map(low_);
  var map = {
    'lid': p.id, 'cid': p.id, 'name': p.name, 'phone number': p.phone, 'shelf number': p.shelf,
    'batch name': p.batch, 'number': p.number, 'blood group': p.blood, 'email': p.email, 'photo link': p.photo
  };
  var row = heads.map(function (h) { return map[h] === undefined ? '' : norm_(map[h]); });
  var r = sh.getLastRow() + 1;
  var rng = sh.getRange(r, 1, 1, row.length);
  rng.setNumberFormat('@');
  rng.setValues([row]);
  return { ok: true, status: 'ok' };
}

/** Run manually from the editor if a student changes phone: resetDevice('coaching','CID001') */
function resetDevice(type, id) {
  PropertiesService.getScriptProperties().deleteProperty('dev_' + type + '_' + low_(id));
}
