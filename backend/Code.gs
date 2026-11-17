/**
 * DSL Lab Project Submission Portal: secure backend API
 * Ajeenkya D Y Patil University · School of Engineering · Second Year
 *
 * Frontend (GitHub Pages) → this Google Apps Script web app (JSON API) → your Google Sheet + Drive.
 * Students: open access (no login) to submit and track.
 * Faculty: username + password (salted & hashed), role-based access, division scoping,
 *          masked contact numbers, private Drive files, full audit log, login lockout.
 *
 * Nothing secret is stored in this file, so it is safe to keep in a public GitHub repo.
 * See README.md for the 5-minute deployment steps.
 */

const CONFIG = {
  UNIVERSITY: 'Ajeenkya D Y Patil University',
  SCHOOL: 'School of Engineering',
  LAB_NAME: 'Data Structures Laboratory (DSL)',
  CLASS_NAME: 'Second Year B.Tech',
  ACADEMIC_YEAR: '2026-27',
  DEADLINE: '2026-12-15T23:59:00+05:30',   // '' = no deadline
  ACCEPT_LATE: true,                        // false = close submissions after the deadline
  MAX_FILE_MB: 10,
  DIVISIONS: 'ABCDEFGHIJKLM'.split(''),
  PROGRAMS: [
    'B.Tech Computer Engineering',
    'B.Tech Computer Science & Engineering',
    'B.Tech CSE (Artificial Intelligence & Machine Learning)',
    'B.Tech CSE (Data Science)',
    'B.Tech CSE (Cyber Security)',
    'B.Tech Artificial Intelligence & Data Science',
    'B.Tech Information Technology',
    'B.Tech Electronics & Telecommunication Engineering',
    'B.Tech Robotics & Automation',
    'B.Tech Mechanical Engineering',
    'B.Tech Civil Engineering'
  ],
  DOMAINS: ['Arrays & Strings', 'Linked Lists', 'Stacks', 'Queues', 'Trees & BST', 'Graphs', 'Hashing',
            'Searching & Sorting', 'Heaps & Priority Queues', 'Tries', 'Real-world Application', 'Other'],
  BLOCK_DUPLICATE_URN: true,
  STATUSES: ['Submitted', 'Under Review', 'Approved', 'Needs Revision'],
  SESSION_MINUTES: 120,
  MAX_LOGIN_FAILS: 5,
  LOCK_MINUTES: 15
};

const TZ = 'Asia/Kolkata';
const SH_SUB = 'Submissions', SH_FAC = 'Faculty', SH_LOG = 'Audit Log';
const HEADERS = [
  'Submission ID', 'Submitted On', 'Program', 'Division', 'Project Title', 'Domain', 'Abstract', 'Code / Repo Link', 'Team Size',
  'S1 Name (Leader)', 'S1 URN', 'S1 Contact', 'S2 Name', 'S2 URN', 'S2 Contact',
  'S3 Name', 'S3 URN', 'S3 Contact', 'S4 Name', 'S4 URN', 'S4 Contact',
  'File Name', 'File Link', 'File ID', 'Late', 'Status', 'Faculty Remarks', 'Reviewed On', 'Reviewed By'
];
const FAC_HEADERS = ['Username', 'Name', 'Role', 'Divisions', 'Salt', 'Hash', 'Active', 'Created On', 'Last Login'];
const LOG_HEADERS = ['Time', 'User', 'Action', 'Details'];
const COL = {}; HEADERS.forEach((h, i) => COL[h] = i);

/* ================================================================ SETUP */

/** Run once from the Apps Script editor. Prints your admin username + a one-time random password. */
function setup() {
  const p = PropertiesService.getScriptProperties();
  let ss = p.getProperty('SHEET_ID') ? SpreadsheetApp.openById(p.getProperty('SHEET_ID')) : null;
  if (!ss) { ss = SpreadsheetApp.create('DSL Lab Submissions – AY ' + CONFIG.ACADEMIC_YEAR); p.setProperty('SHEET_ID', ss.getId()); }
  const mk = (name, headers) => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      const first = ss.getSheets()[0];
      sh = (name === SH_SUB && ss.getSheets().length === 1 && first.getLastRow() === 0) ? first.setName(name) : ss.insertSheet(name);
    }
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#111827').setFontColor('#fff');
      sh.setFrozenRows(1);
      sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    }
    return sh;
  };
  mk(SH_SUB, HEADERS); const fac = mk(SH_FAC, FAC_HEADERS); mk(SH_LOG, LOG_HEADERS);
  ss.getSheets().forEach(s => { if (s.getName().indexOf('Sheet') === 0 && s.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(s); });

  if (!p.getProperty('FOLDER_ID')) p.setProperty('FOLDER_ID', DriveApp.createFolder('DSL Lab Submissions – AY ' + CONFIG.ACADEMIC_YEAR).getId());

  if (fac.getLastRow() < 2) {
    const pw = randomPassword_();
    const salt = Utilities.getUuid();
    fac.appendRow(['admin', 'Lab Coordinator', 'admin', 'ALL', salt, hash_(pw, salt), 'YES', now_(), '']);
    Logger.log('================  SAVE THIS  ================');
    Logger.log('Admin username : admin');
    Logger.log('Admin password : ' + pw + '   (change it after first login)');
    Logger.log('=============================================');
  }
  Logger.log('Google Sheet : ' + ss.getUrl());
  Logger.log('Drive folder : ' + DriveApp.getFolderById(p.getProperty('FOLDER_ID')).getUrl());
}

/** Emergency: run from the editor to reset the "admin" password if it is forgotten. */
function resetAdminPassword() {
  const sh = db_().ss.getSheetByName(SH_FAC), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === 'admin') {
    const pw = randomPassword_(), salt = Utilities.getUuid();
    sh.getRange(i + 1, 5, 1, 3).setValues([[salt, hash_(pw, salt), 'YES']]);
    Logger.log('New admin password: ' + pw);
    return;
  }
}

function db_() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SHEET_ID') || !p.getProperty('FOLDER_ID')) setup();
  const ss = SpreadsheetApp.openById(p.getProperty('SHEET_ID'));
  return { p, ss, sub: ss.getSheetByName(SH_SUB), folder: DriveApp.getFolderById(p.getProperty('FOLDER_ID')) };
}

/* ================================================================ HTTP */

function doGet() { return json_({ ok: true, service: 'DSL Lab Submission API', time: new Date().toISOString() }); }

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const fn = ROUTES[req.action];
    if (!fn) throw new Error('Unknown action');
    out = { ok: true, data: fn(req) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return json_(out);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

const ROUTES = {
  // ---- public (students)
  init: () => ({ config: publicConfig_(), stats: stats_() }),
  submit: r => submit_(r.payload || {}),
  track: r => track_(r.id, r.urn),
  // ---- faculty (authenticated)
  login: r => login_(r.username, r.password),
  logout: r => { CacheService.getScriptCache().remove('s_' + r.token); return true; },
  dashboard: r => dashboard_(auth_(r)),
  review: r => review_(auth_(r), r.id, r.status, r.remarks),
  reveal: r => reveal_(auth_(r), r.id),
  file: r => file_(auth_(r), r.id),
  changePassword: r => changePassword_(auth_(r), r.oldPassword, r.newPassword),
  // ---- admin only
  export: r => export_(auth_(r, true)),
  logs: r => logs_(auth_(r, true)),
  facultyList: r => facultyList_(auth_(r, true)),
  facultySave: r => facultySave_(auth_(r, true), r.account || {}),
  facultyDelete: r => facultyDelete_(auth_(r, true), r.username),
  links: r => { auth_(r, true); const d = db_(); return { sheetUrl: d.ss.getUrl(), folderUrl: d.folder.getUrl() }; }
};

/* ================================================================ HELPERS */

const now_ = () => Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm:ss');
const clean_ = (s, max) => String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 300);
const safe_ = s => /^[=+\-@]/.test(String(s)) ? "'" + s : String(s);   // block spreadsheet formula injection
const un_ = v => String(v == null ? '' : v).replace(/^'/, '');
const mask_ = c => c ? c.slice(0, 2) + 'XXXXX' + c.slice(-3) : '';
function randomPassword_() {
  const a = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789@#$%';
  let s = ''; const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now());
  for (let i = 0; i < 14; i++) s += a[(b[i] + 256) % a.length];
  return s;
}
function hash_(pw, salt) {
  let h = String(pw);
  for (let i = 0; i < 1000; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + h + '|' + i));
  return h;
}
function log_(user, action, details) {
  try { db_().ss.getSheetByName(SH_LOG).appendRow([now_(), user, action, safe_(clean_(details, 500))]); } catch (e) {}
}
function publicConfig_() {
  const c = {}; ['UNIVERSITY', 'SCHOOL', 'LAB_NAME', 'CLASS_NAME', 'ACADEMIC_YEAR', 'DEADLINE', 'ACCEPT_LATE', 'MAX_FILE_MB', 'DIVISIONS', 'PROGRAMS', 'DOMAINS', 'STATUSES'].forEach(k => c[k] = CONFIG[k]);
  c.serverNow = new Date().toISOString();
  return c;
}
function rows_() {
  const sh = db_().sub, n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, HEADERS.length).getValues();
}
function rowObj_(r, full) {
  const members = [1, 2, 3, 4].map(i => ({
    name: un_(r[COL['S' + i + (i === 1 ? ' Name (Leader)' : ' Name')]]),
    urn: un_(r[COL['S' + i + ' URN']]),
    contact: full ? un_(r[COL['S' + i + ' Contact']]) : mask_(un_(r[COL['S' + i + ' Contact']]))
  })).filter(m => m.name || m.urn);
  return {
    id: un_(r[0]), submittedOn: un_(r[COL['Submitted On']]), program: un_(r[COL['Program']]), division: un_(r[COL['Division']]),
    title: un_(r[COL['Project Title']]), domain: un_(r[COL['Domain']]), abstract: un_(r[COL['Abstract']]), repo: un_(r[COL['Code / Repo Link']]),
    members, fileName: un_(r[COL['File Name']]), late: r[COL['Late']] === 'YES',
    status: un_(r[COL['Status']]) || 'Submitted', remarks: un_(r[COL['Faculty Remarks']]),
    reviewedOn: un_(r[COL['Reviewed On']]), reviewedBy: un_(r[COL['Reviewed By']])
  };
}

/* ================================================================ PUBLIC: stats */

function stats_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('stats');
  if (hit) return JSON.parse(hit);
  const byDiv = {}; CONFIG.DIVISIONS.forEach(d => byDiv[d] = 0);
  let students = 0;
  const rs = rows_();
  rs.forEach(r => { byDiv[r[COL['Division']]] = (byDiv[r[COL['Division']]] || 0) + 1; students += Number(r[COL['Team Size']]) || 0; });
  const s = { total: rs.length, students, byDiv, updated: new Date().toISOString() };
  cache.put('stats', JSON.stringify(s), 30);
  return s;
}

/* ================================================================ PUBLIC: submit */

const RX_PHONE = /^[6-9]\d{9}$/, RX_URN = /^[A-Za-z0-9\/\-]{4,25}$/;
const EXT = { pdf: 'application/pdf', doc: 'application/msword', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };

function validate_(p) {
  if (p.website) throw new Error('Rejected.');               // honeypot for bots
  const o = {};
  o.program = clean_(p.program, 120); if (CONFIG.PROGRAMS.indexOf(o.program) < 0) throw new Error('Please select a valid program.');
  o.division = clean_(p.division, 2).toUpperCase(); if (CONFIG.DIVISIONS.indexOf(o.division) < 0) throw new Error('Please select a valid division.');
  o.title = clean_(p.title, 200); if (o.title.length < 5) throw new Error('Project title must be at least 5 characters.');
  o.domain = clean_(p.domain, 60); o.abstract = clean_(p.abstract, 600); o.repo = clean_(p.repo, 300);
  if (o.repo && !/^https?:\/\//i.test(o.repo)) throw new Error('Code / repo link must start with https://');
  o.members = []; const seen = {};
  (p.members || []).slice(0, 4).forEach((m, i) => {
    const name = clean_(m && m.name, 80), urn = clean_(m && m.urn, 25).toUpperCase(), contact = clean_(m && m.contact, 10);
    if (i > 0 && !(name || urn || contact)) return;
    const L = i === 0 ? 'Student 1 (Team Leader)' : 'Student ' + (i + 1);
    if (name.length < 2) throw new Error(L + ': name is required.');
    if (!RX_URN.test(urn)) throw new Error(L + ': enter a valid URN.');
    if (i === 0 && !contact) throw new Error(L + ': contact number is required.');
    if (contact && !RX_PHONE.test(contact)) throw new Error(L + ': contact must be a 10-digit mobile number.');
    if (seen[urn]) throw new Error('URN ' + urn + ' is entered twice.');
    seen[urn] = 1; o.members.push({ name, urn, contact });
  });
  if (!o.members.length) throw new Error('Student 1 details are required.');
  const f = p.file || {}, ext = String(f.name || '').split('.').pop().toLowerCase();
  if (!EXT[ext]) throw new Error('Only PDF, DOC or DOCX files are allowed.');
  if (!f.data) throw new Error('Please attach your project file.');
  const bytes = Utilities.base64Decode(f.data);
  if (bytes.length > CONFIG.MAX_FILE_MB * 1048576) throw new Error('File is larger than ' + CONFIG.MAX_FILE_MB + ' MB.');
  if (ext === 'pdf' && String.fromCharCode.apply(null, bytes.slice(0, 4)) !== '%PDF') throw new Error('This does not look like a valid PDF file.');
  o.file = { bytes, ext, mime: EXT[ext], original: clean_(f.name, 150) };
  return o;
}

function submit_(payload) {
  const cache = CacheService.getScriptCache();
  const rate = Number(cache.get('rate') || 0);
  if (rate > 60) throw new Error('Server is busy. Please try again in a minute.');
  cache.put('rate', String(rate + 1), 60);

  const d = validate_(payload);
  const late = !!(CONFIG.DEADLINE && new Date() > new Date(CONFIG.DEADLINE));
  if (late && !CONFIG.ACCEPT_LATE) throw new Error('The submission deadline has passed. Please contact your lab faculty.');

  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const s = db_(), rs = rows_();
    if (CONFIG.BLOCK_DUPLICATE_URN) {
      for (const r of rs) for (const k of ['S1 URN', 'S2 URN', 'S3 URN', 'S4 URN']) {
        const u = String(r[COL[k]] || '').toUpperCase();
        if (u && d.members.some(m => m.urn === u)) throw new Error('URN ' + u + ' is already part of submission ' + r[0] + '. Each student can be in only one team. Contact your lab faculty for changes.');
      }
    }
    const queue = rs.filter(r => r[COL['Division']] === d.division && (r[COL['Status']] === 'Submitted' || r[COL['Status']] === 'Under Review')).length + 1;
    const seq = Number(s.p.getProperty('SEQ_' + d.division) || 0) + 1;
    const id = 'DSL-' + d.division + '-' + ('000' + seq).slice(-3);
    const it = s.folder.getFoldersByName('Division ' + d.division);
    const folder = it.hasNext() ? it.next() : s.folder.createFolder('Division ' + d.division);
    const fname = id + '_' + d.members[0].urn + '_' + d.title.replace(/[^A-Za-z0-9 ]+/g, '').trim().replace(/\s+/g, '_').slice(0, 40) + '.' + d.file.ext;
    const file = folder.createFile(Utilities.newBlob(d.file.bytes, d.file.mime, fname));

    const row = new Array(HEADERS.length).fill('');
    const set = (k, v) => row[COL[k]] = v;
    set('Submission ID', id); set('Submitted On', now_()); set('Program', d.program); set('Division', d.division);
    set('Project Title', safe_(d.title)); set('Domain', safe_(d.domain)); set('Abstract', safe_(d.abstract)); set('Code / Repo Link', safe_(d.repo));
    set('Team Size', String(d.members.length));
    d.members.forEach((m, i) => { const n = 'S' + (i + 1); set(n + (i ? ' Name' : ' Name (Leader)'), safe_(m.name)); set(n + ' URN', m.urn); set(n + ' Contact', m.contact); });
    set('File Name', fname); set('File Link', file.getUrl()); set('File ID', file.getId()); set('Late', late ? 'YES' : 'NO'); set('Status', 'Submitted');
    s.sub.getRange(s.sub.getLastRow() + 1, 1, 1, HEADERS.length).setNumberFormat('@').setValues([row]);
    s.p.setProperty('SEQ_' + d.division, String(seq));
    CacheService.getScriptCache().remove('stats');
    return { id, late, queue, submittedOn: row[COL['Submitted On']], program: d.program, division: d.division, title: d.title,
             members: d.members.map(m => ({ name: m.name, urn: m.urn })), fileName: fname };
  } finally { lock.releaseLock(); }
}

function track_(id, urn) {
  id = clean_(id, 20).toUpperCase(); urn = clean_(urn, 25).toUpperCase();
  if (!id || !urn) throw new Error('Enter both Submission ID and URN.');
  for (const r of rows_()) {
    if (String(r[0]).toUpperCase() !== id) continue;
    const o = rowObj_(r, false);
    if (!o.members.some(m => m.urn.toUpperCase() === urn)) break;
    return { id: o.id, title: o.title, division: o.division, program: o.program, submittedOn: o.submittedOn, status: o.status,
             remarks: o.remarks, reviewedOn: o.reviewedOn, late: o.late, members: o.members.map(m => m.name) };
  }
  throw new Error('No submission found for this ID and URN.');
}

/* ================================================================ FACULTY: auth */

function facRows_() { const sh = db_().ss.getSheetByName(SH_FAC); return { sh, v: sh.getDataRange().getValues() }; }

function login_(username, password) {
  username = clean_(username, 40).toLowerCase(); password = String(password || '');
  const cache = CacheService.getScriptCache(), lk = 'f_' + username, fails = Number(cache.get(lk) || 0);
  if (fails >= CONFIG.MAX_LOGIN_FAILS) throw new Error('Account temporarily locked after repeated failed attempts. Try again in ' + CONFIG.LOCK_MINUTES + ' minutes.');
  Utilities.sleep(400);                                    // slow down guessing
  const { sh, v } = facRows_();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() !== username) continue;
    if (v[i][6] !== 'YES' || hash_(password, v[i][4]) !== v[i][5]) break;
    cache.remove(lk);
    const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
    const sess = { u: v[i][0], name: v[i][1], role: v[i][2], divs: String(v[i][3] || 'ALL') };
    cache.put('s_' + token, JSON.stringify(sess), CONFIG.SESSION_MINUTES * 60);
    sh.getRange(i + 1, 9).setValue(now_());
    log_(sess.u, 'LOGIN', 'success');
    return { token, user: sess, minutes: CONFIG.SESSION_MINUTES };
  }
  cache.put(lk, String(fails + 1), CONFIG.LOCK_MINUTES * 60);
  log_(username || '?', 'LOGIN_FAILED', 'attempt ' + (fails + 1));
  throw new Error('Invalid username or password.');
}

function auth_(r, adminOnly) {
  const cache = CacheService.getScriptCache(), raw = r && r.token && cache.get('s_' + r.token);
  if (!raw) throw new Error('SESSION_EXPIRED');
  cache.put('s_' + r.token, raw, CONFIG.SESSION_MINUTES * 60);   // sliding expiry
  const s = JSON.parse(raw);
  if (adminOnly && s.role !== 'admin') throw new Error('Only the lab coordinator (admin) can do this.');
  return s;
}
const inScope_ = (s, div) => s.divs === 'ALL' || s.divs.split(',').map(x => x.trim()).indexOf(div) >= 0;
function findRow_(s, id) {
  const rs = rows_();
  for (let i = 0; i < rs.length; i++) if (String(rs[i][0]) === id) {
    if (!inScope_(s, rs[i][COL['Division']])) throw new Error('This division is outside your access.');
    return { r: rs[i], rowNum: i + 2 };
  }
  throw new Error('Submission not found.');
}

/* ================================================================ FACULTY: data */

function dashboard_(s) {
  return { user: s, config: publicConfig_(), rows: rows_().filter(r => inScope_(s, r[COL['Division']])).map(r => rowObj_(r, false)) };
}
function review_(s, id, status, remarks) {
  if (CONFIG.STATUSES.indexOf(status) < 0) throw new Error('Invalid status.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const f = findRow_(s, String(id));
    const when = Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm');
    db_().sub.getRange(f.rowNum, COL['Status'] + 1, 1, 4).setValues([[status, safe_(clean_(remarks, 500)), when, s.name || s.u]]);
    log_(s.u, 'REVIEW', id + ' → ' + status);
    return { reviewedOn: when, reviewedBy: s.name || s.u };
  } finally { lock.releaseLock(); }
}
function reveal_(s, id) {
  const f = findRow_(s, String(id));
  log_(s.u, 'VIEW_CONTACTS', id);
  return rowObj_(f.r, true).members;
}
function file_(s, id) {
  const f = findRow_(s, String(id));
  const blob = DriveApp.getFileById(f.r[COL['File ID']]).getBlob();
  log_(s.u, 'DOWNLOAD_FILE', id);
  return { name: f.r[COL['File Name']], mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}
function export_(s) {
  const q = v => '"' + un_(v).replace(/"/g, '""') + '"';
  const keep = HEADERS.filter(h => h !== 'File ID');
  const csv = [keep.map(q).join(',')].concat(rows_().map(r => keep.map(h => q(r[COL[h]])).join(','))).join('\n');
  log_(s.u, 'EXPORT', 'full CSV export');
  return { filename: 'DSL_Submissions_' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HHmm') + '.csv', csv };
}
function logs_() {
  const sh = db_().ss.getSheetByName(SH_LOG), n = sh.getLastRow();
  if (n < 2) return [];
  const start = Math.max(2, n - 299);
  return sh.getRange(start, 1, n - start + 1, 4).getValues().reverse().map(r => ({ time: r[0], user: r[1], action: r[2], details: un_(r[3]) }));
}
function changePassword_(s, oldPw, newPw) {
  newPw = String(newPw || '');
  if (newPw.length < 8 || !/[A-Za-z]/.test(newPw) || !/\d/.test(newPw)) throw new Error('New password needs 8+ characters with letters and numbers.');
  const { sh, v } = facRows_();
  for (let i = 1; i < v.length; i++) if (v[i][0] === s.u) {
    if (hash_(String(oldPw || ''), v[i][4]) !== v[i][5]) throw new Error('Current password is incorrect.');
    const salt = Utilities.getUuid(); sh.getRange(i + 1, 5, 1, 2).setValues([[salt, hash_(newPw, salt)]]);
    log_(s.u, 'PASSWORD_CHANGED', ''); return true;
  }
  throw new Error('Account not found.');
}

/* ================================================================ ADMIN: faculty accounts */

function facultyList_() {
  return facRows_().v.slice(1).map(r => ({ username: r[0], name: r[1], role: r[2], divisions: r[3], active: r[6] === 'YES', created: r[7], lastLogin: r[8] }));
}
function facultySave_(s, a) {
  const username = clean_(a.username, 40).toLowerCase();
  if (!/^[a-z0-9._-]{3,40}$/.test(username)) throw new Error('Username: 3–40 characters (letters, numbers, . _ -).');
  const role = a.role === 'admin' ? 'admin' : 'faculty';
  let divs = clean_(a.divisions, 60).toUpperCase().replace(/\s+/g, '');
  if (divs !== 'ALL') {
    divs = divs.split(',').filter(x => CONFIG.DIVISIONS.indexOf(x) >= 0).join(',');
    if (!divs) throw new Error('Choose at least one division (or ALL).');
  }
  const pw = String(a.password || '');
  if (pw && (pw.length < 8 || !/[A-Za-z]/.test(pw) || !/\d/.test(pw))) throw new Error('Password needs 8+ characters with letters and numbers.');
  const { sh, v } = facRows_();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).toLowerCase() === username) {
    if (username === s.u && (role !== 'admin' || a.active === false)) throw new Error('You cannot remove your own admin access.');
    sh.getRange(i + 1, 2, 1, 3).setValues([[safe_(clean_(a.name, 80)), role, divs]]);
    sh.getRange(i + 1, 7).setValue(a.active === false ? 'NO' : 'YES');
    if (pw) { const salt = Utilities.getUuid(); sh.getRange(i + 1, 5, 1, 2).setValues([[salt, hash_(pw, salt)]]); }
    log_(s.u, 'FACULTY_UPDATED', username + ' role=' + role + ' div=' + divs + (pw ? ' (password reset)' : ''));
    return true;
  }
  if (!pw) throw new Error('Set a password for the new account.');
  const salt = Utilities.getUuid();
  sh.appendRow([username, safe_(clean_(a.name, 80)), role, divs, salt, hash_(pw, salt), 'YES', now_(), '']);
  log_(s.u, 'FACULTY_CREATED', username + ' role=' + role + ' div=' + divs);
  return true;
}
function facultyDelete_(s, username) {
  username = clean_(username, 40).toLowerCase();
  if (username === s.u) throw new Error('You cannot delete your own account.');
  const { sh, v } = facRows_();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).toLowerCase() === username) { sh.deleteRow(i + 1); log_(s.u, 'FACULTY_DELETED', username); return true; }
  throw new Error('Account not found.');
}
