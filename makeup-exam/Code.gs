/**
 * Make-up Midterm Exam Application  (student page + admin page)
 * Student page : <web app URL>
 * Admin page   : <web app URL>?page=admin   (password = Script Property ADMIN_PASSWORD)
 */
const DEFAULTS = {
  title: 'Make-up Midterm Exam Application',
  courses: ['CVLE211', 'CVLE222', 'CVLE372', 'CVLE472'],
  deadline: '',            // 'yyyy-MM-ddTHH:mm' in the script time zone; empty = no deadline
  open: true,              // master switch
  announcement: '',
  maxMb: 5,
  requireCode: false       // true = student must enter a code e-mailed to their roster e-mail
};
const HARD_MAX_MB = 5;
const FOLDER_NAME = 'MakeUp Exam Petitions';
const DB_NAME = 'MakeUp Exam Database';
const ALLOWED_TYPES = ['application/pdf', 'image/jpeg', 'image/png'];
const APP_HEADERS = ['Timestamp', 'Student Number', 'Name', 'Surname', 'Courses', 'No. of Courses',
                     'Petition File', 'Reference', 'Status', 'File ID', 'Withdrawn At'];
const ROSTER_HEADERS = ['Student Number', 'First Name', 'Surname', 'Course', 'Email'];

/* ===================== web entry ===================== */
function doGet(e) {
  const admin = e && e.parameter && e.parameter.page === 'admin';
  const t = HtmlService.createTemplateFromFile(admin ? 'Admin' : 'Index');
  t.info = JSON.stringify(publicInfo_()).replace(/</g, '\\u003c');
  return t.evaluate().setTitle(admin ? 'Admin' : getSettings_().title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ===================== settings ===================== */
function getSettings_() {
  let s = {};
  try { s = JSON.parse(PropertiesService.getScriptProperties().getProperty('SETTINGS') || '{}'); } catch (e) {}
  const o = Object.assign({}, DEFAULTS, s);
  o.maxMb = Math.min(Number(o.maxMb) || HARD_MAX_MB, HARD_MAX_MB);
  return o;
}
function deadlineMs_(s) {
  if (!s.deadline) return null;
  try { return Utilities.parseDate(s.deadline, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm").getTime(); }
  catch (e) { return null; }
}
function windowState_(s) {
  if (!s.open) return { open: false, reason: 'Applications are currently closed.' };
  const d = deadlineMs_(s);
  if (d !== null && Date.now() > d) return { open: false, reason: 'The deadline has passed. Applications are closed.' };
  return { open: true, reason: '' };
}
function publicInfo_() {
  const s = getSettings_(), d = deadlineMs_(s), w = windowState_(s), tz = Session.getScriptTimeZone();
  return {
    title: s.title, announcement: s.announcement, maxMb: s.maxMb, requireCode: s.requireCode,
    courses: s.courses, open: w.open, reason: w.reason, serverNow: Date.now(), deadlineMs: d,
    deadlineText: d === null ? '' : Utilities.formatDate(new Date(d), tz, "EEEE, d MMMM yyyy 'at' HH:mm") + ' (' + tz + ')'
  };
}

/* ===================== storage ===================== */
let _ss = null;
function db_() {
  if (_ss) return _ss;
  const p = PropertiesService.getScriptProperties();
  let ss, id = p.getProperty('DB_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) { ss = SpreadsheetApp.create(DB_NAME); p.setProperty('DB_ID', ss.getId()); ss.getSheets()[0].setName('Applications'); }
  initSheet_(ss, 'Applications', APP_HEADERS);
  initSheet_(ss, 'Roster', ROSTER_HEADERS);
  return (_ss = ss);
}
function initSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() > 0) return sh;
  sh.appendRow(headers);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8eefc');
  sh.getRange('A:B').setNumberFormat('@');          // student numbers stay text (keeps leading zeros)
  return sh;
}
function sheet_(n) { return db_().getSheetByName(n); }
function getFolder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}
function courseFolder_(code) {
  const root = getFolder_(), it = root.getFoldersByName(code);
  return it.hasNext() ? it.next() : root.createFolder(code);
}
/** Run once from the editor: authorizes and creates the Sheet + Folder. */
function setup() {
  Logger.log('Database: ' + db_().getUrl());
  Logger.log('Folder: ' + getFolder_().getUrl());
  Logger.log('Admin password set: ' + !!PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD'));
}

/* ===================== helpers ===================== */
function normNum_(v) { return String(v == null ? '' : v).trim().toUpperCase(); }
function validNum_(n) { return /^[A-Z0-9]{4,20}$/.test(n); }
function safe_(v) { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; }  // formula-injection guard
function fmt_(d) { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }

function rosterRows_() {
  const sh = sheet_('Roster'), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 5).getValues().map(r => ({
    num: normNum_(r[0]), first: String(r[1]), last: String(r[2]),
    course: String(r[3]).trim().toUpperCase(), email: String(r[4]).trim()
  }));
}
function eligible_(num, s) {
  const mine = rosterRows_().filter(r => r.num === num);
  const set = {}; let email = '';
  mine.forEach(r => { set[r.course] = true; if (!email && r.email) email = r.email; });
  return { courses: s.courses.filter(c => set[c]), email: email };
}
function allApps_() {
  const sh = sheet_('Applications'), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, APP_HEADERS.length).getValues().map((v, i) => ({
    row: i + 2, when: fmt_(v[0]), num: normNum_(v[1]), name: v[2] + ' ' + v[3], courses: String(v[4]),
    url: v[6], ref: v[7], status: v[8], fileId: v[9]
  })).reverse();
}
function findActive_(num) {
  const a = allApps_().filter(x => x.num === num && x.status === 'Active');
  return a.length ? a[0] : null;
}
function withdrawRow_(app) {
  const sh = sheet_('Applications');
  sh.getRange(app.row, 9).setValue('Withdrawn');
  sh.getRange(app.row, 11).setValue(new Date());
  String(app.fileId || '').split(',').forEach(id => { try { if (id) DriveApp.getFileById(id).setTrashed(true); } catch (e) {} });
}

/* ===================== student API ===================== */
function lookupStudent(numRaw) {
  const num = normNum_(numRaw);
  if (!validNum_(num)) throw new Error('Enter a valid student number.');
  const s = getSettings_(), w = windowState_(s);
  if (!w.open) throw new Error(w.reason);
  const el = eligible_(num, s);
  if (!el.courses.length) return { status: 'NOT_LISTED' };
  if (s.requireCode) return sendCode_(num, el.email);
  return studentState_(num, s, el, '');
}

function sendCode_(num, email) {
  if (!email) return { status: 'NO_EMAIL' };
  const c = CacheService.getScriptCache();
  if (c.get('cs_' + num)) throw new Error('A code was just sent. Wait 60 seconds before requesting another.');
  const code = ('' + (100000 + Math.floor(Math.random() * 900000)));
  c.put('code_' + num, code, 600); c.put('att_' + num, '0', 600); c.put('cs_' + num, '1', 60);
  MailApp.sendEmail(email, 'Make-up exam application - verification code',
    'Your verification code is: ' + code + '\nIt is valid for 10 minutes. If you did not request it, ignore this e-mail.');
  const parts = email.split('@');
  return { status: 'CODE_SENT', masked: parts[0].charAt(0) + '***@' + (parts[1] || '') };
}

function verifyCode(numRaw, code) {
  const num = normNum_(numRaw), c = CacheService.getScriptCache(), s = getSettings_();
  const att = Number(c.get('att_' + num) || 0), real = c.get('code_' + num);
  if (!real || att >= 5) throw new Error('Code expired or too many attempts. Request a new code.');
  if (String(code).trim() !== real) { c.put('att_' + num, String(att + 1), 600); throw new Error('Wrong code.'); }
  c.remove('code_' + num);
  const token = Utilities.getUuid();
  c.put('stu_' + num, token, 1800);
  return studentState_(num, s, eligible_(num, s), token);
}

function authStudent_(num, token, s) {
  if (s.requireCode && (!token || CacheService.getScriptCache().get('stu_' + num) !== token))
    throw new Error('Session expired. Please start again.');
}

function studentState_(num, s, el, token) {
  const app = findActive_(num), w = windowState_(s);
  return { status: 'READY', token: token, courses: el.courses, open: w.open, reason: w.reason,
           existing: app ? { ref: app.ref, courses: app.courses, when: app.when } : null };
}

function submitApplication(d) {
  const s = getSettings_(), w = windowState_(s);
  if (!w.open) throw new Error(w.reason);
  const num = normNum_(d.studentNumber);
  const first = String(d.firstName || '').trim(), last = String(d.lastName || '').trim();
  const courses = Array.isArray(d.courses) ? d.courses.filter((c, i, a) => a.indexOf(c) === i) : [];
  if (!validNum_(num)) throw new Error('Invalid student number.');
  authStudent_(num, d.token, s);
  if (first.length < 2 || last.length < 2 || first.length > 60 || last.length > 60) throw new Error('Invalid name or surname.');
  if (!courses.length) throw new Error('Select at least one course.');
  const el = eligible_(num, s);
  courses.forEach(c => { if (el.courses.indexOf(c) < 0) throw new Error('You are not on the eligible list for ' + c + '.'); });
  if (ALLOWED_TYPES.indexOf(d.mimeType) < 0) throw new Error('Petition must be PDF, JPG or PNG.');
  if (!d.base64) throw new Error('Petition file is missing.');
  const bytes = Utilities.base64Decode(d.base64);
  if (bytes.length > s.maxMb * 1024 * 1024) throw new Error('File is larger than ' + s.maxMb + ' MB.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (findActive_(num)) throw new Error('You already have an application. Remove it first to submit again.');
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const ext = { 'application/pdf': 'pdf', 'image/jpeg': 'jpg', 'image/png': 'png' }[d.mimeType];
    const ref = 'MU-' + stamp + '-' + num;
    const ids = [], urls = [];
    try {
      let master = null;
      courses.forEach((c, i) => {
        const name = num + '_' + c + '_Petition_' + stamp + '.' + ext;   // one copy in each course folder
        const f = i === 0 ? courseFolder_(c).createFile(Utilities.newBlob(bytes, d.mimeType, name))
                          : master.makeCopy(name, courseFolder_(c));
        if (i === 0) master = f;
        ids.push(f.getId()); urls.push(f.getUrl());
      });
    } catch (e) {
      ids.forEach(id => { try { DriveApp.getFileById(id).setTrashed(true); } catch (x) {} });
      throw new Error('Could not save your petition. Please try again.');
    }
    sheet_('Applications').appendRow([new Date(), num, safe_(first), safe_(last), courses.join('; '), courses.length,
                                      urls.join('\n'), ref, 'Active', ids.join(','), '']);
    return { ref: ref };
  } finally { lock.releaseLock(); }
}

function withdrawApplication(numRaw, token) {
  const s = getSettings_(), w = windowState_(s), num = normNum_(numRaw);
  if (!w.open) throw new Error(w.reason);
  if (!validNum_(num)) throw new Error('Invalid student number.');
  authStudent_(num, token, s);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const app = findActive_(num);
    if (!app) throw new Error('No active application found.');
    withdrawRow_(app);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ===================== admin API ===================== */
function adminLogin(pw) {
  const c = CacheService.getScriptCache();
  const fails = Number(c.get('adm_fail') || 0);
  if (fails >= 5) throw new Error('Too many wrong attempts. Try again in 15 minutes.');
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real) throw new Error('Admin password is not set (Project Settings > Script properties > ADMIN_PASSWORD).');
  if (String(pw) !== real) { c.put('adm_fail', String(fails + 1), 900); throw new Error('Wrong password.'); }
  const t = Utilities.getUuid();
  c.put('adm_' + t, '1', 21600);
  return t;
}
function adminLogout(t) {
  if (t) CacheService.getScriptCache().remove('adm_' + t);
  return true;
}
function adm_(t) {
  if (!t || !CacheService.getScriptCache().get('adm_' + t)) throw new Error('Session expired. Log in again.');
}

function adminOverview(t) {
  adm_(t);
  const rc = {}, ac = {};
  rosterRows_().forEach(r => rc[r.course] = (rc[r.course] || 0) + 1);
  const apps = allApps_();
  apps.filter(a => a.status === 'Active').forEach(a => a.courses.split('; ').forEach(c => ac[c] = (ac[c] || 0) + 1));
  return { settings: getSettings_(), tz: Session.getScriptTimeZone(), sheetUrl: db_().getUrl(),
           folderUrl: getFolder_().getUrl(), rosterCounts: rc, appCounts: ac,
           apps: apps.map(a => ({ ref: a.ref, num: a.num, name: a.name, courses: a.courses, when: a.when, status: a.status, url: a.url })) };
}

function adminSaveSettings(t, o) {
  adm_(t);
  const courses = [];
  String(o.courses || '').split(/[\s,;]+/).forEach(c => {
    c = c.trim().toUpperCase();
    if (c && courses.indexOf(c) < 0) { if (!/^[A-Z0-9]{3,15}$/.test(c)) throw new Error('Bad course code: ' + c); courses.push(c); }
  });
  if (!courses.length) throw new Error('Add at least one course.');
  const deadline = String(o.deadline || '');
  if (deadline && !/^\d{4}-\d\d-\d\dT\d\d:\d\d$/.test(deadline)) throw new Error('Bad deadline format.');
  const s = {
    title: String(o.title || DEFAULTS.title).slice(0, 120), courses: courses, deadline: deadline,
    open: !!o.open, announcement: String(o.announcement || '').slice(0, 1000),
    maxMb: Math.max(1, Math.min(HARD_MAX_MB, Number(o.maxMb) || HARD_MAX_MB)), requireCode: !!o.requireCode
  };
  PropertiesService.getScriptProperties().setProperty('SETTINGS', JSON.stringify(s));
  return s;
}

function adminWithdraw(t, ref) {
  adm_(t);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const app = allApps_().filter(a => a.ref === ref && a.status === 'Active')[0];
    if (!app) throw new Error('Application not found or already withdrawn.');
    withdrawRow_(app);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/** text = rows pasted from Excel/CSV. Column 1 = student number. If a "Course Name" column exists it is used,
 *  otherwise defaultCourse. mode 'replace' = overwrite the lists of the courses found in the paste; 'add' = append. */
function adminImportRoster(t, text, defaultCourse, mode) {
  adm_(t);
  const incoming = parseRoster_(text, String(defaultCourse || '').trim().toUpperCase());
  if (!incoming.length) throw new Error('No valid rows found. Column 1 must be the student number; pick a course if there is no "Course Name" column.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = sheet_('Roster'), n = sh.getLastRow();
    let rows = n > 1 ? sh.getRange(2, 1, n - 1, 5).getValues() : [];
    const touched = {};
    incoming.forEach(r => touched[r[3]] = (touched[r[3]] || 0) + 1);
    if (mode === 'replace') rows = rows.filter(r => !touched[String(r[3]).trim().toUpperCase()]);
    const seen = {};
    rows.forEach(r => seen[normNum_(r[0]) + '|' + String(r[3]).trim().toUpperCase()] = true);
    incoming.forEach(r => { const k = r[0] + '|' + r[3]; if (!seen[k]) { seen[k] = true; rows.push(r); } });
    if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
    if (sh.getMaxRows() > 1) sh.getRange(2, 1, sh.getMaxRows() - 1, 5).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);
    const known = getSettings_().courses;
    return { imported: touched, total: rows.length,
             unknown: Object.keys(touched).filter(c => known.indexOf(c) < 0) };
  } finally { lock.releaseLock(); }
}

function parseRoster_(text, defaultCourse) {
  const lines = String(text || '').split(/\r?\n/).filter(l => l.trim());
  if (!lines.length) return [];
  const h = lines[0];
  const delim = h.indexOf('\t') >= 0 ? '\t' : (h.split(';').length > h.split(',').length ? ';' : ',');
  let rows = lines.map(l => l.split(delim).map(c => c.replace(/^"|"$/g, '').trim()));
  const idx = { first: 1, last: 2, course: -1, email: -1 };
  if (!/\d/.test(rows[0][0])) {                       // header row present
    const head = rows[0].map(c => c.toLowerCase());
    idx.course = head.findIndex(c => c === 'course name' || c === 'course' || c === 'course code');
    idx.email = head.findIndex(c => c === 'e-mail' || c === 'email');
    const f = head.indexOf('first name'), l = head.indexOf('surname');
    if (f >= 0) idx.first = f; if (l >= 0) idx.last = l;
    rows = rows.slice(1);
  }
  const out = [];
  rows.forEach(r => {
    const num = normNum_(r[0]);
    const course = (idx.course >= 0 ? String(r[idx.course] || '') : defaultCourse).trim().toUpperCase();
    if (!validNum_(num) || !course) return;
    const email = idx.email >= 0 ? (r[idx.email] || '') : (r.filter(c => c.indexOf('@') > 0)[0] || '');
    out.push([num, safe_(r[idx.first] || ''), safe_(r[idx.last] || ''), course, email]);
  });
  return out;
}
