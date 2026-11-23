/** ===== SETTINGS: edit this part ===== */
const CONFIG = {
  TITLE: 'Make-up Midterm Exam Application',
  FOLDER_NAME: 'MakeUp Exam Petitions',        // created automatically in your Drive
  SHEET_NAME: 'MakeUp Exam Applications',      // created automatically in your Drive
  MAX_FILE_MB: 5,
  COURSES: [                                   // <-- put your real courses here
    'MATH101 - Calculus I',
    'PHYS101 - Physics I',
    'CENG201 - Data Structures',
    'ENGR101 - Intro to Engineering'
  ]
};
const ALLOWED_TYPES = ['application/pdf', 'image/jpeg', 'image/png'];

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.courses = CONFIG.COURSES;
  t.title = CONFIG.TITLE;
  t.maxMb = CONFIG.MAX_FILE_MB;
  return t.evaluate()
    .setTitle(CONFIG.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Called from the page. Validates again on the server (never trust the browser). */
function submitApplication(d) {
  const num = String(d.studentNumber || '').trim();
  const first = String(d.firstName || '').trim();
  const last = String(d.lastName || '').trim();
  const courses = Array.isArray(d.courses) ? d.courses : [];

  if (!/^[A-Za-z0-9]{4,20}$/.test(num)) throw new Error('Invalid student number.');
  if (first.length < 2 || last.length < 2 || first.length > 60 || last.length > 60) throw new Error('Invalid name or surname.');
  if (!courses.length) throw new Error('Select at least one course.');
  courses.forEach(c => { if (CONFIG.COURSES.indexOf(c) < 0) throw new Error('Unknown course: ' + c); });
  if (ALLOWED_TYPES.indexOf(d.mimeType) < 0) throw new Error('Petition must be PDF, JPG or PNG.');
  if (!d.base64) throw new Error('Petition file is missing.');

  const bytes = Utilities.base64Decode(d.base64);
  if (bytes.length > CONFIG.MAX_FILE_MB * 1024 * 1024) throw new Error('File is larger than ' + CONFIG.MAX_FILE_MB + ' MB.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const ext = { 'application/pdf': 'pdf', 'image/jpeg': 'jpg', 'image/png': 'png' }[d.mimeType];
    const fileName = num + '_Petition_' + stamp + '.' + ext;     // student number in the file name

    const file = getFolder_().createFile(Utilities.newBlob(bytes, d.mimeType, fileName));
    const ref = 'MU-' + stamp + '-' + num;

    getSheet_().appendRow([
      new Date(), safe_(num), safe_(first), safe_(last),
      courses.join('; '), courses.length, file.getUrl(), ref
    ]);
    return { ref: ref, fileName: fileName };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- helpers ---------- */
function safe_(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }  // blocks spreadsheet formula injection

function getFolder_() {
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(CONFIG.FOLDER_NAME);
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}

function getSheet_() {
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('SHEET_ID');
  let ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.SHEET_NAME);
    p.setProperty('SHEET_ID', ss.getId());
    const sh = ss.getSheets()[0];
    sh.appendRow(['Timestamp', 'Student Number', 'Name', 'Surname', 'Courses', 'No. of Courses', 'Petition File', 'Reference']);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold').setBackground('#e8eefc');
    sh.getRange('B:B').setNumberFormat('@');     // keep leading zeros in student numbers
  }
  return ss.getSheets()[0];
}

/** Run this ONCE from the editor to authorize and create the Sheet + Folder; check the log for links. */
function setup() {
  Logger.log('Folder: ' + getFolder_().getUrl());
  Logger.log('Sheet: ' + getSheet_().getParent().getUrl());
}
