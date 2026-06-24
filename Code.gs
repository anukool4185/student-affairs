/**
 * ระบบจัดเก็บข้อมูลและรูปภาพนักเรียน (v3)
 * Backend: Google Apps Script
 *
 * เวิร์กโฟลว์รูปภาพ:
 *   1) แอดมิน/ครู อัปโหลดรูป -> เก็บในโฟลเดอร์ "ชั่วคราว" (สถานะ: pending) ยังไม่จัดหมวดวันที่/ชั้น/ห้อง
 *   2) แอดมินเข้าแท็บ "จัดการส่งข้อมูล" เลือกรายการที่จะเผยแพร่ + เลือกวันที่ (หรือไม่เลือกก็ได้) แล้วกด "เผยแพร่"
 *      -> ระบบย้ายไฟล์ไปโฟลเดอร์จริง: [วันที่ถ้ามี]/[ระดับชั้น]/ห้อง [ห้อง]/รหัส.ext (สถานะ: published)
 *
 * สิทธิ์:
 *   - แอดมิน      : ทำทุกอย่างได้ทันที (เพิ่ม/อัปโหลด/ยืนยันรหัส/ย้ายห้อง/ลบ/เผยแพร่) + จัดการบัญชีครู/บริษัท
 *   - ครูที่ปรึกษา : เพิ่ม/อัปโหลด/ยืนยันรหัส ได้ทันที, ย้ายห้อง/ลบ ต้อง "แจ้งคำขอ" ให้แอดมินอนุมัติก่อน
 *   - นักเรียน    : login = รหัสนักเรียนปัจจุบัน / password = เบอร์โทร เห็น+แก้ "เบอร์โทร" ของตัวเองได้เท่านั้น
 *                   (เปลี่ยนรูปไม่ได้ แก้ชั้น/ห้อง/รหัสไม่ได้)
 *   - บริษัท      : ค้นหา (กรองวันที่/ระดับชั้น/ห้อง) ติ๊กเลือกรายการ แล้วดาวน์โหลด ZIP (รูป+CSV) เฉพาะที่เลือก
 *
 * อ้างอิงนักเรียนด้วย RecordId (UUID) ที่ไม่เปลี่ยนแปลง ไม่ใช้เลขแถว
 * เพื่อให้ระบบคำขอ/อนุมัติถูกต้อง แม้มีการลบ/แก้ไขแถวอื่นแทรกกลางคัน
 */

// ========================= CONFIG =========================
const SHEET_NAME = 'นักเรียน';
const USERS_SHEET_NAME = 'ผู้ใช้งานระบบ';
const REQUESTS_SHEET_NAME = 'คำขอ';
const ROOT_FOLDER_NAME = 'รูปภาพนักเรียน';
const TEMP_FOLDER_NAME = 'รูปภาพนักเรียน_รอเผยแพร่';
const SESSION_TTL_SECONDS = 21600; // 6 ชั่วโมง

const DEFAULT_USERS = [
  { role: 'admin', username: 'admin', password: 'admin123', displayName: 'ผู้ดูแลระบบ' },
  { role: 'teacher', username: 'tadmin', password: 'tadmin1', displayName: 'ครูที่ปรึกษา' },
  { role: 'company', username: 'company', password: 'company1', displayName: 'บริษัทภายนอก' }
];

const COLUMNS = [
  'RecordId', 'รหัสนักเรียนจริง', 'รหัสชั่วคราว', 'คำนำหน้า', 'ชื่อ', 'นามสกุล',
  'ระดับชั้น', 'ห้อง', 'เลขที่', 'เบอร์โทรนักเรียน',
  'ชื่อไฟล์รูปปัจจุบัน', 'DriveFileId', 'สถานะรูป', 'วันที่ของรูป(ถ้ามี)',
  'วันที่นำเข้า', 'หมายเหตุ'
];
const COL = {
  RECORD_ID: 0, REAL_ID: 1, TEMP_ID: 2, PREFIX: 3, FNAME: 4, LNAME: 5,
  GRADE: 6, ROOM: 7, NO: 8, PHONE: 9,
  PHOTO_FILE: 10, FILE_ID: 11, PHOTO_STATUS: 12, PHOTO_DATE: 13,
  IMPORT_DATE: 14, NOTE: 15
};
// PHOTO_STATUS: '' (ไม่มีรูป) | 'pending' (รอเผยแพร่) | 'published' (เผยแพร่แล้ว)

// ========================= WEB APP ENTRY =========================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบจัดเก็บข้อมูลและรูปภาพนักเรียน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ========================= AUTH: PASSWORD HASHING =========================
function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password) + '::' + String(salt));
  return bytes.map(function (b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

// ========================= USERS SHEET =========================
function getUsersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.appendRow(['บทบาท', 'Username', 'PasswordHash', 'Salt', 'ชื่อที่แสดง', 'แก้ไขล่าสุด']);
    sheet.setFrozenRows(1);
    DEFAULT_USERS.forEach(function (u) {
      const salt = Utilities.getUuid();
      const hash = hashPassword_(u.password, salt);
      sheet.appendRow([u.role, u.username, hash, salt, u.displayName, new Date()]);
    });
  }
  return sheet;
}

// ========================= RATE LIMIT =========================
function checkRateLimit_(username) {
  const cache = CacheService.getScriptCache();
  const count = parseInt(cache.get('fail_' + username) || '0', 10);
  if (count >= 5) throw new Error('พยายามเข้าสู่ระบบผิดหลายครั้งเกินไป กรุณารอประมาณ 5 นาทีแล้วลองใหม่');
}
function recordFailedLogin_(username) {
  const cache = CacheService.getScriptCache();
  const count = parseInt(cache.get('fail_' + username) || '0', 10) + 1;
  cache.put('fail_' + username, String(count), 300);
}
function clearFailedLogin_(username) {
  CacheService.getScriptCache().remove('fail_' + username);
}

// ========================= SESSION =========================
function createSession_(role, username, studentId, displayName) {
  const token = Utilities.getUuid();
  const session = { role: role, username: username, studentId: studentId || null, displayName: displayName || username, createdAt: Date.now() };
  CacheService.getScriptCache().put(token, JSON.stringify(session), SESSION_TTL_SECONDS);
  return { token: token, role: role, displayName: session.displayName, studentId: session.studentId };
}
function validateSession_(token) {
  if (!token) throw new Error('กรุณาเข้าสู่ระบบ');
  const raw = CacheService.getScriptCache().get(token);
  if (!raw) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  return JSON.parse(raw);
}
function requireRole_(session, roles) {
  if (roles.indexOf(session.role) === -1) throw new Error('คุณไม่มีสิทธิ์ทำรายการนี้');
}

// ========================= LOGIN / LOGOUT =========================
function login(username, password) {
  username = (username || '').toString().trim();
  password = (password || '').toString();
  if (!username || !password) throw new Error('กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');
  checkRateLimit_(username);

  const studentSheet = getSheet();
  const last = studentSheet.getLastRow();
  if (last >= 2) {
    const data = studentSheet.getRange(2, 1, last - 1, COLUMNS.length).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const code = activeCode_(row[COL.REAL_ID], row[COL.TEMP_ID]);
      const phone = String(row[COL.PHONE] || '').trim();
      if (code && code === username) {
        if (phone && phone === password) {
          clearFailedLogin_(username);
          const fullName = (row[COL.PREFIX] + ' ' + row[COL.FNAME] + ' ' + row[COL.LNAME]).trim();
          return createSession_('student', username, row[COL.RECORD_ID], fullName);
        }
        recordFailedLogin_(username);
        throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
      }
    }
  }

  const usersSheet = getUsersSheet_();
  const lastU = usersSheet.getLastRow();
  if (lastU >= 2) {
    const udata = usersSheet.getRange(2, 1, lastU - 1, 6).getValues();
    for (let i = 0; i < udata.length; i++) {
      const role = udata[i][0], uname = udata[i][1], hash = udata[i][2], salt = udata[i][3], displayName = udata[i][4];
      if (uname === username) {
        const computed = hashPassword_(password, salt);
        if (computed === hash) {
          clearFailedLogin_(username);
          return createSession_(role, username, null, displayName);
        }
        recordFailedLogin_(username);
        throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
      }
    }
  }

  recordFailedLogin_(username);
  throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
}
function logout(token) {
  if (token) CacheService.getScriptCache().remove(token);
  return { ok: true };
}

// ========================= ACCOUNT MANAGEMENT (admin/teacher/company) =========================
function changeOwnPassword(token, oldPassword, newPassword) {
  const session = validateSession_(token);
  requireRole_(session, ['admin', 'teacher', 'company']);
  newPassword = (newPassword || '').toString();
  if (newPassword.length < 4) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 4 ตัวอักษร');
  const sheet = getUsersSheet_();
  const last = sheet.getLastRow();
  const data = sheet.getRange(2, 1, last - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === session.username) {
      const computed = hashPassword_(oldPassword, data[i][3]);
      if (computed !== data[i][2]) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
      const newSalt = Utilities.getUuid();
      const newHash = hashPassword_(newPassword, newSalt);
      sheet.getRange(i + 2, 3, 1, 2).setValues([[newHash, newSalt]]);
      sheet.getRange(i + 2, 6).setValue(new Date());
      return { ok: true };
    }
  }
  throw new Error('ไม่พบบัญชีผู้ใช้');
}

function adminListUsers(token) {
  const session = validateSession_(token);
  requireRole_(session, ['admin']);
  const sheet = getUsersSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 6).getValues().map(function (r) {
    return { role: r[0], username: r[1], displayName: r[4] };
  });
}

function adminUpdateCredential(token, targetUsername, newUsername, newPassword) {
  const session = validateSession_(token);
  requireRole_(session, ['admin']);
  const sheet = getUsersSheet_();
  const last = sheet.getLastRow();
  const data = sheet.getRange(2, 1, last - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === targetUsername) {
      if (newUsername && newUsername.trim()) sheet.getRange(i + 2, 2).setValue(newUsername.trim());
      if (newPassword && newPassword.trim()) {
        const newSalt = Utilities.getUuid();
        const newHash = hashPassword_(newPassword.trim(), newSalt);
        sheet.getRange(i + 2, 3, 1, 2).setValues([[newHash, newSalt]]);
      }
      sheet.getRange(i + 2, 6).setValue(new Date());
      return { ok: true };
    }
  }
  throw new Error('ไม่พบบัญชีผู้ใช้');
}

function adminAddUser(token, role, username, password, displayName) {
  const session = validateSession_(token);
  requireRole_(session, ['admin']);
  role = (role || '').trim();
  username = (username || '').trim();
  password = (password || '').toString();
  if (['admin', 'teacher', 'company'].indexOf(role) === -1) throw new Error('บทบาทไม่ถูกต้อง');
  if (!username || !password) throw new Error('กรุณากรอก username และ password');
  if (password.length < 4) throw new Error('รหัสผ่านต้องมีอย่างน้อย 4 ตัวอักษร');
  const sheet = getUsersSheet_();
  const last = sheet.getLastRow();
  const data = last >= 2 ? sheet.getRange(2, 1, last - 1, 6).getValues() : [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === username) throw new Error('Username นี้มีอยู่แล้ว');
  }
  const salt = Utilities.getUuid();
  const hash = hashPassword_(password, salt);
  sheet.appendRow([role, username, hash, salt, displayName || username, new Date()]);
  return { ok: true };
}

function adminDeleteUser(token, username) {
  const session = validateSession_(token);
  requireRole_(session, ['admin']);
  if (username === session.username) throw new Error('ไม่สามารถลบบัญชีของตัวเองได้');
  const sheet = getUsersSheet_();
  const last = sheet.getLastRow();
  const data = sheet.getRange(2, 1, last - 1, 6).getValues();
  let adminCount = 0;
  data.forEach(function (r) { if (r[0] === 'admin') adminCount++; });
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === username) {
      if (data[i][0] === 'admin' && adminCount <= 1) throw new Error('ต้องมีบัญชีแอดมินเหลืออย่างน้อย 1 บัญชี');
      sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  throw new Error('ไม่พบบัญชีผู้ใช้');
}

// ========================= STUDENT SHEET HELPERS =========================
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
  }
  return sheet;
}

function findRowByRecordId_(id) {
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last < 2) throw new Error('ไม่พบข้อมูลนักเรียน');
  const ids = sheet.getRange(2, COL.RECORD_ID + 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2;
  }
  throw new Error('ไม่พบข้อมูลนักเรียน (อาจถูกลบไปแล้ว)');
}

function activeCode_(realId, tempId) {
  realId = (realId || '').toString().trim();
  tempId = (tempId || '').toString().trim();
  return realId !== '' ? realId : tempId;
}
function sanitizeFileCode_(code) {
  return String(code).replace(/[\\/:*?"<>|]/g, '-').trim();
}

// ========================= DRIVE HELPERS =========================
function getRootFolder_() {
  const folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(ROOT_FOLDER_NAME);
}
function getTempFolder_() {
  const folders = DriveApp.getFoldersByName(TEMP_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(TEMP_FOLDER_NAME);
}
function getOrCreateSubFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}
/** ระบุ dateStr -> รูปภาพนักเรียน/[dateStr]/[grade]/ห้อง [room] | ไม่ระบุ -> รูปภาพนักเรียน/[grade]/ห้อง [room] */
function getClassFolder_(grade, room, dateStr) {
  const root = getRootFolder_();
  let base = root;
  if (dateStr && String(dateStr).trim()) base = getOrCreateSubFolder_(root, String(dateStr).trim());
  const gradeFolder = getOrCreateSubFolder_(base, grade);
  return getOrCreateSubFolder_(gradeFolder, 'ห้อง ' + room);
}

function generateTempId() {
  const sheet = getSheet();
  const last = sheet.getLastRow();
  let max = 0;
  if (last >= 2) {
    sheet.getRange(2, COL.TEMP_ID + 1, last - 1, 1).getValues().forEach(function (r) {
      const m = String(r[0]).match(/^TMP-(\d+)$/);
      if (m) { const n = parseInt(m[1], 10); if (n > max) max = n; }
    });
  }
  return 'TMP-' + String(max + 1).padStart(5, '0');
}

function rowToObject_(row) {
  return {
    id: row[COL.RECORD_ID],
    realId: row[COL.REAL_ID] || '',
    tempId: row[COL.TEMP_ID] || '',
    prefix: row[COL.PREFIX] || '',
    firstName: row[COL.FNAME] || '',
    lastName: row[COL.LNAME] || '',
    grade: row[COL.GRADE] || '',
    room: row[COL.ROOM] || '',
    no: row[COL.NO] || '',
    phone: row[COL.PHONE] || '',
    photoFile: row[COL.PHOTO_FILE] || '',
    fileId: row[COL.FILE_ID] || '',
    photoStatus: row[COL.PHOTO_STATUS] || '',
    photoDate: row[COL.PHOTO_DATE] || '',
    importDate: row[COL.IMPORT_DATE] ? new Date(row[COL.IMPORT_DATE]).toLocaleDateString('th-TH') : '',
    note: row[COL.NOTE] || '',
    activeCode: activeCode_(row[COL.REAL_ID], row[COL.TEMP_ID]),
    photoUrl: row[COL.FILE_ID] ? ('https://drive.google.com/thumbnail?id=' + row[COL.FILE_ID] + '&sz=w200') : ''
  };
}

// ========================= STUDENT CRUD =========================
function listStudents(token, filter) {
  const session = validateSession_(token);
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, 1, last - 1, COLUMNS.length).getValues();

  if (session.role === 'student') {
    const row = data.filter(function (r) { return r[COL.RECORD_ID] === session.studentId; })[0];
    return row ? [rowToObject_(row)] : [];
  }

  requireRole_(session, ['admin', 'teacher']);
  const result = [];
  data.forEach(function (row) {
    if (filter) {
      if (filter.grade && row[COL.GRADE] !== filter.grade) return;
      if (filter.room && String(row[COL.ROOM]) !== String(filter.room)) return;
      if (filter.onlyTemp && row[COL.REAL_ID]) return;
      if (filter.onlyNoPhoto && row[COL.PHOTO_STATUS]) return;
      if (filter.query) {
        const q = filter.query.toString().toLowerCase();
        const hay = [row[COL.FNAME], row[COL.LNAME], row[COL.REAL_ID], row[COL.TEMP_ID]].join(' ').toLowerCase();
        if (hay.indexOf(q) === -1) return;
      }
    }
    result.push(rowToObject_(row));
  });
  return result;
}

function getMyInfo(token) {
  const session = validateSession_(token);
  requireRole_(session, ['student']);
  const row = findRowByRecordId_(session.studentId);
  const sheet = getSheet();
  return rowToObject_(sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0]);
}

function addStudent(token, info) {
  const session = validateSession_(token);
  requireRole_(session, ['admin', 'teacher']);
  if (!info || !info.firstName || !info.grade || !info.room) {
    throw new Error('กรุณาระบุ ชื่อ, ระดับชั้น และ ห้อง ให้ครบ');
  }
  const sheet = getSheet();
  const tempId = generateTempId();
  const recordId = Utilities.getUuid();
  sheet.appendRow([
    recordId, (info.realId || '').trim(), tempId, info.prefix || '', info.firstName || '', info.lastName || '',
    info.grade || '', info.room || '', info.no || '', info.phone || '',
    '', '', '', '', new Date(), info.note || ''
  ]);
  return { id: recordId, tempId: tempId };
}

/**
 * info: { prefix, firstName, lastName, no, phone, note }
 * student: แก้ได้แค่ของตัวเอง และแค่ phone/note
 */
function updateStudentInfo(token, id, info) {
  const session = validateSession_(token);
  info = info || {};
  const sheet = getSheet();

  if (session.role === 'student') {
    if (session.studentId !== id) throw new Error('ไม่มีสิทธิ์แก้ไขข้อมูลของผู้อื่น');
    const row = findRowByRecordId_(id);
    if (typeof info.phone !== 'undefined') sheet.getRange(row, COL.PHONE + 1).setValue(info.phone);
    if (typeof info.note !== 'undefined') sheet.getRange(row, COL.NOTE + 1).setValue(info.note);
    return { ok: true };
  }

  requireRole_(session, ['admin', 'teacher']);
  const row = findRowByRecordId_(id);
  sheet.getRange(row, COL.PREFIX + 1, 1, 4).setValues([[info.prefix || '', info.firstName || '', info.lastName || '', info.no || '']]);
  if (typeof info.phone !== 'undefined') sheet.getRange(row, COL.PHONE + 1).setValue(info.phone || '');
  if (typeof info.note !== 'undefined') sheet.getRange(row, COL.NOTE + 1).setValue(info.note || '');
  return { ok: true };
}

/** อัปโหลดรูป -> เก็บในโฟลเดอร์ชั่วคราวเสมอ (สถานะ pending) รอแอดมินเผยแพร่ — admin/teacher เท่านั้น */
function uploadPhotoForStudent(token, id, base64Data, mimeType, originalName) {
  const session = validateSession_(token);
  requireRole_(session, ['admin', 'teacher']);

  const row = findRowByRecordId_(id);
  const sheet = getSheet();
  const rowVals = sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0];
  const oldFileId = rowVals[COL.FILE_ID];

  const code = sanitizeFileCode_(activeCode_(rowVals[COL.REAL_ID], rowVals[COL.TEMP_ID]));
  const folder = getTempFolder_();

  if (oldFileId) { try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e) { /* ignore */ } }

  const extMatch = String(originalName || '').match(/\.[^.]+$/);
  const ext = extMatch ? extMatch[0] : '.jpg';
  const newFileName = code + ext;
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType, newFileName);
  const file = folder.createFile(blob);
  file.setName(newFileName);

  sheet.getRange(row, COL.PHOTO_FILE + 1, 1, 4).setValues([[newFileName, file.getId(), 'pending', '']]);

  return {
    fileName: newFileName, fileId: file.getId(), status: 'pending',
    photoUrl: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w200'
  };
}

/** ยืนยันรหัสจริง -> เปลี่ยนชื่อไฟล์เดิม (ตำแหน่งเดิม ไม่ว่าจะอยู่โฟลเดอร์ชั่วคราวหรือเผยแพร่แล้ว) */
function confirmRealId(token, id, realId) {
  const session = validateSession_(token);
  requireRole_(session, ['admin', 'teacher']);
  realId = (realId || '').toString().trim();
  if (!realId) throw new Error('กรุณาระบุรหัสนักเรียนจริง');

  const row = findRowByRecordId_(id);
  const sheet = getSheet();
  const rowVals = sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0];
  const fileId = rowVals[COL.FILE_ID];

  sheet.getRange(row, COL.REAL_ID + 1).setValue(realId);

  if (fileId) {
    try {
      const file = DriveApp.getFileById(fileId);
      const extMatch = file.getName().match(/\.[^.]+$/);
      const ext = extMatch ? extMatch[0] : '.jpg';
      const newName = sanitizeFileCode_(realId) + ext;
      file.setName(newName);
      sheet.getRange(row, COL.PHOTO_FILE + 1).setValue(newName);
    } catch (e) { /* ไฟล์อาจถูกลบไปแล้ว */ }
  }
  return { ok: true };
}

// ========================= PUBLISH (admin only) =========================
function listPendingPhotos(token) {
  const session = validateSession_(token);
  requireRole_(session, ['admin']);
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, 1, last - 1, COLUMNS.length).getValues();
  const result = [];
  data.forEach(function (row) {
    if (row[COL.PHOTO_STATUS] === 'pending') result.push(rowToObject_(row));
  });
  return result;
}

/** เผยแพร่รายการที่เลือก: ย้ายไฟล์จากโฟลเดอร์ชั่วคราว ไปโฟลเดอร์ [วันที่ถ้ามี]/ระดับชั้น/ห้อง */
function publishPhotos(token, ids, date) {
  const session = validateSession_(token);
  requireRole_(session, ['admin']);
  if (!ids || !ids.length) throw new Error('กรุณาเลือกรายการที่ต้องการเผยแพร่อย่างน้อย 1 รายการ');
  const dateStr = (date || '').toString().trim();
  const sheet = getSheet();
  let count = 0;

  ids.forEach(function (id) {
    try {
      const row = findRowByRecordId_(id);
      const rowVals = sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0];
      const fileId = rowVals[COL.FILE_ID];
      const grade = rowVals[COL.GRADE], room = rowVals[COL.ROOM];
      if (!fileId || !grade || !room) return;

      const file = DriveApp.getFileById(fileId);
      const folder = getClassFolder_(grade, room, dateStr);
      const parentIds = [];
      const parents = file.getParents();
      while (parents.hasNext()) parentIds.push(parents.next().getId());
      folder.addFile(file);
      parentIds.forEach(function (pid) {
        if (pid !== folder.getId()) { try { DriveApp.getFolderById(pid).removeFile(file); } catch (e) { /* ignore */ } }
      });

      sheet.getRange(row, COL.PHOTO_STATUS + 1, 1, 2).setValues([['published', dateStr]]);
      count++;
    } catch (e) { /* ข้ามรายการที่มีปัญหา ทำรายการอื่นต่อ */ }
  });

  return { ok: true, count: count };
}

// ========================= MOVE / DELETE (internal, ใช้ทั้งทางตรงและผ่านคำขอ) =========================
function moveStudentClassInternal_(id, newGrade, newRoom, photoDateOverride) {
  const row = findRowByRecordId_(id);
  const sheet = getSheet();
  const rowVals = sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0];
  const fileId = rowVals[COL.FILE_ID];
  const status = rowVals[COL.PHOTO_STATUS];
  let photoDate = String(rowVals[COL.PHOTO_DATE] || '').trim();
  if (photoDateOverride !== undefined && photoDateOverride !== null && String(photoDateOverride).trim() !== '') {
    photoDate = String(photoDateOverride).trim();
  }

  if (fileId && status === 'published') {
    try {
      const file = DriveApp.getFileById(fileId);
      const newFolder = getClassFolder_(newGrade, newRoom, photoDate);
      const parentIds = [];
      const parents = file.getParents();
      while (parents.hasNext()) parentIds.push(parents.next().getId());
      newFolder.addFile(file);
      parentIds.forEach(function (pid) {
        if (pid !== newFolder.getId()) { try { DriveApp.getFolderById(pid).removeFile(file); } catch (e) { /* ignore */ } }
      });
      sheet.getRange(row, COL.PHOTO_DATE + 1).setValue(photoDate);
    } catch (e) { /* ไฟล์อาจถูกลบไปแล้ว ยังอัปเดตชีทต่อไปได้ */ }
  }
  // ถ้ารูปยังอยู่โฟลเดอร์ชั่วคราว (pending) หรือยังไม่มีรูป ไม่ต้องย้ายไฟล์ แค่เปลี่ยนชั้น/ห้องในชีท

  sheet.getRange(row, COL.GRADE + 1, 1, 2).setValues([[newGrade, newRoom]]);
}

function deleteStudentInternal_(id) {
  const row = findRowByRecordId_(id);
  const sheet = getSheet();
  const rowVals = sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0];
  const fileId = rowVals[COL.FILE_ID];
  if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* ignore */ } }
  sheet.deleteRow(row);
}

/** ย้ายห้องทันที — admin เท่านั้น (ครูต้องใช้ teacherRequestMove แทน) */
function moveStudentClass(token, id, newGrade, newRoom, photoDateOverride) {
  const session = validateSession_(token);
  requireRole_(session, ['admin']);
  if (!newGrade || !newRoom) throw new Error('กรุณาระบุระดับชั้นและห้องใหม่');
  moveStudentClassInternal_(id, newGrade, newRoom, photoDateOverride);
  return { ok: true };
}

/** ลบทันที — admin เท่านั้น (ครูต้องใช้ teacherRequestDelete แทน) */
function deleteStudent(token, id) {
  const session = validateSession_(token);
  requireRole_(session, ['admin']);
  deleteStudentInternal_(id);
  return { ok: true };
}

// ========================= คำขอจากครู (ย้ายห้อง / ลบ) =========================
function getRequestsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(REQUESTS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(REQUESTS_SHEET_NAME);
    sheet.appendRow(['ประเภท', 'StudentRecordId', 'ชื่อนักเรียน', 'รายละเอียด', 'ผู้ขอ', 'วันที่ขอ', 'สถานะ']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function teacherRequestMove(token, id, newGrade, newRoom) {
  const session = validateSession_(token);
  requireRole_(session, ['teacher', 'admin']);
  if (!newGrade || !newRoom) throw new Error('กรุณาระบุระดับชั้นและห้องใหม่');
  const row = findRowByRecordId_(id);
  const sheet = getSheet();
  const rowVals = sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0];
  const name = (rowVals[COL.PREFIX] + ' ' + rowVals[COL.FNAME] + ' ' + rowVals[COL.LNAME]).trim();
  const details = JSON.stringify({ newGrade: newGrade, newRoom: newRoom });
  getRequestsSheet_().appendRow(['ย้ายห้อง', id, name, details, session.username, new Date(), 'รออนุมัติ']);
  return { ok: true };
}

function teacherRequestDelete(token, id) {
  const session = validateSession_(token);
  requireRole_(session, ['teacher', 'admin']);
  const row = findRowByRecordId_(id);
  const sheet = getSheet();
  const rowVals = sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0];
  const name = (rowVals[COL.PREFIX] + ' ' + rowVals[COL.FNAME] + ' ' + rowVals[COL.LNAME]).trim();
  getRequestsSheet_().appendRow(['ลบ', id, name, '', session.username, new Date(), 'รออนุมัติ']);
  return { ok: true };
}

function adminListRequests(token) {
  const session = validateSession_(token);
  requireRole_(session, ['admin']);
  const sheet = getRequestsSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, 1, last - 1, 7).getValues();
  const result = [];
  data.forEach(function (r, idx) {
    if (r[6] !== 'รออนุมัติ') return;
    result.push({
      requestRow: idx + 2,
      type: r[0],
      studentId: r[1],
      studentName: r[2],
      details: r[3] ? JSON.parse(r[3]) : null,
      requestedBy: r[4],
      requestedAt: r[5] ? new Date(r[5]).toLocaleString('th-TH') : '',
      status: r[6]
    });
  });
  return result;
}

/** decision: 'approve' | 'reject', photoDate: ใช้เฉพาะกรณีอนุมัติคำขอย้ายห้อง (ไม่บังคับ) */
function adminHandleRequest(token, requestRow, decision, photoDate) {
  const session = validateSession_(token);
  requireRole_(session, ['admin']);
  const reqSheet = getRequestsSheet_();
  const reqVals = reqSheet.getRange(requestRow, 1, 1, 7).getValues()[0];
  const type = reqVals[0], studentId = reqVals[1], detailsRaw = reqVals[3];

  if (decision === 'approve') {
    try {
      if (type === 'ย้ายห้อง') {
        const details = JSON.parse(detailsRaw);
        moveStudentClassInternal_(studentId, details.newGrade, details.newRoom, photoDate);
      } else if (type === 'ลบ') {
        deleteStudentInternal_(studentId);
      }
      reqSheet.getRange(requestRow, 7).setValue('อนุมัติแล้ว');
    } catch (e) {
      reqSheet.getRange(requestRow, 7).setValue('ผิดพลาด: ' + e.message);
      throw e;
    }
  } else {
    reqSheet.getRange(requestRow, 7).setValue('ปฏิเสธ');
  }
  return { ok: true };
}

// ========================= บริษัท: ค้นหา + ดาวน์โหลดเฉพาะที่เลือก =========================
function companySearchStudents(token, filters) {
  const session = validateSession_(token);
  requireRole_(session, ['company', 'admin', 'teacher']);
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, 1, last - 1, COLUMNS.length).getValues();

  const result = [];
  data.forEach(function (row) {
    if (row[COL.PHOTO_STATUS] !== 'published') return; // บริษัทเห็นแค่รูปที่เผยแพร่แล้ว
    if (filters) {
      if (filters.grade && row[COL.GRADE] !== filters.grade) return;
      if (filters.room && String(row[COL.ROOM]) !== String(filters.room)) return;
      if (filters.date) {
        const rowDate = String(row[COL.PHOTO_DATE] || '').trim();
        if (rowDate !== filters.date) return;
      }
    }
    const obj = rowToObject_(row);
    delete obj.phone;
    result.push(obj);
  });
  return result;
}

/** ids: รายการ RecordId ที่ติ๊กเลือกไว้ -> สร้าง ZIP (รูปที่เลือก + ข้อมูล.csv) */
function companyDownloadSelected(token, ids) {
  const session = validateSession_(token);
  requireRole_(session, ['company', 'admin', 'teacher']);
  if (!ids || !ids.length) throw new Error('กรุณาเลือกรูปที่ต้องการดาวน์โหลดอย่างน้อย 1 รายการ');

  const sheet = getSheet();
  const blobs = [];
  const csvRows = [['รหัสนักเรียน', 'ชื่อ-สกุล', 'ระดับชั้น', 'ห้อง', 'วันที่ของรูป', 'ชื่อไฟล์']];

  ids.forEach(function (id) {
    try {
      const row = findRowByRecordId_(id);
      const rowVals = sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0];
      if (rowVals[COL.PHOTO_STATUS] !== 'published' || !rowVals[COL.FILE_ID]) return;
      const file = DriveApp.getFileById(rowVals[COL.FILE_ID]);
      blobs.push(file.getBlob());
      csvRows.push([
        activeCode_(rowVals[COL.REAL_ID], rowVals[COL.TEMP_ID]),
        (rowVals[COL.PREFIX] + ' ' + rowVals[COL.FNAME] + ' ' + rowVals[COL.LNAME]).trim(),
        rowVals[COL.GRADE], rowVals[COL.ROOM], rowVals[COL.PHOTO_DATE] || '', rowVals[COL.PHOTO_FILE]
      ]);
    } catch (e) { /* ข้ามรายการที่มีปัญหา */ }
  });

  if (!blobs.length) throw new Error('ไม่พบไฟล์รูปที่เลือก');

  const csvContent = '\uFEFF' + csvRows.map(function (r) {
    return r.map(function (c) { return '"' + String(c).replace(/"/g, '""') + '"'; }).join(',');
  }).join('\n');
  blobs.push(Utilities.newBlob(csvContent, 'text/csv', 'ข้อมูล.csv'));

  const zipBlob = Utilities.zip(blobs, 'student_photos.zip');
  return { base64: Utilities.base64Encode(zipBlob.getBytes()), filename: 'student_photos_' + new Date().getTime() + '.zip' };
}

// ========================= MENU (เปิดจากหน้า Google Sheet) =========================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ระบบรูปภาพนักเรียน')
    .addItem('วิธีเปิดเว็บแอป', 'showWebAppInfo_')
    .addToUi();
}
function showWebAppInfo_() {
  const url = ScriptApp.getService().getUrl();
  const msg = url ? ('เปิดเว็บแอประบบได้ที่:\n' + url) : 'ยังไม่ได้ Deploy เว็บแอป กรุณาไปที่ Deploy > New deployment > Web app';
  SpreadsheetApp.getUi().alert(msg);
}
