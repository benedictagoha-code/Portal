/**
 * psyc@UNIDEL 3.10 — Courseware, Messaging & Grade Compilation
 * Department of Psychology, University of Delta, Agbor
 * Existing installations: run upgradeSystem39() once; never rerun setupSystem().
 */

const PROP_DB_ID = 'PSYC_UNIDEL_V3_DB_ID';
const PROP_DRIVE_ID = 'PSYC_UNIDEL_V3_DRIVE_ID';

const SHEETS = {
  SETTINGS: 'Settings',
  USERS: 'Users',
  STUDENTS: 'Students',
  LECTURERS: 'Lecturers',
  COURSES: 'Courses',
  COURSE_LECTURERS: 'CourseLecturers',
  ENROLMENTS: 'Enrolments',
  MATERIALS: 'Materials',
  ATTENDANCE: 'Attendance',
  TESTS: 'Tests',
  QUESTIONS: 'Questions',
  ATTEMPTS: 'Attempts',
  RESULTS: 'Results',
  ANNOUNCEMENTS: 'Announcements',
  DISCUSSIONS: 'Discussions',
  REPLIES: 'DiscussionReplies',
  TEST_BANK: 'TestBank',
  TEST_SESSIONS: 'TestSessions'
};

/* ============================================================
   1. ONE-CLICK SYSTEM INITIALISATION & SAFE UPGRADE
   ============================================================ */
function setupSystem() {
  const props = PropertiesService.getScriptProperties();
  let dbId = props.getProperty(PROP_DB_ID);
  let ss;

  if (dbId) {
    try { ss = SpreadsheetApp.openById(dbId); } catch (e) { ss = null; }
  }
  if (dbId && !ss) {
    throw new Error('The configured psyc@UNIDEL database could not be opened. Setup stopped to protect existing records. Restore the correct Script Property database ID; do not initialise a new database over a live installation.');
  }
  if (!ss) {
    ss = SpreadsheetApp.create('psyc@UNIDEL_3.5_Database');
    dbId = ss.getId();
    props.setProperty(PROP_DB_ID, dbId);
  }

  let folderId = props.getProperty(PROP_DRIVE_ID);
  if (!folderId) {
    const folder = DriveApp.createFolder('psyc@UNIDEL_Courseware_Repository');
    folderId = folder.getId();
    props.setProperty(PROP_DRIVE_ID, folderId);
  }

  const schema = {
    [SHEETS.SETTINGS]: ['Key', 'Value'],
    [SHEETS.USERS]: ['UserID', 'Name', 'Role', 'PasswordHash', 'Active', 'CreatedAt'],
    [SHEETS.STUDENTS]: ['StudentID', 'MatricNo', 'Surname', 'FirstName', 'OtherNames', 'Gender', 'Programme', 'Level', 'Email', 'Phone', 'CreatedAt'],
    [SHEETS.LECTURERS]: ['LecturerID', 'StaffID', 'Title', 'Surname', 'FirstName', 'Rank', 'Email', 'Status'],
    [SHEETS.COURSES]: ['CourseID', 'CourseCode', 'CourseTitle', 'Level', 'Semester', 'CreditUnits', 'LecturerID'],
    [SHEETS.COURSE_LECTURERS]: ['CourseID', 'LecturerID', 'AssignedAt', 'AssignedBy'],
    [SHEETS.ENROLMENTS]: ['EnrolmentID', 'CourseID', 'StudentID', 'MatricNo', 'StudentName', 'EnrolledAt'],
    [SHEETS.MATERIALS]: ['MaterialID', 'CourseID', 'CourseCode', 'Title', 'MaterialType', 'Description', 'FileURL', 'UploadedBy', 'UploadedAt'],
    [SHEETS.ATTENDANCE]: ['AttendanceID', 'CourseID', 'LectureDate', 'Topic', 'StudentID', 'MatricNo', 'Status', 'RecordedBy'],
    [SHEETS.TESTS]: ['TestID', 'CourseID', 'CourseCode', 'Kind', 'Title', 'DurationMinutes', 'TotalMarks', 'MaxAttempts', 'ScoringMethod', 'FeedbackMode', 'OpenDateTime', 'CloseDateTime', 'Shuffle', 'Instructions', 'Status', 'CreatedAt', 'ComponentGroup', 'ComponentGradingMethod', 'SelectedTestID', 'QuestionCount', 'ShuffleOptions'],
    [SHEETS.QUESTIONS]: ['QuestionID', 'TestID', 'QuestionType', 'QuestionText', 'OptionA', 'OptionB', 'OptionC', 'OptionD', 'CorrectAnswer', 'Marks', 'Explanation'],
    [SHEETS.ATTEMPTS]: ['AttemptID', 'TestID', 'StudentID', 'MatricNo', 'AttemptNumber', 'Score', 'MaxScore', 'SubmittedAt'],
    [SHEETS.RESULTS]: ['ResultID', 'CourseID', 'CourseCode', 'CourseTitle', 'CreditUnits', 'StudentID', 'MatricNo', 'StudentName', 'CAScore', 'ExamScore', 'TotalScore', 'Grade', 'GradePoint', 'UpdatedAt'],
    [SHEETS.ANNOUNCEMENTS]: ['AnnouncementID', 'Title', 'Audience', 'Body', 'AuthorName', 'AuthorRole', 'DatePosted'],
    [SHEETS.DISCUSSIONS]: ['TopicID', 'CourseID', 'CourseCode', 'Title', 'Body', 'AuthorName', 'AuthorRole', 'DateCreated'],
    [SHEETS.REPLIES]: ['ReplyID', 'TopicID', 'Body', 'AuthorName', 'AuthorRole', 'DatePosted'],
    [SHEETS.TEST_BANK]: ['BankQuestionID','CourseID','OwnerID','Sharing','Topic','Difficulty','QuestionType','QuestionText','OptionA','OptionB','OptionC','OptionD','CorrectAnswer','Marks','Explanation','Status','CreatedAt','UpdatedAt'],
    [SHEETS.TEST_SESSIONS]: ['SessionID','TestID','StudentID','QuestionIDs','OptionOrderJSON','StartedAt','ExpiresAt','Submitted']
  };

  Object.keys(schema).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const targetHeaders = schema[name];
    if (sh.getLastRow() === 0) {
      sh.appendRow(targetHeaders);
      sh.getRange(1, 1, 1, targetHeaders.length).setFontWeight('bold');
    } else {
      const currentHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const missingHeaders = targetHeaders.filter(h => currentHeaders.indexOf(h) === -1);
      if (missingHeaders.length) {
        sh.getRange(1, currentHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]).setFontWeight('bold');
      }
    }
  });

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch (e) {}
  }

  const setSh = ss.getSheetByName(SHEETS.SETTINGS);
  if (setSh.getLastRow() === 1) {
    [
      ['SystemName', 'psyc@UNIDEL'],
      ['Version', '3.9'],
      ['Institution', 'University of Delta, Agbor'],
      ['Department', 'Department of Psychology'],
      ['AcademicSession', '2025/2026'],
      ['Semester', 'First Semester'],
      ['MinimumAttendance', '75'],
      ['ResultPublication', 'PUBLISHED']
    ].forEach(r => setSh.appendRow(r));
  }

  const usrSh = ss.getSheetByName(SHEETS.USERS);
  if (usrSh.getLastRow() === 1) {
    usrSh.appendRow(['ADMIN001', 'Departmental Administrator', 'Administrator', hashPassword_('Admin@1234'), 'YES', new Date().toISOString()]);
    usrSh.appendRow(['LEC-STF01', 'Dr. Sample Lecturer', 'Lecturer', hashPassword_('Lect@1234'), 'YES', new Date().toISOString()]);
  }

  const lecSh = ss.getSheetByName(SHEETS.LECTURERS);
  if (lecSh.getLastRow() === 1) {
    lecSh.appendRow(['LEC-STF01', 'STF01', 'Dr.', 'Sample', 'Lecturer', 'Senior Lecturer', 'lecturer@unidel.edu.ng', 'ACTIVE']);
  }

  const crsSh = ss.getSheetByName(SHEETS.COURSES);
  if (crsSh.getLastRow() === 1) {
    crsSh.appendRow(['CRS-PSY101', 'PSY 101', 'Introduction to Psychology I', '100', 'First Semester', 3, 'LEC-STF01']);
    crsSh.appendRow(['CRS-PSY201', 'PSY 201', 'Cognitive Psychology', '200', 'First Semester', 2, 'LEC-STF01']);
    crsSh.appendRow(['CRS-PSY301', 'PSY 301', 'Psychological Testing & Assessment', '300', 'First Semester', 3, 'LEC-STF01']);
    crsSh.appendRow(['CRS-PSY401', 'PSY 401', 'Clinical Psychology & Psychotherapy', '400', 'First Semester', 3, 'LEC-STF01']);
  }

  Logger.log('SETUP COMPLETE! Spreadsheet URL: ' + ss.getUrl());
  return { success: true, spreadsheetUrl: ss.getUrl() };
}

/* ============================================================
   1A. NON-DESTRUCTIVE 3.6 DATABASE MIGRATION
   ============================================================ */
function upgradeSystem36() {
  const ss = getDB_();
  const sh = ss.getSheetByName(SHEETS.TESTS);
  if (!sh) throw new Error('Tests sheet is missing. No changes were made.');
  const required = ['ComponentGroup', 'ComponentGradingMethod', 'SelectedTestID'];
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const missing = required.filter(h => headers.indexOf(h) === -1);
  if (missing.length) sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  const rows = getRows_(SHEETS.SETTINGS);
  const version = rows.find(r => String(r.key).toLowerCase() === 'version');
  if (version) settings.getRange(version._row, 2).setValue('3.6');
  else settings.appendRow(['Version', '3.6']);
  const minAtt = rows.find(r => String(r.key).toLowerCase() === 'minimumattendance');
  if (!minAtt) settings.appendRow(['MinimumAttendance', '75']);
  return { success: true, version: '3.6', addedColumns: missing, message: '3.6 database migration completed. Existing records were retained.' };
}


/**
 * Non-destructive 3.7 migration.
 * No database tables or existing records are recreated or cleared.
 */
function upgradeSystem37() {
  const ss = getDB_();
  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (!settings) throw new Error('Settings sheet not found. Existing database was not changed.');
  const rows = getRows_(SHEETS.SETTINGS);
  const version = rows.find(r => String(r.key).trim().toLowerCase() === 'version');
  if (version) settings.getRange(version._row, 2).setValue('3.7');
  else settings.appendRow(['Version', '3.7']);
  const attendance = rows.find(r => String(r.key).trim().toLowerCase() === 'minimumattendance');
  if (!attendance) settings.appendRow(['MinimumAttendance', 75]);
  return { success: true, version: '3.7', message: '3.7 upgrade completed. Existing student, attendance, assessment and result records were retained.' };
}

/** Non-destructive 3.8 migration for optional multi-lecturer course assignment. */
function upgradeSystem38() {
  const ss = getDB_();
  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (!ss.getSheetByName(SHEETS.COURSES) || !settings) throw new Error('Courses or Settings sheet is missing.');
  let sh = ss.getSheetByName(SHEETS.COURSE_LECTURERS);
  if (!sh) { sh = ss.insertSheet(SHEETS.COURSE_LECTURERS); sh.appendRow(['CourseID','LecturerID','AssignedAt','AssignedBy']); sh.getRange(1,1,1,4).setFontWeight('bold'); }
  const existing = getRows_(SHEETS.COURSE_LECTURERS);
  const keys = new Set(existing.map(r => String(r.courseid).trim() + '|' + String(r.lecturerid).trim().toUpperCase()));
  const additions = [];
  getRows_(SHEETS.COURSES).forEach(c => { const cid=String(c.courseid||'').trim(), lid=String(c.lecturerid||'').trim().toUpperCase(); if(cid&&lid&&!keys.has(cid+'|'+lid)){additions.push([cid,lid,new Date().toISOString(),'SYSTEM-MIGRATION']);keys.add(cid+'|'+lid);} });
  if (additions.length) sh.getRange(sh.getLastRow()+1,1,additions.length,4).setValues(additions);
  const rows=getRows_(SHEETS.SETTINGS), v=rows.find(r=>String(r.key).trim().toLowerCase()==='version');
  if(v) settings.getRange(v._row,2).setValue('3.8'); else settings.appendRow(['Version','3.8']);
  return {success:true,version:'3.8',migratedAssignments:additions.length,message:'3.8 migration completed. Existing records were retained.'};
}

/** Non-destructive 3.9 migration: question bank and resumable randomised attempts. */
function upgradeSystem39() {
  const ss = getDB_(), settings = ss.getSheetByName(SHEETS.SETTINGS), tests = ss.getSheetByName(SHEETS.TESTS);
  if (!settings || !tests || !ss.getSheetByName(SHEETS.QUESTIONS)) throw new Error('Required assessment sheets are missing. No migration was performed.');
  const bankHeaders = ['BankQuestionID','CourseID','OwnerID','Sharing','Topic','Difficulty','QuestionType','QuestionText','OptionA','OptionB','OptionC','OptionD','CorrectAnswer','Marks','Explanation','Status','CreatedAt','UpdatedAt'];
  let bank = ss.getSheetByName(SHEETS.TEST_BANK);
  if (!bank) { bank = ss.insertSheet(SHEETS.TEST_BANK); bank.appendRow(bankHeaders); bank.getRange(1,1,1,bankHeaders.length).setFontWeight('bold'); }
  const sessionHeaders = ['SessionID','TestID','StudentID','QuestionIDs','OptionOrderJSON','StartedAt','ExpiresAt','Submitted'];
  let sessions = ss.getSheetByName(SHEETS.TEST_SESSIONS);
  if (!sessions) { sessions = ss.insertSheet(SHEETS.TEST_SESSIONS); sessions.appendRow(sessionHeaders); sessions.getRange(1,1,1,sessionHeaders.length).setFontWeight('bold'); }
  const lastCol = Math.max(1, tests.getLastColumn()), headers = tests.getRange(1,1,1,lastCol).getValues()[0].map(v=>String(v).trim());
  const missing = ['QuestionCount','ShuffleOptions'].filter(h=>headers.indexOf(h)<0);
  const qsh=ss.getSheetByName(SHEETS.QUESTIONS), qlast=Math.max(1,qsh.getLastColumn()), qheaders=qsh.getRange(1,1,1,qlast).getValues()[0].map(v=>String(v).trim());
  if(qheaders.indexOf('Explanation')<0) qsh.getRange(1,qlast+1).setValue('Explanation').setFontWeight('bold');
  if (missing.length) tests.getRange(1,lastCol+1,1,missing.length).setValues([missing]).setFontWeight('bold');
  const rows = getRows_(SHEETS.SETTINGS), version = rows.find(r=>String(r.key).trim().toLowerCase()==='version');
  if (version) settings.getRange(version._row,2).setValue('3.9'); else settings.appendRow(['Version','3.9']);
  return {success:true,version:'3.9',addedTestColumns:missing,message:'3.9 upgrade completed. Existing records were retained.'};
}

/* ============================================================
   2. CORE HELPERS & SECURITY
   ============================================================ */
function upgradeSystem391() {
  const ss=getDB_(), qsh=ss.getSheetByName(SHEETS.QUESTIONS), bsh=ss.getSheetByName(SHEETS.TEST_BANK);
  if(!qsh||!bsh) throw new Error('Required assessment and Test Bank sheets are missing. No migration was performed.');
  [qsh,bsh].forEach(sh=>{const last=Math.max(1,sh.getLastColumn()),hs=sh.getRange(1,1,1,last).getValues()[0].map(v=>String(v).trim());if(hs.indexOf('AnswerDataJSON')<0)sh.getRange(1,last+1).setValue('AnswerDataJSON').setFontWeight('bold');});
  const settings=ss.getSheetByName(SHEETS.SETTINGS), rows=getRows_(SHEETS.SETTINGS), version=rows.find(r=>String(r.key).trim().toLowerCase()==='version');
  if(version)settings.getRange(version._row,2).setValue('3.9.1');else settings.appendRow(['Version','3.9.1']);
  return {success:true,version:'3.9.1',message:'3.9.1 upgrade completed. Existing questions, records and attempts were retained.'};
}
function parseAnswerData_(v){if(!v)return{};try{const x=JSON.parse(String(v));return x&&typeof x==='object'?x:{};}catch(e){return{};}}
function normaliseAnswerText_(v){return String(v==null?'':v).trim().replace(/\s+/g,' ').toLowerCase();}
function buildQuestionAnswerData_(p,t){
  p=p||{};const d={};
  if(t==='FILL_GAP'){const a=Array.isArray(p.acceptedAnswers)?p.acceptedAnswers:String(p.correctAnswer||'').split('|');d.acceptedAnswers=a.map(normaliseAnswerText_).filter(Boolean);}
  if(t==='MATCHING'){const a=Array.isArray(p.matchingPairs)?p.matchingPairs:[];d.matchingPairs=a.map(x=>({left:String(x.left||'').trim(),right:String(x.right||'').trim()})).filter(x=>x.left&&x.right);}
  if(t==='MULTI_CORRECT'||t==='MULTI_INCORRECT'){const a=Array.isArray(p.correctKeys)?p.correctKeys:String(p.correctAnswer||'').split(',');d.correctKeys=a.map(x=>String(x).trim().toLowerCase()).filter(x=>/^[a-d]$/.test(x));d.markingMode=String(p.markingMode||'PARTIAL').toUpperCase()==='EXACT'?'EXACT':'PARTIAL';}
  return d;
}

function getDB_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_DB_ID);
  if (!id) throw new Error('Database not initialised. Run setupSystem() once in Code.gs.');
  return SpreadsheetApp.openById(id);
}

function getRows_(sheetName) {
  const sh = getDB_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  return data.slice(1).map((row, idx) => {
    const obj = { _row: idx + 2 };
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function hashPassword_(raw) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'UNIDEL_SALT_' + String(raw));
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function makeId_(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function getSessionUser_(token) {
  if (!token) return null;
  const cached = CacheService.getScriptCache().get('TOK_' + token);
  if (!cached) return null;
  try { return JSON.parse(cached); } catch (e) { return null; }
}

function requireUser_(token, allowedRoles) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Session expired. Please sign in again.');
  if (allowedRoles && allowedRoles.length && allowedRoles.indexOf(user.role) === -1) {
    throw new Error('Access denied for role: ' + user.role);
  }
  return user;
}

function normaliseMatric_(value) {
  return String(value || '').replace(/\s+/g, '').trim().toUpperCase();
}

function isStudentEnrolled_(matric, courseID) {
  const key = normaliseMatric_(matric);
  return getRows_(SHEETS.ENROLMENTS).some(e => normaliseMatric_(e.matricno) === key && String(e.courseid) === String(courseID));
}

function getCourseLecturerIds_(course) {
  const ids=[]; const primary=String(course && course.lecturerid || '').trim().toUpperCase();
  if(primary) ids.push(primary);
  getRows_(SHEETS.COURSE_LECTURERS).forEach(r=>{if(String(r.courseid)===String(course.courseid)){const id=String(r.lecturerid||'').trim().toUpperCase();if(id&&ids.indexOf(id)<0)ids.push(id);}});
  return ids;
}
function getLecturerCourseIds_(lecturerID) {
  const uid=String(lecturerID||'').trim().toUpperCase();
  return new Set(getRows_(SHEETS.COURSES).filter(c=>getCourseLecturerIds_(c).indexOf(uid)>=0).map(c=>String(c.courseid)));
}
function requireCourseAccess_(user, courseID) {
  const course=getRows_(SHEETS.COURSES).find(c=>String(c.courseid)===String(courseID));
  if(!course) throw new Error('Course not found.');
  if(user.role==='Administrator') return course;
  if(user.role!=='Lecturer'||getCourseLecturerIds_(course).indexOf(String(user.userID||'').trim().toUpperCase())<0) throw new Error('Access denied. You are not assigned to this course.');
  return course;
}

function escapeHtml_(value) {
  return String(value == null ? '' : value).replace(/[&<>\"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;',"'":'&#39;'}[ch]));
}

function getSetting_(key, fallback) {
  const row = getRows_(SHEETS.SETTINGS).find(r => String(r.key).trim().toLowerCase() === String(key).trim().toLowerCase());
  return row ? row.value : fallback;
}

function setCellByHeader_(sheetName, rowNumber, header, value) {
  const sh = getDB_().getSheetByName(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim().toLowerCase());
  const idx = headers.indexOf(String(header).toLowerCase());
  if (idx < 0) throw new Error('Missing database column: ' + header);
  sh.getRange(rowNumber, idx + 1).setValue(value);
}

/* ============================================================
   3. AUTHENTICATION, STUDENT & LECTURER SELF-REGISTRATION
   ============================================================ */
function getPublicInfo() {
  try {
    const rows = getRows_(SHEETS.SETTINGS);
    const map = {};
    rows.forEach(r => { map[r.key] = r.value; });
    return {
      success: true,
      systemName: map.SystemName || 'psyc@UNIDEL',
      version: map.Version || '3.10',
      institution: map.Institution || 'University of Delta, Agbor',
      department: map.Department || 'Department of Psychology',
      academicSession: map.AcademicSession || '2025/2026',
      semester: map.Semester || 'First Semester'
    };
  } catch (e) {
    return { success: true, systemName: 'psyc@UNIDEL', version: '3.10', institution: 'University of Delta, Agbor', department: 'Department of Psychology' };
  }
}

function loginUser(userID, password) {
  const cleanId = String(userID || '').trim().toUpperCase();
  const hash = hashPassword_(String(password || ''));
  const users = getRows_(SHEETS.USERS);
  const match = users.find(u => String(u.userid).trim().toUpperCase() === cleanId);

  if (!match) return { success: false, message: 'Account not found. Check your User ID or Matriculation Number.' };
  if (String(match.active).toUpperCase() !== 'YES') return { success: false, message: 'Account is inactive. Contact the Departmental Administrator.' };
  if (String(match.passwordhash) !== hash) return { success: false, message: 'Incorrect password.' };

  const token = Utilities.getUuid();
  const sessionUser = { userID: match.userid, name: match.name, role: match.role };
  CacheService.getScriptCache().put('TOK_' + token, JSON.stringify(sessionUser), 21600);
  return { success: true, token: token, user: sessionUser };
}

function validateSession(token) {
  const user = getSessionUser_(token);
  return user ? { success: true, user: user } : { success: false };
}

function logoutUser(token) {
  if (token) CacheService.getScriptCache().remove('TOK_' + token);
  return { success: true };
}

function registerStudent(details, password) {
  details = details || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const matric = normaliseMatric_(details.matricNo);
    const surname = String(details.surname || '').trim();
    const firstName = String(details.firstName || '').trim();
    const level = String(details.level || '').trim();
    const gender = String(details.gender || 'OTHER').trim().toUpperCase();
    if (!matric || !surname || !firstName || !level || !password) {
      return { success: false, message: 'Please complete all required fields.' };
    }
    if (String(password).length < 6) return { success: false, message: 'Password must be at least 6 characters.' };
    if (['MALE','FEMALE','OTHER'].indexOf(gender) === -1) return { success: false, message: 'Please select a valid gender.' };

    // Check both tables: a partial/legacy record must never permit a second registration.
    const students = getRows_(SHEETS.STUDENTS);
    const users = getRows_(SHEETS.USERS);
    if (students.some(st => normaliseMatric_(st.matricno) === matric) ||
        users.some(u => normaliseMatric_(u.userid) === matric)) {
      return { success: false, message: 'Matriculation Number already registered. Please sign in or contact the administrator if you cannot access the account.' };
    }

    const fullName = [surname, firstName, String(details.otherNames || '').trim()].filter(Boolean).join(' ');
    const now = new Date().toISOString();
    const studentID = makeId_('STU');
    const studentSheet = getDB_().getSheetByName(SHEETS.STUDENTS);
    const userSheet = getDB_().getSheetByName(SHEETS.USERS);
    const studentRow = [studentID, matric, surname, firstName, String(details.otherNames || '').trim(), gender,
      details.programme || 'B.Sc. Psychology', level, String(details.email || '').trim(), String(details.phone || '').trim(), now];
    const userRow = [matric, fullName, 'Student', hashPassword_(password), 'YES', now];
    studentSheet.appendRow(studentRow);
    try {
      userSheet.appendRow(userRow);
    } catch (writeError) {
      studentSheet.deleteRow(studentSheet.getLastRow());
      throw writeError;
    }
    return { success: true, userID: matric, message: 'Student registration successful! Sign in with Matric No: ' + matric };
  } finally {
    lock.releaseLock();
  }
}

function registerLecturer(details, password) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const staffId = String(details.staffID || '').trim().toUpperCase();
    const surname = String(details.surname || '').trim();
    const firstName = String(details.firstName || '').trim();
    const email = String(details.email || '').trim().toLowerCase();

    if (!staffId || !surname || !firstName || !email || !password) {
      return { success: false, message: 'Staff ID, Title, Surname, First Name, Email, and Password are required.' };
    }
    if (String(password).length < 6) {
      return { success: false, message: 'Password must be at least 6 characters.' };
    }

    const userId = staffId.indexOf('LEC-') === 0 ? staffId : 'LEC-' + staffId;
    const users = getRows_(SHEETS.USERS);
    if (users.some(u => String(u.userid).trim().toUpperCase() === userId)) {
      return { success: false, message: 'Lecturer account already exists. Sign in with User ID: ' + userId };
    }

    const title = String(details.title || 'Dr.').trim();
    const rank = String(details.rank || 'Lecturer').trim();
    const fullName = [title, surname, firstName].filter(Boolean).join(' ').trim();

    const lecs = getRows_(SHEETS.LECTURERS);
    if (!lecs.some(l => String(l.staffid).trim().toUpperCase() === staffId)) {
      getDB_().getSheetByName(SHEETS.LECTURERS).appendRow([
        userId, staffId, title, surname, firstName, rank, email, 'ACTIVE'
      ]);
    }

    getDB_().getSheetByName(SHEETS.USERS).appendRow([
      userId, fullName, 'Lecturer', hashPassword_(password), 'YES', new Date().toISOString()
    ]);

    return { success: true, userID: userId, message: 'Lecturer self-registration complete! Sign in with User ID: ' + userId };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
   4. COURSES & SELF-ENROLMENT
   ============================================================ */
function getCourses(token) {
  const user = requireUser_(token);
  const all = getRows_(SHEETS.COURSES);

  if (user.role === 'Lecturer') {
    const uid = String(user.userID).toUpperCase();
    return {
      success: true,
      courses: all.filter(c => getCourseLecturerIds_(c).indexOf(uid) >= 0),
      allCourses: all.map(c => Object.assign({}, c, { assignedLecturers: getCourseLecturerIds_(c) })),
      lecturerOptions: getRows_(SHEETS.USERS).filter(u => String(u.role).toLowerCase() === 'lecturer').map(u => ({ userID: u.userid, name: u.name }))
    };
  }

  if (user.role === 'Student') {
    const myEnrols = getRows_(SHEETS.ENROLMENTS).filter(e => String(e.matricno).toUpperCase() === String(user.userID).toUpperCase());
    const enrolledIds = new Set(myEnrols.map(e => e.courseid));
    return {
      success: true,
      courses: all.filter(c => enrolledIds.has(c.courseid)),
      catalogue: all.map(c => Object.assign({}, c, { isEnrolled: enrolledIds.has(c.courseid) }))
    };
  }

  return { success: true, courses: all, allCourses: all.map(c => Object.assign({}, c, { assignedLecturers: getCourseLecturerIds_(c) })), lecturerOptions: getRows_(SHEETS.USERS).filter(u => String(u.role).toLowerCase() === 'lecturer').map(u => ({ userID: u.userid, name: u.name })) };
}

function createCourse(token, payload) {
  requireUser_(token, ['Administrator']);
  const code = String(payload.courseCode || '').trim().toUpperCase();
  const title = String(payload.courseTitle || '').trim();
  if (!code || !title) return { success: false, message: 'Course code and title are required.' };
  if (getRows_(SHEETS.COURSES).some(c => String(c.coursecode).trim().toUpperCase() === code)) {
    return { success: false, message: 'A course with this course code already exists.' };
  }

  const cid = 'CRS-' + code.replace(/\s+/g, '');
  getDB_().getSheetByName(SHEETS.COURSES).appendRow([
    cid, code, title, String(payload.level || '100'), payload.semester || 'First Semester',
    Number(payload.creditUnits || 2), String(payload.lecturerID || '').trim().toUpperCase()
  ]);
  return { success: true, message: 'Course ' + code + ' created.' };
}

function claimCourseAsLecturer(token, courseID) {
  const user=requireUser_(token,['Lecturer']);
  const target=getRows_(SHEETS.COURSES).find(c=>String(c.courseid)===String(courseID));
  if(!target) return {success:false,message:'Course not found.'};
  const claimant=String(user.userID||'').trim().toUpperCase();
  if(getCourseLecturerIds_(target).indexOf(claimant)>=0) return {success:true,message:'You are already assigned to '+target.coursecode+'.'};
  const sh=getDB_().getSheetByName(SHEETS.COURSE_LECTURERS);
  if(!sh) throw new Error('Course assignment table is missing. Ask the Administrator to run upgradeSystem38().');
  sh.appendRow([target.courseid,claimant,new Date().toISOString(),user.userID]);
  return {success:true,message:'You have been added to '+target.coursecode+'. Existing lecturer assignments were retained.'};
}

function manageCourseLecturer(token, courseID, lecturerID, action) {
  const admin=requireUser_(token,['Administrator']);
  const target=getRows_(SHEETS.COURSES).find(c=>String(c.courseid)===String(courseID));
  if(!target) return {success:false,message:'Course not found.'};
  const lid=String(lecturerID||'').trim().toUpperCase();
  if(!lid) return {success:false,message:'Select a lecturer.'};
  const lecturer=getRows_(SHEETS.USERS).find(u=>String(u.userid).trim().toUpperCase()===lid&&String(u.role).toLowerCase()==='lecturer');
  if(!lecturer) return {success:false,message:'The selected account is not a lecturer account.'};
  const sh=getDB_().getSheetByName(SHEETS.COURSE_LECTURERS);
  if(!sh) throw new Error('Course assignment table is missing. Run upgradeSystem38() first.');
  const assigned=getCourseLecturerIds_(target);
  const mapped=getRows_(SHEETS.COURSE_LECTURERS).find(r=>String(r.courseid)===String(courseID)&&String(r.lecturerid).trim().toUpperCase()===lid);
  const primary=String(target.lecturerid||'').trim().toUpperCase();
  if(String(action).toLowerCase()==='add'){
    if(assigned.indexOf(lid)>=0) return {success:true,message:lecturer.name+' is already assigned.'};
    sh.appendRow([courseID,lid,new Date().toISOString(),admin.userID]);
    return {success:true,message:lecturer.name+' added to '+target.coursecode+'.'};
  }
  if(String(action).toLowerCase()==='remove'){
    if(primary===lid) setCellByHeader_(SHEETS.COURSES,target._row,'LecturerID','');
    if(mapped) sh.deleteRow(mapped._row);
    return {success:true,message:lecturer.name+' removed from '+target.coursecode+'.'};
  }
  return {success:false,message:'Unknown assignment action.'};
}

function enrolStudentInCourse(token, courseID) {
  const user = requireUser_(token, ['Student']);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const matric = normaliseMatric_(user.userID);
    const course = getRows_(SHEETS.COURSES).find(c => String(c.courseid) === String(courseID));
    if (!course) return { success: false, message: 'Course not found. Please select a valid course.' };
    if (isStudentEnrolled_(matric, courseID)) return { success: false, message: 'You are already enrolled in this course.' };
    const stu = getRows_(SHEETS.STUDENTS).find(s => normaliseMatric_(s.matricno) === matric);
    if (!stu) return { success: false, message: 'Student profile not found. Please contact the administrator.' };
    getDB_().getSheetByName(SHEETS.ENROLMENTS).appendRow([
      makeId_('ENR'), courseID, stu.studentid, matric,
      [stu.surname, stu.firstname, stu.othernames].filter(Boolean).join(' '), new Date().toISOString()
    ]);
    return { success: true, message: 'Course enrolment confirmed!' };
  } finally { lock.releaseLock(); }
}

function getCourseRoster(token, courseID) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  requireCourseAccess_(user, courseID);
  return { success: true, students: getRows_(SHEETS.ENROLMENTS).filter(e => e.courseid === courseID) };
}

/* ============================================================
   5. COURSEWARE MATERIALS
   ============================================================ */
function uploadMaterial(token, payload) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  payload = payload || {};
  const crs = requireCourseAccess_(user, payload.courseID);
  const title = String(payload.title || '').trim();
  if (!title || title.length > 200) throw new Error('Enter a title of no more than 200 characters.');

  let fileUrl = String(payload.externalLink || '').trim();
  if (fileUrl) {
    if (fileUrl.length > 2048 || !/^https?:\/\//i.test(fileUrl)) {
      throw new Error('External resource links must be valid HTTP or HTTPS URLs.');
    }
    const urlMatch = fileUrl.match(/^https?:\/\/([^\/?#]+)(?:[\/?#]|$)/i);
    if (!urlMatch || !urlMatch[1] || /[\\s@]/.test(urlMatch[1])) {
      throw new Error('The external resource link is not valid. Use a complete HTTP or HTTPS link.');
    }
  }

  if (payload.base64Data || payload.fileName) {
    if (!payload.base64Data || !payload.fileName) throw new Error('The uploaded file data is incomplete.');
    const encoded = String(payload.base64Data);
    if (encoded.length > 28 * 1024 * 1024) throw new Error('Direct uploads are limited to 20 MB. Use an approved cloud link for larger media.');
    const bytes = Utilities.base64Decode(encoded);
    if (bytes.length > 20 * 1024 * 1024) throw new Error('Direct uploads are limited to 20 MB. Use an approved cloud link for larger media.');
    const safeName = String(payload.fileName).replace(/[\\\\/:*?"<>|\\r\\n]/g, '_').slice(0, 180);
    const allowed = /\\.(pdf|doc|docx|ppt|pptx|mp4|webm|mov|mp3|wav|m4a)$/i;
    if (!allowed.test(safeName)) throw new Error('Unsupported file type. Upload PDF, Office, video or audio files only.');
    const folderId = PropertiesService.getScriptProperties().getProperty(PROP_DRIVE_ID);
    if (!folderId) throw new Error('Courseware storage is not configured. Contact the system administrator.');
    const folder = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(bytes, String(payload.mimeType || 'application/octet-stream').slice(0, 120), safeName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileUrl = file.getUrl();
  }
  if (!fileUrl) return { success: false, message: 'Attach a file or enter an external link.' };

  getDB_().getSheetByName(SHEETS.MATERIALS).appendRow([
    makeId_('MAT'), crs.courseid, crs.coursecode, title, String(payload.materialType || 'Lecture Note').slice(0, 80),
    String(payload.description || '').slice(0, 2000), fileUrl, user.name, new Date().toISOString()
  ]);
  return { success: true, message: 'Course material published.' };
}

function getMaterials(token) {
  const user = requireUser_(token);
  const all = getRows_(SHEETS.MATERIALS);
  if (user.role === 'Student') {
    const myCourses = new Set(getRows_(SHEETS.ENROLMENTS)
      .filter(e => normaliseMatric_(e.matricno) === normaliseMatric_(user.userID))
      .map(e => e.courseid));
    return { success: true, materials: all.filter(m => myCourses.has(m.courseid)) };
  }
  if (user.role === 'Lecturer') {
    const myCourses = new Set(getRows_(SHEETS.COURSES)
      .filter(c => getCourseLecturerIds_(c).indexOf(String(user.userID).trim().toUpperCase()) >= 0)
      .map(c => c.courseid));
    return { success: true, materials: all.filter(m => myCourses.has(m.courseid)) };
  }
  return { success: true, materials: all };
}

/* ============================================================
   6. ATTENDANCE TOTALS, PERCENTAGE & ELIGIBLE CANDIDATES PDF (>= 75%)
   ============================================================ */
function recordAttendance(token, payload) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  payload = payload || {};
  const crs = requireCourseAccess_(user, payload.courseID);
  const lectureDate = String(payload.lectureDate || '').trim();
  const entries = Array.isArray(payload.entries) ? payload.entries : [];
  if (!lectureDate || !entries.length) return { success: false, message: 'Lecture date and attendance entries are required.' };
  const enrolments = getRows_(SHEETS.ENROLMENTS).filter(e => String(e.courseid) === String(crs.courseid));
  const sh = getDB_().getSheetByName(SHEETS.ATTENDANCE);
  const existing = getRows_(SHEETS.ATTENDANCE);
  let saved = 0;
  entries.forEach(ent => {
    const matric = normaliseMatric_(ent.matricNo);
    const enrol = enrolments.find(e => normaliseMatric_(e.matricno) === matric);
    const status = String(ent.status || '').trim().toUpperCase();
    if (!enrol || ['PRESENT','ABSENT','LATE'].indexOf(status) === -1) return;
    const old = existing.find(a => String(a.courseid) === String(crs.courseid) &&
      String(a.lecturedate).trim() === lectureDate && normaliseMatric_(a.matricno) === matric);
    const row = [old ? old.attendanceid : makeId_('ATT'), crs.courseid, lectureDate,
      String(payload.topic || 'Lecture'), enrol.studentid || ent.studentID || '', matric, status, user.userID];
    if (old) sh.getRange(old._row, 1, 1, row.length).setValues([row]);
    else {
      sh.appendRow(row);
      existing.push({ attendanceid: row[0], courseid: row[1], lecturedate: row[2], matricno: row[5], status: row[6], _row: sh.getLastRow() });
    }
    saved++;
  });
  return { success: true, message: 'Attendance saved/updated for ' + saved + ' enrolled student(s).', saved: saved };
}


/**
 * Import a complete attendance sheet (CSV is parsed by the mobile frontend).
 * Validates every matriculation number and status before writing any records.
 * Existing course/date/student records are updated rather than duplicated.
 */
function bulkUploadAttendance(token, payload) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  payload = payload || {};
  const crs = requireCourseAccess_(user, payload.courseID);
  const date = String(payload.lectureDate || '').trim();
  const entries = Array.isArray(payload.entries) ? payload.entries : [];
  if (!date || !entries.length) return { success: false, message: 'Select a lecture date and provide attendance rows.' };

  const enrolled = new Set(getRows_(SHEETS.ENROLMENTS)
    .filter(e => String(e.courseid) === String(crs.courseid))
    .map(e => normaliseMatric_(e.matricno)));
  const invalid = [];
  const clean = [];
  const seen = new Set();
  entries.forEach((entry, i) => {
    const matric = normaliseMatric_(entry.matricNo);
    const status = String(entry.status || '').trim().toUpperCase();
    if (!matric) invalid.push('Row ' + (i + 1) + ': missing matriculation number');
    else if (seen.has(matric)) invalid.push('Row ' + (i + 1) + ': duplicate matriculation number ' + matric);
    else if (!enrolled.has(matric)) invalid.push('Row ' + (i + 1) + ': ' + matric + ' is not enrolled in ' + crs.coursecode);
    else if (['PRESENT', 'ABSENT', 'LATE'].indexOf(status) === -1) invalid.push('Row ' + (i + 1) + ': invalid status for ' + matric);
    else { seen.add(matric); clean.push({ matricNo: matric, status: status }); }
  });
  if (invalid.length) return { success: false, message: 'No attendance was imported. Correct these entries: ' + invalid.slice(0, 12).join('; ') + (invalid.length > 12 ? '; and ' + (invalid.length - 12) + ' more.' : '') };
  return recordAttendance(token, {
    courseID: crs.courseid,
    lectureDate: date,
    topic: String(payload.topic || 'Lecture'),
    entries: clean
  });
}

function computeCourseAttendanceStats_(courseID) {
  const crs = getRows_(SHEETS.COURSES).find(c => c.courseid === courseID) || {};
  const roster = getRows_(SHEETS.ENROLMENTS).filter(e => e.courseid === courseID);
  const logs = getRows_(SHEETS.ATTENDANCE).filter(a => a.courseid === courseID);

  // Distinct lecture dates held for this course
  const distinctDates = new Set(logs.map(l => String(l.lecturedate).trim()).filter(Boolean));
  const totalLecturesHeld = distinctDates.size;

  const list = roster.map(stu => {
    const matric = normaliseMatric_(stu.matricno);
    const stuLogs = logs.filter(l => normaliseMatric_(l.matricno) === matric);
    const byDate = {};
    stuLogs.forEach(l => { byDate[String(l.lecturedate).trim()] = String(l.status || '').toUpperCase(); });
    const attended = Object.keys(byDate).filter(d => byDate[d] === 'PRESENT' || byDate[d] === 'LATE').length;
    const totalRecords = totalLecturesHeld;
    const pct = totalRecords > 0 ? Number(((attended / totalRecords) * 100).toFixed(1)) : 0;
    return {
      studentID: stu.studentid,
      matricNo: stu.matricno,
      studentName: stu.studentname,
      totalAttended: attended,
      totalLectures: totalRecords,
      percentage: pct,
      eligible: pct >= 75
    };
  });

  return { course: crs, totalLecturesHeld: totalLecturesHeld, summary: list };
}

function getCourseAttendanceSummary(token, courseID) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  requireCourseAccess_(user, courseID);
  const data = computeCourseAttendanceStats_(courseID);
  return {
    success: true,
    courseCode: data.course.coursecode || courseID,
    courseTitle: data.course.coursetitle || '',
    totalLecturesHeld: data.totalLecturesHeld,
    summary: data.summary
  };
}

function getAttendanceData(token, courseID) {
  const user = requireUser_(token);
  const all = getRows_(SHEETS.ATTENDANCE);
  const courses = getRows_(SHEETS.COURSES);

  if (user.role === 'Student') {
    const matric = String(user.userID).toUpperCase();
    const myEnrols = getRows_(SHEETS.ENROLMENTS).filter(e => String(e.matricno).toUpperCase() === matric);
    const summary = myEnrols.map(en => {
      const crs = courses.find(c => c.courseid === en.courseid) || {};
      const stats = computeCourseAttendanceStats_(en.courseid);
      const mine = stats.summary.find(s => String(s.matricNo).toUpperCase() === matric) || { totalAttended: 0, totalLectures: 0, percentage: 0, eligible: false };
      return {
        courseCode: crs.coursecode || en.courseid,
        courseTitle: crs.coursetitle || '',
        totalClasses: mine.totalLectures,
        presentCount: mine.totalAttended,
        percentage: mine.percentage,
        eligible: mine.percentage >= 75
      };
    });
    return { success: true, summary: summary };
  }

  let allowedCourseIds;
  if (user.role === 'Lecturer') {
    allowedCourseIds = new Set(courses.filter(c => getCourseLecturerIds_(c).indexOf(String(user.userID).trim().toUpperCase()) >= 0).map(c => c.courseid));
  }
  const targetCrs = all.filter(a => (!courseID || a.courseid === courseID) &&
    (user.role !== 'Lecturer' || allowedCourseIds.has(a.courseid)));
  if (courseID && user.role === 'Lecturer') requireCourseAccess_(user, courseID);
  return { success: true, records: targetCrs.slice(-100).reverse() };
}

function generateEligibleCandidatesPDF(token, courseID) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  requireCourseAccess_(user, courseID);
  const data = computeCourseAttendanceStats_(courseID);
  if (!data.course || !data.course.courseid) return { success: false, message: 'Select a valid course.' };
  const threshold = Number(getSetting_('MinimumAttendance', 75)) || 75;
  const eligible = data.summary.filter(s => s.percentage >= threshold).length;
  const rowsHtml = data.summary.length ? data.summary.map((s, i) => `
    <tr><td>${i + 1}</td><td><strong>${escapeHtml_(s.matricNo)}</strong></td>
    <td>${escapeHtml_(s.studentName)}</td><td>${s.totalAttended} / ${s.totalLectures}</td>
    <td><strong>${s.percentage}%</strong></td><td style="font-weight:bold;color:${s.percentage >= threshold ? '#166534' : '#991b1b'}">${s.percentage >= threshold ? 'ELIGIBLE' : 'NOT ELIGIBLE'}</td></tr>`).join('')
    : '<tr><td colspan="6">No students are enrolled in this course.</td></tr>';
  const html = `<html><body style="font-family:Arial,sans-serif;padding:24px;color:#111;">
    <h2 style="margin-bottom:4px;color:#123b5d;">University of Delta, Agbor — Department of Psychology</h2>
    <h3 style="margin-top:0;">Examination Attendance Eligibility Register</h3>
    <p><strong>Course:</strong> ${escapeHtml_(data.course.coursecode)} — ${escapeHtml_(data.course.coursetitle)}<br>
    <strong>Attendance threshold:</strong> ${threshold}% &nbsp; | &nbsp; <strong>Eligible:</strong> ${eligible} &nbsp; | &nbsp; <strong>Not eligible:</strong> ${data.summary.length - eligible} &nbsp; | &nbsp; <strong>Enrolled:</strong> ${data.summary.length}</p>
    <table border="1" cellspacing="0" cellpadding="8" width="100%" style="border-collapse:collapse;font-size:12px;">
      <thead><tr style="background:#e6eef5;"><th>S/N</th><th>Matriculation No</th><th>Student Name</th><th>Classes Attended</th><th>Attendance %</th><th>Examination Status</th></tr></thead>
      <tbody>${rowsHtml}</tbody></table>
    <p style="margin-top:18px;font-size:11px;color:#555;">Eligibility is calculated from attendance recorded for this course only. Generated via psyc@UNIDEL on ${new Date().toLocaleString()}.</p>
    <p style="margin-top:24px;">______________________________<br>Authorised Lecturer / Departmental Officer</p>
  </body></html>`;
  const pdfBlob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
  return { success: true, fileName: `${String(data.course.coursecode).replace(/\s+/g, '_')}_Exam_Eligibility_Register.pdf`, base64Pdf: Utilities.base64Encode(pdfBlob.getBytes()), count: eligible, total: data.summary.length };
}

/* ============================================================
   7. CBT ASSESSMENTS & EXAMS (SCHEDULE, ATTEMPTS, SCORING MODE, FEEDBACK)
   ============================================================ */
function createTest(token, payload) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  const crs = requireCourseAccess_(user, payload.courseID);
  if (!String(payload.title || '').trim()) return { success: false, message: 'Test title is required.' };

  const testID = makeId_('TST');
  const testSheet = getDB_().getSheetByName(SHEETS.TESTS);
  testSheet.appendRow([
    testID, crs.courseid, crs.coursecode, payload.kind || 'ASSESSMENT', payload.title,
    Number(payload.durationMinutes || 30), Number(payload.totalMarks || 30),
    Number(payload.maxAttempts != null ? payload.maxAttempts : 1),
    payload.scoringMethod || 'HIGHEST', payload.feedbackMode || 'SCORE_AND_ANSWERS',
    payload.openDateTime || '', payload.closeDateTime || '', payload.shuffleQuestions === 'YES' ? 'YES' : (payload.shuffle || 'NO'),
    payload.instructions || '', 'OPEN', new Date().toISOString()
  ]);
  const newRow = testSheet.getLastRow();
  setCellByHeader_(SHEETS.TESTS, newRow, 'ComponentGroup', String(payload.componentGroup || '').trim());
  setCellByHeader_(SHEETS.TESTS, newRow, 'ComponentGradingMethod', payload.componentGradingMethod || 'AVERAGE');
  setCellByHeader_(SHEETS.TESTS, newRow, 'SelectedTestID', payload.selectedTestID || '');
  setCellByHeader_(SHEETS.TESTS, newRow, 'QuestionCount', Math.max(0, Number(payload.questionCount || 0)));
  setCellByHeader_(SHEETS.TESTS, newRow, 'ShuffleOptions', payload.shuffleOptions === 'YES' ? 'YES' : 'NO');
  return { success: true, message: 'Quiz / CBT Examination created and scheduled.' };
}

function updateTestSettings(token, testID, payload) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  const tests = getRows_(SHEETS.TESTS);
  const t = tests.find(x => x.testid === testID);
  if (!t) return { success: false, message: 'Test not found.' };
  requireCourseAccess_(user, t.courseid);

  const sh = getDB_().getSheetByName(SHEETS.TESTS);
  sh.getRange(t._row, 8).setValue(Number(payload.maxAttempts != null ? payload.maxAttempts : t.maxattempts));
  sh.getRange(t._row, 9).setValue(payload.scoringMethod || t.scoringmethod || 'HIGHEST');
  sh.getRange(t._row, 10).setValue(payload.feedbackMode || t.feedbackmode || 'SCORE_AND_ANSWERS');
  if (payload.openDateTime !== undefined) sh.getRange(t._row, 11).setValue(payload.openDateTime);
  if (payload.closeDateTime !== undefined) sh.getRange(t._row, 12).setValue(payload.closeDateTime);
  if (payload.componentGroup !== undefined) setCellByHeader_(SHEETS.TESTS, t._row, 'ComponentGroup', String(payload.componentGroup || '').trim());
  if (payload.componentGradingMethod !== undefined) setCellByHeader_(SHEETS.TESTS, t._row, 'ComponentGradingMethod', payload.componentGradingMethod);
  if (payload.selectedTestID !== undefined) setCellByHeader_(SHEETS.TESTS, t._row, 'SelectedTestID', payload.selectedTestID);
  if (payload.shuffleQuestions !== undefined) setCellByHeader_(SHEETS.TESTS, t._row, 'Shuffle', payload.shuffleQuestions === 'YES' ? 'YES' : 'NO');
  if (payload.shuffleOptions !== undefined) setCellByHeader_(SHEETS.TESTS, t._row, 'ShuffleOptions', payload.shuffleOptions === 'YES' ? 'YES' : 'NO');
  if (payload.questionCount !== undefined) setCellByHeader_(SHEETS.TESTS, t._row, 'QuestionCount', Math.max(0, Number(payload.questionCount || 0)));

  return { success: true, message: 'Test policy, schedule, feedback, and component grading settings updated.' };
}

function addQuestion(token,payload){
 const user=requireUser_(token,['Lecturer','Administrator']);
 if(!payload.testID||!payload.questionText)return{success:false,message:'Test and question text are required.'};
 const test=getRows_(SHEETS.TESTS).find(t=>String(t.testid)===String(payload.testID));if(!test)return{success:false,message:'Test not found.'};requireCourseAccess_(user,test.courseid);
 const type=String(payload.questionType||'MCQ').toUpperCase(),allowed=['MCQ','TRUE_FALSE','FILL_GAP','MATCHING','MULTI_CORRECT','MULTI_INCORRECT'];if(allowed.indexOf(type)<0)return{success:false,message:'Unsupported question type.'};
 const data=buildQuestionAnswerData_(payload,type);
 if(type==='FILL_GAP'&&!data.acceptedAnswers.length)return{success:false,message:'Enter at least one accepted answer.'};
 if(type==='MATCHING'&&data.matchingPairs.length<2)return{success:false,message:'Enter at least two matching pairs.'};
 if((type==='MULTI_CORRECT'||type==='MULTI_INCORRECT')&&!data.correctKeys.length)return{success:false,message:'Select at least one correct key.'};
 const sh=getDB_().getSheetByName(SHEETS.QUESTIONS);sh.appendRow([makeId_('QST'),payload.testID,type,payload.questionText,payload.optionA||'',payload.optionB||'',payload.optionC||'',payload.optionD||'',String(payload.correctAnswer||'').trim(),Number(payload.marks||1)]);
 setCellByHeader_(SHEETS.QUESTIONS,sh.getLastRow(),'AnswerDataJSON',JSON.stringify(data));return{success:true,message:'Question saved.'};
}

function saveBankQuestion(token,payload){
 const user=requireUser_(token,['Lecturer','Administrator']);payload=payload||{};const course=requireCourseAccess_(user,payload.courseID);
 if(!String(payload.questionText||'').trim())return{success:false,message:'Question text is required.'};
 const type=String(payload.questionType||'MCQ').toUpperCase(),allowed=['MCQ','TRUE_FALSE','FILL_GAP','MATCHING','MULTI_CORRECT','MULTI_INCORRECT'];if(allowed.indexOf(type)<0)return{success:false,message:'Unsupported question type.'};
 const data=buildQuestionAnswerData_(payload,type);
 if(type==='FILL_GAP'&&!data.acceptedAnswers.length)return{success:false,message:'Enter at least one accepted answer.'};
 if(type==='MATCHING'&&data.matchingPairs.length<2)return{success:false,message:'Enter at least two matching pairs.'};
 if((type==='MULTI_CORRECT'||type==='MULTI_INCORRECT')&&!data.correctKeys.length)return{success:false,message:'Select at least one correct key.'};
 const id=makeId_('BANK'),now=new Date().toISOString(),sh=getDB_().getSheetByName(SHEETS.TEST_BANK);
 sh.appendRow([id,course.courseid,user.userID,String(payload.sharing||'PRIVATE').toUpperCase()==='SHARED'?'SHARED':'PRIVATE',String(payload.topic||'').trim(),String(payload.difficulty||'MEDIUM').toUpperCase(),type,String(payload.questionText).trim(),payload.optionA||'',payload.optionB||'',payload.optionC||'',payload.optionD||'',String(payload.correctAnswer||'').trim(),Math.max(0,Number(payload.marks||1)),String(payload.explanation||'').trim(),'ACTIVE',now,now]);
 setCellByHeader_(SHEETS.TEST_BANK,sh.getLastRow(),'AnswerDataJSON',JSON.stringify(data));return{success:true,questionID:id,message:'Question saved to the Test Bank.'};
}
function listBankQuestions(token, courseID) {
  const user = requireUser_(token,['Lecturer','Administrator']);
  requireCourseAccess_(user,courseID);
  const rows = getRows_(SHEETS.TEST_BANK).filter(q=>String(q.courseid)===String(courseID) && String(q.status||'ACTIVE').toUpperCase()==='ACTIVE' &&
    (user.role==='Administrator' || String(q.ownerid).toUpperCase()===String(user.userID).toUpperCase() || String(q.sharing).toUpperCase()==='SHARED'));
  return {success:true,questions:rows};
}
function addBankQuestionsToTest(token, testID, bankQuestionIDs) {
  const user = requireUser_(token,['Lecturer','Administrator']);
  const test = getRows_(SHEETS.TESTS).find(t=>String(t.testid)===String(testID));
  if (!test) return {success:false,message:'Test not found.'};
  requireCourseAccess_(user,test.courseid);
  const ids = Array.isArray(bankQuestionIDs) ? bankQuestionIDs.map(String) : [String(bankQuestionIDs||'')];
  const bank = getRows_(SHEETS.TEST_BANK), sh = getDB_().getSheetByName(SHEETS.QUESTIONS);
  let added=0;
  ids.forEach(id=>{
    const q=bank.find(x=>String(x.bankquestionid)===id && String(x.courseid)===String(test.courseid) &&
      (user.role==='Administrator' || String(x.ownerid).toUpperCase()===String(user.userID).toUpperCase() || String(x.sharing).toUpperCase()==='SHARED'));
    if(!q) return;
    sh.appendRow([makeId_('QST'),testID,q.questiontype,q.questiontext,q.optiona,q.optionb,q.optionc,q.optiond,q.correctanswer,Number(q.marks||1)]);
    setCellByHeader_(SHEETS.QUESTIONS,sh.getLastRow(),'Explanation',q.explanation||'');
    setCellByHeader_(SHEETS.QUESTIONS,sh.getLastRow(),'AnswerDataJSON',q.answerdatajson||'{}');
    added++;
  });
  return {success:true,added:added,message:added+' question(s) added to the assessment.'};
}
function updateBankQuestion(token, questionID, payload) {
  const user=requireUser_(token,['Lecturer','Administrator']), rows=getRows_(SHEETS.TEST_BANK);
  const q=rows.find(x=>String(x.bankquestionid)===String(questionID));
  if(!q) return {success:false,message:'Saved question not found.'};
  requireCourseAccess_(user,q.courseid);
  if(user.role!=='Administrator'&&String(q.ownerid).toUpperCase()!==String(user.userID).toUpperCase()) return {success:false,message:'Only the owner or Administrator can edit this question.'};
  payload=payload||{};
  const type=String(payload.questionType||q.questiontype||'MCQ').toUpperCase(),data=buildQuestionAnswerData_(payload,type);
  if(type==='FILL_GAP'&&!data.acceptedAnswers.length)return{success:false,message:'Enter at least one accepted answer.'};
  if(type==='MATCHING'&&data.matchingPairs.length<2)return{success:false,message:'Enter at least two matching pairs.'};
  if((type==='MULTI_CORRECT'||type==='MULTI_INCORRECT')&&!data.correctKeys.length)return{success:false,message:'Select at least one correct key.'};
  const fields={Topic:payload.topic,Difficulty:payload.difficulty,Sharing:payload.sharing,QuestionType:type,QuestionText:payload.questionText,OptionA:payload.optionA,OptionB:payload.optionB,OptionC:payload.optionC,OptionD:payload.optionD,CorrectAnswer:payload.correctAnswer,AnswerDataJSON:JSON.stringify(data),Marks:payload.marks,Explanation:payload.explanation,UpdatedAt:new Date().toISOString()};
  Object.keys(fields).forEach(h=>{if(fields[h]!==undefined)setCellByHeader_(SHEETS.TEST_BANK,q._row,h,fields[h]);});
  return {success:true,message:'Test Bank question updated.'};
}
function archiveBankQuestion(token, questionID) {
  const user=requireUser_(token,['Lecturer','Administrator']), q=getRows_(SHEETS.TEST_BANK).find(x=>String(x.bankquestionid)===String(questionID));
  if(!q) return {success:false,message:'Saved question not found.'};
  requireCourseAccess_(user,q.courseid);
  if(user.role!=='Administrator'&&String(q.ownerid).toUpperCase()!==String(user.userID).toUpperCase()) return {success:false,message:'Only the owner or Administrator can archive this question.'};
  setCellByHeader_(SHEETS.TEST_BANK,q._row,'Status','ARCHIVED');
  return {success:true,message:'Question archived. Existing assessments are not affected.'};
}
function shuffleArray_(items) {
  const a=items.slice();
  for(let i=a.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));const t=a[i];a[i]=a[j];a[j]=t;}
  return a;
}
function getTestWindowState_(t) {
  const now = Date.now();
  const openMs = t.opendatetime ? new Date(t.opendatetime).getTime() : NaN;
  const closeMs = t.closedatetime ? new Date(t.closedatetime).getTime() : NaN;

  if (!isNaN(openMs) && now < openMs) return { isOpen: false, statusLabel: 'UPCOMING' };
  if (!isNaN(closeMs) && now > closeMs) return { isOpen: false, statusLabel: 'CLOSED' };
  return { isOpen: true, statusLabel: 'OPEN' };
}

function computeOfficialScoreForStudent_(test, studentAttempts) {
  if (!studentAttempts || !studentAttempts.length) return null;
  const sorted = studentAttempts.slice().sort((a, b) => Number(a.attemptnumber) - Number(b.attemptnumber));
  const scores = sorted.map(a => Number(a.score || 0));
  const mode = String(test.scoringmethod || 'HIGHEST').toUpperCase();

  let finalScore = scores[0];
  if (mode === 'HIGHEST') finalScore = Math.max.apply(null, scores);
  else if (mode === 'AVERAGE') finalScore = scores.reduce((acc, v) => acc + v, 0) / scores.length;
  else if (mode === 'FIRST') finalScore = scores[0];
  else if (mode === 'SECOND') finalScore = scores.length >= 2 ? scores[1] : scores[0];
  else if (mode === 'LAST') finalScore = scores[scores.length - 1];

  return Number(finalScore.toFixed(2));
}

function getTestsAndAttempts(token) {
  const user = requireUser_(token);
  let tests = getRows_(SHEETS.TESTS).map(t => {
    const win = getTestWindowState_(t);
    return Object.assign({}, t, { isOpen: win.isOpen, windowStatus: win.statusLabel });
  });
  let attempts = getRows_(SHEETS.ATTEMPTS);

  if (user.role === 'Student') {
    const matric = normaliseMatric_(user.userID);
    const myCourses = new Set(getRows_(SHEETS.ENROLMENTS)
      .filter(e => normaliseMatric_(e.matricno) === matric)
      .map(e => e.courseid));
    tests = tests.filter(t => myCourses.has(t.courseid));
    attempts = attempts.filter(a => normaliseMatric_(a.matricno) === matric);
  } else if (user.role === 'Lecturer') {
    const myCourses = new Set(getRows_(SHEETS.COURSES)
      .filter(c => getCourseLecturerIds_(c).indexOf(String(user.userID).trim().toUpperCase()) >= 0)
      .map(c => c.courseid));
    tests = tests.filter(t => myCourses.has(t.courseid));
    attempts = attempts.filter(a => myCourses.has(a.courseid || (tests.find(t => t.testid === a.testid) || {}).courseid));
  }

  // Compute Official Score table using each test's ScoringMethod
  const officialScores = [];
  tests.forEach(t => {
    const tAttempts = attempts.filter(a => a.testid === t.testid);
    const byStudent = {};
    tAttempts.forEach(a => {
      const m = String(a.matricno).toUpperCase();
      if (!byStudent[m]) byStudent[m] = [];
      byStudent[m].push(a);
    });
    Object.keys(byStudent).forEach(matric => {
      const list = byStudent[matric];
      const computed = computeOfficialScoreForStudent_(t, list);
      const fbMode = String(t.feedbackmode || 'SCORE_AND_ANSWERS').toUpperCase();
      const hideFromStudent = user.role === 'Student' && fbMode === 'HIDDEN';

      officialScores.push({
        testID: t.testid,
        courseCode: t.coursecode,
        testTitle: t.title,
        matricNo: matric,
        attemptsTaken: list.length,
        scoringMethod: t.scoringmethod || 'HIGHEST',
        feedbackMode: fbMode,
        officialScore: hideFromStudent ? 'Withheld by Lecturer' : `${computed} / ${t.totalmarks}`
      });
    });
  });

  // Hide raw attempt scores if lecturer chose HIDDEN feedback mode
  const visibleAttempts = attempts.map(a => {
    if (user.role !== 'Student') return a;
    const t = tests.find(x => x.testid === a.testid) || {};
    if (String(t.feedbackmode).toUpperCase() === 'HIDDEN') {
      return Object.assign({}, a, { score: 'Withheld' });
    }
    return a;
  });

  return {
    success: true,
    tests: tests,
    attempts: visibleAttempts,
    officialScores: officialScores
  };
}

function getQuestionsForTest(token, testID) {
  const user = requireUser_(token), test = getRows_(SHEETS.TESTS).find(t => String(t.testid) === String(testID));
  if (!test) return { success: false, message: 'Test not found.' };
  let session = null, optionMap = {};
  if (user.role === 'Student') {
    if (!isStudentEnrolled_(user.userID, test.courseid)) return { success: false, message: 'You are not enrolled in this course.' };
    if (String(test.status || 'OPEN').toUpperCase() !== 'OPEN') return { success: false, message: 'This test is not available.' };
    const win = getTestWindowState_(test);
    if (!win.isOpen) return { success: false, message: win.statusLabel === 'UPCOMING' ? 'This test has not opened yet.' : 'This test window has already closed.' };
    const matric = normaliseMatric_(user.userID);
    const prev = getRows_(SHEETS.ATTEMPTS).filter(a => String(a.testid) === String(testID) && normaliseMatric_(a.matricno) === matric);
    const maxAtt = Number(test.maxattempts != null ? test.maxattempts : 1);
    if (maxAtt > 0 && prev.length >= maxAtt) return { success: false, message: 'You have used all ' + maxAtt + ' allowed attempt(s).' };
    const active = getRows_(SHEETS.TEST_SESSIONS).filter(s=>String(s.testid)===String(testID)&&normaliseMatric_(s.studentid)===matric&&String(s.submitted).toUpperCase()!=='YES'&&new Date(s.expiresat).getTime()>Date.now());
    if (active.length) session=active[active.length-1];
  } else if (user.role === 'Lecturer') requireCourseAccess_(user, test.courseid);
  else if (user.role !== 'Administrator') return { success:false,message:'Access denied.' };
  const all = getRows_(SHEETS.QUESTIONS).filter(q => String(q.testid) === String(testID));
  if (!all.length) return { success:false,message:'No questions have been added to this assessment.' };
  let qs;
  if (user.role==='Student') {
    if (session) {
      let ids=[]; try{ids=JSON.parse(String(session.questionids||'[]'));}catch(e){}
      try{optionMap=JSON.parse(String(session.optionorderjson||'{}'))||{};}catch(e){}
      const map=new Map(all.map(q=>[String(q.questionid),q]));
      qs=ids.map(id=>map.get(String(id))).filter(Boolean);
    } else {
      qs=all.slice();
      const randomOrder=String(test.shuffle||'NO').toUpperCase()==='YES';
      const count=Math.max(0,Number(test.questioncount||0));
      if(count>0&&count<qs.length){qs=shuffleArray_(qs).slice(0,count);if(!randomOrder)qs.sort((a,b)=>all.indexOf(a)-all.indexOf(b));}
      else if(randomOrder) qs=shuffleArray_(qs);
      if(String(test.shuffleoptions||'NO').toUpperCase()==='YES') {
        qs.forEach(q=>{
          const typ=String(q.questiontype||'').toUpperCase();
          if(['MCQ','MULTI_CORRECT','MULTI_INCORRECT'].indexOf(typ)>=0){const opts=['A','B','C','D'].filter(k=>String(q['option'+k]||'').trim());optionMap[String(q.questionid)]=shuffleArray_(opts);}
          else if(typ==='MATCHING'){const ad=parseAnswerData_(q.answerdatajson);optionMap[String(q.questionid)]=shuffleArray_((ad.matchingPairs||[]).map(p=>p.right));}
        });
      }
      const now=new Date(), expiry=new Date(now.getTime()+(Math.max(5,Number(test.durationminutes||30))+10)*60000);
      getDB_().getSheetByName(SHEETS.TEST_SESSIONS).appendRow([makeId_('SES'),testID,normaliseMatric_(user.userID),JSON.stringify(qs.map(q=>q.questionid)),JSON.stringify(optionMap),now.toISOString(),expiry.toISOString(),'NO']);
    }
  } else qs=all;
  const out=qs.map(q=>{
    const obj={questionID:q.questionid,questionType:q.questiontype,questionText:q.questiontext,optionA:q.optiona,optionB:q.optionb,optionC:q.optionc,optionD:q.optiond,marks:q.marks};
    const typ=String(q.questiontype||'').toUpperCase(),ad=parseAnswerData_(q.answerdatajson);
    if(user.role==='Student'&&typ==='MATCHING'){const ps=ad.matchingPairs||[];obj.matchingPrompts=ps.map(p=>p.left);obj.matchingChoices=optionMap[String(q.questionid)]||ps.map(p=>p.right);}
    if(user.role==='Student'&&typ!=='MATCHING'&&optionMap[String(q.questionid)])obj.optionOrder=optionMap[String(q.questionid)].map(k=>({key:k.toLowerCase(),label:q['option'+k]}));
    return obj;
  });
  return {success:true,test:test,questions:out};
}
function submitCBTAttempt(token, testID, answers) {
  const user = requireUser_(token, ['Student']);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const test = getRows_(SHEETS.TESTS).find(t => String(t.testid) === String(testID));
    if (!test) return { success: false, message: 'Test not found.' };
    const matric = normaliseMatric_(user.userID);
    if (!isStudentEnrolled_(matric, test.courseid)) return { success: false, message: 'You are not enrolled in this course.' };
    if (String(test.status || 'OPEN').toUpperCase() !== 'OPEN') return { success: false, message: 'This test is not available.' };
    const win = getTestWindowState_(test);
    if (!win.isOpen) return { success: false, message: win.statusLabel === 'UPCOMING' ? 'This test has not opened yet.' : 'This test window has closed.' };

    const prevAttempts = getRows_(SHEETS.ATTEMPTS).filter(a => String(a.testid) === String(testID) && normaliseMatric_(a.matricno) === matric);
    const maxAtt = Number(test.maxattempts != null ? test.maxattempts : 1);
    if (maxAtt > 0 && prevAttempts.length >= maxAtt) return { success: false, message: 'Maximum allowed attempts (' + maxAtt + ') reached.' };

    const allQuestions = getRows_(SHEETS.QUESTIONS).filter(q => String(q.testid) === String(testID));
    const activeSession = getRows_(SHEETS.TEST_SESSIONS).filter(s=>String(s.testid)===String(testID)&&normaliseMatric_(s.studentid)===matric&&String(s.submitted).toUpperCase()!=='YES').slice(-1)[0];
    let selectedIds=[]; if(activeSession){try{selectedIds=JSON.parse(String(activeSession.questionids||'[]'));}catch(e){}}
    const questionMap=new Map(allQuestions.map(q=>[String(q.questionid),q]));
    const qs=selectedIds.length?selectedIds.map(id=>questionMap.get(String(id))).filter(Boolean):allQuestions;
    if (!qs.length) return { success: false, message: 'This test has no questions yet.' };
    let rawScore = 0, rawMax = 0;
    const reviewItems = [];
    qs.forEach((q, idx) => {
      const mk = Math.max(0, Number(q.marks || 1));
      rawMax += mk;
      const raw=answers&&answers[q.questionid],typ=String(q.questiontype||'').toUpperCase(),ad=parseAnswerData_(q.answerdatajson);let earned=0,isCorrect=false,givenRaw='';
      if(typ==='FILL_GAP'){givenRaw=String(raw==null?'':raw).trim();const accepted=(ad.acceptedAnswers&&ad.acceptedAnswers.length?ad.acceptedAnswers:String(q.correctanswer||'').split('|')).map(normaliseAnswerText_);isCorrect=!!givenRaw&&accepted.indexOf(normaliseAnswerText_(givenRaw))>=0;earned=isCorrect?mk:0;}
      else if(typ==='MATCHING'){const response=raw&&typeof raw==='object'?raw:{},pairs=ad.matchingPairs||[];let matched=0;pairs.forEach((p,i)=>{if(normaliseAnswerText_(response[String(i)])===normaliseAnswerText_(p.right))matched++;});earned=pairs.length?mk*matched/pairs.length:0;isCorrect=pairs.length>0&&matched===pairs.length;givenRaw=Object.keys(response).map(k=>(Number(k)+1)+': '+response[k]).join('; ')||'(No Answer)';}
      else if(typ==='MULTI_CORRECT'||typ==='MULTI_INCORRECT'){const selected=Array.isArray(raw)?raw.map(x=>String(x).toLowerCase()):[],correct=(ad.correctKeys||String(q.correctanswer||'').split(',')).map(x=>String(x).toLowerCase()),chosen=new Set(selected),target=new Set(correct),hits=correct.filter(x=>chosen.has(x)).length,extras=selected.filter(x=>!target.has(x)).length;isCorrect=chosen.size===target.size&&hits===target.size;earned=isCorrect?mk:(ad.markingMode==='EXACT'?0:Math.max(0,mk*(hits-extras)/Math.max(1,target.size)));givenRaw=selected.join(', ')||'(No Answer)';}
      else{givenRaw=String(raw==null?'':raw).trim();isCorrect=!!givenRaw&&givenRaw.toLowerCase()===String(q.correctanswer||'').trim().toLowerCase();earned=isCorrect?mk:0;}
      let keyDisplay=q.correctanswer||'';
      if(typ==='FILL_GAP')keyDisplay=(ad.acceptedAnswers||String(q.correctanswer||'').split('|')).join(' / ');
      else if(typ==='MATCHING')keyDisplay=(ad.matchingPairs||[]).map(p=>p.left+' => '+p.right).join('; ');
      else if(typ==='MULTI_CORRECT'||typ==='MULTI_INCORRECT')keyDisplay=(ad.correctKeys||[]).map(k=>k.toUpperCase()+'. '+String(q['option'+k.toUpperCase()]||'')).join('; ');
      rawScore+=earned;reviewItems.push({number:idx+1,questionText:q.questiontext,yourAnswer:givenRaw||'(No Answer)',correctAnswer:keyDisplay,isCorrect:isCorrect,earned:Number(earned.toFixed(2)),maxMarks:mk});
    });
    const maxScore = Number(test.totalmarks || rawMax);
    const scaledScore = rawMax > 0 ? Number(((rawScore / rawMax) * maxScore).toFixed(2)) : 0;
    const attNum = prevAttempts.length + 1;
    getDB_().getSheetByName(SHEETS.ATTEMPTS).appendRow([
      makeId_('ATT'), testID, user.userID, matric, attNum, scaledScore, maxScore, new Date().toISOString()
    ]);

    if(activeSession) setCellByHeader_(SHEETS.TEST_SESSIONS,activeSession._row,'Submitted','YES');
    const fbMode = String(test.feedbackmode || 'SCORE_AND_ANSWERS').toUpperCase();
    if (fbMode === 'HIDDEN') return { success: true, feedbackMode: 'HIDDEN', message: `Attempt #${attNum} submitted! Your lecturer has set scores to be released later.` };
    return {
      success: true, feedbackMode: fbMode, score: scaledScore, maxScore: maxScore, attemptNumber: attNum,
      review: fbMode === 'SCORE_AND_ANSWERS' ? reviewItems : [],
      message: `Attempt #${attNum} submitted! Score: ${scaledScore} / ${maxScore}`
    };
  } finally { lock.releaseLock(); }
}

/* ============================================================
   8. GRADING, RESULTS, ANNOUNCEMENTS & DISCUSSION FORUM
   ============================================================ */
function computeGrade_(total) {
  if (total >= 70) return ['A', 5];
  if (total >= 60) return ['B', 4];
  if (total >= 50) return ['C', 3];
  if (total >= 45) return ['D', 2];
  if (total >= 40) return ['E', 1];
  return ['F', 0];
}

function saveResult(token, payload) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  const crs = requireCourseAccess_(user, payload.courseID);
  const ca = Number(payload.caScore || 0);
  const exam = Number(payload.examScore || 0);
  if (!isFinite(ca) || !isFinite(exam) || ca < 0 || exam < 0 || ca > 100 || exam > 100) return { success: false, message: 'CA and examination scores must each be between 0 and 100.' };
  const total = Number((ca + exam).toFixed(2));
  const [grade, gp] = computeGrade_(total);
  const matric = normaliseMatric_(payload.matricNo);
  const stu = getRows_(SHEETS.STUDENTS).find(s => normaliseMatric_(s.matricno) === matric);
  if (!stu || !isStudentEnrolled_(matric, crs.courseid)) return { success: false, message: 'The student is not registered and enrolled in this course.' };

  const sh = getDB_().getSheetByName(SHEETS.RESULTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v).trim().toLowerCase());
  const academicSession = String(getSetting_('AcademicSession', '') || '').trim();
  const semester = String(getSetting_('Semester', crs.semester || '') || '').trim();
  const existing = getRows_(SHEETS.RESULTS).find(r =>
    String(r.courseid) === String(crs.courseid) &&
    normaliseMatric_(r.matricno) === matric &&
    String(r.academicsession || '').trim() === academicSession &&
    String(r.semester || '').trim() === semester
  );
  const rowVals = [
    existing ? existing.resultid : makeId_('RES'), crs.courseid, crs.coursecode, crs.coursetitle,
    Number(crs.creditunits || 2), stu.studentid || matric, matric,
    [stu.surname, stu.firstname, stu.othernames].filter(Boolean).join(' '), ca, exam, total, grade, gp, new Date().toISOString()
  ];
  const rowNumber = existing ? existing._row : sh.getLastRow() + 1;
  if (existing) sh.getRange(rowNumber, 1, 1, rowVals.length).setValues([rowVals]);
  else sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).setValues([rowVals.concat(Array(Math.max(0, sh.getLastColumn() - rowVals.length)).fill(''))]);
  const sessionCol = headers.indexOf('academicsession') + 1;
  const semesterCol = headers.indexOf('semester') + 1;
  if (sessionCol > 0) sh.getRange(rowNumber, sessionCol).setValue(academicSession);
  if (semesterCol > 0) sh.getRange(rowNumber, semesterCol).setValue(semester);
  return { success: true, message: `Result saved for ${matric}: ${total} (${grade}).` };
}

function gradeClassFromCGPA_(value) {
  const g = Number(value || 0);
  if (g >= 4.5) return 'First Class';
  if (g >= 3.5) return 'Second Class Upper';
  if (g >= 2.4) return 'Second Class Lower';
  if (g >= 2.0) return 'Third Class';
  if (g >= 1.0) return 'Pass';
  return 'Fail';
}

function gradeAverage_(rows) {
  let units = 0, weighted = 0;
  (rows || []).forEach(r => {
    const cu = Number(r.creditunits || 0), gp = Number(r.gradepoint || 0);
    if (cu > 0 && isFinite(gp)) { units += cu; weighted += cu * gp; }
  });
  return { units: units, points: weighted, value: units ? Number((weighted / units).toFixed(2)) : 0 };
}

function getGradeReportData_(user) {
  if (!user || !['Administrator', 'Lecturer'].includes(user.role)) throw new Error('Access denied. Grade reports are restricted to authorised staff.');
  const session = String(getSetting_('AcademicSession', '') || '').trim();
  const semester = String(getSetting_('Semester', '') || '').trim();
  const courses = getRows_(SHEETS.COURSES);
  const allowedIds = user.role === 'Administrator'
    ? new Set(courses.map(c => String(c.courseid)))
    : getLecturerCourseIds_(user.userID);
  const enrolments = getRows_(SHEETS.ENROLMENTS).filter(e => allowedIds.has(String(e.courseid)));
  const students = getRows_(SHEETS.STUDENTS);
  const results = getRows_(SHEETS.RESULTS);
  const byStudent = new Map();

  enrolments.forEach(e => {
    const matric = normaliseMatric_(e.matricno);
    if (!matric) return;
    if (!byStudent.has(matric)) {
      const st = students.find(s => normaliseMatric_(s.matricno) === matric) || {};
      byStudent.set(matric, {
        matricNo: matric,
        studentName: [st.surname, st.firstname, st.othernames].filter(Boolean).join(' ') || e.studentname || matric,
        programme: st.programme || '',
        level: st.level || '',
        courses: []
      });
    }
    const course = courses.find(c => String(c.courseid) === String(e.courseid)) || {};
    const allForCourse = results.filter(r => String(r.courseid) === String(e.courseid) && normaliseMatric_(r.matricno) === matric);
    const currentResult = allForCourse.filter(r => String(r.academicsession || '').trim() === session && String(r.semester || '').trim() === semester);
    const legacyResult = allForCourse.filter(r => !String(r.academicsession || '').trim() && !String(r.semester || '').trim());
    const result = currentResult.length ? currentResult[currentResult.length - 1] : (legacyResult.length ? legacyResult[legacyResult.length - 1] : null);
    byStudent.get(matric).courses.push({
      courseID: e.courseid, courseCode: course.coursecode || result && result.coursecode || '',
      courseTitle: course.coursetitle || result && result.coursetitle || '',
      creditUnits: Number(course.creditunits || result && result.creditunits || 0),
      caScore: result ? Number(result.cascore || 0) : '',
      examScore: result ? Number(result.examscore || 0) : '',
      totalScore: result ? Number(result.totalscore || 0) : '',
      grade: result ? String(result.grade || '') : '',
      gradePoint: result ? Number(result.gradepoint || 0) : '',
      resultSession: result ? String(result.academicsession || '') : '',
      resultSemester: result ? String(result.semester || '') : '',
      hasResult: !!result
    });
  });

  const report = Array.from(byStudent.values()).sort((a,b) => a.matricNo.localeCompare(b.matricNo));
  report.forEach(st => {
    const ownResults = results.filter(r => normaliseMatric_(r.matricno) === st.matricNo);
    const termResults = ownResults.filter(r =>
      String(r.academicsession || '').trim() === session &&
      String(r.semester || '').trim() === semester &&
      allowedIds.has(String(r.courseid))
    );
    const gpa = gradeAverage_(termResults);
    const cgpa = gradeAverage_(ownResults);
    st.gpa = gpa.value;
    st.gpaUnits = gpa.units;
    st.cgpa = cgpa.value;
    st.cgpaUnits = cgpa.units;
    st.flag = (gpa.value < 1 || cgpa.value < 1) ? 'WARNING: GPA/cGPA BELOW 1.00' : '';
    st.classOfDegree = gradeClassFromCGPA_(cgpa.value);
    st.courses.sort((a,b) => String(a.courseCode).localeCompare(String(b.courseCode)));
  });
  return { academicSession: session, semester: semester, students: report, accessScope: user.role === 'Administrator' ? 'All courses' : 'Courses assigned to this lecturer' };
}

function getGradeReport(token) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  return Object.assign({ success: true }, getGradeReportData_(user));
}

function exportGradeReportToExcel(token) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  const data = getGradeReportData_(user);
  const temp = SpreadsheetApp.create('psyc_UNIDEL_Grade_Report_' + Date.now());
  try {
    const summary = temp.getSheets()[0];
    summary.setName('Student Summary');
    const summaryValues = [['Matriculation Number','Student Name','Programme','Level','Academic Session','Semester','GPA','cGPA','Warning Flag','Class of Degree','GPA Credit Units','cGPA Credit Units']];
    data.students.forEach(st => summaryValues.push([st.matricNo,st.studentName,st.programme,st.level,data.academicSession,data.semester,st.gpa,st.cgpa,st.flag,st.classOfDegree,st.gpaUnits,st.cgpaUnits]));
    summary.getRange(1,1,summaryValues.length,summaryValues[0].length).setValues(summaryValues);
    const detail = temp.insertSheet('Course Grades');
    const detailValues = [['Matriculation Number','Student Name','Programme','Level','Course Code','Course Title','Credit Units','CA Score','Examination Score','Total Score','Grade','Grade Point','Result Session','Result Semester','GPA','cGPA','Warning Flag','Class of Degree']];
    data.students.forEach(st => st.courses.forEach(cr => detailValues.push([st.matricNo,st.studentName,st.programme,st.level,cr.courseCode,cr.courseTitle,cr.creditUnits,cr.caScore,cr.examScore,cr.totalScore,cr.grade,cr.gradePoint,cr.resultSession,cr.resultSemester,st.gpa,st.cgpa,st.flag,st.classOfDegree])));
    detail.getRange(1,1,detailValues.length,detailValues[0].length).setValues(detailValues);
    [summary,detail].forEach(sh => {
      sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold').setBackground('#123b5d').setFontColor('#ffffff').setWrap(true);
      sh.setFrozenRows(1); sh.autoResizeColumns(1,sh.getLastColumn());
      sh.getDataRange().setVerticalAlignment('middle');
    });
    summary.getRange(1,1,summaryValues.length,summaryValues[0].length).createFilter();
    const url = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=xlsx';
    const response = UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
    if (response.getResponseCode() !== 200) throw new Error('Grade report Excel export failed. Check Google Drive/Sheets authorisation and try again.');
    const blob=response.getBlob().setName('psyc_UNIDEL_Grade_Report.xlsx');
    return {success:true,fileName:blob.getName(),base64:Utilities.base64Encode(blob.getBytes()),students:data.students.length,courses:data.students.reduce((n,s)=>n+s.courses.length,0)};
  } finally {
    DriveApp.getFileById(temp.getId()).setTrashed(true);
  }
}

function getResults(token) {
  const user = requireUser_(token);
  const all = getRows_(SHEETS.RESULTS);
  if (user.role === 'Student') {
    const publication = String(getSetting_('ResultPublication', 'PUBLISHED')).trim().toUpperCase();
    if (['PUBLISHED','YES','TRUE','ON'].indexOf(publication) === -1) return { success: true, results: [], message: 'Results have not yet been published.' };
    const matric = normaliseMatric_(user.userID);
    return { success: true, results: all.filter(r => normaliseMatric_(r.matricno) === matric) };
  }
  if (user.role === 'Lecturer') {
    const myCourses = new Set(getRows_(SHEETS.COURSES)
      .filter(c => getCourseLecturerIds_(c).indexOf(String(user.userID).trim().toUpperCase()) >= 0)
      .map(c => c.courseid));
    return { success: true, results: all.filter(r => myCourses.has(r.courseid)) };
  }
  return { success: true, results: all };
}

function exportResultsToExcel(token, courseID) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  let rows = getRows_(SHEETS.RESULTS);
  if (user.role === 'Lecturer') {
    const myCourseIds = new Set(getRows_(SHEETS.COURSES).filter(c => getCourseLecturerIds_(c).indexOf(String(user.userID).trim().toUpperCase()) >= 0).map(c => String(c.courseid)));
    rows = rows.filter(r => myCourseIds.has(String(r.courseid)));
  }
  if (courseID) {
    const course = requireCourseAccess_(user, courseID);
    rows = rows.filter(r => String(r.courseid) === String(courseID));
  }
  const values = [['Course Code','Course Title','Credit Units','Matriculation Number','Student Name','CA Score','Examination Score','Total Score','Grade','Grade Point','Last Updated']];
  rows.forEach(r => values.push([r.coursecode, r.coursetitle, r.creditunits, r.matricno, r.studentname, r.cascore, r.examscore, r.totalscore, r.grade, r.gradepoint, r.updatedat]));
  const temp = SpreadsheetApp.create('psyc_UNIDEL_Results_Export_' + Date.now());
  try {
    const sh = temp.getSheets()[0]; sh.setName('Results');
    sh.getRange(1, 1, values.length, values[0].length).setValues(values);
    sh.getRange(1, 1, 1, values[0].length).setFontWeight('bold').setBackground('#123b5d').setFontColor('#ffffff');
    sh.setFrozenRows(1); sh.autoResizeColumns(1, values[0].length);
    const url = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=xlsx';
    const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) throw new Error('Excel export failed. Check Google Drive/Sheets authorisation and try again.');
    const blob = response.getBlob().setName('psyc_UNIDEL_Results.xlsx');
    return { success: true, fileName: blob.getName(), base64: Utilities.base64Encode(blob.getBytes()), count: rows.length };
  } finally {
    DriveApp.getFileById(temp.getId()).setTrashed(true);
  }
}

function getComponentGroupScores(token, courseID, groupName) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  requireCourseAccess_(user, courseID);
  const tests = getRows_(SHEETS.TESTS).filter(t => String(t.courseid) === String(courseID) && String(t.componentgroup || '').trim() === String(groupName || '').trim());
  const attempts = getRows_(SHEETS.ATTEMPTS);
  const students = getRows_(SHEETS.ENROLMENTS).filter(e => String(e.courseid) === String(courseID));
  const configured = tests.find(t => String(t.componentgradingmethod || '').trim()) || tests[0] || {};
  const method = String(configured.componentgradingmethod || 'AVERAGE').toUpperCase();
  const selectedId = String(tests.map(t => t.selectedtestid).find(Boolean) || '');
  return { success: true, groupName: groupName, gradingMethod: method, tests: tests.map(t => ({testID:t.testid,title:t.title,totalMarks:t.totalmarks})), scores: students.map(st => {
    const mat = normaliseMatric_(st.matricno);
    const perTest = tests.map(t => {
      const aa = attempts.filter(a => String(a.testid) === String(t.testid) && normaliseMatric_(a.matricno) === mat);
      return aa.length ? { testID:t.testid, score:computeOfficialScoreForStudent_(t,aa), maxScore:Number(t.totalmarks||0) } : null;
    }).filter(Boolean);
    let chosen = perTest;
    if (method === 'SPECIFIC') chosen = perTest.filter(x => x.testID === selectedId);
    const pct = chosen.length ? chosen.reduce((sum,x)=>sum+(x.maxScore ? x.score/x.maxScore : 0),0)/chosen.length : 0;
    return { matricNo:st.matricno, studentName:st.studentname, percentage:Number((pct*100).toFixed(2)), testsCount:chosen.length };
  }) };
}

function postAnnouncement(token, payload) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  getDB_().getSheetByName(SHEETS.ANNOUNCEMENTS).appendRow([
    makeId_('ANN'), payload.title, payload.audience || 'ALL', payload.body, user.name, user.role, new Date().toISOString()
  ]);
  return { success: true, message: 'Announcement published.' };
}

function getAnnouncements(token) {
  requireUser_(token);  return { success: true, announcements: getRows_(SHEETS.ANNOUNCEMENTS).reverse() };
}

function createDiscussionTopic(token, courseID, title, body) {
  const user = requireUser_(token);
  const crs = courseID && courseID !== 'GENERAL' ? getRows_(SHEETS.COURSES).find(c => String(c.courseid) === String(courseID)) : {};
  if (courseID && courseID !== 'GENERAL' && !crs) return { success: false, message: 'Course not found.' };
  if (user.role === 'Lecturer' && courseID && courseID !== 'GENERAL') requireCourseAccess_(user, courseID);
  if (user.role === 'Student' && courseID && courseID !== 'GENERAL' && !isStudentEnrolled_(user.userID, courseID)) return { success: false, message: 'You are not enrolled in this course.' };
  getDB_().getSheetByName(SHEETS.DISCUSSIONS).appendRow([
    makeId_('TOP'), courseID || 'GENERAL', crs.coursecode || 'GENERAL', title, body, user.name, user.role, new Date().toISOString()
  ]);
  return { success: true, message: 'Discussion topic posted.' };
}

function replyDiscussionTopic(token, topicID, body) {
  const user = requireUser_(token);
  const topic = getRows_(SHEETS.DISCUSSIONS).find(t => String(t.topicid) === String(topicID));
  if (!topic) return { success: false, message: 'Discussion topic not found.' };
  if (topic.courseid && topic.courseid !== 'GENERAL') {
    if (user.role === 'Lecturer') requireCourseAccess_(user, topic.courseid);
    if (user.role === 'Student' && !isStudentEnrolled_(user.userID, topic.courseid)) return { success: false, message: 'You are not enrolled in this course.' };
  }
  getDB_().getSheetByName(SHEETS.REPLIES).appendRow([
    makeId_('REP'), topicID, body, user.name, user.role, new Date().toISOString()
  ]);
  return { success: true, message: 'Reply posted.' };
}

function getDiscussions(token) {
  const user = requireUser_(token);
  let topics = getRows_(SHEETS.DISCUSSIONS);
  if (user.role === 'Student') {
    const myCourses = new Set(getRows_(SHEETS.ENROLMENTS).filter(e => normaliseMatric_(e.matricno) === normaliseMatric_(user.userID)).map(e => e.courseid));
    topics = topics.filter(t => !t.courseid || t.courseid === 'GENERAL' || myCourses.has(t.courseid));
  } else if (user.role === 'Lecturer') {
    const myCourses = new Set(getRows_(SHEETS.COURSES).filter(c => getCourseLecturerIds_(c).indexOf(String(user.userID).trim().toUpperCase()) >= 0).map(c => c.courseid));
    topics = topics.filter(t => !t.courseid || t.courseid === 'GENERAL' || myCourses.has(t.courseid));
  }
  topics = topics.reverse();
  const topicIds = new Set(topics.map(t => t.topicid));
  const replies = getRows_(SHEETS.REPLIES).filter(r => topicIds.has(r.topicid));
  return {
    success: true,
    topics: topics.map(t => Object.assign({}, t, {
      replies: replies.filter(r => r.topicid === t.topicid)
    }))
  };
}

/* ============================================================
   9. ADMIN DIRECTORY & GOVERNANCE
   ============================================================ */
function getAdminOverview(token) {
  requireUser_(token, ['Administrator']);
  return {
    success: true,
    counts: {
      students: getRows_(SHEETS.STUDENTS).length,
      lecturers: getRows_(SHEETS.LECTURERS).length,
      courses: getRows_(SHEETS.COURSES).length,
      enrolments: getRows_(SHEETS.ENROLMENTS).length,
      materials: getRows_(SHEETS.MATERIALS).length,
      tests: getRows_(SHEETS.TESTS).length,
      results: getRows_(SHEETS.RESULTS).length
    },
    users: getRows_(SHEETS.USERS).map(u => ({ userID: u.userid, name: u.name, role: u.role, active: u.active }))
  };
}

function toggleUserStatus(token, targetUserID, newStatus) {
  requireUser_(token, ['Administrator']);
  const users = getRows_(SHEETS.USERS);
  const found = users.find(u => String(u.userid).toUpperCase() === String(targetUserID).toUpperCase());
  if (!found) return { success: false, message: 'User not found.' };
  const status = String(newStatus || '').trim().toUpperCase();
  if (['YES','NO'].indexOf(status) === -1) return { success: false, message: 'Status must be YES or NO.' };
  getDB_().getSheetByName(SHEETS.USERS).getRange(found._row, 5).setValue(status);
  return { success: true, message: 'Updated status for ' + targetUserID + ' to ' + newStatus };
}


/* ============================================================
   VERSION 3.10 — EMAIL MESSAGING (NO MESSAGE-BODY STORAGE)
   ============================================================ */
function validEmailAddress_(value) {
  const s = String(value || '').trim();
  return /^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(s);
}

function resolvePortalEmailRecipient_(identifier) {
  const key = String(identifier || '').trim();
  if (!key) return null;
  const users = getRows_(SHEETS.USERS);
  const students = getRows_(SHEETS.STUDENTS);
  const lecturers = getRows_(SHEETS.LECTURERS);
  const user = users.find(x => String(x.userid || '').trim().toUpperCase() === key.toUpperCase());
  if (user) {
    if (String(user.active || '').toUpperCase() !== 'YES') return null;
    const student = students.find(x => normaliseMatric_(x.matricno) === normaliseMatric_(user.userid));
    const lecturer = lecturers.find(x => String(x.lecturerid || '').trim().toUpperCase() === key.toUpperCase());
    const email = String(user.email || student && student.email || lecturer && lecturer.email || '').trim();
    return email ? { email: email, name: user.name || key, registered: true } : null;
  }
  const student = students.find(x => normaliseMatric_(x.matricno) === normaliseMatric_(key));
  if (student) return validEmailAddress_(student.email) ? {email:String(student.email).trim(),name:[student.surname,student.firstname].filter(Boolean).join(' ') || key,registered:true} : null;
  const lecturer = lecturers.find(x => String(x.lecturerid || '').trim().toUpperCase() === key.toUpperCase() || String(x.staffid || '').trim().toUpperCase() === key.toUpperCase());
  if (lecturer) {
    const linked = users.find(x => String(x.userid || '').trim().toUpperCase() === String(lecturer.lecturerid || '').trim().toUpperCase() && String(x.active || '').toUpperCase() === 'YES');
    return linked && validEmailAddress_(lecturer.email) ? {email:String(lecturer.email).trim(),name:[lecturer.title,lecturer.surname,lecturer.firstname].filter(Boolean).join(' ') || key,registered:true} : null;
  }
  if (validEmailAddress_(key)) {
    const emailKey = key.toLowerCase();
    const registeredUser = users.find(x => String(x.email || '').trim().toLowerCase() === emailKey && String(x.active || '').toUpperCase() === 'YES');
    if (registeredUser) return {email:key,name:registeredUser.name || key,registered:true};
    const registeredStudent = students.find(x => String(x.email || '').trim().toLowerCase() === emailKey);
    if (registeredStudent) {
      const linked = users.find(x => normaliseMatric_(x.userid) === normaliseMatric_(registeredStudent.matricno) && String(x.active || '').toUpperCase() === 'YES');
      if (linked) return {email:key,name:[registeredStudent.surname,registeredStudent.firstname].filter(Boolean).join(' ') || key,registered:true};
    }
    const registeredLecturer = lecturers.find(x => String(x.email || '').trim().toLowerCase() === emailKey);
    if (registeredLecturer) {
      const linked = users.find(x => String(x.userid || '').trim().toUpperCase() === String(registeredLecturer.lecturerid || '').trim().toUpperCase() && String(x.active || '').toUpperCase() === 'YES');
      if (linked) return {email:key,name:[registeredLecturer.title,registeredLecturer.surname,registeredLecturer.firstname].filter(Boolean).join(' ') || key,registered:true};
    }
    return {email:key,name:key,registered:false};
  }
  return null;
}

function getMessagingSettings(token) {
  const user = requireUser_(token);
  return {
    success:true,
    allowExternalRecipients:String(getSetting_('AllowExternalEmailRecipients','YES')).trim().toUpperCase() !== 'NO',
    isAdministrator:user.role === 'Administrator'
  };
}

function setAllowExternalEmailRecipients(token, enabled) {
  requireUser_(token, ['Administrator']);
  const value = enabled ? 'YES' : 'NO';
  const ss = getDB_(), sh = ss.getSheetByName(SHEETS.SETTINGS), rows = getRows_(SHEETS.SETTINGS);
  const row = rows.find(r => String(r.key).trim().toLowerCase() === 'allowexternalemailrecipients');
  if (row) sh.getRange(row._row,2).setValue(value);
  else sh.appendRow(['AllowExternalEmailRecipients',value]);
  return {success:true,allowExternalRecipients:enabled,message:'External email recipients are now '+(enabled?'allowed.':'disabled.')};
}

function logEmailDelivery_(user, recipientCount, status) {
  try {
    const sh = getDB_().getSheetByName('EmailLog');
    if (sh) sh.appendRow([new Date().toISOString(),String(user.userID),String(user.name),String(user.role),Number(recipientCount || 0),String(status)]);
  } catch (e) { /* Delivery logging must never store message content or block delivery. */ }
}

function sendPortalEmail(token, payload) {
  const user = requireUser_(token);
  if (user.role === 'Student' && !getRows_(SHEETS.ENROLMENTS).some(e => normaliseMatric_(e.matricno) === normaliseMatric_(user.userID))) {
    return {success:false,message:'Only students enrolled in at least one course may use portal email messaging.'};
  }
  payload = payload || {};
  const subject = String(payload.subject || '').trim();
  const body = String(payload.body || '').trim();
  const rawRecipients = String(payload.recipients || '').split(/[;,\\n]+/).map(x => x.trim()).filter(Boolean);
  if (!subject || subject.length > 200) return {success:false,message:'Enter a subject of no more than 200 characters.'};
  if (!body || body.length > 10000) return {success:false,message:'Enter a message of no more than 10,000 characters.'};
  if (!rawRecipients.length) return {success:false,message:'Enter at least one recipient.'};
  if (rawRecipients.length > 10) return {success:false,message:'A maximum of 10 recipients is allowed per message.'};

  const allowExternal = String(getSetting_('AllowExternalEmailRecipients','YES')).trim().toUpperCase() !== 'NO';
  const recipients = [], seen = new Set();
  for (const item of rawRecipients) {
    const resolved = resolvePortalEmailRecipient_(item);
    if (!resolved) return {success:false,message:'Recipient not found or has no valid registered email: '+item};
    if (!resolved.registered && !allowExternal) return {success:false,message:'External email recipients have been disabled by the administrator.'};
    const emailKey = resolved.email.toLowerCase();
    if (!seen.has(emailKey)) { seen.add(emailKey); recipients.push(resolved); }
  }
  if (!recipients.length) return {success:false,message:'No valid unique recipients were supplied.'};

  const dateKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Africa/Lagos', 'yyyyMMdd');
  const propKey = 'PSYC_MAIL_DAILY_' + String(user.userID).replace(/[^A-Za-z0-9_-]/g,'_') + '_' + dateKey;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const sentToday = Number(props.getProperty(propKey) || 0);
    if (sentToday >= 20) return {success:false,message:'Daily sending limit reached (20 messages per user). Try again tomorrow.'};
    const senderEmail = (getRows_(SHEETS.USERS).find(x => String(x.userid).toUpperCase() === String(user.userID).toUpperCase()) || {}).email ||
      (getRows_(SHEETS.STUDENTS).find(x => normaliseMatric_(x.matricno) === normaliseMatric_(user.userID)) || {}).email ||
      (getRows_(SHEETS.LECTURERS).find(x => String(x.lecturerid).toUpperCase() === String(user.userID).toUpperCase()) || {}).email || '';
    const options = {name:String(user.name || 'psyc@UNIDEL')+' — psyc@UNIDEL'};
    if (validEmailAddress_(senderEmail)) options.replyTo = String(senderEmail).trim();
    let sent = 0;
    try {
      recipients.forEach(recipient => {
        MailApp.sendEmail({to:recipient.email,subject:subject,body:body,name:options.name,replyTo:options.replyTo});
        sent++;
      });
      props.setProperty(propKey,String(sentToday + 1));
      logEmailDelivery_(user, recipients.length, 'SENT');
      return {success:true,message:'Email sent successfully to '+sent+' recipient(s). Message content was not stored in the portal.',sent:sent};
    } catch (err) {
      logEmailDelivery_(user, sent, 'PARTIAL_OR_FAILED');
      return {success:false,message:'Email delivery stopped after '+sent+' recipient(s): '+err.message};
    }
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
   10. WEB APP ENTRY POINTS (BOTH DIRECT & GITHUB PAGES)
   ============================================================ */
function doGet(e) {
  if (e && e.parameter && e.parameter.bridge === '1') {
    return HtmlService
      .createHtmlOutputFromFile('Bridge')
      .setTitle('psyc@UNIDEL API Bridge')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('psyc@UNIDEL 3.10 — Department of Psychology')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const body = JSON.parse(raw);
    const fn = String(body.fn || '').trim();
    const args = Array.isArray(body.args) ? body.args : [];

    const ROUTES = {
      getPublicInfo: getPublicInfo,
      loginUser: loginUser,
      validateSession: validateSession,
      logoutUser: logoutUser,
      registerStudent: registerStudent,
      registerLecturer: registerLecturer,
      getCourses: getCourses,
      createCourse: createCourse,
      claimCourseAsLecturer: claimCourseAsLecturer,
      manageCourseLecturer: manageCourseLecturer,
      enrolStudentInCourse: enrolStudentInCourse,
      getCourseRoster: getCourseRoster,
      uploadMaterial: uploadMaterial,
      getMaterials: getMaterials,
      recordAttendance: recordAttendance,
      bulkUploadAttendance: bulkUploadAttendance,
      getCourseAttendanceSummary: getCourseAttendanceSummary,
      getAttendanceData: getAttendanceData,
      generateEligibleCandidatesPDF: generateEligibleCandidatesPDF,
      exportResultsToExcel: exportResultsToExcel,
      getComponentGroupScores: getComponentGroupScores,
      createTest: createTest,
      updateTestSettings: updateTestSettings,
      addQuestion: addQuestion,
      getTestsAndAttempts: getTestsAndAttempts,
      getQuestionsForTest: getQuestionsForTest,
      submitCBTAttempt: submitCBTAttempt,
      saveResult: saveResult,
      getResults: getResults,
      postAnnouncement: postAnnouncement,
      getAnnouncements: getAnnouncements,
      createDiscussionTopic: createDiscussionTopic,
      replyDiscussionTopic: replyDiscussionTopic,
      getDiscussions: getDiscussions,
      getAdminOverview: getAdminOverview,
      toggleUserStatus: toggleUserStatus,
      getGradeReport: getGradeReport,
      exportGradeReportToExcel: exportGradeReportToExcel,
      getMessagingSettings: getMessagingSettings,
      setAllowExternalEmailRecipients: setAllowExternalEmailRecipients,
      sendPortalEmail: sendPortalEmail,
      startClassroom: startClassroom,
      getClassrooms: getClassrooms,
      joinClassroom: joinClassroom,
      leaveClassroom: leaveClassroom,
      updateClassroomParticipation: updateClassroomParticipation,
      endClassroom: endClassroom
    };

    if (!Object.prototype.hasOwnProperty.call(ROUTES, fn)) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, message: 'Unknown endpoint: ' + fn }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const result = ROUTES[fn].apply(null, args);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err && err.message ? err.message : 'Server error.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function checkDatabaseConnection() {
  const id = PropertiesService.getScriptProperties()
    .getProperty('PSYC_UNIDEL_V3_DB_ID');

  Logger.log('Database ID: ' + (id || 'NOT FOUND'));

  if (!id) {
    Logger.log('The database ID is missing.');
    return;
  }

  const ss = SpreadsheetApp.openById(id);
  Logger.log('Database name: ' + ss.getName());
  Logger.log('Database sheets: ' +
    ss.getSheets().map(s => s.getName()).join(', '));
}

// Version 3.10 — non-destructive grade-report schema upgrade.
function upgradeSystem310() {
  const ss = getDB_();
  const sh = ss.getSheetByName(SHEETS.RESULTS);
  if (!sh) throw new Error('Results sheet is missing. No changes were made.');
  const headers = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(v => String(v).trim());
  ['AcademicSession','Semester'].forEach(name => {
    if (headers.map(x => x.toLowerCase()).indexOf(name.toLowerCase()) < 0) {
      sh.getRange(1,sh.getLastColumn()+1).setValue(name).setFontWeight('bold');
    }
  });
  let emailLog = ss.getSheetByName('EmailLog');
  if (!emailLog) {
    emailLog = ss.insertSheet('EmailLog');
    emailLog.appendRow(['Timestamp','SenderID','SenderName','SenderRole','RecipientCount','Status']);
    emailLog.getRange(1,1,1,6).setFontWeight('bold').setBackground('#123b5d').setFontColor('#ffffff');
  }
  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  if (!usersSheet) throw new Error('Users sheet is missing. No migration was performed.');
  const userHeaders = usersSheet.getRange(1,1,1,Math.max(1,usersSheet.getLastColumn())).getValues()[0].map(v => String(v).trim().toLowerCase());
  if (userHeaders.indexOf('email') < 0) usersSheet.getRange(1,usersSheet.getLastColumn()+1).setValue('Email').setFontWeight('bold');
  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  const rows = getRows_(SHEETS.SETTINGS);
  const version = rows.find(r => String(r.key).trim().toLowerCase() === 'version');
  if (version) settings.getRange(version._row,2).setValue('3.10');
  else settings.appendRow(['Version','3.10']);
  const external = rows.find(r => String(r.key).trim().toLowerCase() === 'allowexternalemailrecipients');
  if (!external) settings.appendRow(['AllowExternalEmailRecipients','YES']);
  return {success:true,version:'3.10',message:'Version 3.10 schema is ready. Grade-report columns and minimal EmailLog were added; existing result records and message privacy are preserved.'};
}


/* ============================================================
   VERSION 11 — INTERACTIVE VIRTUAL CLASSROOM
   Non-destructive migration and authenticated classroom APIs.
   Live media transport is provided by Jitsi, not Apps Script.
   ============================================================ */

const V11_CLASSROOMS_SHEET = 'ClassroomsV11';
const V11_CLASSROOM_EVENTS_SHEET = 'ClassroomEventsV11';
const V11_PARTICIPATION_MODES = [
  'VIEW_ONLY',
  'AUDIO_PARTICIPATION',
  'VIDEO_AUDIO',
  'RAISE_HAND',
  'OPEN_INTERACTIVE'
];

function upgradeSystem311Classroom() {
  const ss = getDB_();
  const definitions = [
    [V11_CLASSROOMS_SHEET, [
      'SessionID','RoomName','CourseID','CourseCode','CourseTitle','Title',
      'LecturerID','LecturerName','ParticipationMode','Status','StartedAt','EndedAt'
    ]],
    [V11_CLASSROOM_EVENTS_SHEET, [
      'EventID','SessionID','UserID','UserName','UserRole','EventType','OccurredAt'
    ]]
  ];
  const created = [];
  definitions.forEach(def => {
    let sh = ss.getSheetByName(def[0]);
    if (!sh) {
      sh = ss.insertSheet(def[0]);
      sh.appendRow(def[1]);
      sh.getRange(1,1,1,def[1].length)
        .setFontWeight('bold').setBackground('#123b5d').setFontColor('#ffffff');
      created.push(def[0]);
    } else {
      const lastCol = Math.max(1, sh.getLastColumn());
      const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(v => String(v).trim());
      const missing = def[1].filter(h => headers.indexOf(h) < 0);
      if (missing.length) {
        sh.getRange(1,lastCol+1,1,missing.length).setValues([missing])
          .setFontWeight('bold').setBackground('#123b5d').setFontColor('#ffffff');
      }
    }
  });
  return {
    success:true,
    version:'11-classroom-schema',
    createdSheets:created,
    message:'Classroom schema is ready. Existing portal sheets and records were not modified.'
  };
}

function v11ClassroomSheet_(name) {
  const sh = getDB_().getSheetByName(name);
  if (!sh) throw new Error('Classroom database is not prepared. An administrator must run upgradeSystem311Classroom() first.');
  return sh;
}

function v11ParticipationMode_(value) {
  const mode = String(value || 'RAISE_HAND').trim().toUpperCase();
  if (V11_PARTICIPATION_MODES.indexOf(mode) < 0) throw new Error('Choose a valid student participation mode.');
  return mode;
}

function v11FindClassroom_(sessionID) {
  const id = String(sessionID || '').trim();
  if (!id) throw new Error('Classroom session is required.');
  const row = getRows_(V11_CLASSROOMS_SHEET).find(r => String(r.sessionid) === id);
  if (!row) throw new Error('Classroom session was not found.');
  return row;
}

function v11CanAccessClassroom_(user, classroom) {
  if (user.role === 'Administrator') return true;
  if (user.role === 'Lecturer') {
    requireCourseAccess_(user, classroom.courseid);
    return true;
  }
  if (user.role === 'Student') {
    if (!isStudentEnrolled_(user.userID, classroom.courseid)) {
      throw new Error('You are not enrolled in this course.');
    }
    return true;
  }
  throw new Error('Your account cannot access the virtual classroom.');
}

function startClassroom(token, payload) {
  const user = requireUser_(token, ['Lecturer','Administrator']);
  payload = payload || {};
  const course = requireCourseAccess_(user, payload.courseID);
  const title = String(payload.title || '').trim().slice(0,160);
  if (!title) throw new Error('Enter a lecture title or topic.');
  const mode = v11ParticipationMode_(payload.participationMode);
  const sh = v11ClassroomSheet_(V11_CLASSROOMS_SHEET);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const active = getRows_(V11_CLASSROOMS_SHEET).find(r =>
      String(r.courseid) === String(course.courseid) && String(r.status).toUpperCase() === 'LIVE'
    );
    if (active) throw new Error('A live classroom is already open for this course.');
    const sessionID = makeId_('CLS');
    const roomName = 'psycUNIDEL-' + Utilities.getUuid().replace(/-/g,'');
    const now = new Date().toISOString();
    sh.appendRow([
      sessionID, roomName, course.courseid, course.coursecode, course.coursetitle,
      title, user.userID, user.name, mode, 'LIVE', now, ''
    ]);
    return {
      success:true, sessionID:sessionID, roomName:roomName,
      courseID:course.courseid, courseCode:course.coursecode, courseTitle:course.coursetitle,
      title:title, participationMode:mode, status:'LIVE', startedAt:now,
      message:'Virtual classroom opened. Students enrolled in this course can now join.'
    };
  } finally {
    lock.releaseLock();
  }
}

function getClassrooms(token, courseID) {
  const user = requireUser_(token);
  let rows = getRows_(V11_CLASSROOMS_SHEET);
  if (user.role === 'Student') {
    const enrolled = new Set(getRows_(SHEETS.ENROLMENTS)
      .filter(e => normaliseMatric_(e.matricno) === normaliseMatric_(user.userID))
      .map(e => String(e.courseid)));
    rows = rows.filter(r => enrolled.has(String(r.courseid)));
  } else if (user.role === 'Lecturer') {
    const assigned = getLecturerCourseIds_(user.userID);
    rows = rows.filter(r => assigned.has(String(r.courseid)));
  }
  if (courseID) rows = rows.filter(r => String(r.courseid) === String(courseID));
  rows.sort((a,b) => String(b.startedat).localeCompare(String(a.startedat)));
  return {success:true,classrooms:rows.map(r => ({
    sessionID:r.sessionid, courseID:r.courseid, courseCode:r.coursecode,
    courseTitle:r.coursetitle, title:r.title, lecturerName:r.lecturername,
    participationMode:r.participationmode, status:r.status,
    startedAt:r.startedat, endedAt:r.endedat
  }))};
}

function joinClassroom(token, sessionID) {
  const user = requireUser_(token);
  const classroom = v11FindClassroom_(sessionID);
  v11CanAccessClassroom_(user, classroom);
  if (String(classroom.status).toUpperCase() !== 'LIVE') throw new Error('This classroom is not currently live.');
  const now = new Date().toISOString();
  const eventID = makeId_('EVT');
  v11ClassroomSheet_(V11_CLASSROOM_EVENTS_SHEET).appendRow([
    eventID, classroom.sessionid, user.userID, user.name, user.role, 'JOIN', now
  ]);
  return {
    success:true, eventID:eventID, sessionID:classroom.sessionid,
    roomName:classroom.roomname, title:classroom.title,
    courseCode:classroom.coursecode, lecturerName:classroom.lecturername,
    participationMode:classroom.participationmode, userID:user.userID,
    displayName:user.name, role:user.role, joinedAt:now
  };
}

function leaveClassroom(token, sessionID, eventID) {
  const user = requireUser_(token);
  const classroom = v11FindClassroom_(sessionID);
  v11CanAccessClassroom_(user, classroom);
  v11ClassroomSheet_(V11_CLASSROOM_EVENTS_SHEET).appendRow([
    makeId_('EVT'), classroom.sessionid, user.userID, user.name, user.role,
    'LEAVE:' + String(eventID || '').slice(0,80), new Date().toISOString()
  ]);
  return {success:true,message:'Classroom departure recorded.'};
}

function updateClassroomParticipation(token, sessionID, participationMode) {
  const user = requireUser_(token, ['Lecturer','Administrator']);
  const classroom = v11FindClassroom_(sessionID);
  requireCourseAccess_(user, classroom.courseid);
  if (String(classroom.status).toUpperCase() !== 'LIVE') throw new Error('The classroom has ended.');
  const mode = v11ParticipationMode_(participationMode);
  const sh = v11ClassroomSheet_(V11_CLASSROOMS_SHEET);
  setCellByHeader_(V11_CLASSROOMS_SHEET, classroom._row, 'ParticipationMode', mode);
  return {success:true,participationMode:mode,message:'Student participation settings updated.'};
}

function endClassroom(token, sessionID) {
  const user = requireUser_(token, ['Lecturer','Administrator']);
  const classroom = v11FindClassroom_(sessionID);
  requireCourseAccess_(user, classroom.courseid);
  if (String(classroom.status).toUpperCase() !== 'LIVE') {
    return {success:true,message:'This classroom has already ended.'};
  }
  const sh = v11ClassroomSheet_(V11_CLASSROOMS_SHEET);
  const endedAt = new Date().toISOString();
  setCellByHeader_(V11_CLASSROOMS_SHEET, classroom._row, 'Status', 'ENDED');
  setCellByHeader_(V11_CLASSROOMS_SHEET, classroom._row, 'EndedAt', endedAt);
  return {success:true,message:'The classroom session has been closed in the portal. Participants already connected to the meeting may remain connected until they leave.'};
}
