/**
 * psyc@UNIDEL 3.7 — Enhanced Academic & Courseware Management System
 * Department of Psychology, University of Delta, Agbor
 * Run setupSystem() ONCE from the editor, then Deploy -> New version.
 */

const PROP_DB_ID = 'PSYC_UNIDEL_V3_DB_ID';
const PROP_DRIVE_ID = 'PSYC_UNIDEL_V3_DRIVE_ID';

const SHEETS = {
  SETTINGS: 'Settings',
  USERS: 'Users',
  STUDENTS: 'Students',
  LECTURERS: 'Lecturers',
  COURSES: 'Courses',
  ENROLMENTS: 'Enrolments',
  MATERIALS: 'Materials',
  ATTENDANCE: 'Attendance',
  TESTS: 'Tests',
  QUESTIONS: 'Questions',
  ATTEMPTS: 'Attempts',
  RESULTS: 'Results',
  ANNOUNCEMENTS: 'Announcements',
  DISCUSSIONS: 'Discussions',
  REPLIES: 'DiscussionReplies'
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
    [SHEETS.ENROLMENTS]: ['EnrolmentID', 'CourseID', 'StudentID', 'MatricNo', 'StudentName', 'EnrolledAt'],
    [SHEETS.MATERIALS]: ['MaterialID', 'CourseID', 'CourseCode', 'Title', 'MaterialType', 'Description', 'FileURL', 'UploadedBy', 'UploadedAt'],
    [SHEETS.ATTENDANCE]: ['AttendanceID', 'CourseID', 'LectureDate', 'Topic', 'StudentID', 'MatricNo', 'Status', 'RecordedBy'],
    [SHEETS.TESTS]: ['TestID', 'CourseID', 'CourseCode', 'Kind', 'Title', 'DurationMinutes', 'TotalMarks', 'MaxAttempts', 'ScoringMethod', 'FeedbackMode', 'OpenDateTime', 'CloseDateTime', 'Shuffle', 'Instructions', 'Status', 'CreatedAt', 'ComponentGroup', 'ComponentGradingMethod', 'SelectedTestID'],
    [SHEETS.QUESTIONS]: ['QuestionID', 'TestID', 'QuestionType', 'QuestionText', 'OptionA', 'OptionB', 'OptionC', 'OptionD', 'CorrectAnswer', 'Marks'],
    [SHEETS.ATTEMPTS]: ['AttemptID', 'TestID', 'StudentID', 'MatricNo', 'AttemptNumber', 'Score', 'MaxScore', 'SubmittedAt'],
    [SHEETS.RESULTS]: ['ResultID', 'CourseID', 'CourseCode', 'CourseTitle', 'CreditUnits', 'StudentID', 'MatricNo', 'StudentName', 'CAScore', 'ExamScore', 'TotalScore', 'Grade', 'GradePoint', 'UpdatedAt'],
    [SHEETS.ANNOUNCEMENTS]: ['AnnouncementID', 'Title', 'Audience', 'Body', 'AuthorName', 'AuthorRole', 'DatePosted'],
    [SHEETS.DISCUSSIONS]: ['TopicID', 'CourseID', 'CourseCode', 'Title', 'Body', 'AuthorName', 'AuthorRole', 'DateCreated'],
    [SHEETS.REPLIES]: ['ReplyID', 'TopicID', 'Body', 'AuthorName', 'AuthorRole', 'DatePosted']
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
      ['Version', '3.7'],
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

/* ============================================================
   2. CORE HELPERS & SECURITY
   ============================================================ */
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

function requireCourseAccess_(user, courseID) {
  const course = getRows_(SHEETS.COURSES).find(c => String(c.courseid) === String(courseID));
  if (!course) throw new Error('Course not found.');
  if (user.role === 'Administrator') return course;
  if (user.role !== 'Lecturer' || String(course.lecturerid || '').trim().toUpperCase() !== String(user.userID || '').trim().toUpperCase()) {
    throw new Error('Access denied. You are not assigned to this course.');
  }
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
      version: map.Version || '3.7',
      institution: map.Institution || 'University of Delta, Agbor',
      department: map.Department || 'Department of Psychology',
      academicSession: map.AcademicSession || '2025/2026',
      semester: map.Semester || 'First Semester'
    };
  } catch (e) {
    return { success: true, systemName: 'psyc@UNIDEL', version: '3.7', institution: 'University of Delta, Agbor', department: 'Department of Psychology' };
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
      courses: all.filter(c => String(c.lecturerid).toUpperCase() === uid),
      allCourses: all
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

  return { success: true, courses: all, allCourses: all };
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
  // Lecturers may claim unassigned courses for themselves. Administrators
  // retain the ability to assign/reassign a course through this action.
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  const courses = getRows_(SHEETS.COURSES);
  const target = courses.find(c => String(c.courseid) === String(courseID));
  if (!target) return { success: false, message: 'Course not found.' };

  const claimant = String(user.userID || '').trim().toUpperCase();
  const currentLecturer = String(target.lecturerid || '').trim().toUpperCase();

  if (user.role === 'Lecturer' && currentLecturer && currentLecturer !== claimant) {
    return { success: false, message: 'This course is already assigned to another lecturer. Please contact the Administrator.' };
  }

  if (currentLecturer === claimant) {
    return { success: true, message: `You are already assigned to ${target.coursecode}.` };
  }

  getDB_().getSheetByName(SHEETS.COURSES).getRange(target._row, 7).setValue(claimant);
  return { success: true, message: `You are now assigned to ${target.coursecode}.` };
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
  const crs = requireCourseAccess_(user, payload.courseID);

  let fileUrl = String(payload.externalLink || '').trim();
  if (payload.base64Data && payload.fileName) {
    const folderId = PropertiesService.getScriptProperties().getProperty(PROP_DRIVE_ID);
    const folder = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(Utilities.base64Decode(payload.base64Data), payload.mimeType || 'application/octet-stream', payload.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileUrl = file.getUrl();
  }
  if (!fileUrl) return { success: false, message: 'Attach a file or enter an external link.' };

  getDB_().getSheetByName(SHEETS.MATERIALS).appendRow([
    makeId_('MAT'), crs.courseid, crs.coursecode, payload.title, payload.materialType || 'Lecture Note',
    payload.description || '', fileUrl, user.name, new Date().toISOString()
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
      .filter(c => String(c.lecturerid).toUpperCase() === String(user.userID).toUpperCase())
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
    allowedCourseIds = new Set(courses.filter(c => String(c.lecturerid).toUpperCase() === String(user.userID).toUpperCase()).map(c => c.courseid));
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
    payload.openDateTime || '', payload.closeDateTime || '', payload.shuffle || 'YES',
    payload.instructions || '', 'OPEN', new Date().toISOString()
  ]);
  const newRow = testSheet.getLastRow();
  setCellByHeader_(SHEETS.TESTS, newRow, 'ComponentGroup', String(payload.componentGroup || '').trim());
  setCellByHeader_(SHEETS.TESTS, newRow, 'ComponentGradingMethod', payload.componentGradingMethod || 'AVERAGE');
  setCellByHeader_(SHEETS.TESTS, newRow, 'SelectedTestID', payload.selectedTestID || '');
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

  return { success: true, message: 'Test policy, schedule, feedback, and component grading settings updated.' };
}

function addQuestion(token, payload) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  if (!payload.testID || !payload.questionText) return { success: false, message: 'Test and question text are required.' };
  const test = getRows_(SHEETS.TESTS).find(t => String(t.testid) === String(payload.testID));
  if (!test) return { success: false, message: 'Test not found.' };
  requireCourseAccess_(user, test.courseid);

  getDB_().getSheetByName(SHEETS.QUESTIONS).appendRow([
    makeId_('QST'), payload.testID, payload.questionType || 'MCQ', payload.questionText,
    payload.optionA || '', payload.optionB || '', payload.optionC || '', payload.optionD || '',
    String(payload.correctAnswer || '').trim(), Number(payload.marks || 1)
  ]);
  return { success: true, message: 'Question saved.' };
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
      .filter(c => String(c.lecturerid).toUpperCase() === String(user.userID).toUpperCase())
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
  const user = requireUser_(token);
  const test = getRows_(SHEETS.TESTS).find(t => t.testid === testID);
  if (!test) return { success: false, message: 'Test not found.' };

  if (user.role === 'Student') {
    if (!isStudentEnrolled_(user.userID, test.courseid)) return { success: false, message: 'You are not enrolled in this course.' };
    if (String(test.status || 'OPEN').toUpperCase() !== 'OPEN') return { success: false, message: 'This test is not available.' };
    const win = getTestWindowState_(test);
    if (!win.isOpen) {
      return {
        success: false,
        message: win.statusLabel === 'UPCOMING'
          ? `This test opens at ${new Date(test.opendatetime).toLocaleString()}.`
          : 'This test window has already closed.'
      };
    }
    const matric = String(user.userID).toUpperCase();
    const prev = getRows_(SHEETS.ATTEMPTS).filter(a => a.testid === testID && String(a.matricno).toUpperCase() === matric);
    const maxAtt = Number(test.maxattempts != null ? test.maxattempts : 1);
    if (maxAtt > 0 && prev.length >= maxAtt) {
      return { success: false, message: `You have used all ${maxAtt} allowed attempt(s).` };
    }
  } else if (user.role === 'Lecturer') {
    requireCourseAccess_(user, test.courseid);
  } else if (user.role !== 'Administrator') {
    return { success: false, message: 'Access denied.' };
  }

  let qs = getRows_(SHEETS.QUESTIONS).filter(q => q.testid === testID).map(q => ({
    questionID: q.questionid,
    questionType: q.questiontype,
    questionText: q.questiontext,
    optionA: q.optiona,
    optionB: q.optionb,
    optionC: q.optionc,
    optionD: q.optiond,
    marks: q.marks
  }));

  if (String(test.shuffle).toUpperCase() === 'YES' && user.role === 'Student') {
    qs = qs.sort(() => Math.random() - 0.5);
  }
  return { success: true, test: test, questions: qs };
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

    const qs = getRows_(SHEETS.QUESTIONS).filter(q => String(q.testid) === String(testID));
    if (!qs.length) return { success: false, message: 'This test has no questions yet.' };
    let rawScore = 0, rawMax = 0;
    const reviewItems = [];
    qs.forEach((q, idx) => {
      const mk = Math.max(0, Number(q.marks || 1));
      rawMax += mk;
      const givenRaw = String((answers && answers[q.questionid]) || '').trim();
      const given = givenRaw.toLowerCase();
      const expected = String(q.correctanswer || '').trim().toLowerCase();
      const accepted = q.questiontype === 'FILL_GAP' ? expected.split('|').map(x => x.trim()) : [expected];
      const isCorrect = Boolean(given && accepted.indexOf(given) !== -1);
      if (isCorrect) rawScore += mk;
      reviewItems.push({ number: idx + 1, questionText: q.questiontext, yourAnswer: givenRaw || '(No Answer)', correctAnswer: q.correctanswer, isCorrect: isCorrect });
    });
    const maxScore = Number(test.totalmarks || rawMax);
    const scaledScore = rawMax > 0 ? Number(((rawScore / rawMax) * maxScore).toFixed(2)) : 0;
    const attNum = prevAttempts.length + 1;
    getDB_().getSheetByName(SHEETS.ATTEMPTS).appendRow([
      makeId_('ATT'), testID, user.userID, matric, attNum, scaledScore, maxScore, new Date().toISOString()
    ]);

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
  const matric = String(payload.matricNo || '').trim().toUpperCase();
  const stu = getRows_(SHEETS.STUDENTS).find(s => normaliseMatric_(s.matricno) === matric);
  if (!stu || !isStudentEnrolled_(matric, crs.courseid)) return { success: false, message: 'The student is not registered and enrolled in this course.' };

  const sh = getDB_().getSheetByName(SHEETS.RESULTS);
  const existing = getRows_(SHEETS.RESULTS).find(r => r.courseid === crs.courseid && String(r.matricno).toUpperCase() === matric);
  const rowVals = [
    existing ? existing.resultid : makeId_('RES'), crs.courseid, crs.coursecode, crs.coursetitle,
    Number(crs.creditunits || 2), stu ? stu.studentid : matric, matric,
    stu ? `${stu.surname} ${stu.firstname}` : matric, ca, exam, total, grade, gp, new Date().toISOString()
  ];

  if (existing) sh.getRange(existing._row, 1, 1, rowVals.length).setValues([rowVals]);
  else sh.appendRow(rowVals);

  return { success: true, message: `Result saved for ${matric}: ${total} (${grade}).` };
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
      .filter(c => String(c.lecturerid).toUpperCase() === String(user.userID).toUpperCase())
      .map(c => c.courseid));
    return { success: true, results: all.filter(r => myCourses.has(r.courseid)) };
  }
  return { success: true, results: all };
}

function exportResultsToExcel(token, courseID) {
  const user = requireUser_(token, ['Lecturer', 'Administrator']);
  let rows = getRows_(SHEETS.RESULTS);
  if (user.role === 'Lecturer') {
    const myCourseIds = new Set(getRows_(SHEETS.COURSES).filter(c => String(c.lecturerid).toUpperCase() === String(user.userID).toUpperCase()).map(c => String(c.courseid)));
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
    const myCourses = new Set(getRows_(SHEETS.COURSES).filter(c => String(c.lecturerid).toUpperCase() === String(user.userID).toUpperCase()).map(c => c.courseid));
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
   10. WEB APP ENTRY POINTS (BOTH DIRECT & GITHUB PAGES)
   ============================================================ */
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('psyc@UNIDEL 3.7 — Department of Psychology')
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
      toggleUserStatus: toggleUserStatus
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