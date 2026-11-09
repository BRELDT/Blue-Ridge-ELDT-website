// Blue Ridge Entry Level Driver Training
// Google Apps Script backend source
// Current curriculum + secure sessions + best-score retention + request-level Sheets caching.

const STATUS_SHEET = "Status";
const STUDENTS_SHEET = "Students";
const ADMINS_SHEET = "Admins";
const CLASSES_SHEET = "Classes";
const MODULES_SHEET = "Modules";
const TEST_QUESTIONS_SHEET = "TestQuestions";
const STUDENT_CLASSES_SHEET = "StudentClasses";
const PROGRESS_SHEET = "Progress";
const TEST_RESULTS_SHEET = "TestResults";
const SIGNUP_REQUESTS_SHEET = "SignupRequests";

const DEFAULT_CLASS_ID = "class-a-b";
const REQUEST_CLASS_IDS = ["class-a-b", "class-b-to-a"];
const BACKEND_VERSION = "blue-ridge-v7-secure";
const AUTH_TTL_SECONDS = 21600;

function property_(name) {
  return PropertiesService.getScriptProperties().getProperty(name);
}

function adminEmail_() {
  return property_("BLUE_RIDGE_ADMIN_EMAIL");
}

function passwordHash_(value, legacy) {
  const password = String(value || "");
  if (!password || bcrypt.truncates(password) || (!legacy && password.length < 12)) {
    throw new Error("Use at least 12 characters and no more than 72 UTF-8 bytes.");
  }
  bcrypt.setRandomFallback(function(length) {
    return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      Utilities.getUuid() + Utilities.getUuid()).slice(0, length).map(function(byte) {
      return byte & 255;
    });
  });
  return bcrypt.hashSync(password, 12);
}

function isPasswordHash_(value) {
  return /^\$2[aby]\$12\$[./A-Za-z0-9]{53}$/.test(String(value || ""));
}

function passwordMatches_(password, hash) {
  return isPasswordHash_(hash) && String(password).length <= 256 &&
    !bcrypt.truncates(String(password)) && bcrypt.compareSync(String(password), String(hash));
}

// Run manually in the Apps Script editor, after backup and retiring old versions.
function migrateBlueRidgePasswords() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return withRequestCache_(function() {
      let migrated = 0;
      [ADMINS_SHEET, STUDENTS_SHEET].forEach(function(name) {
        rowObjs_(name).filter(function(row) {
          return row.obj.password && !isPasswordHash_(row.obj.password) && !bcrypt.truncates(String(row.obj.password));
        }).slice(0, 25).forEach(function(row) {
          setField_(sh_(name), row.row, "password", passwordHash_(row.obj.password, true));
          migrated++;
        });
      });
      SpreadsheetApp.flush();
      console.log("Credentials migrated in this batch: " + migrated);
    });
  } finally {
    lock.releaseLock();
  }
}

const DEFAULT_CLASSES = [
  [DEFAULT_CLASS_ID, "Class A and B", "CDL Class A/B ELDT training", 80, 0.9, 1, true],
  ["class-b-to-a", "Class B to A Upgrade", "Upgrade training from Class B to Class A", 80, 0.9, 2, true],
  ["passenger", "Passenger and School Bus Endorsement", "Passenger and school bus endorsement training", 80, 0.9, 3, true],
  ["school-bus", "School Bus Endorsement", "School bus endorsement training", 80, 0.9, 4, false],
  ["tanker", "Tanker Endorsement", "Tanker endorsement training", 80, 0.9, 5, false],
  ["hazmat", "Hazmat Endorsement", "Hazmat endorsement training", 80, 0.9, 6, true]
];

const DEFAULT_MODULES = [
  ["1", DEFAULT_CLASS_ID, "Module 1 — Introduction", "-qXt8htJ9h4", 1, 0.9, true],
  ["2", DEFAULT_CLASS_ID, "Module 2 — Safety & Inspection", "RS4K5FCL988", 2, 0.9, true],
  ["3", DEFAULT_CLASS_ID, "Module 3 — Basic Operations", "TLeq0WikSmU", 3, 0.9, true],
  ["4", DEFAULT_CLASS_ID, "Module 4 — Advanced Driving", "cMML4tQdVvY", 4, 0.9, true],
  ["8", "hazmat", "Hazmat Module 1", "g8WOxP_PDJ8", 1, 0.9, true],
  ["9", "hazmat", "Hazmat Module 2", "CLIhc8MWFJ8", 2, 0.9, true],
  ["10", "class-b-to-a", "B to A Upgrade", "zeaHTafu4CY", 1, 0.9, true],
  ["11", "passenger", "Passenger and School Bus Module 1", "ocQxZ3-fk1M", 1, 0.9, true],
  ["12", "passenger", "Passenger and School Bus Module 2", "Z0V1nlzn2ks", 2, 0.9, true]
];

let REQUEST_CACHE_ = null;

function withRequestCache_(callback) {
  REQUEST_CACHE_ = { ss: null, sheets: {}, headers: {}, rows: {} };
  try {
    return callback();
  } finally {
    REQUEST_CACHE_ = null;
  }
}

function cache_() {
  if (!REQUEST_CACHE_) REQUEST_CACHE_ = { ss: null, sheets: {}, headers: {}, rows: {} };
  return REQUEST_CACHE_;
}

function clearSheetCache_(name) {
  const c = cache_();
  delete c.headers[name];
  delete c.rows[name];
}

function doGet(e) {
  return withRequestCache_(function() {
    try {
      const p = (e && e.parameter) || {};
      const action = String(p.action || "").toLowerCase();
      let result;

      switch (action) {
        case "getversion":
          result = { ok: true, version: BACKEND_VERSION, intakeOpen: property_("BLUE_RIDGE_INTAKE_OPEN") === "true" && !!adminEmail_() };
          break;
        case "listclasses":
          result = listClasses_(p.activeOnly);
          break;
        default:
          result = { ok: false, error: "Unknown action" };
      }
      return json_(result);
    } catch (err) {
      return json_({ ok: false, error: String(err && err.message ? err.message : err) });
    }
  });
}

function doPost(e) {
  return withRequestCache_(function() {
    try {
      const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
      const action = String(data.action || "").toLowerCase();
      let result;

      switch (action) {
        case "validatelogin":
          result = validateLogin_(data.username, data.password);
          break;
        case "adminlogin":
          result = adminLogin_(data.username, data.password);
          break;
        case "liststudents":
          requireAdminToken_(data.adminToken);
          result = listStudents_();
          break;
        case "getstatus":
          requireStudentToken_(data.studentToken, data.username);
          result = getStatus_(data.username, data.classId);
          break;
        case "getstudentdashboard":
          requireStudentToken_(data.studentToken, data.username);
          result = getStudentDashboard_(data.username);
          break;
        case "listmodules":
          requireAdminToken_(data.adminToken);
          result = listModules_(data.classId, data.activeOnly);
          break;
        case "listtestquestions":
          result = testQuestionsForSession_(data);
          break;
        case "addstudent":
          requireAdminToken_(data.adminToken);
          result = addStudent_(data);
          break;
        case "updatestudent":
          requireAdminToken_(data.adminToken);
          result = updateStudent_(data);
          break;
        case "deletestudent":
        case "archivestudent":
          requireAdminToken_(data.adminToken);
          result = archiveStudent_(data.username);
          break;
        case "approvestudent":
          requireAdminToken_(data.adminToken);
          result = approveStudent_(data.username);
          break;
        case "beginmodule":
          requireStudentToken_(data.studentToken, data.username);
          result = beginModule_(data);
          break;
        case "logmodule":
          requireStudentToken_(data.studentToken, data.username);
          requireModuleTime_(data);
          result = logModule_(data.username, data.classId, data.moduleId);
          break;
        case "logtest":
          requireStudentToken_(data.studentToken, data.username);
          result = gradeTest_(data.username, data.classId, data.answers);
          break;
        case "saveclass":
          requireAdminToken_(data.adminToken);
          result = saveClass_(data);
          break;
        case "deleteclass":
        case "deactivateclass":
          requireAdminToken_(data.adminToken);
          result = deactivateById_(CLASSES_SHEET, data.id);
          break;
        case "savemodule":
          requireAdminToken_(data.adminToken);
          result = saveModule_(data);
          break;
        case "deletemodule":
        case "deactivatemodule":
          requireAdminToken_(data.adminToken);
          result = deactivateById_(MODULES_SHEET, data.id);
          break;
        case "savetestquestion":
          requireAdminToken_(data.adminToken);
          result = saveTestQuestion_(data);
          break;
        case "deletetestquestion":
        case "deactivatetestquestion":
          requireAdminToken_(data.adminToken);
          result = deactivateById_(TEST_QUESTIONS_SHEET, data.id);
          break;
        case "submitsignuprequest":
          result = submitSignupRequest_(data);
          break;
        default:
          result = { ok: false, error: "Unknown action" };
      }
      return json_(result);
    } catch (err) {
      return json_({ ok: false, error: String(err && err.message ? err.message : err) });
    }
  });
}

function issueToken_(kind, username) {
  const token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  CacheService.getScriptCache().put("auth:v7:" + kind + ":" + token, String(username || ""), AUTH_TTL_SECONDS);
  return token;
}

function tokenUser_(kind, token) {
  if (!token) return "";
  return CacheService.getScriptCache().get("auth:v7:" + kind + ":" + String(token)) || "";
}

function requireAdminToken_(token) {
  const username = tokenUser_("admin", token);
  if (!username) throw new Error("Admin session expired. Please log in again.");
  return username;
}

function requireStudentToken_(token, username) {
  const tokenUsername = tokenUser_("student", token);
  if (!tokenUsername ||
      String(tokenUsername).trim().toLowerCase() !== String(username || "").trim().toLowerCase()) {
    throw new Error("Student session expired. Please log in again.");
  }
  requireActiveStudent_(tokenUsername);
  return tokenUsername;
}

function ss_() {
  const c = cache_();
  if (!c.ss) {
    const id = property_("BLUE_RIDGE_SPREADSHEET_ID");
    if (!id) throw new Error("Blue Ridge database is not configured.");
    c.ss = SpreadsheetApp.openById(id);
  }
  return c.ss;
}

function sh_(name) {
  const c = cache_();
  if (!c.sheets[name]) c.sheets[name] = ss_().getSheetByName(name);
  if (!c.sheets[name]) throw new Error("Missing configured sheet: " + name);
  return c.sheets[name];
}

function headers_(sheet) {
  const name = sheet.getName();
  const c = cache_();
  if (c.headers[name]) return c.headers[name];
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return [];
  c.headers[name] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(value) {
    return String(value || "").trim();
  });
  return c.headers[name];
}

function ensureHeaders_(name, requiredHeaders) {
  const c = cache_();
  const sheet = c.sheets[name] || ss_().getSheetByName(name) || ss_().insertSheet(name);
  c.sheets[name] = sheet;
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    clearSheetCache_(name);
    return sheet;
  }

  let current = headers_(sheet).slice();
  let changed = false;
  requiredHeaders.forEach(function(header) {
    if (current.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      current.push(header);
      changed = true;
    }
  });
  if (changed) clearSheetCache_(name);
  return sheet;
}

function rowObjs_(name) {
  const c = cache_();
  if (c.rows[name]) return c.rows[name];

  const sheet = sh_(name);
  const headers = headers_(sheet);
  const lastRow = sheet.getLastRow();
  if (!headers.length || lastRow < 2) {
    c.rows[name] = [];
    return c.rows[name];
  }

  c.rows[name] = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues().map(function(row, index) {
    const obj = {};
    headers.forEach(function(header, col) { obj[header] = row[col]; });
    return { row: index + 2, obj: obj };
  });
  return c.rows[name];
}

function appendObject_(sheetName, obj) {
  const sheet = sh_(sheetName);
  const headers = headers_(sheet);
  if (!headers.length) throw new Error("Missing headers for " + sheetName);
  sheet.appendRow(headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(obj, header) ? sheetValue_(obj[header]) : "";
  }));
  const newRow = sheet.getLastRow();
  clearSheetCache_(sheetName);
  return newRow;
}

function sheetValue_(value) {
  return typeof value === "string" && /^\s*=/.test(value) ? "'" + value : value;
}

function setField_(sheet, row, header, value) {
  const headers = headers_(sheet);
  const index = headers.indexOf(header);
  if (index === -1) throw new Error("Missing header: " + header);
  sheet.getRange(row, index + 1).setValue(sheetValue_(value));
  clearSheetCache_(sheet.getName());
}

function cloneObj_(obj) {
  return Object.assign({}, obj || {});
}

function active_(value) {
  if (value === "" || value === null || value === undefined) return true;
  if (value === true) return true;
  const text = String(value).trim().toLowerCase();
  return text === "true" || text === "yes" || text === "active" || text === "complete";
}

function activeOnly_(value) {
  return value === true || String(value || "").trim().toLowerCase() === "true";
}

function complete_(value) {
  if (value === true) return true;
  const text = String(value || "").trim().toLowerCase();
  return text === "true" || text === "complete" || text === "yes";
}

function normalizeWatchPercent_(value) {
  const n = Number(value);
  if (!n || isNaN(n)) return 0.9;
  return n > 0 && n <= 100 ? (n > 1 ? n / 100 : n) : 0.9;
}

function setupSheets_() {
  ensureHeaders_(STUDENTS_SHEET, ["username", "password", "updatedAt", "fullNameOnLicense", "licenseNumber", "dob", "active", "archivedAt", "preferredContact", "licenseState"]);
  ensureHeaders_(ADMINS_SHEET, ["username", "password"]);
  ensureHeaders_(STATUS_SHEET, ["username", "m1", "m2", "m3", "m4", "m5", "m6", "m7", "m8", "m9", "m10", "testComplete", "testScore", "updatedAt"]);
  ensureHeaders_(CLASSES_SHEET, ["id", "title", "description", "passingScore", "requiredWatchPercent", "sortOrder", "active", "updatedAt"]);
  ensureHeaders_(MODULES_SHEET, ["id", "classId", "title", "youtubeId", "sortOrder", "requiredWatchPercent", "active", "updatedAt"]);
  ensureHeaders_(TEST_QUESTIONS_SHEET, ["id", "classId", "question", "optionA", "optionB", "optionC", "optionD", "correctIndex", "sortOrder", "active", "updatedAt"]);
  ensureHeaders_(STUDENT_CLASSES_SHEET, ["username", "classId", "active", "updatedAt"]);
  ensureHeaders_(PROGRESS_SHEET, ["username", "classId", "moduleId", "complete", "updatedAt"]);
  ensureHeaders_(TEST_RESULTS_SHEET, ["username", "classId", "complete", "score", "passed", "updatedAt"]);
  ensureHeaders_(SIGNUP_REQUESTS_SHEET, ["createdAt", "fullNameOnLicense", "licenseNumber", "dob", "requestedClassId", "requestedClassTitle", "status", "preferredContact", "licenseState"]);
  seedRows_(CLASSES_SHEET, DEFAULT_CLASSES);
  seedRows_(MODULES_SHEET, DEFAULT_MODULES);
  return { ok: true, version: BACKEND_VERSION };
}

function seedRows_(sheetName, rows) {
  const existing = {};
  rowObjs_(sheetName).forEach(function(row) { existing[String(row.obj.id)] = true; });
  rows.forEach(function(row) {
    if (!existing[String(row[0])]) {
      const headers = headers_(sh_(sheetName));
      const obj = {};
      headers.forEach(function(header, index) {
        if (index < row.length) obj[header] = row[index];
      });
      if (headers.indexOf("updatedAt") !== -1) obj.updatedAt = new Date();
      appendObject_(sheetName, obj);
      existing[String(row[0])] = true;
    }
  });
}

function findStudent_(username) {
  const key = String(username || "").trim().toLowerCase();
  if (!key) return null;
  return rowObjs_(STUDENTS_SHEET).find(function(row) {
    return String(row.obj.username || "").trim().toLowerCase() === key;
  }) || null;
}

function requireActiveStudent_(username) {
  const student = findStudent_(username);
  if (!student || !active_(student.obj.active)) {
    throw new Error("Student is archived, inactive, or not found");
  }
  return student;
}

function classById_(classId, requireActive) {
  const id = String(classId || DEFAULT_CLASS_ID);
  const row = rowObjs_(CLASSES_SHEET).find(function(item) {
    return String(item.obj.id) === id;
  });
  if (!row) throw new Error("Class not found");
  const obj = cloneObj_(row.obj);
  if (requireActive && !active_(obj.active)) throw new Error("Class is inactive");
  obj.passingScore = Number(obj.passingScore || 80);
  obj.requiredWatchPercent = normalizeWatchPercent_(obj.requiredWatchPercent);
  obj.sortOrder = Number(obj.sortOrder || 0);
  obj.active = active_(obj.active);
  return obj;
}

function assignmentRows_(username) {
  const key = String(username || "").trim().toLowerCase();
  return rowObjs_(STUDENT_CLASSES_SHEET).filter(function(row) {
    return String(row.obj.username || "").trim().toLowerCase() === key;
  });
}

function assignedClassIds_(username) {
  const rows = assignmentRows_(username);
  if (!rows.length) return [DEFAULT_CLASS_ID];
  return rows.filter(function(row) {
    return active_(row.obj.active);
  }).map(function(row) {
    return String(row.obj.classId);
  });
}

function requireAssignedClass_(username, classId) {
  requireActiveStudent_(username);
  const id = String(classId || DEFAULT_CLASS_ID);
  if (assignedClassIds_(username).indexOf(id) === -1) {
    throw new Error("Student is not assigned to this class");
  }
  return classById_(id, true);
}

function listClasses_(activeOnly) {
  const onlyActive = activeOnly_(activeOnly);
  const classes = rowObjs_(CLASSES_SHEET).map(function(row) {
    const obj = cloneObj_(row.obj);
    obj.passingScore = Number(obj.passingScore || 80);
    obj.requiredWatchPercent = normalizeWatchPercent_(obj.requiredWatchPercent);
    obj.sortOrder = Number(obj.sortOrder || 0);
    obj.active = active_(obj.active);
    return obj;
  }).filter(function(obj) {
    return !onlyActive || obj.active;
  }).sort(function(a, b) {
    return a.sortOrder - b.sortOrder;
  });
  return { ok: true, classes: classes };
}

function listModules_(classId, activeOnly) {
  const onlyActive = activeOnly_(activeOnly);
  const id = String(classId || "");
  const modules = rowObjs_(MODULES_SHEET).map(function(row) {
    const obj = cloneObj_(row.obj);
    obj.id = String(obj.id);
    obj.classId = String(obj.classId);
    obj.sortOrder = Number(obj.sortOrder || 0);
    obj.requiredWatchPercent = normalizeWatchPercent_(obj.requiredWatchPercent);
    obj.active = active_(obj.active);
    return obj;
  }).filter(function(obj) {
    return (!id || obj.classId === id) && (!onlyActive || obj.active);
  }).sort(function(a, b) {
    return a.sortOrder - b.sortOrder;
  });
  return { ok: true, modules: modules };
}

function listTestQuestions_(classId, activeOnly) {
  const onlyActive = activeOnly_(activeOnly);
  const id = String(classId || "");
  const questions = rowObjs_(TEST_QUESTIONS_SHEET).map(function(row) {
    const obj = cloneObj_(row.obj);
    obj.id = String(obj.id);
    obj.classId = String(obj.classId);
    obj.correctIndex = Number(obj.correctIndex || 0);
    obj.sortOrder = Number(obj.sortOrder || 0);
    obj.active = active_(obj.active);
    return obj;
  }).filter(function(obj) {
    return (!id || obj.classId === id) && (!onlyActive || obj.active);
  }).sort(function(a, b) {
    return a.sortOrder - b.sortOrder;
  });
  return { ok: true, questions: questions };
}

function limitLoginAttempts_(kind, username) {
  const key = "login:v7:" + kind + ":" + String(username || "").trim().toLowerCase().slice(0, 80);
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const cache = CacheService.getScriptCache();
    const count = Number(cache.get(key) || 0);
    if (count >= 8) throw new Error("Too many sign-in attempts. Try again in 10 minutes.");
    cache.put(key, String(count + 1), 600);
  } finally {
    lock.releaseLock();
  }
}

function validateLogin_(username, password) {
  if (!username || !password) return { ok: false, error: "Missing username or password" };
  limitLoginAttempts_("student", username);
  const student = findStudent_(username);
  if (!student ||
      !passwordMatches_(password, student.obj.password) ||
      !active_(student.obj.active)) {
    return { ok: false, error: "Invalid username or password" };
  }
  return {
    ok: true,
    username: student.obj.username,
    token: issueToken_("student", student.obj.username)
  };
}

function adminLogin_(username, password) {
  if (!username || !password) return { ok: false, error: "Missing username or password" };
  limitLoginAttempts_("admin", username);
  const key = String(username).trim().toLowerCase();
  const admin = rowObjs_(ADMINS_SHEET).find(function(row) {
    return String(row.obj.username || "").trim().toLowerCase() === key &&
      passwordMatches_(password, row.obj.password);
  });
  if (!admin) return { ok: false, error: "Invalid admin login" };
  return {
    ok: true,
    username: admin.obj.username,
    token: issueToken_("admin", admin.obj.username)
  };
}

function legacyStatus_(username) {
  const key = String(username || "").trim().toLowerCase();
  const row = rowObjs_(STATUS_SHEET).find(function(item) {
    return String(item.obj.username || "").trim().toLowerCase() === key;
  });
  const out = { modules: {}, testComplete: false, testScore: "" };
  for (let i = 1; i <= 10; i++) out.modules["m" + i] = false;
  if (!row) return out;
  for (let i = 1; i <= 10; i++) out.modules["m" + i] = complete_(row.obj["m" + i]);
  out.testComplete = complete_(row.obj.testComplete);
  out.testScore = row.obj.testScore || "";
  return out;
}

function getStatus_(username, classId) {
  const id = String(classId || DEFAULT_CLASS_ID);
  const cls = requireAssignedClass_(username, id);
  const modules = listModules_(id, true).modules;
  const progressRows = rowObjs_(PROGRESS_SHEET);
  const testRows = rowObjs_(TEST_RESULTS_SHEET);
  const legacy = id === DEFAULT_CLASS_ID ? legacyStatus_(username) : null;
  const key = String(username || "").trim().toLowerCase();

  const statusModules = modules.map(function(module) {
    const progressComplete = progressRows.some(function(row) {
      return String(row.obj.username || "").trim().toLowerCase() === key &&
        String(row.obj.classId) === id &&
        String(row.obj.moduleId) === String(module.id) &&
        complete_(row.obj.complete);
    });
    return {
      id: String(module.id),
      classId: id,
      title: module.title,
      youtubeId: module.youtubeId,
      sortOrder: Number(module.sortOrder || 0),
      requiredWatchPercent: normalizeWatchPercent_(module.requiredWatchPercent || cls.requiredWatchPercent),
      active: true,
      complete: !!progressComplete || !!(legacy && legacy.modules["m" + module.id])
    };
  });

  const matches = testRows.filter(function(row) {
    return String(row.obj.username || "").trim().toLowerCase() === key &&
      String(row.obj.classId) === id;
  });

  let bestScore = "";
  let passedEver = false;
  const passingScore = Number(cls.passingScore || 80);

  matches.forEach(function(row) {
    const score = Number(row.obj.score);
    if (!isNaN(score) && (bestScore === "" || score > Number(bestScore))) bestScore = score;
    if (complete_(row.obj.passed) ||
        (complete_(row.obj.complete) && !isNaN(score) && score >= passingScore)) passedEver = true;
  });

  if (legacy) {
    const legacyScore = Number(legacy.testScore);
    if (legacy.testScore !== "" && !isNaN(legacyScore) && (bestScore === "" || legacyScore > Number(bestScore))) bestScore = legacyScore;
    if (legacy.testComplete) passedEver = true;
  }

  return {
    ok: true,
    username: username,
    classId: id,
    classInfo: cls,
    modules: statusModules,
    testComplete: passedEver,
    testScore: bestScore,
    testPassed: passedEver
  };
}

function getStudentDashboard_(username) {
  requireActiveStudent_(username);
  const assigned = assignedClassIds_(username);
  const classes = listClasses_(true).classes.filter(function(cls) {
    return assigned.indexOf(String(cls.id)) !== -1;
  }).map(function(cls) {
    const copy = cloneObj_(cls);
    copy.status = getStatus_(username, cls.id);
    return copy;
  });
  return { ok: true, username: username, classes: classes };
}

function listStudents_() {
  const all = rowObjs_(STUDENTS_SHEET).map(function(row) {
    const obj = cloneObj_(row.obj);
    delete obj.password;
    obj.classes = assignedClassIds_(obj.username);
    return obj;
  });
  return {
    ok: true,
    students: all.filter(function(obj) { return active_(obj.active); }),
    pendingStudents: all.filter(function(obj) {
      return !active_(obj.active) && !String(obj.archivedAt || "").trim();
    }),
    backendVersion: BACKEND_VERSION
  };
}

function requestedAssignments_(data) {
  const classes = data.classes === undefined ? [DEFAULT_CLASS_ID] : data.classes;
  if (!Array.isArray(classes) || !classes.length || classes.some(function(id) { return !id; })) {
    throw new Error("Assign at least one course before saving the student.");
  }
  classes.forEach(function(id) { classById_(id, true); });
  return classes;
}

function addStudent_(data) {
  if (!data.username || !data.password) return { ok: false, error: "Missing username or password" };
  if (findStudent_(data.username)) return { ok: false, error: "Username already exists or is archived" };

  const classes = requestedAssignments_(data);
  appendObject_(STUDENTS_SHEET, {
    username: String(data.username).trim(),
    password: passwordHash_(data.password),
    updatedAt: new Date(),
    fullNameOnLicense: data.fullNameOnLicense || "",
    licenseNumber: data.licenseNumber || "",
    dob: data.dob || "",
    active: true,
    archivedAt: "",
    preferredContact: data.preferredContact || "",
    licenseState: normalizeLicenseState_(data.licenseState)
  });
  saveAssignments_(data.username, classes);
  ensureStatusRow_(data.username);
  return { ok: true };
}

function updateStudent_(data) {
  const student = findStudent_(data.username);
  if (!student) return { ok: false, error: "Student not found" };
  if (student.obj.archivedAt) throw new Error("Archived profiles cannot be edited.");
  const classes = requestedAssignments_(data);

  const sheet = sh_(STUDENTS_SHEET);
  if (data.password) setField_(sheet, student.row, "password", passwordHash_(data.password));
  setField_(sheet, student.row, "updatedAt", new Date());
  setField_(sheet, student.row, "fullNameOnLicense", data.fullNameOnLicense || "");
  setField_(sheet, student.row, "licenseNumber", data.licenseNumber || "");
  setField_(sheet, student.row, "dob", data.dob || "");
  setField_(sheet, student.row, "preferredContact", data.preferredContact || "");
  if (data.licenseState !== undefined) setField_(sheet, student.row, "licenseState", normalizeLicenseState_(data.licenseState));
  setField_(sheet, student.row, "active", true);
  setField_(sheet, student.row, "archivedAt", "");
  saveAssignments_(data.username, classes);
  ensureStatusRow_(data.username);
  return { ok: true };
}

function approveStudent_(username) {
  const student = findStudent_(username);
  if (!student) return { ok: false, error: "Student not found" };
  if (String(student.obj.archivedAt || "").trim()) return { ok: false, error: "Archived profiles cannot be approved" };
  const sheet = sh_(STUDENTS_SHEET);
  setField_(sheet, student.row, "active", true);
  setField_(sheet, student.row, "updatedAt", new Date());
  setField_(sheet, student.row, "archivedAt", "");
  ensureStatusRow_(student.obj.username);
  return { ok: true };
}

function archiveStudent_(username) {
  const student = findStudent_(username);
  if (!student) return { ok: false, error: "Student not found" };
  const sheet = sh_(STUDENTS_SHEET);
  setField_(sheet, student.row, "active", false);
  setField_(sheet, student.row, "archivedAt", new Date());
  setField_(sheet, student.row, "updatedAt", new Date());
  saveAssignments_(username, []);
  return { ok: true };
}

function saveAssignments_(username, classIds) {
  const sheet = sh_(STUDENT_CLASSES_SHEET);
  const rows = assignmentRows_(username);
  const unique = Array.from(new Set((classIds || []).filter(Boolean).map(String)));
  unique.forEach(function(classId) { classById_(classId, false); });
  rows.forEach(function(row) {
    const active = unique.includes(String(row.obj.classId));
    setField_(sheet, row.row, "active", active);
    setField_(sheet, row.row, "updatedAt", new Date());
  });
  unique.forEach(function(classId) {
    const existing = rows.find(function(row) { return String(row.obj.classId) === classId; });
    if (!existing) {
      appendObject_(STUDENT_CLASSES_SHEET, { username: username, classId: classId, active: true, updatedAt: new Date() });
    }
  });
}

function ensureStatusRow_(username) {
  const key = String(username || "").trim().toLowerCase();
  const existing = rowObjs_(STATUS_SHEET).find(function(row) {
    return String(row.obj.username || "").trim().toLowerCase() === key;
  });
  return existing ? existing.row : appendObject_(STATUS_SHEET, { username: username, updatedAt: new Date() });
}

function activeModuleForClass_(classId, moduleId) {
  const module = listModules_(classId, true).modules.find(function(item) {
    return String(item.id) === String(moduleId);
  });
  if (!module) throw new Error("Module not found for this class");
  return module;
}

function moduleTiming_(data) {
  const id = String(data.classId || DEFAULT_CLASS_ID);
  const cls = requireAssignedClass_(data.username, id);
  const module = activeModuleForClass_(id, data.moduleId);
  const durations = JSON.parse(property_("BLUE_RIDGE_VIDEO_SECONDS") || "{}");
  const duration = Number(durations[module.youtubeId]);
  if (!Number.isFinite(duration) || duration <= 0 || duration > 21600) {
    throw new Error("Viewing time has not been configured for this lesson.");
  }
  const seconds = duration * normalizeWatchPercent_(module.requiredWatchPercent || cls.requiredWatchPercent);
  return { key: "lesson:v7:" + data.studentToken + ":" + module.youtubeId + ":" + id + ":" + module.id, seconds: seconds };
}

function beginModule_(data) {
  const timing = moduleTiming_(data);
  const cache = CacheService.getScriptCache();
  if (!cache.get(timing.key)) cache.put(timing.key, String(Date.now()), AUTH_TTL_SECONDS);
  return { ok: true };
}

function requireModuleTime_(data) {
  const timing = moduleTiming_(data);
  const started = Number(CacheService.getScriptCache().get(timing.key));
  if (!started || Date.now() - started < timing.seconds * 1000) {
    throw new Error("The minimum viewing time has not elapsed. Continue watching the lesson.");
  }
}

function logModule_(username, classId, moduleId) {
  const id = String(classId || DEFAULT_CLASS_ID);
  requireAssignedClass_(username, id);
  const module = activeModuleForClass_(id, moduleId);
  const sheet = sh_(PROGRESS_SHEET);
  const key = String(username || "").trim().toLowerCase();
  const existing = rowObjs_(PROGRESS_SHEET).find(function(row) {
    return String(row.obj.username || "").trim().toLowerCase() === key &&
      String(row.obj.classId) === id &&
      String(row.obj.moduleId) === String(module.id);
  });

  if (existing) {
    setField_(sheet, existing.row, "complete", true);
    setField_(sheet, existing.row, "updatedAt", new Date());
  } else {
    appendObject_(PROGRESS_SHEET, { username: username, classId: id, moduleId: String(module.id), complete: true, updatedAt: new Date() });
  }

  if (id === DEFAULT_CLASS_ID) {
    const statusRow = ensureStatusRow_(username);
    const statusSheet = sh_(STATUS_SHEET);
    const field = "m" + module.id;
    if (headers_(statusSheet).indexOf(field) !== -1) setField_(statusSheet, statusRow, field, "complete");
    setField_(statusSheet, statusRow, "updatedAt", new Date());
  }
  return getStatus_(username, id);
}

function allModulesComplete_(modules) {
  return Array.isArray(modules) && modules.length > 0 && modules.every(function(module) { return !!module.complete; });
}

function testQuestionsForSession_(data) {
  if (tokenUser_("admin", data.adminToken)) return listTestQuestions_(data.classId, data.activeOnly);
  const username = requireStudentToken_(data.studentToken, data.username);
  const status = getStatus_(username, data.classId);
  if (!allModulesComplete_(status.modules)) throw new Error("Finish all lessons before opening the test.");
  const result = listTestQuestions_(data.classId || DEFAULT_CLASS_ID, true);
  result.questions = result.questions.map(function(q) {
    delete q.correctIndex;
    return q;
  });
  return result;
}

function gradeTest_(username, classId, answers) {
  const id = String(classId || DEFAULT_CLASS_ID);
  requireAssignedClass_(username, id);
  const questions = listTestQuestions_(id, true).questions;
  if (!answers || !questions.length || Object.keys(answers).length !== questions.length ||
      !questions.every(function(q) { return Number.isInteger(answers[q.id]) && answers[q.id] >= 0 && answers[q.id] < 4; })) {
    throw new Error("The test has changed or an answer is missing. Reload the test and try again.");
  }
  const correct = questions.filter(function(q) { return answers[q.id] === q.correctIndex; }).length;
  const attemptScore = Math.floor(correct * 10000 / questions.length) / 100;
  return Object.assign(logTest_(username, id, true, attemptScore), { attemptScore: attemptScore, correct: correct });
}

function logTest_(username, classId, complete, score) {
  const id = String(classId || DEFAULT_CLASS_ID);
  const cls = requireAssignedClass_(username, id);
  const status = getStatus_(username, id);
  if (!allModulesComplete_(status.modules)) throw new Error("All modules for this class must be complete before logging the test");

  const numericScore = Number(score);
  if (isNaN(numericScore)) throw new Error("Missing score");
  const passed = !!complete && numericScore >= Number(cls.passingScore || 80);

  appendObject_(TEST_RESULTS_SHEET, { username: username, classId: id, complete: !!complete, score: numericScore, passed: passed, updatedAt: new Date() });

  if (id === DEFAULT_CLASS_ID) {
    const row = ensureStatusRow_(username);
    const sheet = sh_(STATUS_SHEET);
    const priorScore = status.testScore === "" ? null : Number(status.testScore);
    const bestScore = priorScore === null || isNaN(priorScore) ? numericScore : Math.max(priorScore, numericScore);
    const passedEver = !!status.testPassed || passed;
    setField_(sheet, row, "testComplete", passedEver ? "complete" : "");
    setField_(sheet, row, "testScore", bestScore);
    setField_(sheet, row, "updatedAt", new Date());
  }

  if (passed && !status.testPassed) {
    try { sendTestEmail_(username, id, cls, numericScore, new Date()); }
    catch (err) { console.error("Completion email failed"); }
  }
  return getStatus_(username, id);
}

function saveClass_(data) {
  return upsertById_(CLASSES_SHEET, {
    id: data.id || Utilities.getUuid(),
    title: data.title || "Untitled Class",
    description: data.description || "",
    passingScore: Number(data.passingScore || 80),
    requiredWatchPercent: normalizeWatchPercent_(data.requiredWatchPercent),
    sortOrder: Number(data.sortOrder || 99),
    active: data.active === false ? false : active_(data.active)
  }, ["id", "title", "description", "passingScore", "requiredWatchPercent", "sortOrder", "active"]);
}

function saveModule_(data) {
  const classId = String(data.classId || DEFAULT_CLASS_ID);
  classById_(classId, false);
  return upsertById_(MODULES_SHEET, {
    id: data.id || Utilities.getUuid(),
    classId: classId,
    title: data.title || "Untitled Module",
    youtubeId: extractYouTubeId_(data.youtubeId || data.youtubeUrl || ""),
    sortOrder: Number(data.sortOrder || 99),
    requiredWatchPercent: normalizeWatchPercent_(data.requiredWatchPercent),
    active: data.active === false ? false : active_(data.active)
  }, ["id", "classId", "title", "youtubeId", "sortOrder", "requiredWatchPercent", "active"]);
}

function saveTestQuestion_(data) {
  const classId = String(data.classId || DEFAULT_CLASS_ID);
  classById_(classId, false);
  return upsertById_(TEST_QUESTIONS_SHEET, {
    id: data.id || Utilities.getUuid(),
    classId: classId,
    question: data.question || "",
    optionA: data.optionA || "",
    optionB: data.optionB || "",
    optionC: data.optionC || "",
    optionD: data.optionD || "",
    correctIndex: Number(data.correctIndex || 0),
    sortOrder: Number(data.sortOrder || 99),
    active: data.active === false ? false : active_(data.active)
  }, ["id", "classId", "question", "optionA", "optionB", "optionC", "optionD", "correctIndex", "sortOrder", "active"]);
}

function upsertById_(sheetName, data, fields) {
  const sheet = sh_(sheetName);
  const existing = rowObjs_(sheetName).find(function(row) { return String(row.obj.id) === String(data.id); });
  if (existing) {
    fields.forEach(function(field) { setField_(sheet, existing.row, field, data[field]); });
    if (headers_(sheet).indexOf("updatedAt") !== -1) setField_(sheet, existing.row, "updatedAt", new Date());
  } else {
    appendObject_(sheetName, Object.assign({}, data, { updatedAt: new Date() }));
  }
  return { ok: true, id: data.id };
}

function deactivateById_(sheetName, id) {
  if (!id) return { ok: false, error: "Missing id" };
  const sheet = sh_(sheetName);
  const existing = rowObjs_(sheetName).find(function(row) { return String(row.obj.id) === String(id); });
  if (!existing) return { ok: false, error: "Not found" };
  setField_(sheet, existing.row, "active", false);
  if (headers_(sheet).indexOf("updatedAt") !== -1) setField_(sheet, existing.row, "updatedAt", new Date());
  return { ok: true };
}

function submitSignupRequest_(data) {
  const recipient = adminEmail_();
  if (property_("BLUE_RIDGE_INTAKE_OPEN") !== "true" || !recipient) {
    throw new Error("Online training requests are currently unavailable.");
  }
  const username = String(data.username || "").trim();
  const name = String(data.fullNameOnLicense || "").trim();
  const contact = String(data.preferredContact || "").trim();
  if (!/^[a-zA-Z0-9][a-zA-Z0-9._-]{2,39}$/.test(username)) throw new Error("Choose a username of 3–40 letters, numbers, dots, underscores or hyphens.");
  if (!name || name.length > 120 || /[\r\n]/.test(name)) throw new Error("Enter your full name.");
  if (contact.length > 254 || !validPreferredContact_(contact)) throw new Error("Enter a valid email address or cell phone number.");
  if (!REQUEST_CLASS_IDS.includes(data.requestedClassId)) throw new Error("Select an available training course.");
  const cls = classById_(data.requestedClassId, true);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    clearSheetCache_(STUDENTS_SHEET);
    if (findStudent_(username)) throw new Error("That username is already in use. Please choose another.");

    const password = passwordHash_(data.password);
    appendObject_(STUDENTS_SHEET, {
      username: username,
      password: password,
      updatedAt: new Date(),
      fullNameOnLicense: name,
      active: false,
      preferredContact: contact
    });
    saveAssignments_(username, [data.requestedClassId]);
    appendObject_(SIGNUP_REQUESTS_SHEET, {
      createdAt: new Date(),
      fullNameOnLicense: name,
      requestedClassId: data.requestedClassId,
      requestedClassTitle: cls.title || data.requestedClassId,
      status: "new",
      preferredContact: contact
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  try {
    MailApp.sendEmail({
      to: recipient,
      subject: "New Blue Ridge ELDT training request",
      name: "Blue Ridge ELDT Website",
      body: "A new request is awaiting review in the Blue Ridge ELDT admin dashboard. Course: " + cls.title
    });
  } catch (err) {
    console.error("Training request notification failed");
  }

  return { ok: true, username: username, pendingApproval: true, backendVersion: BACKEND_VERSION };
}

function normalizeLicenseState_(value) {
  return String(value || "").trim().toUpperCase();
}

function validPreferredContact_(value) {
  const contact = String(value || "").trim();
  const emailOk = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(contact);
  const digits = contact.replace(/\D/g, "");
  const phoneOk = /^\+?[\d\s().-]+$/.test(contact) && digits.length >= 10 && digits.length <= 15;
  return emailOk || phoneOk;
}

function formatEmailDate_(value) {
  const date = value instanceof Date ? value : new Date(value);
  const timezone = Session.getScriptTimeZone() || "America/New_York";
  return Utilities.formatDate(date, timezone, "MMMM d, yyyy 'at' h:mm a z");
}

function sendTestEmail_(username, classId, cls, score, completedAt) {
  const recipient = adminEmail_();
  if (!recipient) return;

  MailApp.sendEmail({
    to: recipient,
    subject: "Blue Ridge ELDT training completed",
    name: "Blue Ridge ELDT Website",
    body: [
      "A student completed their assigned training and passed the final test.",
      "",
      "Completed: " + formatEmailDate_(completedAt),
      "Username: " + String(username || ""),
      "Training completed: " + String((cls && cls.title) || classId || ""),
      "Class ID: " + String(classId || ""),
      "Final test score: " + String(score) + "%",
      "Passing score: " + String((cls && cls.passingScore) || 80) + "%",
      "Result: Passed",
      "",
      "For security, the student's password is not included in this email."
    ].join("\n")
  });
}

function extractYouTubeId_(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  if (/^[a-zA-Z0-9_-]{11}$/.test(text)) return text;
  const match = text.match(/(?:youtu\.be\/|[?&]v=|embed\/|shorts\/)([a-zA-Z0-9_-]{11})/);
  return match ? match[1] : text;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
