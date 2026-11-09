const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");
const vm = require("node:vm");

const sheets = new Map();
function makeSheet(name) {
  const rows = [];
  return {
    rows,
    getName: () => name,
    getLastRow: () => rows.length,
    getLastColumn: () => rows[0]?.length || 0,
    appendRow: (row) => rows.push(row),
    getRange(row, col, height = 1, width = 1) {
      return {
        getValues: () => Array.from({ length: height }, (_, y) =>
          Array.from({ length: width }, (_, x) => rows[row + y - 1]?.[col + x - 1] ?? "")),
        setValues(values) {
          values.forEach((line, y) => line.forEach((value, x) => {
            const target = rows[row + y - 1] ??= [];
            target[col + x - 1] = value;
          }));
        },
        setValue(value) { this.setValues([[value]]); }
      };
    }
  };
}

const properties = {
  BLUE_RIDGE_SPREADSHEET_ID: "synthetic-business-sheet",
  BLUE_RIDGE_ADMIN_EMAIL: "owner@example.test",
  BLUE_RIDGE_VIDEO_SECONDS: '{"zeaHTafu4CY":60,"-qXt8htJ9h4":60,"RS4K5FCL988":60,"TLeq0WikSmU":60,"cMML4tQdVvY":60}'
};
const cache = new Map();
const emails = [];
let now = 1700000000000;
let failMail = false;
const context = vm.createContext({
  console: { log() {}, error() {} },
  Date: class extends Date { static now() { return now; } },
  PropertiesService: { getScriptProperties: () => ({ getProperty: (key) => properties[key] }) },
  SpreadsheetApp: {
    openById(id) {
      assert.equal(id, properties.BLUE_RIDGE_SPREADSHEET_ID);
      return {
        getSheetByName: (name) => sheets.get(name),
        insertSheet: (name) => {
          const sheet = makeSheet(name);
          sheets.set(name, sheet);
          return sheet;
        }
      };
    },
    flush() {}
  },
  Utilities: {
    getUuid: crypto.randomUUID,
    DigestAlgorithm: { SHA_256: "sha256" },
    computeDigest: (algorithm, value) => Array.from(crypto.createHash(algorithm).update(value).digest()),
    formatDate: () => "Test date"
  },
  Session: { getScriptTimeZone: () => "America/New_York" },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  CacheService: { getScriptCache: () => ({
    put: (key, value, ttl) => cache.set(key, { value, expires: now + ttl * 1000 }),
    get: (key) => cache.get(key)?.expires > now ? cache.get(key).value : null
  }) },
  ContentService: { MimeType: { JSON: "json" }, createTextOutput: (value) => ({ setMimeType: () => value }) },
  MailApp: { sendEmail: (mail) => { if (failMail) throw Error("quota"); emails.push(mail); } }
});
for (const name of ["bcrypt.js", "app.js"]) vm.runInContext(fs.readFileSync(name, "utf8"), context);
const run = (code) => vm.runInContext(code, context);
const get = (action, parameters = {}) => JSON.parse(context.doGet({ parameter: { action, ...parameters } }));
const post = (action, body = {}) => JSON.parse(context.doPost({ postData: { contents: JSON.stringify({ action, ...body }) } }));
const ok = (response) => { assert.equal(response.ok, true, response.error); return response; };
const rejected = (response, error) => { assert.equal(response.ok, false); assert.match(response.error, error); };

run('withRequestCache_(setupSheets_)');
sheets.get("Admins").appendRow(["qa-admin", "legacy-admin"]);
sheets.get("Students").appendRow(["qa-existing", "legacy-student", "", "Existing", "", "", true]);
rejected(get("adminLogin", { username: "qa-admin", password: "legacy-admin" }), /unknown action/i);
rejected(post("adminLogin", { username: "qa-admin", password: "legacy-admin" }), /invalid/i);
run("migrateBlueRidgePasswords()");
const firstHash = sheets.get("Admins").rows[1][1];
assert.match(firstHash, /^\$2b\$12\$/);
run("migrateBlueRidgePasswords()");
assert.equal(sheets.get("Admins").rows[1][1], firstHash);
ok(post("validateLogin", { username: "qa-existing", password: "legacy-student" }));
const admin = { adminToken: ok(post("adminLogin", { username: "qa-admin", password: "legacy-admin" })).token };
assert.equal(get("getVersion").intakeOpen, false);
rejected(get("listStudents", admin), /unknown action/i);
rejected(get("listModules"), /unknown action/i);
rejected(post("listStudents"), /expired/i);
const signup = {
  username: "qa-upgrade",
  password: "qa-only-password-123",
  fullNameOnLicense: "=HYPERLINK(\"https://example.test\")",
  licenseNumber: "sensitive-data-test",
  dob: "2000-01-01",
  preferredContact: "student@example.test",
  requestedClassId: "class-b-to-a"
};
rejected(post("submitSignupRequest", signup), /unavailable/i);
properties.BLUE_RIDGE_INTAKE_OPEN = "true";
assert.equal(get("getVersion").intakeOpen, true);
rejected(post("submitSignupRequest", { ...signup, requestedClassId: "hazmat" }), /available/i);
ok(post("submitSignupRequest", signup));
assert.equal(emails.length, 1);
assert.equal(emails[0].to, properties.BLUE_RIDGE_ADMIN_EMAIL);
assert.doesNotMatch(JSON.stringify(emails[0]), /qa-only|sensitive-data|2000-01-01|student@example/);
const requestRows = sheets.get("SignupRequests").rows;
assert.equal(requestRows.length, 2);
assert.doesNotMatch(JSON.stringify(requestRows), /sensitive-data|2000-01-01/);
assert.ok(requestRows[1].includes("'" + signup.fullNameOnLicense));
const pending = ok(post("listStudents", admin)).pendingStudents[0];
assert.equal(pending.username, signup.username);
assert.deepEqual(pending.classes, ["class-b-to-a"]);
assert.equal("password" in pending, false);
rejected(post("validateLogin", signup), /invalid/i);
rejected(post("approveStudent", { username: signup.username }), /expired/i);
ok(post("updateStudent", { ...admin, username: signup.username, fullNameOnLicense: "QA Upgrade",
  preferredContact: signup.preferredContact, classes: ["class-b-to-a"] }));
const student = { username: signup.username, studentToken: ok(post("validateLogin", signup)).token };
assert.deepEqual(ok(post("getStudentDashboard", student)).classes.map((course) => course.id), ["class-b-to-a"]);
rejected(post("getStatus", { ...student, classId: "class-a-b" }), /assign/i);
rejected(post("getStudentDashboard", { ...student, username: "qa-existing" }), /expired/i);
const course = { ...student, classId: "class-b-to-a" };
rejected(post("listTestQuestions", course), /finish/i);
const lesson = { ...course, moduleId: "10" };
rejected(post("logModule", lesson), /minimum/i);
ok(post("beginModule", lesson));
rejected(post("logModule", lesson), /minimum/i);
now += 54_001;
assert.equal(ok(post("logModule", lesson)).modules[0].complete, true);
assert.equal(sheets.get("Progress").rows.length, 2);
ok(post("logModule", lesson));
assert.equal(sheets.get("Progress").rows.length, 2);
ok(post("saveTestQuestion", { ...admin, id: "q-1", classId: course.classId,
  question: "QA question", optionA: "A", optionB: "B", correctIndex: 1, active: true }));
assert.equal(ok(post("listTestQuestions", course)).questions[0].correctIndex, undefined);
assert.equal(ok(post("listTestQuestions", { ...admin, classId: course.classId })).questions[0].correctIndex, 1);
rejected(post("logTest", { ...course, complete: true, score: 100 }), /answer/i);
assert.equal(ok(post("logTest", { ...course, answers: { "q-1": 0 } })).testPassed, false);
assert.equal(ok(post("logTest", { ...course, answers: { "q-1": 1 } })).attemptScore, 100);
const best = ok(post("logTest", { ...course, answers: { "q-1": 0 } }));
assert.equal(best.testPassed, true);
assert.equal(best.testScore, 100);
assert.equal(emails.length, 2);

failMail = true;
const initial = { ...signup, username: "qa-initial", requestedClassId: "class-a-b" };
ok(post("submitSignupRequest", initial));
ok(post("approveStudent", { ...admin, username: initial.username }));
const initialStudent = { username: initial.username,
  studentToken: ok(post("validateLogin", initial)).token, classId: initial.requestedClassId };
const initialCourse = ok(post("getStudentDashboard", initialStudent)).classes[0];
assert.equal(initialCourse.id, "class-a-b");
for (const module of initialCourse.status.modules) {
  const lesson = { ...initialStudent, moduleId: module.id };
  ok(post("beginModule", lesson));
  now += 54_001;
  ok(post("logModule", lesson));
}
ok(post("saveTestQuestion", { ...admin, id: "qa-initial-1", classId: "class-a-b",
  question: "QA initial question", optionA: "A", correctIndex: 0, active: true }));
assert.equal(ok(post("logTest", { ...initialStudent, answers: { "qa-initial-1": 0 } })).testPassed, true);
ok(post("archiveStudent", { ...admin, username: initial.username }));
rejected(post("getStudentDashboard", initialStudent), /archived/i);
console.log("Passed synthetic Apps Script workflows for initial Class A/B and B-to-A, including access and negative checks.");
