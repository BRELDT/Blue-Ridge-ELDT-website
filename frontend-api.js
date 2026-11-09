const AUTH_STORAGE_KEY = "cdl_username";
const STUDENT_TOKEN_KEY = "cdl_student_token";
const ADMIN_TOKEN_KEY = "cdl_admin_token";
const DEFAULT_CLASS_ID = "class-a-b";

const DEFAULT_CLASSES = [
  { id: DEFAULT_CLASS_ID, title: "Class A and B", description: "CDL Class A/B ELDT training", passingScore: 80, requiredWatchPercent: 0.9, sortOrder: 1, active: true },
  { id: "class-b-to-a", title: "Class B to A Upgrade", description: "Upgrade training from Class B to Class A", passingScore: 80, requiredWatchPercent: 0.9, sortOrder: 2, active: true },
  { id: "passenger", title: "Passenger and School Bus Endorsement", description: "Passenger and school bus endorsement training", passingScore: 80, requiredWatchPercent: 0.9, sortOrder: 3, active: true },
  { id: "school-bus", title: "School Bus Endorsement", description: "School bus endorsement training", passingScore: 80, requiredWatchPercent: 0.9, sortOrder: 4, active: false },
  { id: "tanker", title: "Tanker Endorsement", description: "Tanker endorsement training", passingScore: 80, requiredWatchPercent: 0.9, sortOrder: 5, active: false },
  { id: "hazmat", title: "Hazmat Endorsement", description: "Hazmat endorsement training", passingScore: 80, requiredWatchPercent: 0.9, sortOrder: 6, active: true }
];

function qs(param) {
  return new URLSearchParams(location.search).get(param);
}

function normalizeWatchPercent(value) {
  const n = Number(value);
  if (!n || Number.isNaN(n)) return 0.9;
  return n > 1 ? n / 100 : n;
}

function displayPercent(value) {
  return Math.round(normalizeWatchPercent(value) * 100);
}

function extractYouTubeId(value) {
  const text = String(value || "").trim();
  if (!text) return "";

  const direct = text.match(/^[a-zA-Z0-9_-]{11}$/);
  if (direct) return text;

  const patterns = [
    /youtu\.be\/([a-zA-Z0-9_-]{11})/,
    /[?&]v=([a-zA-Z0-9_-]{11})/,
    /embed\/([a-zA-Z0-9_-]{11})/,
    /shorts\/([a-zA-Z0-9_-]{11})/
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return match[1];
  }
  return text;
}

function authParams(action) {
  if (["adminLogin", "validateLogin", "submitSignupRequest"].includes(action)) return {};
  const student = ["getStatus", "getStudentDashboard", "beginModule", "logModule", "logTest"].includes(action) ||
    (action === "listTestQuestions" && !location.pathname.endsWith("/admin.html"));
  return student
    ? { studentToken: localStorage.getItem(STUDENT_TOKEN_KEY) }
    : { adminToken: sessionStorage.getItem(ADMIN_TOKEN_KEY) };
}

function isSessionExpiredResponse(data) {
  return !!(data && !data.ok && /session expired/i.test(String(data.error || "")));
}

function handleSessionExpired(data) {
  if (!isSessionExpiredResponse(data)) return false;
  localStorage.removeItem(AUTH_STORAGE_KEY);
  localStorage.removeItem(STUDENT_TOKEN_KEY);
  sessionStorage.removeItem(ADMIN_TOKEN_KEY);
  alert("Your session expired. Please sign in again.");
  location.href = "index.html";
  return true;
}

async function apiGet(action, params = {}) {
  const url = new URL(APP_SCRIPT_URL);
  url.searchParams.set("action", action);

  Object.entries(params).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") {
      url.searchParams.set(key, value);
    }
  });

  const response = await fetch(url.toString(), { cache: "no-store" });
  const data = await response.json();
  return data;
}

let verifiedBackend;
async function secureBackend() {
  if (!verifiedBackend) {
    const version = await apiGet("getVersion");
    if (version.version !== "blue-ridge-v7-secure") {
      throw new Error("Blue Ridge ELDT online services are temporarily unavailable.");
    }
    verifiedBackend = version;
  }
  return verifiedBackend;
}

async function apiPost(action, body = {}) {
  await secureBackend();
  const response = await fetch(APP_SCRIPT_URL, {
    method: "POST",
    redirect: "follow",
    body: JSON.stringify({ ...body, ...authParams(action), action })
  });

  const data = await response.json();
  if (action === "adminLogin" && data.ok && data.token) {
    sessionStorage.setItem(ADMIN_TOKEN_KEY, data.token);
  }
  handleSessionExpired(data);
  return data;
}

async function login(username, password) {
  try {
    const data = await apiPost("validateLogin", { username, password });
    if (data && data.ok) {
      localStorage.setItem(AUTH_STORAGE_KEY, data.username || username);
      if (data.token) localStorage.setItem(STUDENT_TOKEN_KEY, data.token);
      return true;
    }
    localStorage.removeItem(STUDENT_TOKEN_KEY);
    return false;
  } catch (err) {
    console.error(err);
    localStorage.removeItem(STUDENT_TOKEN_KEY);
    return false;
  }
}

function requireAuth() {
  const username = localStorage.getItem(AUTH_STORAGE_KEY);
  if (!username || !localStorage.getItem(STUDENT_TOKEN_KEY)) {
    location.href = "index.html";
    return "";
  }
  return username;
}

function logout() {
  localStorage.removeItem(AUTH_STORAGE_KEY);
  localStorage.removeItem(STUDENT_TOKEN_KEY);
  sessionStorage.removeItem(ADMIN_TOKEN_KEY);
  location.href = "index.html";
}

async function apiGetStatus(username, classId = DEFAULT_CLASS_ID) {
  return apiPost("getStatus", { username, classId });
}

async function apiGetStudentDashboard(username) {
  return apiPost("getStudentDashboard", { username });
}

async function apiListClasses(activeOnly = false) {
  try {
    const data = await apiGet("listClasses", { activeOnly: activeOnly ? "true" : "" });
    if (data && data.ok && Array.isArray(data.classes)) return data;
  } catch (err) {
    console.warn(err);
  }

  return {
    ok: true,
    classes: DEFAULT_CLASSES.filter(cls => !activeOnly || cls.active !== false)
  };
}

async function apiListModules(classId = "", activeOnly = false) {
  return apiPost("listModules", { classId, activeOnly: activeOnly ? "true" : "" });
}

async function apiListTestQuestions(classId = "", activeOnly = false) {
  return apiPost("listTestQuestions", {
    classId,
    username: localStorage.getItem(AUTH_STORAGE_KEY),
    activeOnly: activeOnly ? "true" : ""
  });
}

async function apiLogModule(username, moduleId, classId = DEFAULT_CLASS_ID) {
  return apiPost("logModule", { username, moduleId, classId });
}

async function apiLogTest(username, answers, classId = DEFAULT_CLASS_ID) {
  return apiPost("logTest", { username, answers, classId });
}

async function apiSubmitSignupRequest(payload) {
  return apiPost("submitSignupRequest", payload);
}

function allModulesComplete(status) {
  if (!status || !Array.isArray(status.modules)) return false;
  return status.modules.length > 0 && status.modules.every(module => !!module.complete);
}
