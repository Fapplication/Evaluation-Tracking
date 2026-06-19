// ============================================================
//  apps.js — Optimised API Bridge (parallel calls + cache)
// ============================================================

const APP_URL = "https://script.google.com/macros/s/AKfycbwZy4NZo-oC0Sd8yuX1QP1jZAR3kgagioWvu8ll1BZ5fMX7_OVInmSsGBTAXDqB9wFK/exec";

// ── Simple in-memory cache (TTL 60s) ──
const _cache = {};
function _cacheGet(key) {
  const e = _cache[key];
  if (e && Date.now() - e.ts < 60000) return e.val;
  return null;
}
function _cacheSet(key, val) { _cache[key] = { val, ts: Date.now() }; }
function _cacheClear(key)    { delete _cache[key]; }

// ── Core fetch with timeout ──
async function callAPI(action, data = {}, useCache = false) {
  const cacheKey = action + JSON.stringify(data);
  if (useCache) {
    const cached = _cacheGet(cacheKey);
    if (cached) return cached;
  }
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 30000);
    const response = await fetch(APP_URL, {
      method : "POST",
      headers: { "Content-Type": "text/plain" },
      body   : JSON.stringify({ action, data }),
      signal : controller.signal
    });
    clearTimeout(timer);
    if (!response.ok) throw new Error("Network error: " + response.status);
    const result = await response.json();
    if (useCache && result.success) _cacheSet(cacheKey, result);
    return result;
  } catch (err) {
    if (err.name === 'AbortError') return { success: false, message: 'Request timed out. Please try again.' };
    console.error("API Error [" + action + "]:", err);
    return { success: false, message: err.message };
  }
}

// ── Session ──
const Session = {
  set(sessionId, user) {
    localStorage.setItem("sessionId", sessionId);
    localStorage.setItem("user", JSON.stringify(user));
  },
  getSessionId() { return localStorage.getItem("sessionId"); },
  getUser() {
    try { return JSON.parse(localStorage.getItem("user")); }
    catch { return null; }
  },
  clear()      { localStorage.removeItem("sessionId"); localStorage.removeItem("user"); },
  isLoggedIn() { return !!localStorage.getItem("sessionId"); },
  guard()      { if (!Session.isLoggedIn()) window.location.href = "login.html"; }
};

// ── Auth API ──
const Auth = {
  async login(username, password) {
    return callAPI("login", { username, password });
  },
  async logout() {
    const sessionId = Session.getSessionId();
    Session.clear();
    // Fire-and-forget logout — don't wait for server
    callAPI("logout", { sessionId }).catch(() => {});
    return { success: true };
  },
  async registerUser(userData) {
    return callAPI("registerUser", userData);
  },
  async checkUsername(username) {
    return callAPI("checkUsername", { username });
  }
};

// ── Bid API ──
const BidAPI = {
  async registerBid(bidData) {
    _cacheClear("getAllBids");
    _cacheClear("getSummaryReport");
    return callAPI("registerBid", { bidData, sessionId: Session.getSessionId() });
  },
  async getAllBids() {
    return callAPI("getAllBids", { sessionId: Session.getSessionId() }, true);
  },
  async getBidById(bidId) {
    return callAPI("getBidById", { bidId, sessionId: Session.getSessionId() }, true);
  },
  async getSummaryReport() {
    return callAPI("getSummaryReport", { sessionId: Session.getSessionId() }, true);
  },
  invalidateCache() {
    _cacheClear("getAllBids");
    _cacheClear("getSummaryReport");
  }
};

// ── Bidder API — parallel batch save ──
const BidderAPI = {
  async addBidder(bidderData) {
    return callAPI("addBidder", { bidderData, sessionId: Session.getSessionId() });
  },
  // Save ALL bidders in parallel instead of sequentially
  async addBiddersBatch(bidderList, bidId) {
    const promises = bidderList.map(b =>
      callAPI("addBidder", { bidderData: { ...b, bidId }, sessionId: Session.getSessionId() })
    );
    return Promise.all(promises);
  },
  async updateBidder(bidderData) {
    return callAPI("updateBidder", { bidderData, sessionId: Session.getSessionId() });
  }
};

// ── Clarification API — parallel batch save ──
const ClarifAPI = {
  async addClarification(clarifData) {
    return callAPI("addClarification", { clarifData, sessionId: Session.getSessionId() });
  },
  // Save ALL clarifications in parallel
  async addClarifsBatch(clarifList, bidId) {
    const promises = clarifList.map(c =>
      callAPI("addClarification", { clarifData: { ...c, bidId }, sessionId: Session.getSessionId() })
    );
    return Promise.all(promises);
  },
  async updateClarification(clarifData) {
    return callAPI("updateClarification", { clarifData, sessionId: Session.getSessionId() });
  }
};

// ── Activity API ──
const ActivityAPI = {
  async updateTechActivities(actData) {
    BidAPI.invalidateCache();
    return callAPI("updateTechActivities", { actData, sessionId: Session.getSessionId() });
  },
  async updateFinActivities(actData) {
    BidAPI.invalidateCache();
    return callAPI("updateFinActivities", { actData, sessionId: Session.getSessionId() });
  }
};

// ── UI Utilities ──
const UI = {
  showToast(msg, type = "info") {
    const t    = document.getElementById("toast");
    const icon = document.getElementById("toastIcon");
    const txt  = document.getElementById("toastMsg");
    if (!t) return;
    icon.textContent = type === "error" ? "❌" : type === "success" ? "✅" : "ℹ️";
    txt.textContent  = msg;
    t.classList.add("show");
    setTimeout(() => t.classList.remove("show"), 3200);
  },
  initUserBar() {
    const user = Session.getUser();
    if (!user) return;
    const avatar = document.getElementById("userAvatar");
    const nameEl = document.getElementById("userNameSidebar");
    const roleEl = document.getElementById("userRoleSidebar");
    if (avatar) avatar.textContent = (user.fullName || user.username || "?").charAt(0).toUpperCase();
    if (nameEl) nameEl.textContent = user.fullName || user.username;
    if (roleEl) roleEl.textContent = user.role;
  },
  async doLogout() {
    await Auth.logout();
    window.location.href = "login.html";
  },
  formatDate(val) {
    if (!val) return "—";
    const d = new Date(val);
    return isNaN(d) ? "—" : d.toLocaleDateString("en-GB");
  }
};
