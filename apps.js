// ============================================================
//  apps.js — API Bridge: GitHub Pages → Apps Script Web App
//  Include this in every HTML page via <script src="apps.js">
// ============================================================

const APP_URL = "https://script.google.com/macros/s/AKfycbzpAR52VWNqAlW7N7aqC3StIaTAGmWTE1tDfI0a5YRs6GRB7jqjksYj6HR0eCgwnV0n/exec";
// ↑ Replace YOUR_DEPLOYMENT_ID with your actual deployment ID after deploying

// ──────────────────────────────────────────────
//  Core fetch wrapper
// ──────────────────────────────────────────────
async function callAPI(action, data = {}) {
  try {
    const response = await fetch(APP_URL, {
      method  : "POST",
      headers : { "Content-Type": "text/plain" }, // required for Apps Script CORS
      body    : JSON.stringify({ action, data })
    });
    if (!response.ok) throw new Error("Network error: " + response.status);
    return await response.json();
  } catch (err) {
    console.error("API Error [" + action + "]:", err);
    return { success: false, message: err.message };
  }
}

// ──────────────────────────────────────────────
//  Session helpers (localStorage)
// ──────────────────────────────────────────────
const Session = {
  set(sessionId, user) {
    localStorage.setItem("sessionId", sessionId);
    localStorage.setItem("user", JSON.stringify(user));
  },
  getSessionId() { return localStorage.getItem("sessionId"); },
  getUser()      {
    try { return JSON.parse(localStorage.getItem("user")); }
    catch { return null; }
  },
  clear()        { localStorage.removeItem("sessionId"); localStorage.removeItem("user"); },
  isLoggedIn()   { return !!localStorage.getItem("sessionId"); },
  guard()        { if (!Session.isLoggedIn()) { window.location.href = "login.html"; } }
};

// ──────────────────────────────────────────────
//  Auth API
// ──────────────────────────────────────────────
const Auth = {
  async login(username, password) {
    return await callAPI("login", { username, password });
  },
  async logout() {
    const sessionId = Session.getSessionId();
    const result    = await callAPI("logout", { sessionId });
    Session.clear();
    return result;
  },
  async registerUser(userData) {
    return await callAPI("registerUser", userData);
  },
  async checkUsername(username) {
    return await callAPI("checkUsername", { username });
  }
};

// ──────────────────────────────────────────────
//  Bid API
// ──────────────────────────────────────────────
const BidAPI = {
  async registerBid(bidData) {
    return await callAPI("registerBid", {
      bidData,
      sessionId: Session.getSessionId()
    });
  },
  async getAllBids() {
    return await callAPI("getAllBids", { sessionId: Session.getSessionId() });
  },
  async getBidById(bidId) {
    return await callAPI("getBidById", { bidId, sessionId: Session.getSessionId() });
  },
  async getSummaryReport() {
    return await callAPI("getSummaryReport", { sessionId: Session.getSessionId() });
  }
};

// ──────────────────────────────────────────────
//  Bidder API
// ──────────────────────────────────────────────
const BidderAPI = {
  async addBidder(bidderData) {
    return await callAPI("addBidder", {
      bidderData,
      sessionId: Session.getSessionId()
    });
  },
  async updateBidder(bidderData) {
    return await callAPI("updateBidder", {
      bidderData,
      sessionId: Session.getSessionId()
    });
  }
};

// ──────────────────────────────────────────────
//  Clarification API
// ──────────────────────────────────────────────
const ClarifAPI = {
  async addClarification(clarifData) {
    return await callAPI("addClarification", {
      clarifData,
      sessionId: Session.getSessionId()
    });
  },
  async updateClarification(clarifData) {
    return await callAPI("updateClarification", {
      clarifData,
      sessionId: Session.getSessionId()
    });
  }
};

// ──────────────────────────────────────────────
//  Activity API
// ──────────────────────────────────────────────
const ActivityAPI = {
  async updateTechActivities(actData) {
    return await callAPI("updateTechActivities", {
      actData,
      sessionId: Session.getSessionId()
    });
  },
  async updateFinActivities(actData) {
    return await callAPI("updateFinActivities", {
      actData,
      sessionId: Session.getSessionId()
    });
  }
};

// ──────────────────────────────────────────────
//  UI Utilities (shared across pages)
// ──────────────────────────────────────────────
const UI = {
  showToast(msg, type = "info") {
    const t    = document.getElementById("toast");
    const icon = document.getElementById("toastIcon");
    const txt  = document.getElementById("toastMsg");
    if (!t) return;
    icon.textContent = type === "error" ? "❌ " : type === "success" ? "✅ " : "ℹ️ ";
    txt.textContent  = msg;
    t.classList.add("show");
    setTimeout(() => t.classList.remove("show"), 3200);
  },

  showAlert(elId, msg, type) {
    const el = document.getElementById(elId);
    if (!el) return;
    el.className     = "alert " + type;
    const icon = el.querySelector(".alert-icon");
    const txt  = el.querySelector(".alert-msg");
    if (icon) icon.textContent = type === "error" ? "⚠️" : "✅";
    if (txt)  txt.textContent  = msg;
  },

  hideAlert(elId) {
    const el = document.getElementById(elId);
    if (el) el.className = "alert hidden";
  },

  setLoading(btnId, txtId, spinnerId, on) {
    const btn = document.getElementById(btnId);
    const txt = document.getElementById(txtId);
    const spn = document.getElementById(spinnerId);
    if (btn) btn.disabled          = on;
    if (txt) txt.style.display     = on ? "none" : "inline";
    if (spn) spn.style.display     = on ? "inline-block" : "none";
  },

  initUserBar() {
    const user = Session.getUser();
    if (!user) return;
    const avatar   = document.getElementById("userAvatar");
    const nameEl   = document.getElementById("userNameSidebar");
    const roleEl   = document.getElementById("userRoleSidebar");
    if (avatar) avatar.textContent = (user.fullName || user.username).charAt(0).toUpperCase();
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
