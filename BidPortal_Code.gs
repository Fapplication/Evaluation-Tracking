// ═══════════════════════════════════════════════════════════════════
//  BID EVALUATION PORTAL — Google Apps Script Backend
//  Deploy as Web App: Execute as Me, Anyone can access (or org only)
// ═══════════════════════════════════════════════════════════════════

const CONFIG = {
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',
  SESSION_HOURS:  8,
  ORG_NAME:       'Bid Evaluation Team',

  // ── SLA in working days per activity ──────────────────────────
  SLA: {
    // Technical
    'tech_doc_collection':    5,
    'clarification_request':  3,
    'tech_eval_progress':    10,
    'tech_report_submittal':  4,
    // Financial
    'fin_doc_collection':     4,
    'engineering_estimation': 5,
    'doc_authentication':     3,
    'fin_report_evaluation':  7,
  },
};

const SHEETS = {
  USERS:          'Users',
  BIDS:           'Bids',
  BIDDERS:        'Bidders',
  ACTIVITIES:     'Activities',
  CLARIFICATIONS: 'Clarifications',
  SESSIONS:       'Sessions',
};

// ════════════════════════════════════════════════════════════════════
//  ENTRY POINTS
// ════════════════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const token  = e.parameter.token  || '';

    if (action === 'ping') return jsonOk({ status: 'ok', time: new Date().toISOString() });
    if (action === 'login') return handleLogin(e);

    switch (action) {
      case 'getDashboard':         return getDashboard(e);
      case 'getBids':              return getBids(e);
      case 'getBidDetail':         return getBidDetail(e);
      case 'getActivities':        return getActivities(e);
      case 'getClarifications':    return getClarifications(e);
      case 'getDelayReport':       return getDelayReport(e);
      default:                     return jsonError('Unknown action: ' + action);
    }
  } catch (err) {
    return jsonError('Server error: ' + err.message);
  }
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents || '{}');
    const action = body.action || '';

    if (action === 'login')    return handleLoginPost(body);
    if (action === 'register') return registerUser(body);

    switch (action) {
      // Bids
      case 'createBid':            return createBid(body);
      case 'updateBid':            return updateBid(body);
      case 'deleteBid':            return deleteBid(body);

      // Bidders
      case 'addBidder':            return addBidder(body);
      case 'updateBidder':         return updateBidder(body);
      case 'removeBidder':         return removeBidder(body);

      // Activities
      case 'startActivity':        return startActivity(body);
      case 'updateActivityStatus': return updateActivityStatus(body);
      case 'completeActivity':     return completeActivity(body);

      // Clarifications
      case 'addClarification':     return addClarification(body);
      case 'updateClarification':  return updateClarification(body);

      default: return jsonError('Unknown POST action: ' + action);
    }
  } catch (err) {
    return jsonError('Server error: ' + err.message);
  }
}

// ════════════════════════════════════════════════════════════════════
//  SPREADSHEET HELPERS
// ════════════════════════════════════════════════════════════════════

function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getOrCreateSheet(name) {
  const ss    = getSpreadsheet();
  let   sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function sheetToObjects(sheetName) {
  const sheet = getOrCreateSheet(sheetName);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1)
    .filter(row => row.some(c => c !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function appendRow(sheetName, rowObj) {
  const sheet   = getOrCreateSheet(sheetName);
  const data    = sheet.getDataRange().getValues();
  const headers = data.length > 0 ? data[0].map(h => String(h).trim()) : [];
  if (headers.length === 0) {
    const keys = Object.keys(rowObj);
    sheet.appendRow(keys);
    sheet.appendRow(keys.map(k => rowObj[k] !== undefined ? rowObj[k] : ''));
    return;
  }
  sheet.appendRow(headers.map(h => rowObj[h] !== undefined ? rowObj[h] : ''));
}

function updateRowByKey(sheetName, keyCol, keyVal, updates) {
  const sheet   = getOrCreateSheet(sheetName);
  const data    = sheet.getDataRange().getValues();
  if (data.length < 2) return false;
  const headers = data[0].map(h => String(h).trim());
  const keyIdx  = headers.indexOf(keyCol);
  if (keyIdx === -1) return false;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyIdx]) === String(keyVal)) {
      Object.entries(updates).forEach(([col, val]) => {
        const cIdx = headers.indexOf(col);
        if (cIdx !== -1) sheet.getRange(i + 1, cIdx + 1).setValue(val);
      });
      return true;
    }
  }
  return false;
}

function deleteRowByKey(sheetName, keyCol, keyVal) {
  const sheet   = getOrCreateSheet(sheetName);
  const data    = sheet.getDataRange().getValues();
  if (data.length < 2) return false;
  const headers = data[0].map(h => String(h).trim());
  const col     = headers.indexOf(keyCol);
  if (col === -1) return false;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][col]) === String(keyVal)) { sheet.deleteRow(i + 1); return true; }
  }
  return false;
}

// ════════════════════════════════════════════════════════════════════
//  INITIALISATION
// ════════════════════════════════════════════════════════════════════

function initializeSheets() {
  const ss = getSpreadsheet();

  const schema = {
    [SHEETS.USERS]: [
      'UserID','FullName','Email','Role','PasswordHash','Status','CreatedAt','LastLogin'
    ],
    [SHEETS.BIDS]: [
      'BidID','BidNumber','BidName','PAC','PurchaseMethod',
      'Phase','PhaseStatus','CreatedBy','CreatedAt','UpdatedAt','Notes'
    ],
    [SHEETS.BIDDERS]: [
      'BidderID','BidID','CompanyName','TechnicalScore','FinancialScore',
      'TotalScore','Status','Notes','AddedAt'
    ],
    [SHEETS.ACTIVITIES]: [
      'ActivityID','BidID','Phase','ActivityKey','ActivityName',
      'Status','StartDate','DueDate','CompletedDate','CompletedBy',
      'SLADays','IsDelayed','Notes','UpdatedAt'
    ],
    [SHEETS.CLARIFICATIONS]: [
      'ClarificationID','BidID','Subject','RequestedDate',
      'ResponseDate','Status','Notes','CreatedBy','CreatedAt'
    ],
    [SHEETS.SESSIONS]: [
      'Token','UserID','CreatedAt','ExpiresAt'
    ],
  };

  Object.entries(schema).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#1a3a5c')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  });

  // Seed admin user
  const users = sheetToObjects(SHEETS.USERS);
  if (!users.find(u => u['Email'] === 'admin@bid.gov')) {
    appendRow(SHEETS.USERS, {
      UserID:       'USR001',
      FullName:     'System Admin',
      Email:        'admin@bid.gov',
      Role:         'admin',
      PasswordHash: hashPassword('admin123'),
      Status:       'active',
      CreatedAt:    new Date().toISOString(),
      LastLogin:    '',
    });
  }

  return jsonOk({ message: 'Sheets initialized. Admin: admin@bid.gov / admin123' });
}

// ════════════════════════════════════════════════════════════════════
//  AUTH
// ════════════════════════════════════════════════════════════════════

function handleLogin(e) {
  return processLogin(e.parameter.email, e.parameter.pw);
}

function handleLoginPost(body) {
  return processLogin(body.email, body.pw);
}

function processLogin(email, pw) {
  if (!email || !pw) return jsonError('Email and password are required.');
  const users = sheetToObjects(SHEETS.USERS);
  const user  = users.find(u => String(u['Email']).toLowerCase() === String(email).toLowerCase());
  if (!user) return jsonError('No account found with this email.');
  const stored = user['PasswordHash'] || '';
  if (stored !== pw && stored !== hashPassword(pw)) return jsonError('Incorrect password.');
  if (user['Status'] !== 'active') return jsonError('Account is inactive. Contact admin.');

  const token  = generateToken();
  const expiry = new Date(Date.now() + CONFIG.SESSION_HOURS * 3600 * 1000).toISOString();
  appendRow(SHEETS.SESSIONS, {
    Token: token, UserID: user['UserID'],
    CreatedAt: new Date().toISOString(), ExpiresAt: expiry
  });
  updateRowByKey(SHEETS.USERS, 'UserID', user['UserID'], { LastLogin: new Date().toISOString() });

  return jsonOk({
    success: true, token,
    user: { id: user['UserID'], name: user['FullName'], email: user['Email'], role: user['Role'] }
  });
}

function registerUser(body) {
  const { name, email, pw, role } = body;
  if (!name || !email || !pw) return jsonError('Name, email and password are required.');
  const users = sheetToObjects(SHEETS.USERS);
  if (users.find(u => String(u['Email']).toLowerCase() === String(email).toLowerCase())) {
    return jsonError('An account with this email already exists.');
  }
  const uid = 'USR' + Date.now().toString(36).toUpperCase();
  appendRow(SHEETS.USERS, {
    UserID:       uid,
    FullName:     name,
    Email:        email,
    Role:         role || 'evaluator',
    PasswordHash: hashPassword(pw),
    Status:       'active',
    CreatedAt:    new Date().toISOString(),
    LastLogin:    '',
  });
  return jsonOk({ success: true, message: 'Account created. You can now log in.' });
}

function validateSession(token) {
  if (!token) return null;
  const rows = sheetToObjects(SHEETS.SESSIONS);
  const s    = rows.find(r => r['Token'] === token);
  if (!s) return null;
  if (new Date(s['ExpiresAt']) < new Date()) return null;
  return s;
}

function generateToken() {
  return Utilities.getUuid().replace(/-/g, '') + Date.now().toString(36);
}

function hashPassword(pw) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ════════════════════════════════════════════════════════════════════
//  BID MANAGEMENT
// ════════════════════════════════════════════════════════════════════

const ALL_ACTIVITIES = {
  technical: [
    { key: 'tech_doc_collection',   name: 'Technical Document Collection' },
    { key: 'clarification_request', name: 'Clarification Request'          },
    { key: 'tech_eval_progress',    name: 'Technical Evaluation Progress'  },
    { key: 'tech_report_submittal', name: 'Technical Evaluation Report Submittal' },
  ],
  financial: [
    { key: 'fin_doc_collection',    name: 'Financial Document Collection'  },
    { key: 'engineering_estimation',name: 'Engineering Estimation & Approved Budget' },
    { key: 'doc_authentication',    name: 'Document Authentication'        },
    { key: 'fin_report_evaluation', name: 'Financial Report Evaluation'    },
  ],
};

function createBid(body) {
  const { bidNumber, bidName, pac, purchaseMethod, notes, createdBy } = body;
  if (!bidNumber || !bidName || !pac || !purchaseMethod) {
    return jsonError('Bid Number, Name, PAC and Purchase Method are required.');
  }

  const bids = sheetToObjects(SHEETS.BIDS);
  if (bids.find(b => String(b['BidNumber']) === String(bidNumber))) {
    return jsonError('A bid with this number already exists.');
  }

  const bidId   = 'BID' + Date.now().toString(36).toUpperCase();
  const now     = new Date().toISOString();

  appendRow(SHEETS.BIDS, {
    BidID:          bidId,
    BidNumber:      bidNumber,
    BidName:        bidName,
    PAC:            pac,
    PurchaseMethod: purchaseMethod,
    Phase:          'technical',
    PhaseStatus:    'in_progress',
    CreatedBy:      createdBy || '',
    CreatedAt:      now,
    UpdatedAt:      now,
    Notes:          notes || '',
  });

  // Auto-create all activities for TECHNICAL phase
  createActivitiesForPhase(bidId, 'technical');

  return jsonOk({ success: true, bidId, message: `Bid ${bidNumber} created with technical phase activities.` });
}

function updateBid(body) {
  const { bidId, bidName, pac, purchaseMethod, notes } = body;
  if (!bidId) return jsonError('Bid ID required.');
  updateRowByKey(SHEETS.BIDS, 'BidID', bidId, {
    BidName: bidName, PAC: pac,
    PurchaseMethod: purchaseMethod, Notes: notes,
    UpdatedAt: new Date().toISOString(),
  });
  return jsonOk({ success: true, message: 'Bid updated.' });
}

function deleteBid(body) {
  const { bidId } = body;
  if (!bidId) return jsonError('Bid ID required.');
  deleteRowByKey(SHEETS.BIDS, 'BidID', bidId);
  // Remove related records
  const actSheet = getOrCreateSheet(SHEETS.ACTIVITIES);
  const actData  = actSheet.getDataRange().getValues();
  const actHeaders = actData[0].map(h => String(h).trim());
  const bidIdx     = actHeaders.indexOf('BidID');
  for (let i = actData.length - 1; i >= 1; i--) {
    if (String(actData[i][bidIdx]) === String(bidId)) actSheet.deleteRow(i + 1);
  }
  const clarSheet = getOrCreateSheet(SHEETS.CLARIFICATIONS);
  const clarData  = clarSheet.getDataRange().getValues();
  if (clarData.length > 1) {
    const clarHeaders = clarData[0].map(h => String(h).trim());
    const cBidIdx     = clarHeaders.indexOf('BidID');
    for (let i = clarData.length - 1; i >= 1; i--) {
      if (String(clarData[i][cBidIdx]) === String(bidId)) clarSheet.deleteRow(i + 1);
    }
  }
  return jsonOk({ success: true, message: 'Bid and all related records deleted.' });
}

function createActivitiesForPhase(bidId, phase) {
  const activities = ALL_ACTIVITIES[phase] || [];
  const now        = new Date();
  let   startDate  = new Date(now);

  activities.forEach((act, i) => {
    const sla     = CONFIG.SLA[act.key] || 5;
    const due     = addWorkingDays(new Date(startDate), sla);
    const actId   = `ACT_${phase.substring(0,3).toUpperCase()}_${Date.now().toString(36)}_${i}`;

    appendRow(SHEETS.ACTIVITIES, {
      ActivityID:    actId,
      BidID:         bidId,
      Phase:         phase,
      ActivityKey:   act.key,
      ActivityName:  act.name,
      Status:        i === 0 ? 'in_progress' : 'pending',
      StartDate:     i === 0 ? now.toISOString() : '',
      DueDate:       i === 0 ? due.toISOString() : '',
      CompletedDate: '',
      CompletedBy:   '',
      SLADays:       sla,
      IsDelayed:     'false',
      Notes:         '',
      UpdatedAt:     now.toISOString(),
    });
    // Sequential: next activity starts after this one's due
    if (i === 0) startDate = new Date(due);
  });
}

function addWorkingDays(date, days) {
  let d = new Date(date);
  let added = 0;
  while (added < days) {
    d.setDate(d.getDate() + 1);
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6) added++; // skip weekends
  }
  return d;
}

// ════════════════════════════════════════════════════════════════════
//  BIDDER MANAGEMENT
// ════════════════════════════════════════════════════════════════════

function addBidder(body) {
  const { bidId, companyName, technicalScore, financialScore, notes } = body;
  if (!bidId || !companyName) return jsonError('Bid ID and Company Name are required.');

  const techScore = parseFloat(technicalScore) || 0;
  const finScore  = parseFloat(financialScore)  || 0;
  const total     = Math.round((techScore * 0.6 + finScore * 0.4) * 100) / 100;

  const bidderId = 'BDR' + Date.now().toString(36).toUpperCase();
  appendRow(SHEETS.BIDDERS, {
    BidderID:       bidderId,
    BidID:          bidId,
    CompanyName:    companyName,
    TechnicalScore: techScore,
    FinancialScore: finScore,
    TotalScore:     total,
    Status:         'active',
    Notes:          notes || '',
    AddedAt:        new Date().toISOString(),
  });

  return jsonOk({ success: true, bidderId, totalScore: total, message: `${companyName} added.` });
}

function updateBidder(body) {
  const { bidderId, companyName, technicalScore, financialScore, status, notes } = body;
  if (!bidderId) return jsonError('Bidder ID required.');
  const techScore = parseFloat(technicalScore) || 0;
  const finScore  = parseFloat(financialScore)  || 0;
  const total     = Math.round((techScore * 0.6 + finScore * 0.4) * 100) / 100;
  updateRowByKey(SHEETS.BIDDERS, 'BidderID', bidderId, {
    CompanyName: companyName, TechnicalScore: techScore,
    FinancialScore: finScore, TotalScore: total,
    Status: status, Notes: notes,
  });
  return jsonOk({ success: true, totalScore: total });
}

function removeBidder(body) {
  const { bidderId } = body;
  if (!bidderId) return jsonError('Bidder ID required.');
  deleteRowByKey(SHEETS.BIDDERS, 'BidderID', bidderId);
  return jsonOk({ success: true });
}

// ════════════════════════════════════════════════════════════════════
//  ACTIVITY MANAGEMENT
// ════════════════════════════════════════════════════════════════════

function startActivity(body) {
  const { activityId, bidId } = body;
  if (!activityId) return jsonError('Activity ID required.');

  const now    = new Date();
  const acts   = sheetToObjects(SHEETS.ACTIVITIES);
  const act    = acts.find(a => a['ActivityID'] === activityId);
  if (!act) return jsonError('Activity not found.');

  const sla = parseInt(act['SLADays']) || 5;
  const due = addWorkingDays(now, sla);

  updateRowByKey(SHEETS.ACTIVITIES, 'ActivityID', activityId, {
    Status:    'in_progress',
    StartDate: now.toISOString(),
    DueDate:   due.toISOString(),
    UpdatedAt: now.toISOString(),
  });
  return jsonOk({ success: true, dueDate: due.toISOString() });
}

function updateActivityStatus(body) {
  const { activityId, status, notes } = body;
  if (!activityId || !status) return jsonError('Activity ID and status required.');

  const now  = new Date();
  const acts = sheetToObjects(SHEETS.ACTIVITIES);
  const act  = acts.find(a => a['ActivityID'] === activityId);
  if (!act) return jsonError('Activity not found.');

  const isDelayed = checkIfDelayed(act, status);
  updateRowByKey(SHEETS.ACTIVITIES, 'ActivityID', activityId, {
    Status:    status,
    IsDelayed: String(isDelayed),
    Notes:     notes || act['Notes'],
    UpdatedAt: now.toISOString(),
  });
  return jsonOk({ success: true, isDelayed });
}

function completeActivity(body) {
  const { activityId, completedBy, notes, bidId } = body;
  if (!activityId) return jsonError('Activity ID required.');

  const now  = new Date();
  const acts = sheetToObjects(SHEETS.ACTIVITIES);
  const act  = acts.find(a => a['ActivityID'] === activityId);
  if (!act) return jsonError('Activity not found.');

  const isDelayed = act['DueDate'] ? now > new Date(act['DueDate']) : false;

  updateRowByKey(SHEETS.ACTIVITIES, 'ActivityID', activityId, {
    Status:        'completed',
    CompletedDate: now.toISOString(),
    CompletedBy:   completedBy || '',
    IsDelayed:     String(isDelayed),
    Notes:         notes || act['Notes'],
    UpdatedAt:     now.toISOString(),
  });

  // Auto-start next pending activity in same phase
  const phaseActs = acts
    .filter(a => a['BidID'] === act['BidID'] && a['Phase'] === act['Phase'])
    .sort((a, b) => {
      const phaseKeys = ALL_ACTIVITIES[act['Phase']] ? ALL_ACTIVITIES[act['Phase']].map(x => x.key) : [];
      return phaseKeys.indexOf(a['ActivityKey']) - phaseKeys.indexOf(b['ActivityKey']);
    });

  const completedIdx = phaseActs.findIndex(a => a['ActivityID'] === activityId);
  if (completedIdx >= 0 && completedIdx < phaseActs.length - 1) {
    const next = phaseActs[completedIdx + 1];
    if (next['Status'] === 'pending') {
      const sla = parseInt(next['SLADays']) || 5;
      const due = addWorkingDays(now, sla);
      updateRowByKey(SHEETS.ACTIVITIES, 'ActivityID', next['ActivityID'], {
        Status:    'in_progress',
        StartDate: now.toISOString(),
        DueDate:   due.toISOString(),
        UpdatedAt: now.toISOString(),
      });
    }
  }

  // If all technical activities done, prompt financial phase
  const allDone = phaseActs.every(a => a['ActivityID'] === activityId || a['Status'] === 'completed');
  if (allDone && act['Phase'] === 'technical') {
    // Auto-create financial phase activities
    const bids = sheetToObjects(SHEETS.BIDS);
    const bid  = bids.find(b => b['BidID'] === act['BidID']);
    if (bid && bid['Phase'] === 'technical') {
      updateRowByKey(SHEETS.BIDS, 'BidID', act['BidID'], {
        Phase: 'financial', PhaseStatus: 'in_progress', UpdatedAt: now.toISOString()
      });
      createActivitiesForPhase(act['BidID'], 'financial');
      return jsonOk({ success: true, phaseTransition: true, message: 'Technical phase complete! Financial phase started.' });
    }
  }

  // If all financial activities done, mark bid complete
  if (allDone && act['Phase'] === 'financial') {
    updateRowByKey(SHEETS.BIDS, 'BidID', act['BidID'], {
      Phase: 'completed', PhaseStatus: 'completed', UpdatedAt: now.toISOString()
    });
    return jsonOk({ success: true, bidCompleted: true, message: 'All phases complete! Bid evaluation finished.' });
  }

  return jsonOk({ success: true, isDelayed, message: 'Activity marked complete.' });
}

function checkIfDelayed(act, newStatus) {
  if (newStatus === 'completed') {
    return act['DueDate'] ? new Date() > new Date(act['DueDate']) : false;
  }
  if (act['DueDate'] && (newStatus === 'in_progress' || newStatus === 'pending')) {
    return new Date() > new Date(act['DueDate']);
  }
  return false;
}

// ════════════════════════════════════════════════════════════════════
//  CLARIFICATIONS
// ════════════════════════════════════════════════════════════════════

function addClarification(body) {
  const { bidId, subject, requestedDate, notes, createdBy } = body;
  if (!bidId || !subject || !requestedDate) return jsonError('Bid ID, Subject and Requested Date are required.');

  const clarId = 'CLR' + Date.now().toString(36).toUpperCase();
  appendRow(SHEETS.CLARIFICATIONS, {
    ClarificationID: clarId,
    BidID:           bidId,
    Subject:         subject,
    RequestedDate:   requestedDate,
    ResponseDate:    '',
    Status:          'pending',
    Notes:           notes || '',
    CreatedBy:       createdBy || '',
    CreatedAt:       new Date().toISOString(),
  });
  return jsonOk({ success: true, clarId, message: 'Clarification request recorded.' });
}

function updateClarification(body) {
  const { clarId, responseDate, status, notes } = body;
  if (!clarId) return jsonError('Clarification ID required.');
  updateRowByKey(SHEETS.CLARIFICATIONS, 'ClarificationID', clarId, {
    ResponseDate: responseDate || '',
    Status:       status || 'pending',
    Notes:        notes || '',
  });
  return jsonOk({ success: true });
}

function getClarifications(e) {
  const bidId = e.parameter.bidId || '';
  let clars   = sheetToObjects(SHEETS.CLARIFICATIONS);
  if (bidId) clars = clars.filter(c => c['BidID'] === bidId);
  clars = clars.sort((a, b) => new Date(b['CreatedAt']) - new Date(a['CreatedAt']));
  return jsonOk({ success: true, clarifications: clars });
}

// ════════════════════════════════════════════════════════════════════
//  QUERIES
// ════════════════════════════════════════════════════════════════════

function getBids(e) {
  const pac    = e.parameter.pac    || '';
  const method = e.parameter.method || '';
  const phase  = e.parameter.phase  || '';

  let bids = sheetToObjects(SHEETS.BIDS)
    .sort((a, b) => new Date(b['CreatedAt']) - new Date(a['CreatedAt']));

  if (pac)    bids = bids.filter(b => b['PAC']            === pac);
  if (method) bids = bids.filter(b => b['PurchaseMethod'] === method);
  if (phase)  bids = bids.filter(b => b['Phase']          === phase);

  // Enrich with delay flag
  const activities = sheetToObjects(SHEETS.ACTIVITIES);
  refreshDelayFlags(activities);
  const clars      = sheetToObjects(SHEETS.CLARIFICATIONS);

  const enriched = bids.map(b => {
    const bActs  = activities.filter(a => a['BidID'] === b['BidID']);
    const bClars = clars.filter(c => c['BidID'] === b['BidID']);
    const delayed = bActs.some(a => isActivityDelayed(a));
    const done    = bActs.filter(a => a['Status'] === 'completed').length;
    const total   = bActs.length;
    return {
      ...b,
      activityProgress: total > 0 ? Math.round(done / total * 100) : 0,
      totalActivities:  total,
      doneActivities:   done,
      hasDelay:         delayed,
      clarificationCount: bClars.length,
    };
  });

  return jsonOk({ success: true, bids: enriched });
}

function getBidDetail(e) {
  const bidId = e.parameter.bidId || '';
  if (!bidId) return jsonError('Bid ID required.');

  const bids = sheetToObjects(SHEETS.BIDS);
  const bid  = bids.find(b => b['BidID'] === bidId);
  if (!bid) return jsonError('Bid not found.');

  const bidders = sheetToObjects(SHEETS.BIDDERS)
    .filter(b => b['BidID'] === bidId)
    .sort((a, b) => parseFloat(b['TotalScore']) - parseFloat(a['TotalScore']));

  const activities = sheetToObjects(SHEETS.ACTIVITIES)
    .filter(a => a['BidID'] === bidId);

  const clars = sheetToObjects(SHEETS.CLARIFICATIONS)
    .filter(c => c['BidID'] === bidId)
    .sort((a, b) => new Date(b['RequestedDate']) - new Date(a['RequestedDate']));

  // Annotate activities with current delay status
  const now = new Date();
  const annotated = activities.map(a => ({
    ...a,
    isDelayed:    isActivityDelayed(a),
    daysOverdue:  a['DueDate'] && a['Status'] !== 'completed'
      ? Math.max(0, Math.floor((now - new Date(a['DueDate'])) / 86400000))
      : 0,
    daysRemaining: a['DueDate'] && a['Status'] !== 'completed'
      ? Math.max(0, Math.floor((new Date(a['DueDate']) - now) / 86400000))
      : 0,
  }));

  return jsonOk({ success: true, bid, bidders, activities: annotated, clarifications: clars });
}

function getActivities(e) {
  const bidId = e.parameter.bidId || '';
  const phase = e.parameter.phase || '';
  let acts    = sheetToObjects(SHEETS.ACTIVITIES);
  if (bidId) acts = acts.filter(a => a['BidID'] === bidId);
  if (phase) acts = acts.filter(a => a['Phase'] === phase);
  return jsonOk({ success: true, activities: acts });
}

function getDashboard(e) {
  const bids       = sheetToObjects(SHEETS.BIDS);
  const activities = sheetToObjects(SHEETS.ACTIVITIES);
  const clars      = sheetToObjects(SHEETS.CLARIFICATIONS);
  const now        = new Date();

  const totalBids      = bids.length;
  const techPhase      = bids.filter(b => b['Phase'] === 'technical').length;
  const finPhase       = bids.filter(b => b['Phase'] === 'financial').length;
  const completed      = bids.filter(b => b['Phase'] === 'completed').length;

  const delayedActs    = activities.filter(a => isActivityDelayed(a));
  const delayedBidIds  = [...new Set(delayedActs.map(a => a['BidID']))];

  const pendingClars   = clars.filter(c => c['Status'] === 'pending').length;

  // PAC breakdown
  const pacBreakdown = ['PAC ONE','PAC TWO','PAC THREE','PAC FOUR'].map(pac => ({
    pac,
    count: bids.filter(b => b['PAC'] === pac).length,
  }));

  // Method breakdown
  const methodBreakdown = ['open bid','direct purchase','proforma'].map(m => ({
    method: m,
    count: bids.filter(b => (b['PurchaseMethod']||'').toLowerCase() === m).length,
  }));

  // Upcoming due activities (next 3 working days)
  const soon = activities
    .filter(a => a['Status'] === 'in_progress' && a['DueDate'])
    .map(a => {
      const daysLeft = Math.floor((new Date(a['DueDate']) - now) / 86400000);
      return { ...a, daysLeft };
    })
    .filter(a => a.daysLeft >= 0 && a.daysLeft <= 3)
    .sort((a, b) => a.daysLeft - b.daysLeft)
    .slice(0, 10);

  return jsonOk({
    success: true,
    summary: { totalBids, techPhase, finPhase, completed, delayedBids: delayedBidIds.length, pendingClars },
    pacBreakdown,
    methodBreakdown,
    soonDue: soon,
    recentBids: bids.sort((a,b)=>new Date(b['CreatedAt'])-new Date(a['CreatedAt'])).slice(0, 5),
  });
}

function getDelayReport(e) {
  const activities = sheetToObjects(SHEETS.ACTIVITIES);
  const bids       = sheetToObjects(SHEETS.BIDS);
  const now        = new Date();

  const delayed = activities
    .filter(a => isActivityDelayed(a))
    .map(a => {
      const bid      = bids.find(b => b['BidID'] === a['BidID']) || {};
      const overdue  = a['DueDate'] ? Math.floor((now - new Date(a['DueDate'])) / 86400000) : 0;
      return {
        bidId:        a['BidID'],
        bidNumber:    bid['BidNumber'] || '',
        bidName:      bid['BidName']   || '',
        pac:          bid['PAC']       || '',
        phase:        a['Phase'],
        activityName: a['ActivityName'],
        dueDate:      a['DueDate'],
        daysOverdue:  overdue,
        status:       a['Status'],
      };
    })
    .sort((a, b) => b.daysOverdue - a.daysOverdue);

  return jsonOk({ success: true, delayedActivities: delayed, total: delayed.length });
}

function refreshDelayFlags(activities) {
  const now = new Date();
  activities.forEach(a => {
    if (a['Status'] !== 'completed' && a['DueDate']) {
      const delayed = now > new Date(a['DueDate']);
      if (String(a['IsDelayed']) !== String(delayed)) {
        updateRowByKey(SHEETS.ACTIVITIES, 'ActivityID', a['ActivityID'], { IsDelayed: String(delayed) });
        a['IsDelayed'] = String(delayed);
      }
    }
  });
}

function isActivityDelayed(a) {
  if (a['Status'] === 'completed') return false;
  if (!a['DueDate']) return false;
  return new Date() > new Date(a['DueDate']);
}

// ════════════════════════════════════════════════════════════════════
//  RESPONSE HELPERS
// ════════════════════════════════════════════════════════════════════

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonError(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════
//  SETUP — run ONCE from Apps Script editor
// ════════════════════════════════════════════════════════════════════

function SETUP_RUN_ONCE() {
  Logger.log('=== Bid Evaluation Portal Setup ===');
  initializeSheets();
  Logger.log('All sheets created. Admin: admin@bid.gov / admin123');
  Logger.log('Web App URL: ' + ScriptApp.getService().getUrl());
}
