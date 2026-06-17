/**
 * Bid Evaluation Tracking Portal
 * Google Apps Script - Main Server-Side Code
 * 
 * Features:
 * - User Authentication (Admin/Evaluator roles)
 * - Bid Registration and Management
 * - Technical & Financial Evaluation Tracking
 * - SLA-based Delay Detection
 * - Clarification Management
 * - PDF Reports and Email Notifications
 */

// ============================================================================
// CONFIGURATION CONSTANTS
// ============================================================================

const CONFIG = {
  // SLA Days for each activity type
  SLA_DAYS: {
    // Technical Evaluation Activities
    'tech_doc_collection': 3,
    'tech_clarification': 5,
    'tech_evaluation': 7,
    'tech_report_submit': 3,
    // Financial Evaluation Activities
    'fin_doc_collection': 3,
    'fin_eng_estimation': 5,
    'fin_doc_authentication': 2,
    'fin_evaluation': 5
  },
  
  // Purchase Approving Committees
  PAC_OPTIONS: ['PAC ONE', 'PAC TWO', 'PAC THREE', 'PAC FOUR'],
  
  // Purchase Methods
  PURCHASE_METHODS: ['Open Bid', 'Direct Purchase', 'Proforma'],
  
  // Activity Statuses
  STATUS: {
    PENDING: 'Pending',
    IN_PROGRESS: 'In Progress',
    COMPLETED: 'Completed',
    DELAYED: 'Delayed'
  },
  
  // User Roles
  ROLES: {
    ADMIN: 'Admin',
    EVALUATOR: 'Evaluator'
  }
};

// ============================================================================
// SHEET INITIALIZATION
// ============================================================================

/**
 * Get or create the spreadsheet and sheets
 */
function getSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet found. Please run this script from a Google Sheet.');
  }
  return ss;
}

/**
 * Initialize all required sheets with headers
 */
function initializeSheets() {
  const ss = getSpreadsheet();
  
  // Define all sheets with their headers
  const sheetsConfig = [
    {
      name: 'Users',
      headers: ['ID', 'Username', 'Email', 'Password', 'Role', 'Full Name', 'Created Date', 'Status']
    },
    {
      name: 'Bids',
      headers: ['ID', 'Bid Number', 'Bid Name', 'PAC', 'Purchase Method', 'Status', 
                'Created Date', 'Created By', 'Technical Status', 'Financial Status', 'Overall Status']
    },
    {
      name: 'Bidders',
      headers: ['ID', 'Bid ID', 'Bidder Name', 'Technical Score', 'Technical Status', 
                'Financial Score', 'Rank', 'Remarks']
    },
    {
      name: 'Activities',
      headers: ['ID', 'Bid ID', 'Phase', 'Activity Type', 'Activity Name', 'Status',
                'Start Date', 'Target Date', 'Completion Date', 'SLA Days', 'Is Delayed', 
                'Delay Days', 'Assigned To', 'Remarks', 'Updated By', 'Updated Date']
    },
    {
      name: 'Clarifications',
      headers: ['ID', 'Bid ID', 'Phase', 'Subject', 'Request Date', 'Response Date', 
                'Status', 'Requested By', 'Response Details']
    },
    {
      name: 'ActivityLog',
      headers: ['ID', 'Bid ID', 'Activity', 'Action', 'Previous Value', 'New Value', 
                'Changed By', 'Changed Date']
    }
  ];
  
  sheetsConfig.forEach(config => {
    let sheet = ss.getSheetByName(config.name);
    if (!sheet) {
      sheet = ss.insertSheet(config.name);
      sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
      sheet.getRange(1, 1, 1, config.headers.length)
        .setBackground('#1a73e8')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setFrozen(true);
    }
  });
  
  // Create default admin user if not exists
  createDefaultAdmin();
  
  return { success: true, message: 'Sheets initialized successfully' };
}

/**
 * Create default admin user
 */
function createDefaultAdmin() {
  const sheet = getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    const adminId = generateId('USR');
    const now = new Date();
    sheet.appendRow([
      adminId,
      'admin',
      'admin@bidportal.com',
      Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, 'admin123').map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join(''),
      'Admin',
      'System Administrator',
      now,
      'Active'
    ]);
  }
}

/**
 * Get sheet by name
 */
function getSheetByName(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    initializeSheets();
    sheet = ss.getSheetByName(name);
  }
  return sheet;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Generate unique ID with prefix
 */
function generateId(prefix) {
  const timestamp = Date.now().toString(36);
  const random = Math.random().toString(36).substring(2, 8);
  return prefix + '_' + timestamp + random;
}

/**
 * Hash password
 */
function hashPassword(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, password);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * Format date to string
 */
function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Format date time to string
 */
function formatDateTime(date) {
  if (!date) return '';
  const d = new Date(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Calculate delay status
 */
function calculateDelayStatus(startDate, targetDate, completionDate, slaDays) {
  const today = new Date();
  const target = new Date(targetDate);
  
  if (completionDate) {
    const completion = new Date(completionDate);
    if (completion > target) {
      const delayDays = Math.ceil((completion - target) / (1000 * 60 * 60 * 24));
      return { isDelayed: true, delayDays: delayDays };
    }
    return { isDelayed: false, delayDays: 0 };
  }
  
  if (today > target) {
    const delayDays = Math.ceil((today - target) / (1000 * 60 * 60 * 24));
    return { isDelayed: true, delayDays: delayDays };
  }
  
  return { isDelayed: false, delayDays: 0 };
}

/**
 * Calculate target date based on SLA
 */
function calculateTargetDate(startDate, slaDays) {
  const start = new Date(startDate);
  const target = new Date(start);
  target.setDate(target.getDate() + slaDays);
  return target;
}

// ============================================================================
// USER AUTHENTICATION
// ============================================================================

/**
 * User login
 */
function login(username, password) {
  const sheet = getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const hashedPassword = hashPassword(password);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username && data[i][3] === hashedPassword && data[i][7] === 'Active') {
      return {
        success: true,
        user: {
          id: data[i][0],
          username: data[i][1],
          email: data[i][2],
          role: data[i][4],
          fullName: data[i][5]
        }
      };
    }
  }
  
  return { success: false, message: 'Invalid username or password' };
}

/**
 * Register new user (Admin only)
 */
function registerUser(userData, currentUser) {
  if (currentUser.role !== CONFIG.ROLES.ADMIN) {
    return { success: false, message: 'Only administrators can register new users' };
  }
  
  const sheet = getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  // Check if username exists
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userData.username) {
      return { success: false, message: 'Username already exists' };
    }
  }
  
  const userId = generateId('USR');
  const hashedPassword = hashPassword(userData.password);
  
  sheet.appendRow([
    userId,
    userData.username,
    userData.email,
    hashedPassword,
    userData.role,
    userData.fullName,
    new Date(),
    'Active'
  ]);
  
  return { success: true, message: 'User registered successfully', userId: userId };
}

/**
 * Get all users (Admin only)
 */
function getAllUsers(currentUser) {
  if (currentUser.role !== CONFIG.ROLES.ADMIN) {
    return { success: false, message: 'Access denied' };
  }
  
  const sheet = getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const users = [];
  
  for (let i = 1; i < data.length; i++) {
    users.push({
      id: data[i][0],
      username: data[i][1],
      email: data[i][2],
      role: data[i][4],
      fullName: data[i][5],
      createdDate: formatDate(data[i][6]),
      status: data[i][7]
    });
  }
  
  return { success: true, users: users };
}

/**
 * Update user status (Admin only)
 */
function updateUserStatus(userId, status, currentUser) {
  if (currentUser.role !== CONFIG.ROLES.ADMIN) {
    return { success: false, message: 'Access denied' };
  }
  
  const sheet = getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      sheet.getRange(i + 1, 8).setValue(status);
      return { success: true, message: 'User status updated' };
    }
  }
  
  return { success: false, message: 'User not found' };
}

// ============================================================================
// BID MANAGEMENT
// ============================================================================

/**
 * Create new bid
 */
function createBid(bidData, currentUser) {
  const sheet = getSheetByName('Bids');
  const data = sheet.getDataRange().getValues();
  
  // Check if bid number exists
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === bidData.bidNumber) {
      return { success: false, message: 'Bid number already exists' };
    }
  }
  
  const bidId = generateId('BID');
  const now = new Date();
  
  sheet.appendRow([
    bidId,
    bidData.bidNumber,
    bidData.bidName,
    bidData.pac,
    bidData.purchaseMethod,
    CONFIG.STATUS.PENDING,
    now,
    currentUser.id,
    CONFIG.STATUS.PENDING,
    CONFIG.STATUS.PENDING,
    CONFIG.STATUS.PENDING
  ]);
  
  // Create default activities for both phases
  createDefaultActivities(bidId, currentUser);
  
  // Log activity
  logActivity(bidId, 'Bid Creation', 'Created', '', bidData.bidNumber, currentUser.id);
  
  return { success: true, message: 'Bid created successfully', bidId: bidId };
}

/**
 * Create default activities for a new bid
 */
function createDefaultActivities(bidId, currentUser) {
  const sheet = getSheetByName('Activities');
  const now = new Date();
  
  const technicalActivities = [
    { type: 'tech_doc_collection', name: 'Technical Document Collection' },
    { type: 'tech_clarification', name: 'Clarification Request' },
    { type: 'tech_evaluation', name: 'Technical Evaluation Progress' },
    { type: 'tech_report_submit', name: 'Technical Evaluation Report Submittal' }
  ];
  
  const financialActivities = [
    { type: 'fin_doc_collection', name: 'Financial Document Collection' },
    { type: 'fin_eng_estimation', name: 'Engineering Estimation & Budget Approval' },
    { type: 'fin_doc_authentication', name: 'Document Authentication' },
    { type: 'fin_evaluation', name: 'Financial Evaluation Report' }
  ];
  
  const allActivities = [
    ...technicalActivities.map(a => ({ ...a, phase: 'Technical' })),
    ...financialActivities.map(a => ({ ...a, phase: 'Financial' }))
  ];
  
  allActivities.forEach(activity => {
    const slaDays = CONFIG.SLA_DAYS[activity.type];
    const targetDate = calculateTargetDate(now, slaDays);
    const activityId = generateId('ACT');
    
    sheet.appendRow([
      activityId,
      bidId,
      activity.phase,
      activity.type,
      activity.name,
      CONFIG.STATUS.PENDING,
      now,
      targetDate,
      '',
      slaDays,
      false,
      0,
      '',
      '',
      currentUser.id,
      now
    ]);
  });
}

/**
 * Get all bids
 */
function getAllBids(filters = {}) {
  const sheet = getSheetByName('Bids');
  const data = sheet.getDataRange().getValues();
  const bids = [];
  
  for (let i = 1; i < data.length; i++) {
    const bid = {
      id: data[i][0],
      bidNumber: data[i][1],
      bidName: data[i][2],
      pac: data[i][3],
      purchaseMethod: data[i][4],
      status: data[i][5],
      createdDate: formatDate(data[i][6]),
      createdBy: data[i][7],
      technicalStatus: data[i][8],
      financialStatus: data[i][9],
      overallStatus: data[i][10]
    };
    
    // Apply filters
    if (filters.status && bid.status !== filters.status) continue;
    if (filters.pac && bid.pac !== filters.pac) continue;
    if (filters.purchaseMethod && bid.purchaseMethod !== filters.purchaseMethod) continue;
    if (filters.search && !bid.bidNumber.toLowerCase().includes(filters.search.toLowerCase()) && 
        !bid.bidName.toLowerCase().includes(filters.search.toLowerCase())) continue;
    
    // Get bidders for this bid
    bid.bidders = getBiddersByBidId(bid.id);
    
    // Get delay info
    bid.delayInfo = getBidDelayInfo(bid.id);
    
    bids.push(bid);
  }
  
  return { success: true, bids: bids };
}

/**
 * Get bid by ID
 */
function getBidById(bidId) {
  const sheet = getSheetByName('Bids');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bidId) {
      const bid = {
        id: data[i][0],
        bidNumber: data[i][1],
        bidName: data[i][2],
        pac: data[i][3],
        purchaseMethod: data[i][4],
        status: data[i][5],
        createdDate: formatDate(data[i][6]),
        createdBy: data[i][7],
        technicalStatus: data[i][8],
        financialStatus: data[i][9],
        overallStatus: data[i][10]
      };
      
      bid.bidders = getBiddersByBidId(bid.id);
      bid.activities = getActivitiesByBidId(bid.id);
      bid.clarifications = getClarificationsByBidId(bid.id);
      
      return { success: true, bid: bid };
    }
  }
  
  return { success: false, message: 'Bid not found' };
}

/**
 * Update bid
 */
function updateBid(bidId, bidData, currentUser) {
  const sheet = getSheetByName('Bids');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bidId) {
      if (bidData.bidName) sheet.getRange(i + 1, 3).setValue(bidData.bidName);
      if (bidData.pac) sheet.getRange(i + 1, 4).setValue(bidData.pac);
      if (bidData.purchaseMethod) sheet.getRange(i + 1, 5).setValue(bidData.purchaseMethod);
      if (bidData.status) sheet.getRange(i + 1, 6).setValue(bidData.status);
      if (bidData.technicalStatus) sheet.getRange(i + 1, 9).setValue(bidData.technicalStatus);
      if (bidData.financialStatus) sheet.getRange(i + 1, 10).setValue(bidData.financialStatus);
      if (bidData.overallStatus) sheet.getRange(i + 1, 11).setValue(bidData.overallStatus);
      
      logActivity(bidId, 'Bid Update', 'Updated', '', JSON.stringify(bidData), currentUser.id);
      
      return { success: true, message: 'Bid updated successfully' };
    }
  }
  
  return { success: false, message: 'Bid not found' };
}

// ============================================================================
// BIDDER MANAGEMENT
// ============================================================================

/**
 * Add bidder to bid
 */
function addBidder(bidId, bidderData, currentUser) {
  const sheet = getSheetByName('Bidders');
  const bidderId = generateId('BDR');
  
  sheet.appendRow([
    bidderId,
    bidId,
    bidderData.bidderName,
    bidderData.technicalScore || 0,
    'Pending',
    0,
    0,
    bidderData.remarks || ''
  ]);
  
  logActivity(bidId, 'Bidder Added', 'Added', '', bidderData.bidderName, currentUser.id);
  
  return { success: true, message: 'Bidder added successfully', bidderId: bidderId };
}

/**
 * Get bidders by bid ID
 */
function getBiddersByBidId(bidId) {
  const sheet = getSheetByName('Bidders');
  const data = sheet.getDataRange().getValues();
  const bidders = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === bidId) {
      bidders.push({
        id: data[i][0],
        bidderName: data[i][2],
        technicalScore: data[i][3],
        technicalStatus: data[i][4],
        financialScore: data[i][5],
        rank: data[i][6],
        remarks: data[i][7]
      });
    }
  }
  
  return bidders;
}

/**
 * Update bidder score
 */
function updateBidderScore(bidderId, scoreData, currentUser) {
  const sheet = getSheetByName('Bidders');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bidderId) {
      const bidId = data[i][1];
      
      if (scoreData.technicalScore !== undefined) {
        sheet.getRange(i + 1, 4).setValue(scoreData.technicalScore);
      }
      if (scoreData.technicalStatus) {
        sheet.getRange(i + 1, 5).setValue(scoreData.technicalStatus);
      }
      if (scoreData.financialScore !== undefined) {
        sheet.getRange(i + 1, 6).setValue(scoreData.financialScore);
      }
      if (scoreData.remarks !== undefined) {
        sheet.getRange(i + 1, 8).setValue(scoreData.remarks);
      }
      
      // Recalculate ranks
      recalculateBidderRanks(bidId);
      
      logActivity(bidId, 'Bidder Score Update', 'Updated', '', JSON.stringify(scoreData), currentUser.id);
      
      return { success: true, message: 'Bidder score updated' };
    }
  }
  
  return { success: false, message: 'Bidder not found' };
}

/**
 * Recalculate bidder ranks based on technical score
 */
function recalculateBidderRanks(bidId) {
  const sheet = getSheetByName('Bidders');
  const data = sheet.getDataRange().getValues();
  const bidders = [];
  
  // Collect bidders for this bid
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === bidId) {
      bidders.push({
        rowIndex: i + 1,
        score: data[i][3]
      });
    }
  }
  
  // Sort by score descending and assign ranks
  bidders.sort((a, b) => b.score - a.score);
  bidders.forEach((bidder, index) => {
    sheet.getRange(bidder.rowIndex, 7).setValue(index + 1);
  });
}

/**
 * Delete bidder
 */
function deleteBidder(bidderId, currentUser) {
  const sheet = getSheetByName('Bidders');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bidderId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Bidder deleted' };
    }
  }
  
  return { success: false, message: 'Bidder not found' };
}

// ============================================================================
// ACTIVITY MANAGEMENT
// ============================================================================

/**
 * Get activities by bid ID
 */
function getActivitiesByBidId(bidId) {
  const sheet = getSheetByName('Activities');
  const data = sheet.getDataRange().getValues();
  const activities = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === bidId) {
      const activity = {
        id: data[i][0],
        bidId: data[i][1],
        phase: data[i][2],
        activityType: data[i][3],
        activityName: data[i][4],
        status: data[i][5],
        startDate: formatDate(data[i][6]),
        targetDate: formatDate(data[i][7]),
        completionDate: formatDate(data[i][8]),
        slaDays: data[i][9],
        isDelayed: data[i][10],
        delayDays: data[i][11],
        assignedTo: data[i][12],
        remarks: data[i][13],
        updatedBy: data[i][14],
        updatedDate: formatDateTime(data[i][15])
      };
      
      // Calculate current delay status
      if (activity.status !== CONFIG.STATUS.COMPLETED) {
        const delayInfo = calculateDelayStatus(
          new Date(data[i][6]),
          new Date(data[i][7]),
          null,
          data[i][9]
        );
        activity.isDelayed = delayInfo.isDelayed;
        activity.delayDays = delayInfo.delayDays;
      }
      
      activities.push(activity);
    }
  }
  
  return activities;
}

/**
 * Update activity status
 */
function updateActivityStatus(activityId, statusData, currentUser) {
  const sheet = getSheetByName('Activities');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === activityId) {
      const bidId = data[i][1];
      const previousStatus = data[i][5];
      
      sheet.getRange(i + 1, 6).setValue(statusData.status);
      sheet.getRange(i + 1, 14).setValue(statusData.remarks || '');
      sheet.getRange(i + 1, 15).setValue(currentUser.id);
      sheet.getRange(i + 1, 16).setValue(new Date());
      
      // If completed, set completion date
      if (statusData.status === CONFIG.STATUS.COMPLETED) {
        sheet.getRange(i + 1, 9).setValue(new Date());
        
        // Check if delayed
        const targetDate = new Date(data[i][7]);
        const delayInfo = calculateDelayStatus(new Date(data[i][6]), targetDate, new Date(), data[i][9]);
        sheet.getRange(i + 1, 11).setValue(delayInfo.isDelayed);
        sheet.getRange(i + 1, 12).setValue(delayInfo.delayDays);
      }
      
      // Update bid phase status
      updateBidPhaseStatus(bidId);
      
      logActivity(bidId, 'Activity Update', data[i][4], previousStatus, statusData.status, currentUser.id);
      
      // Send notification if delayed
      if (statusData.status === CONFIG.STATUS.DELAYED) {
        sendDelayNotification(bidId, data[i][4]);
      }
      
      return { success: true, message: 'Activity status updated' };
    }
  }
  
  return { success: false, message: 'Activity not found' };
}

/**
 * Update bid phase status based on activities
 */
function updateBidPhaseStatus(bidId) {
  const activities = getActivitiesByBidId(bidId);
  
  // Technical phase activities
  const techActivities = activities.filter(a => a.phase === 'Technical');
  const techCompleted = techActivities.filter(a => a.status === CONFIG.STATUS.COMPLETED).length;
  const techDelayed = techActivities.filter(a => a.isDelayed).length;
  
  // Financial phase activities
  const finActivities = activities.filter(a => a.phase === 'Financial');
  const finCompleted = finActivities.filter(a => a.status === CONFIG.STATUS.COMPLETED).length;
  const finDelayed = finActivities.filter(a => a.isDelayed).length;
  
  // Determine statuses
  let techStatus = CONFIG.STATUS.PENDING;
  if (techCompleted === techActivities.length) {
    techStatus = CONFIG.STATUS.COMPLETED;
  } else if (techCompleted > 0) {
    techStatus = CONFIG.STATUS.IN_PROGRESS;
  }
  if (techDelayed > 0 && techStatus !== CONFIG.STATUS.COMPLETED) {
    techStatus = CONFIG.STATUS.DELAYED;
  }
  
  let finStatus = CONFIG.STATUS.PENDING;
  if (finCompleted === finActivities.length) {
    finStatus = CONFIG.STATUS.COMPLETED;
  } else if (finCompleted > 0) {
    finStatus = CONFIG.STATUS.IN_PROGRESS;
  }
  if (finDelayed > 0 && finStatus !== CONFIG.STATUS.COMPLETED) {
    finStatus = CONFIG.STATUS.DELAYED;
  }
  
  // Update bid sheet
  const sheet = getSheetByName('Bids');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bidId) {
      sheet.getRange(i + 1, 9).setValue(techStatus);
      sheet.getRange(i + 1, 10).setValue(finStatus);
      
      // Overall status
      let overallStatus = CONFIG.STATUS.PENDING;
      if (techStatus === CONFIG.STATUS.COMPLETED && finStatus === CONFIG.STATUS.COMPLETED) {
        overallStatus = CONFIG.STATUS.COMPLETED;
      } else if (techStatus === CONFIG.STATUS.DELAYED || finStatus === CONFIG.STATUS.DELAYED) {
        overallStatus = CONFIG.STATUS.DELAYED;
      } else if (techStatus === CONFIG.STATUS.IN_PROGRESS || finStatus === CONFIG.STATUS.IN_PROGRESS) {
        overallStatus = CONFIG.STATUS.IN_PROGRESS;
      }
      sheet.getRange(i + 1, 11).setValue(overallStatus);
      sheet.getRange(i + 1, 6).setValue(overallStatus);
      
      break;
    }
  }
}

/**
 * Get bid delay info
 */
function getBidDelayInfo(bidId) {
  const activities = getActivitiesByBidId(bidId);
  const delayedActivities = activities.filter(a => a.isDelayed);
  
  return {
    totalDelayed: delayedActivities.length,
    delayedActivities: delayedActivities.map(a => ({
      name: a.activityName,
      delayDays: a.delayDays,
      phase: a.phase
    }))
  };
}

// ============================================================================
// CLARIFICATION MANAGEMENT
// ============================================================================

/**
 * Add clarification request
 */
function addClarification(bidId, clarificationData, currentUser) {
  const sheet = getSheetByName('Clarifications');
  const clarificationId = generateId('CLR');
  
  sheet.appendRow([
    clarificationId,
    bidId,
    clarificationData.phase,
    clarificationData.subject,
    new Date(clarificationData.requestDate),
    '',
    CONFIG.STATUS.PENDING,
    currentUser.id,
    ''
  ]);
  
  logActivity(bidId, 'Clarification Added', clarificationData.subject, '', '', currentUser.id);
  
  return { success: true, message: 'Clarification request added', clarificationId: clarificationId };
}

/**
 * Get clarifications by bid ID
 */
function getClarificationsByBidId(bidId) {
  const sheet = getSheetByName('Clarifications');
  const data = sheet.getDataRange().getValues();
  const clarifications = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === bidId) {
      clarifications.push({
        id: data[i][0],
        bidId: data[i][1],
        phase: data[i][2],
        subject: data[i][3],
        requestDate: formatDate(data[i][4]),
        responseDate: formatDate(data[i][5]),
        status: data[i][6],
        requestedBy: data[i][7],
        responseDetails: data[i][8]
      });
    }
  }
  
  return clarifications;
}

/**
 * Update clarification
 */
function updateClarification(clarificationId, clarificationData, currentUser) {
  const sheet = getSheetByName('Clarifications');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === clarificationId) {
      if (clarificationData.responseDate) {
        sheet.getRange(i + 1, 6).setValue(new Date(clarificationData.responseDate));
      }
      if (clarificationData.status) {
        sheet.getRange(i + 1, 7).setValue(clarificationData.status);
      }
      if (clarificationData.responseDetails) {
        sheet.getRange(i + 1, 9).setValue(clarificationData.responseDetails);
      }
      
      logActivity(data[i][1], 'Clarification Updated', data[i][3], '', JSON.stringify(clarificationData), currentUser.id);
      
      return { success: true, message: 'Clarification updated' };
    }
  }
  
  return { success: false, message: 'Clarification not found' };
}

// ============================================================================
// ACTIVITY LOG
// ============================================================================

/**
 * Log activity
 */
function logActivity(bidId, activity, action, previousValue, newValue, userId) {
  const sheet = getSheetByName('ActivityLog');
  const logId = generateId('LOG');
  
  sheet.appendRow([
    logId,
    bidId,
    activity,
    action,
    previousValue,
    newValue,
    userId,
    new Date()
  ]);
}

/**
 * Get activity log by bid ID
 */
function getActivityLogByBidId(bidId) {
  const sheet = getSheetByName('ActivityLog');
  const data = sheet.getDataRange().getValues();
  const logs = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === bidId) {
      logs.push({
        id: data[i][0],
        bidId: data[i][1],
        activity: data[i][2],
        action: data[i][3],
        previousValue: data[i][4],
        newValue: data[i][5],
        changedBy: data[i][6],
        changedDate: formatDateTime(data[i][7])
      });
    }
  }
  
  return logs;
}

// ============================================================================
// DASHBOARD STATISTICS
// ============================================================================

/**
 * Get dashboard statistics
 */
function getDashboardStats() {
  const bidsSheet = getSheetByName('Bids');
  const bidsData = bidsSheet.getDataRange().getValues();
  
  const activitiesSheet = getSheetByName('Activities');
  const activitiesData = activitiesSheet.getDataRange().getValues();
  
  let totalBids = 0;
  let pendingBids = 0;
  let inProgressBids = 0;
  let completedBids = 0;
  let delayedBids = 0;
  
  let technicalPhase = { pending: 0, inProgress: 0, completed: 0, delayed: 0 };
  let financialPhase = { pending: 0, inProgress: 0, completed: 0, delayed: 0 };
  
  // Count bids by status
  for (let i = 1; i < bidsData.length; i++) {
    totalBids++;
    const status = bidsData[i][5];
    
    switch (status) {
      case CONFIG.STATUS.PENDING: pendingBids++; break;
      case CONFIG.STATUS.IN_PROGRESS: inProgressBids++; break;
      case CONFIG.STATUS.COMPLETED: completedBids++; break;
      case CONFIG.STATUS.DELAYED: delayedBids++; break;
    }
    
    // Phase counts
    const techStatus = bidsData[i][8];
    const finStatus = bidsData[i][9];
    
    switch (techStatus) {
      case CONFIG.STATUS.PENDING: technicalPhase.pending++; break;
      case CONFIG.STATUS.IN_PROGRESS: technicalPhase.inProgress++; break;
      case CONFIG.STATUS.COMPLETED: technicalPhase.completed++; break;
      case CONFIG.STATUS.DELAYED: technicalPhase.delayed++; break;
    }
    
    switch (finStatus) {
      case CONFIG.STATUS.PENDING: financialPhase.pending++; break;
      case CONFIG.STATUS.IN_PROGRESS: financialPhase.inProgress++; break;
      case CONFIG.STATUS.COMPLETED: financialPhase.completed++; break;
      case CONFIG.STATUS.DELAYED: financialPhase.delayed++; break;
    }
  }
  
  // Count delayed activities
  let delayedActivities = 0;
  let upcomingDeadlines = [];
  const today = new Date();
  const threeDaysLater = new Date(today);
  threeDaysLater.setDate(threeDaysLater.getDate() + 3);
  
  for (let i = 1; i < activitiesData.length; i++) {
    if (activitiesData[i][10] === true) {
      delayedActivities++;
    }
    
    // Check upcoming deadlines
    const targetDate = new Date(activitiesData[i][7]);
    if (activitiesData[i][5] !== CONFIG.STATUS.COMPLETED && targetDate <= threeDaysLater && targetDate >= today) {
      upcomingDeadlines.push({
        bidId: activitiesData[i][1],
        activity: activitiesData[i][4],
        targetDate: formatDate(targetDate),
        phase: activitiesData[i][2]
      });
    }
  }
  
  // PAC distribution
  const pacDistribution = {};
  CONFIG.PAC_OPTIONS.forEach(pac => pacDistribution[pac] = 0);
  
  for (let i = 1; i < bidsData.length; i++) {
    const pac = bidsData[i][3];
    if (pacDistribution[pac] !== undefined) {
      pacDistribution[pac]++;
    }
  }
  
  // Purchase method distribution
  const purchaseMethodDistribution = {};
  CONFIG.PURCHASE_METHODS.forEach(pm => purchaseMethodDistribution[pm] = 0);
  
  for (let i = 1; i < bidsData.length; i++) {
    const pm = bidsData[i][4];
    if (purchaseMethodDistribution[pm] !== undefined) {
      purchaseMethodDistribution[pm]++;
    }
  }
  
  return {
    success: true,
    stats: {
      totalBids,
      pendingBids,
      inProgressBids,
      completedBids,
      delayedBids,
      delayedActivities,
      technicalPhase,
      financialPhase,
      pacDistribution,
      purchaseMethodDistribution,
      upcomingDeadlines
    }
  };
}

// ============================================================================
// REPORTING
// ============================================================================

/**
 * Generate PDF report
 */
function generatePDFReport(bidId) {
  const bidResult = getBidById(bidId);
  if (!bidResult.success) {
    return bidResult;
  }
  
  const bid = bidResult.bid;
  
  // Create a temporary sheet for the report
  const ss = getSpreadsheet();
  const reportSheet = ss.insertSheet('Report_Temp');
  
  // Build report content
  const reportTitle = `Bid Evaluation Report - ${bid.bidNumber}`;
  
  // Header
  reportSheet.getRange('A1:F1').merge();
  reportSheet.getRange('A1').setValue(reportTitle).setFontWeight('bold').setFontSize(16);
  
  // Bid Details
  let row = 3;
  reportSheet.getRange(`A${row}`).setValue('Bid Details').setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  row++;
  reportSheet.getRange(`A${row}`).setValue('Bid Number:');
  reportSheet.getRange(`B${row}`).setValue(bid.bidNumber);
  row++;
  reportSheet.getRange(`A${row}`).setValue('Bid Name:');
  reportSheet.getRange(`B${row}`).setValue(bid.bidName);
  row++;
  reportSheet.getRange(`A${row}`).setValue('PAC:');
  reportSheet.getRange(`B${row}`).setValue(bid.pac);
  row++;
  reportSheet.getRange(`A${row}`).setValue('Purchase Method:');
  reportSheet.getRange(`B${row}`).setValue(bid.purchaseMethod);
  row++;
  reportSheet.getRange(`A${row}`).setValue('Overall Status:');
  reportSheet.getRange(`B${row}`).setValue(bid.overallStatus);
  row++;
  reportSheet.getRange(`A${row}`).setValue('Created Date:');
  reportSheet.getRange(`B${row}`).setValue(bid.createdDate);
  
  // Technical Evaluation Section
  row += 2;
  reportSheet.getRange(`A${row}`).setValue('Technical Evaluation').setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
  row++;
  
  const techActivities = bid.activities.filter(a => a.phase === 'Technical');
  reportSheet.getRange(`A${row}:F${row}`).setValues([['Activity', 'Status', 'Target Date', 'Completion Date', 'Delayed', 'Delay Days']]);
  row++;
  
  techActivities.forEach(activity => {
    reportSheet.getRange(`A${row}:F${row}`).setValues([[
      activity.activityName,
      activity.status,
      activity.targetDate,
      activity.completionDate || '-',
      activity.isDelayed ? 'Yes' : 'No',
      activity.delayDays
    ]]);
    row++;
  });
  
  // Bidders Section
  row += 2;
  reportSheet.getRange(`A${row}`).setValue('Bidders Evaluation').setFontWeight('bold').setBackground('#ea4335').setFontColor('#ffffff');
  row++;
  
  if (bid.bidders.length > 0) {
    reportSheet.getRange(`A${row}:E${row}`).setValues([['Bidder Name', 'Technical Score', 'Status', 'Rank', 'Remarks']]);
    row++;
    
    bid.bidders.forEach(bidder => {
      reportSheet.getRange(`A${row}:E${row}`).setValues([[
        bidder.bidderName,
        bidder.technicalScore,
        bidder.technicalStatus,
        bidder.rank,
        bidder.remarks || '-'
      ]]);
      row++;
    });
  } else {
    reportSheet.getRange(`A${row}`).setValue('No bidders registered yet');
    row++;
  }
  
  // Financial Evaluation Section
  row += 2;
  reportSheet.getRange(`A${row}`).setValue('Financial Evaluation').setFontWeight('bold').setBackground('#fbbc04').setFontColor('#000000');
  row++;
  
  const finActivities = bid.activities.filter(a => a.phase === 'Financial');
  reportSheet.getRange(`A${row}:F${row}`).setValues([['Activity', 'Status', 'Target Date', 'Completion Date', 'Delayed', 'Delay Days']]);
  row++;
  
  finActivities.forEach(activity => {
    reportSheet.getRange(`A${row}:F${row}`).setValues([[
      activity.activityName,
      activity.status,
      activity.targetDate,
      activity.completionDate || '-',
      activity.isDelayed ? 'Yes' : 'No',
      activity.delayDays
    ]]);
    row++;
  });
  
  // Clarifications Section
  if (bid.clarifications.length > 0) {
    row += 2;
    reportSheet.getRange(`A${row}`).setValue('Clarifications').setFontWeight('bold').setBackground('#673ab7').setFontColor('#ffffff');
    row++;
    
    reportSheet.getRange(`A${row}:E${row}`).setValues([['Subject', 'Phase', 'Request Date', 'Response Date', 'Status']]);
    row++;
    
    bid.clarifications.forEach(clarification => {
      reportSheet.getRange(`A${row}:E${row}`).setValues([[
        clarification.subject,
        clarification.phase,
        clarification.requestDate,
        clarification.responseDate || '-',
        clarification.status
      ]]);
      row++;
    });
  }
  
  // Auto-fit columns
  reportSheet.autoResizeColumns(1, 6);
  
  // Export to PDF
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?exportFormat=pdf&format=pdf' +
    '&size=letter' +
    '&portrait=true' +
    '&fitw=true' +
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=true' +
    '&gridlines=false' +
    '&gid=' + reportSheet.getSheetId();
  
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  
  const blob = response.getBlob().setName(`Bid_Evaluation_Report_${bid.bidNumber}.pdf`);
  
  // Delete temporary sheet
  ss.deleteSheet(reportSheet);
  
  // Save to Drive
  const folder = getOrCreateReportsFolder();
  const file = folder.createFile(blob);
  
  return {
    success: true,
    message: 'Report generated successfully',
    fileId: file.getId(),
    fileName: file.getName(),
    fileUrl: file.getUrl()
  };
}

/**
 * Get or create reports folder
 */
function getOrCreateReportsFolder() {
  const folderName = 'Bid Evaluation Reports';
  const folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  }
  
  return DriveApp.createFolder(folderName);
}

// ============================================================================
// EMAIL NOTIFICATIONS
// ============================================================================

/**
 * Send delay notification
 */
function sendDelayNotification(bidId, activityName) {
  const bidResult = getBidById(bidId);
  if (!bidResult.success) return;
  
  const bid = bidResult.bid;
  const usersSheet = getSheetByName('Users');
  const usersData = usersSheet.getDataRange().getValues();
  
  // Get all admin emails
  const adminEmails = [];
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][4] === CONFIG.ROLES.ADMIN && usersData[i][7] === 'Active') {
      adminEmails.push(usersData[i][2]);
    }
  }
  
  if (adminEmails.length === 0) return;
  
  const subject = `⚠️ Delay Alert: ${bid.bidNumber} - ${activityName}`;
  const body = `
Dear Administrator,

This is an automated notification regarding a delayed activity in the Bid Evaluation Portal.

Bid Details:
- Bid Number: ${bid.bidNumber}
- Bid Name: ${bid.bidName}
- PAC: ${bid.pac}
- Purchase Method: ${bid.purchaseMethod}

Delayed Activity: ${activityName}

Please take necessary action to resolve this delay.

Best regards,
Bid Evaluation Portal System
  `;
  
  MailApp.sendEmail({
    to: adminEmails.join(','),
    subject: subject,
    body: body
  });
}

/**
 * Send daily summary report
 */
function sendDailySummary() {
  const stats = getDashboardStats().stats;
  const delayedBids = getAllBids({ status: CONFIG.STATUS.DELAYED }).bids;
  
  const usersSheet = getSheetByName('Users');
  const usersData = usersSheet.getDataRange().getValues();
  
  const adminEmails = [];
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][4] === CONFIG.ROLES.ADMIN && usersData[i][7] === 'Active') {
      adminEmails.push(usersData[i][2]);
    }
  }
  
  if (adminEmails.length === 0) return;
  
  const today = formatDate(new Date());
  const subject = `📊 Daily Summary Report - ${today}`;
  
  let delayedBidsHtml = '';
  if (delayedBids.length > 0) {
    delayedBidsHtml = delayedBids.map(bid => `
      <tr>
        <td>${bid.bidNumber}</td>
        <td>${bid.bidName}</td>
        <td>${bid.pac}</td>
        <td>${bid.delayInfo.totalDelayed} activities delayed</td>
      </tr>
    `).join('');
  }
  
  const htmlBody = `
    <html>
      <body>
        <h2>Bid Evaluation Portal - Daily Summary</h2>
        <p>Date: ${today}</p>
        
        <h3>Overview</h3>
        <table border="1" cellpadding="8" cellspacing="0">
          <tr><td><strong>Total Bids</strong></td><td>${stats.totalBids}</td></tr>
          <tr><td><strong>Pending</strong></td><td>${stats.pendingBids}</td></tr>
          <tr><td><strong>In Progress</strong></td><td>${stats.inProgressBids}</td></tr>
          <tr><td><strong>Completed</strong></td><td>${stats.completedBids}</td></tr>
          <tr style="background-color: #ffdddd;"><td><strong>Delayed</strong></td><td>${stats.delayedBids}</td></tr>
        </table>
        
        ${delayedBids.length > 0 ? `
        <h3>⚠️ Delayed Bids Requiring Attention</h3>
        <table border="1" cellpadding="8" cellspacing="0">
          <tr style="background-color: #1a73e8; color: white;">
            <th>Bid Number</th>
            <th>Bid Name</th>
            <th>PAC</th>
            <th>Delay Details</th>
          </tr>
          ${delayedBidsHtml}
        </table>
        ` : '<p>No delayed bids.</p>'}
        
        <p>Please log in to the portal for detailed information.</p>
        <p>Best regards,<br>Bid Evaluation Portal System</p>
      </body>
    </html>
  `;
  
  MailApp.sendEmail({
    to: adminEmails.join(','),
    subject: subject,
    htmlBody: htmlBody
  });
}

// ============================================================================
// WEB APP HANDLERS
// ============================================================================

/**
 * Serve the main web app
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('login')
    .setTitle('Bid Evaluation Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include HTML file (for templating)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get configuration data for frontend
 */
function getConfig() {
  return {
    pacOptions: CONFIG.PAC_OPTIONS,
    purchaseMethods: CONFIG.PURCHASE_METHODS,
    statusOptions: Object.values(CONFIG.STATUS),
    slaDays: CONFIG.SLA_DAYS
  };
}

// ============================================================================
// PAGE NAVIGATION HELPERS
// ============================================================================

/**
 * Get dashboard page HTML
 */
function getDashboardPage() {
  return HtmlService.createHtmlOutputFromFile('dashboard')
    .setTitle('Bid Evaluation Portal - Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .getContent();
}

/**
 * Get login page HTML
 */
function getLoginPage() {
  return HtmlService.createHtmlOutputFromFile('login')
    .setTitle('Bid Evaluation Portal - Login')
    .getContent();
}

/**
 * Get progress tracking page HTML
 */
function getProgressPage() {
  return HtmlService.createHtmlOutputFromFile('progress-tracking')
    .setTitle('Bid Evaluation Portal - Progress Tracking')
    .getContent();
}

/**
 * Get clarifications page HTML
 */
function getClarificationsPage() {
  return HtmlService.createHtmlOutputFromFile('clarifications')
    .setTitle('Bid Evaluation Portal - Clarifications')
    .getContent();
}

/**
 * Get bidder management page HTML
 */
function getBidderManagementPage() {
  return HtmlService.createHtmlOutputFromFile('bidder-management')
    .setTitle('Bid Evaluation Portal - Bidder Management')
    .getContent();
}
