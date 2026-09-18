/**
 * Serves the HTML file or handles email actions.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    const action = e.parameter.action;
    // Served by the second deployment, the one that runs as the user opening it:
    // a trigger belongs to whoever executes the code, so this is the only way an
    // admin installs their own without going near the spreadsheet.
    if (action === 'authorizeEmail') {
      return authorizeEmailServicePage_();
    }
    if (action === 'record') {
      // Record links are HMAC-signed by sendAssignmentEmails_; refuse altered or forged ones.
      if (!verifyAssignmentLink_(e.parameter)) {
        return assignmentMessagePage_('Link not valid',
          'This coverage link is invalid or has been changed. Ask your administrator to re-send your assignment.');
      }
      return handleAssignmentRecord_(e.parameter);
    }
    if (action === 'accept' || action === 'reject') {
      // Links sent before coverage became an assignment. They cannot be honoured
      // (the signature now covers different fields, and there is no decline path),
      // so say so plainly rather than failing as though they were forged.
      return assignmentMessagePage_('This link is no longer valid',
        'TST coverage is now assigned rather than requested, so this link no longer works. ' +
        'Ask your administrator to re-send your assignment.');
    }
  }

  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('TST Manager')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper to get the current user's context (Role, Building, etc.) securely.
 */
function getUserContext() {
  const userEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName('Staff Directory');
  
  if (!staffSheet) throw new Error("Sheet 'Staff Directory' not found.");
  
  const data = staffSheet.getDataRange().getValues();
  const headers = data.shift(); // Remove headers
  
  // Dynamic Header Lookup
  const emailIdx = headers.findIndex(h => h.toString().toLowerCase().includes('email'));
  const nameIdx = headers.findIndex(h => h.toString().toLowerCase().includes('name'));
  const roleIdx = headers.findIndex(h => h.toString().toLowerCase().includes('role'));
  const buildingIdx = headers.findIndex(h => h.toString().toLowerCase().includes('building'));

  const safeEmailIdx = emailIdx > -1 ? emailIdx : 1;
  const safeNameIdx = nameIdx > -1 ? nameIdx : 0;
  const safeRoleIdx = roleIdx > -1 ? roleIdx : 2;
  const safeBuildingIdx = buildingIdx > -1 ? buildingIdx : 8; // Fallback to Col I (Index 8)

  let currentUserRole = 'Guest';
  let currentUserName = '';
  let currentUserBuilding = DEFAULT_BUILDING;
  let assignedBuildings = [DEFAULT_BUILDING];
  
  const userRow = data.find(r => r[safeEmailIdx].toString().toLowerCase() === userEmail.toLowerCase());
  
  if (userRow) {
    currentUserName = userRow[safeNameIdx];
    currentUserRole = userRow[safeRoleIdx];

    if (safeBuildingIdx > -1 && userRow[safeBuildingIdx]) {
      const rawBuilding = userRow[safeBuildingIdx].toString();
      if (rawBuilding.includes(',')) {
        assignedBuildings = rawBuilding.split(',').map(b => b.trim());
        currentUserBuilding = assignedBuildings[0]; // Default to first
      } else {
        currentUserBuilding = rawBuilding;
        assignedBuildings = [rawBuilding];
      }
    }
  }

  const isSuperAdmin = currentUserRole === 'Super Admin';

  return {
    email: userEmail,
    name: currentUserName,
    role: currentUserRole,
    building: currentUserBuilding,
    buildings: assignedBuildings,
    isSuperAdmin: isSuperAdmin
  };
}

/**
 * Gets the current user's email, determines their role based on the Staff Directory,
 * and fetches necessary initial data.
 */
function getInitialData() {
  const ctx = getUserContext();
  const config = getConfig(); // Load from Sheet

  return {
    email: ctx.email,
    name: ctx.name,
    role: ctx.role,
    building: ctx.building,
    buildings: ctx.buildings, // Pass list to frontend
    isSuperAdmin: ctx.isSuperAdmin,
    config: config,
    defaultBuilding: DEFAULT_BUILDING,
    authorizeUrl: authorizeUrl_(),
    // Same rule as getStaffDirectoryData: admins get the building's balances, a
    // teacher gets their own row plus a name/email roster, and anyone who isn't in
    // the directory gets nothing (the client shows them Access Denied).
    staffData: directoryFor_(ctx, ctx.building)
  };
}

/**
 * "View as" — lets an admin load the app exactly as a teacher in their building
 * sees it. Returns the same shape as getInitialData, but for the target.
 *
 * Only Teachers can be viewed (admins have no teacher view, and viewing as an
 * admin would imply privileges the caller doesn't have). Non-Super-Admins must
 * share a building with the target, and the view is limited to the buildings
 * they share so submissions can't be tagged to a building the caller doesn't
 * administer. `building` is the caller's current building; it is used when
 * allowed, otherwise the target's first allowed building.
 */
function getViewAsData(targetEmail, building) {
  const ctx = getUserContext();
  assertAdmin_(ctx);

  const sheet = getStaffSheet_();
  const idx = getStaffIndices_(sheet);
  const all = sheet.getDataRange().getValues();
  const rowI = findStaffRowByEmail_(all, idx.email, targetEmail || '');
  if (rowI === -1) throw new Error('Staff member not found.');

  const row = all[rowI];
  const buildingCell = row[idx.building];
  assertCanManageRow_(ctx, buildingCell);

  if ((row[idx.role] || '').toString().trim() !== 'Teacher') {
    throw new Error('View as is only available for teachers.');
  }

  const targetBuildings = splitBuildings_(buildingCell);
  if (targetBuildings.length === 0) targetBuildings.push(DEFAULT_BUILDING);
  const allowed = ctx.isSuperAdmin
    ? targetBuildings
    : targetBuildings.filter(b => ctx.buildings.includes(b));
  const activeBuilding = allowed.includes(building) ? building : allowed[0];

  const email = row[idx.email].toString().trim();
  console.log(`[View As] ${ctx.email} is viewing as ${email} (${activeBuilding})`);

  return {
    email: email,
    name: row[idx.name],
    role: 'Teacher',
    building: activeBuilding,
    buildings: allowed,
    isSuperAdmin: false,
    config: getConfig(),
    defaultBuilding: DEFAULT_BUILDING,
    staffData: staffDirectoryData_(activeBuilding, email)
  };
}

/**
 * Loads configuration from 'App Config' sheet.
 * Initializes the sheet with default config if it doesn't exist.
 */
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('App Config');
  
  // Initialize if missing
  if (!sheet) {
    sheet = ss.insertSheet('App Config');
    sheet.appendRow(['Building', 'Config_JSON']);
    
    // Populate with defaults from config.js (BUILDING_CONFIG global)
    // Note: We assume BUILDING_CONFIG is available in the context (from config.js)
    Object.keys(BUILDING_CONFIG).forEach(code => {
      sheet.appendRow([code, JSON.stringify(BUILDING_CONFIG[code], null, 2)]);
    });
    
    return BUILDING_CONFIG;
  }
  
  const data = sheet.getDataRange().getValues();
  data.shift(); // Remove Header
  
  const config = {};
  data.forEach(row => {
    const code = row[0];
    const json = row[1];
    try {
      config[code] = JSON.parse(json);
    } catch (e) {
      console.error(`Error parsing config for ${code}:`, e);
      // Fallback to static if parse fails? Or empty object.
      // If we have a static default available, use it, otherwise empty.
      config[code] = (typeof BUILDING_CONFIG !== 'undefined' && BUILDING_CONFIG[code]) ? BUILDING_CONFIG[code] : {};
    }
  });
  
  // Ensure we at least have defaults if sheet was empty or corrupted
  if (Object.keys(config).length === 0 && typeof BUILDING_CONFIG !== 'undefined') {
    return BUILDING_CONFIG;
  }
  
  return config;
}

/**
 * Saves configuration for a specific building.
 */
function saveBuildingConfig(buildingCode, newConfigObj) {
  const ctx = getUserContext();
  assertAdmin_(ctx);

  // Same building rule as every read: a building admin may edit their own
  // building(s), a Super Admin any configured one. Without this, any admin could
  // rewrite another school's periods, coverage types or calendar.
  if (allowedBuildingFor_(ctx, buildingCode) !== buildingCode) {
    throw new Error('You can only edit settings for your own building(s).');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('App Config');

  if (!sheet) {
    // Should exist, but safety first
    getConfig();
    sheet = ss.getSheetByName('App Config');
  }

  const data = sheet.getDataRange().getValues();
  // Find row index (1-based)
  // Row 1 is header. data index 0 is header.

  let rowIndex = -1;
  let existing = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === buildingCode) {
      rowIndex = i + 1;
      try { existing = JSON.parse(data[i][1]) || {}; } catch (e) { existing = {}; }
      break;
    }
  }

  // The Carry Over cap is owned by Super Admins. A non-Super-Admin's save must
  // never change it, so preserve whatever is already stored (default 12)
  // regardless of the submitted payload — this guards against a crafted call,
  // not just the UI.
  if (!ctx.isSuperAdmin) {
    const storedCap = Number(existing.carryOverMax);
    newConfigObj.carryOverMax = (isFinite(storedCap) && storedCap >= 0) ? storedCap : 12;
  }

  const jsonString = JSON.stringify(newConfigObj, null, 2);

  if (rowIndex > -1) {
    // Update existing
    sheet.getRange(rowIndex, 2).setValue(jsonString);
  } else {
    // Create new
    sheet.appendRow([buildingCode, jsonString]);
  }

  return true;
}

/**
 * Pulls a start/end pair off the end of a label.
 *
 * Handles both shapes the app stores: an OMS-style period label with its times
 * baked in ("Period 8 - 12:37 - 1:08") and a time-range building's period, which
 * is literally the span ("08:30 - 09:15").
 *
 * A one-digit hour comes from a 12-hour label, so 1-6 can only mean the
 * afternoon of a school day; a two-digit hour is already 24-hour (what an
 * <input type="time"> produces) and is left alone. That is what keeps
 * "12:37 - 1:08" from becoming a 12-hour event.
 */
function parseTimeRange_(text) {
  const m = /(\d{1,2}):(\d{2})\s*(?:[-–—]|to)\s*(\d{1,2}):(\d{2})\s*$/.exec(
    (text == null ? '' : text).toString().trim());
  if (!m) return null;

  const pad = n => String(n).padStart(2, '0');
  const to24 = (hRaw, minutes) => {
    const h = Number(hRaw);
    const hour = (hRaw.length === 1 && h >= 1 && h <= 6) ? h + 12 : h;
    return pad(hour) + ':' + pad(Number(minutes));
  };
  return { start: to24(m[1], m[2]), end: to24(m[3], m[4]) };
}

/**
 * The start/end times for a period on a given date, or null if the building has
 * not been told what they are.
 *
 * Checked in order: a day-group override (OHS runs different times on MWF and
 * TTh), the building's default times for that period, then any times written
 * into the label itself. Returning null is meaningful — the calendar reports it
 * rather than inventing a time.
 */
function periodTimesFor_(building, periodLabel, dateStr) {
  const cfg = (getConfig() || {})[building] || {};
  const label = (periodLabel == null ? '' : periodLabel).toString().trim();

  const day = parseYmd_(dateStr);
  if (day && Array.isArray(cfg.dayGroups)) {
    const short = WEEKDAY_NAMES_[day.getDay()].slice(0, 3);
    const group = cfg.dayGroups.find(g => g && Array.isArray(g.days) && g.days.indexOf(short) > -1);
    const t = group && group.times && group.times[label];
    if (t && t.start && t.end) return { start: t.start, end: t.end, source: group.name || 'day schedule' };
  }

  const fallback = cfg.periodTimes && cfg.periodTimes[label];
  if (fallback && fallback.start && fallback.end) {
    return { start: fallback.start, end: fallback.end, source: 'default' };
  }

  const parsed = parseTimeRange_(label);
  if (parsed) return { start: parsed.start, end: parsed.end, source: 'label' };

  return null;
}

/** "13:52" -> "1:52 PM". Times are stored 24-hour; nobody should ever read them that way. */
function formatTime12_(hhmm) {
  const m = /^(\d{1,2}):(\d{2})$/.exec((hhmm == null ? '' : hhmm).toString().trim());
  if (!m) return (hhmm == null ? '' : hhmm).toString();
  let hour = Number(m[1]);
  const suffix = hour >= 12 ? 'PM' : 'AM';
  hour = hour % 12;
  if (hour === 0) hour = 12;
  return hour + ':' + m[2] + ' ' + suffix;
}

/**
 * How a period should read to a person: "Period 3 (10:24 AM - 11:04 AM)".
 *
 * The stored label is not enough on its own. OHS periods are bare ("Period 3"),
 * so without this a teacher is told to cover a period with no hint of when it is
 * — and OHS runs different times on Tue/Thu, so only the date can say. OMS labels
 * do carry times, but in a 24-hour-ish form nobody writes by hand.
 *
 * Time-range buildings store the span itself as the period, so those just get the
 * span back in 12-hour form rather than repeating it.
 */
function periodDisplay_(building, period, dateStr) {
  const raw = (period == null ? '' : period).toString().trim();
  const label = shortPeriodLabel_(raw);
  const times = periodTimesFor_(building, raw, dateStr);
  if (!times) return label || raw;

  const when = formatTime12_(times.start) + ' \u2013 ' + formatTime12_(times.end);
  // The label is nothing but a time range (OIS/SE), so do not say it twice.
  if (label === raw && /^\d{1,2}:\d{2}/.test(label)) return when;
  return label + ' (' + when + ')';
}

/**
 * Reads the per-building Carry Over cap — the maximum hours that roll into the
 * next year when finalizing. Defaults to 12 when unset or invalid so config
 * sheets created before this setting existed behave sensibly without migration.
 */
function carryOverMaxFor_(building) {
  const config = getConfig();
  const c = (config && config[building]) || {};
  const m = Number(c.carryOverMax);
  return (isFinite(m) && m >= 0) ? m : 12;
}

/**
 * Pending counts for the admin badges. Admin only, and counted for exactly the
 * building the queues themselves return, so the badge and the list always agree.
 */
function getDashboardCounts(buildingFilter) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const building = allowedBuildingFor_(ctx, buildingFilter) || ctx.building;

  return {
    earned: pendingEarnedFor_(building).length,
    used: pendingUsedFor_(building).length,
    // Coverage that has already happened with no TST recorded yet.
    assignments: assignmentsFor_(building).filter(a => a.outstanding).length
  };
}

// Batch actions authorize every row up front (authorizeRequestRows_), then apply.
function batchApproveEarned(indices) {
  if (!indices || !Array.isArray(indices)) return;
  const rows = authorizeRequestRows_(getUserContext(), 'earned', indices);
  // Sort descending just in case, though for updates it matters less than deletes
  rows.sort((a, b) => b - a);

  rows.forEach(idx => {
    approveEarnedRow_(idx, { send: false }); // No email for batch
  });
  return true;
}

/**
 * Batch Action: Deny multiple Earned requests.
 */
function batchDenyEarned(indices) {
  if (!indices || !Array.isArray(indices)) return;
  const rows = authorizeRequestRows_(getUserContext(), 'earned', indices);
  rows.sort((a, b) => b - a);

  rows.forEach(idx => {
    denyEarnedRow_(idx, { send: false }); // No email, no specific reason
  });
  return true;
}

/**
 * Batch Action: Approve multiple Used requests.
 */
function batchApproveUsed(indices) {
  if (!indices || !Array.isArray(indices)) return;
  const rows = authorizeRequestRows_(getUserContext(), 'used', indices);
  rows.sort((a, b) => b - a);

  rows.forEach(idx => {
    approveUsedRow_(idx);
  });
  return true;
}

/**
 * Batch Action: Delete multiple Used requests.
 * MUST process descending to preserve indices.
 */
function batchDeleteUsed(indices) {
  if (!indices || !Array.isArray(indices)) return;
  const rows = authorizeRequestRows_(getUserContext(), 'used', indices);
  // Critical: Sort descending
  rows.sort((a, b) => b - a);

  rows.forEach(idx => {
    deleteUsedRow_(idx);
  });
  return true;
}

/**
 * Batch update function for Staff Directory (Carry Over & Paid Out).
 */
function updateStaffBatch(updates, type) {
  // updates: { "email@domain.com": 10.5, ... }
  // type: 'carryOver' | 'paidOut'

  const ctx = getUserContext();
  assertAdmin_(ctx);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Staff Directory');
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; // Row 1 is header

  let targetColIdx = -1;

  if (type === 'carryOver') {
     targetColIdx = headers.findIndex(h => h.toString().toLowerCase().includes('carry'));
     if (targetColIdx === -1) targetColIdx = 5; // Default F (Index 5)
  } else if (type === 'paidOut') {
     targetColIdx = headers.findIndex(h => h.toString().toLowerCase().includes('paid'));
     if (targetColIdx === -1) targetColIdx = 6; // Default G (Index 6)
  }

  if (targetColIdx === -1) throw new Error("Target column not found");

  const emailIdx = headers.findIndex(h => h.toString().toLowerCase().includes('email'));
  const safeEmailIdx = emailIdx > -1 ? emailIdx : 1;
  const buildingIdx = headers.findIndex(h => h.toString().toLowerCase().includes('building'));
  const safeBuildingIdx = buildingIdx > -1 ? buildingIdx : 8;

  const label = type === 'paidOut' ? 'Paid Out' : 'Carry Over';

  // Normalize update keys to lowercase so matching is case-insensitive.
  const updatesLower = {};
  Object.keys(updates || {}).forEach(k => { updatesLower[k.toLowerCase()] = updates[k]; });

  // Carry Over / Paid Out are owned by the primary building. Validate every
  // targeted row up front so the whole batch fails cleanly if any row is not
  // editable by this admin (rather than partially applying).
  const writes = [];
  for (let i = 1; i < data.length; i++) {
    const email = data[i][safeEmailIdx].toString().toLowerCase();
    if (!updatesLower.hasOwnProperty(email)) continue;

    const buildingCell = data[i][safeBuildingIdx];
    assertCanManageRow_(ctx, buildingCell);
    assertPrimaryAdminFor_(ctx, buildingCell, label);

    writes.push({ row: i + 1, value: updatesLower[email] });
  }

  // +1 for 1-based row index, +1 for 1-based col index; targetColIdx is 0-based.
  writes.forEach(w => sheet.getRange(w.row, targetColIdx + 1).setValue(w.value));

  return true;
}

// ===== Staff Management (Add / Edit / Archive / Restore / Delete) =====

function getStaffSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Staff Directory');
  if (!sheet) throw new Error("Sheet 'Staff Directory' not found.");
  return sheet;
}

function getStaffIndices_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const find = (kw, fb) => {
    const i = headers.findIndex(h => h.toString().toLowerCase().includes(kw));
    return i > -1 ? i : fb;
  };
  return {
    name: find('name', 0),
    email: find('email', 1),
    role: find('role', 2),
    carry: find('carry', 5),
    paid: find('paid', 6),
    building: find('building', 8),
    archived: ensureColumn_(sheet, 'archiv', 'Archived'),             // creates J if missing
    lastFinalized: ensureColumn_(sheet, 'last final', 'Last Finalized'), // creates K if missing
    pendingFinalize: ensureColumn_(sheet, 'pending', 'Pending Finalize') // creates L if missing
  };
}

// Parses the Archived (J) cell into a list of building codes the staff member is
// archived from. Handles the legacy boolean form (TRUE = archived everywhere).
function parseArchivedList_(cellValue, userBuildings) {
  const raw = (cellValue || '').toString().trim();
  const low = raw.toLowerCase();
  if (low === 'true') return (userBuildings || []).slice();
  if (low === 'false' || raw === '') return [];
  return raw.split(',').map(b => b.trim()).filter(Boolean);
}

function splitBuildings_(buildingCell) {
  return (buildingCell || '').toString().split(',').map(b => b.trim()).filter(Boolean);
}

function isAdmin_(ctx) {
  return ctx.role === 'Admin' || ctx.role === 'Super Admin';
}

function assertAdmin_(ctx) {
  if (!isAdmin_(ctx)) {
    throw new Error('Unauthorized: admin access required.');
  }
}

/**
 * Resolves which building a read may be scoped to. No request (or an unknown code)
 * means the caller's own (primary) building; otherwise anyone may pick one of their
 * own assigned buildings — a multi-building admin who switched school, or a teacher
 * assigned to two buildings — and Super Admins may pick any configured building.
 * Returns null when the requested building isn't allowed, so callers decide between
 * falling back to ctx.building and refusing.
 */
function allowedBuildingFor_(ctx, requested) {
  if (!requested || requested === ctx.building) return ctx.building;
  // Buildings can be added through Settings without a config.js edit, so an
  // assignment in the directory is authority enough; BUILDING_CONFIG only bounds
  // Super Admins, who aren't limited to their own assignment.
  if (ctx.buildings.includes(requested)) return requested;
  if (ctx.isSuperAdmin && BUILDING_CONFIG.hasOwnProperty(requested)) return requested;
  return null;
}

// Non-super admins may only touch staff who share at least one of their buildings.
function assertCanManageRow_(ctx, buildingCell) {
  if (ctx.isSuperAdmin) return;
  const staffBuildings = (buildingCell || '').toString().split(',').map(b => b.trim()).filter(Boolean);
  if (!staffBuildings.some(b => ctx.buildings.includes(b))) {
    throw new Error('You can only manage staff in your own building(s).');
  }
}

// The PRIMARY building is the first one listed in a staff member's Building cell.
// Carry Over, Paid Out and finalize are "owned" by the primary building's admin.
// Returns true when the caller may edit those owned columns for this person:
// Super Admins always may; otherwise the caller must be assigned to that person's
// primary building. (Earned submit/approve stays open to any building admin.)
function isPrimaryAdminFor_(ctx, buildingCell) {
  if (ctx.isSuperAdmin) return true;
  const primary = splitBuildings_(buildingCell)[0];
  if (!primary) return false;
  return ctx.buildings.includes(primary);
}

function assertPrimaryAdminFor_(ctx, buildingCell, what) {
  if (!isPrimaryAdminFor_(ctx, buildingCell)) {
    const primary = splitBuildings_(buildingCell)[0] || '(unknown)';
    throw new Error('Only the primary building (' + primary + ') administrator can edit ' +
      (what || 'this field') + ' for this staff member.');
  }
}

// Non-super admins may only assign buildings within their own assignment.
function assertCanAssignBuildings_(ctx, buildingStr) {
  if (ctx.isSuperAdmin) return;
  const requested = (buildingStr || '').toString().split(',').map(b => b.trim()).filter(Boolean);
  if (!requested.every(b => ctx.buildings.includes(b))) {
    throw new Error('You can only assign staff to your own building(s).');
  }
}

function findStaffRowByEmail_(allValues, emailIdx, email) {
  const target = email.toString().trim().toLowerCase();
  for (let i = 1; i < allValues.length; i++) {
    if (allValues[i][emailIdx].toString().trim().toLowerCase() === target) {
      return i; // 0-based array index (row number = i + 1)
    }
  }
  return -1;
}

// ===== Request authorization =====
// Every public function here can be called by any signed-in user through
// google.script.run, so each endpoint authorizes on the server. Internal callers
// use the private (_) variants to avoid repeating the checks.

// Request rows belong to the building they were filed under (blank = OMS, matching
// the approval queues). Columns are 1-based.
const REQUEST_SHEETS_ = {
  earned: { name: 'TST Approvals (New)', buildingCol: 14 }, // N
  used: { name: 'TST Usage (New)', buildingCol: 8 }          // H
};

/**
 * Admin gate for request-row actions (approve/deny/delete/edit/revert). Checks every
 * row up front — an integer row index within the sheet (row 1 is the header) whose
 * building is one of the caller's (Super Admins: any) — so a batch either passes
 * entirely or changes nothing. Returns the de-duplicated row numbers.
 */
function authorizeRequestRows_(ctx, type, rowIndices) {
  assertAdmin_(ctx);
  const cfg = REQUEST_SHEETS_[type];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.name);
  if (!sheet) throw new Error("Sheet '" + cfg.name + "' not found.");

  const lastRow = sheet.getLastRow();
  const rows = [];
  (rowIndices || []).forEach(r => {
    const row = Number(r);
    if (!Number.isInteger(row) || row < 2 || row > lastRow) {
      throw new Error('Invalid row ' + r + '. The request may have already been removed; refresh and try again.');
    }
    if (!rows.includes(row)) rows.push(row);
  });

  if (rows.length > 0 && !ctx.isSuperAdmin) {
    const buildings = sheet.getRange(1, cfg.buildingCol, lastRow, 1).getValues();
    rows.forEach(row => {
      const rowBuilding = (buildings[row - 1][0] || 'OMS').toString().trim();
      if (!ctx.buildings.includes(rowBuilding)) {
        throw new Error('You can only manage requests for your own building(s).');
      }
    });
  }
  return rows;
}

// Admin acting for staff: Super Admins for anyone; other admins only for staff who
// share one of their buildings.
function assertCanManageStaffEmails_(ctx, emails) {
  assertAdmin_(ctx);
  if (ctx.isSuperAdmin) return;
  const sheet = getStaffSheet_();
  const idx = getStaffIndices_(sheet);
  const all = sheet.getDataRange().getValues();
  (emails || []).forEach(email => {
    const i = findStaffRowByEmail_(all, idx.email, (email || '').toString());
    if (i === -1) throw new Error('Staff member not found: ' + email);
    assertCanManageRow_(ctx, all[i][idx.building]);
  });
}

// A person may act on their own requests/history; otherwise the caller must be an
// admin who manages them (View As, admin-created requests, staff detail).
function assertSelfOrManagerOf_(targetEmail) {
  const target = (targetEmail || '').toString().trim();
  if (!target) throw new Error('A staff email is required.');
  const sessionEmail = (Session.getActiveUser().getEmail() || '').toString().trim();
  if (sessionEmail && sessionEmail.toLowerCase() === target.toLowerCase()) return;
  assertCanManageStaffEmails_(getUserContext(), [target]);
}

/**
 * Adds a new staff member. Writes individual cells (never appendRow) so the
 * Running Total ARRAYFORMULA in column H is not clobbered.
 */
function addStaffMember(data) {
  const ctx = getUserContext();
  assertAdmin_(ctx);

  const name = (data.name || '').toString().trim();
  const email = (data.email || '').toString().trim();
  if (!name || !email) throw new Error('Name and email are required.');

  const role = (data.role || 'Teacher').toString().trim() || 'Teacher';
  // Building admins always add to their own building; only Super Admins choose.
  const building = ctx.isSuperAdmin
    ? ((data.building || '').toString().trim() || ctx.building)
    : ctx.building;
  const carryOver = Number(data.carryOver) || 0;
  const paidOut = Number(data.paidOut) || 0;

  assertCanAssignBuildings_(ctx, building);

  const sheet = getStaffSheet_();
  const idx = getStaffIndices_(sheet);
  const all = sheet.getDataRange().getValues();

  const existing = findStaffRowByEmail_(all, idx.email, email);
  if (existing > -1) {
    // Super Admins manage existing people via Edit. A building admin re-adding an
    // existing person (e.g. staff who work across buildings) simply gets their
    // building merged into that person's assignment instead of a duplicate row.
    if (ctx.isSuperAdmin) {
      throw new Error('A staff member with that email already exists. Edit them instead.');
    }
    const currentBuildings = (all[existing][idx.building] || '').toString()
      .split(',').map(b => b.trim()).filter(Boolean);
    if (!currentBuildings.includes(building)) {
      currentBuildings.push(building);
      sheet.getRange(existing + 1, idx.building + 1).setValue(currentBuildings.join(', '));
    }
    return true;
  }

  // Insert immediately after the last row that has an email, so the new row
  // stays inside the H column's B2:B ARRAYFORMULA range.
  let lastDataRow = 1; // header row
  for (let i = 1; i < all.length; i++) {
    if (all[i][idx.email].toString().trim() !== '') lastDataRow = i + 1;
  }
  const newRow = lastDataRow + 1;

  sheet.getRange(newRow, idx.name + 1).setValue(name);
  sheet.getRange(newRow, idx.email + 1).setValue(email);
  sheet.getRange(newRow, idx.role + 1).setValue(role);
  sheet.getRange(newRow, idx.carry + 1).setValue(carryOver);
  sheet.getRange(newRow, idx.paid + 1).setValue(paidOut);
  sheet.getRange(newRow, idx.building + 1).setValue(building);
  sheet.getRange(newRow, idx.archived + 1).setValue(false);

  return true;
}

/**
 * Updates an existing staff member matched by (unchanged) email. Email is never
 * rewritten because it keys all transaction history.
 */
function updateStaffMember(email, data) {
  const ctx = getUserContext();
  assertAdmin_(ctx);

  const sheet = getStaffSheet_();
  const idx = getStaffIndices_(sheet);
  const all = sheet.getDataRange().getValues();

  const i = findStaffRowByEmail_(all, idx.email, email);
  if (i === -1) throw new Error('Staff member not found.');

  assertCanManageRow_(ctx, all[i][idx.building]);

  // Only Super Admins reassign buildings. Building admins leave the assignment
  // untouched (their form has no building selector), so we keep the existing
  // value and skip the assignment check — otherwise editing a cross-building
  // person would fail on the buildings they don't own.
  const provided = (data.building || '').toString().trim();
  const newBuilding = provided || all[i][idx.building].toString();
  if (provided) assertCanAssignBuildings_(ctx, newBuilding);

  const row = i + 1;
  if (data.name !== undefined) sheet.getRange(row, idx.name + 1).setValue(data.name.toString().trim());
  if (data.role !== undefined) sheet.getRange(row, idx.role + 1).setValue(data.role);
  sheet.getRange(row, idx.building + 1).setValue(newBuilding);

  // Carry Over / Paid Out are owned by the primary building. Only the primary
  // admin (or a Super Admin) may change them; for non-primary admins we leave
  // the existing values untouched (the edit form makes those fields read-only).
  // Ownership is judged by the person's CURRENT primary building.
  if ((data.carryOver !== undefined || data.paidOut !== undefined) &&
      isPrimaryAdminFor_(ctx, all[i][idx.building])) {
    if (data.carryOver !== undefined) sheet.getRange(row, idx.carry + 1).setValue(Number(data.carryOver) || 0);
    if (data.paidOut !== undefined) sheet.getRange(row, idx.paid + 1).setValue(Number(data.paidOut) || 0);
  }

  return true;
}

/**
 * Archives/restores a staff member for a SINGLE building. Archiving only removes
 * them from that building's directory; they remain active in any other building
 * they're assigned to. Defaults to the caller's current building.
 */
function setStaffArchived_(email, building, archived) {
  const ctx = getUserContext();
  assertAdmin_(ctx);

  const sheet = getStaffSheet_();
  const idx = getStaffIndices_(sheet);
  const all = sheet.getDataRange().getValues();

  const i = findStaffRowByEmail_(all, idx.email, email);
  if (i === -1) throw new Error('Staff member not found.');

  const bldg = (building || ctx.building).toString().trim();
  assertCanManageRow_(ctx, all[i][idx.building]);
  if (!ctx.isSuperAdmin && !ctx.buildings.includes(bldg)) {
    throw new Error('You can only archive within your own building(s).');
  }

  const userBuildings = splitBuildings_(all[i][idx.building]);
  let list = parseArchivedList_(all[i][idx.archived], userBuildings);

  if (archived) {
    if (!list.includes(bldg)) list.push(bldg);
  } else {
    list = list.filter(b => b !== bldg);
  }

  sheet.getRange(i + 1, idx.archived + 1).setValue(list.join(', '));
  return true;
}

function archiveStaffMember(email, building) { return setStaffArchived_(email, building, true); }
function restoreStaffMember(email, building) { return setStaffArchived_(email, building, false); }

/**
 * Permanently removes a staff member's spreadsheet row. Super Admin only, and
 * only for staff archived from EVERY building they're assigned to (fully gone).
 */
function deleteStaffMemberPermanent(email) {
  const ctx = getUserContext();
  if (!ctx.isSuperAdmin) throw new Error('Unauthorized: Super Admin access required.');

  const sheet = getStaffSheet_();
  const idx = getStaffIndices_(sheet);
  const all = sheet.getDataRange().getValues();

  const i = findStaffRowByEmail_(all, idx.email, email);
  if (i === -1) throw new Error('Staff member not found.');

  const userBuildings = splitBuildings_(all[i][idx.building]);
  const list = parseArchivedList_(all[i][idx.archived], userBuildings);
  const fullyArchived = userBuildings.length === 0 || userBuildings.every(b => list.includes(b));
  if (!fullyArchived) {
    throw new Error('Only fully-archived staff (archived from all their buildings) can be permanently deleted.');
  }

  sheet.deleteRow(i + 1);
  return true;
}

/**
 * Returns fully-archived staff (archived from every building) for the Settings
 * permanent-delete list. Super Admin only.
 */
function getArchivedStaff() {
  const ctx = getUserContext();
  if (!ctx.isSuperAdmin) throw new Error('Unauthorized: Super Admin access required.');
  return staffDirectoryData_(null, null, true).filter(s => s.archived);
}

// ===== Year-End Finalize / Archived Years =====

/**
 * Finalizes a school year for one building, PRIMARY-AWARE.
 *
 * Under the ownership model, a person's Carry Over / Paid Out / Used (and the
 * year-end roll) belong to their PRIMARY building (the first listed). Earned can
 * be contributed by any building, and totals are COMBINED across buildings.
 *
 *  - Staff whose PRIMARY is this building ("primary-here"):
 *      * Snapshot their COMBINED end-of-year totals into a new tagged sheet.
 *      * Roll their COMBINED remaining balance into Carry Over and zero Paid Out
 *        (only ONCE per person per school year via Last Finalized).
 *      * Archive ALL their approved transactions across EVERY building so the
 *        combined Earned/Used recompute to 0.
 *      * Clear any pending-finalize flag.
 *  - Staff assigned here but whose PRIMARY is elsewhere ("shared"):
 *      * Are NOT rolled/reset/archived (their data is untouched).
 *      * Get a pending-finalize flag so the UI shows that their primary building
 *        still needs to finalize them. The flag clears when the primary does.
 */
function finalizeSchoolYear(yearName, building, force) {
  const ctx = getUserContext();
  assertAdmin_(ctx);

  const name = (yearName || '').toString().trim();
  if (!name) throw new Error('A name for the archive is required.');

  const bldg = ctx.isSuperAdmin ? (building || ctx.building).toString().trim() : ctx.building;
  if (!ctx.isSuperAdmin && !ctx.buildings.includes(bldg)) {
    throw new Error('You can only finalize your own building.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(name)) {
    throw new Error('A sheet named "' + name + '" already exists. Choose a different name.');
  }

  // Per-building Carry Over cap: any balance that would roll above it is
  // forfeited. round2 keeps float noise from spuriously tripping the cap.
  const cap = carryOverMaxFor_(bldg);
  const round2 = function (n) { return Math.round(n * 100) / 100; };

  // 1. COMBINED balances (across all buildings), computed BEFORE any transactions
  //    are moved. Only the primary building rolls/archives them.
  const balances = calculateDynamicBalances_(null);

  // 2. Build the plan from staff assigned to this building.
  const sheet = getStaffSheet_();
  const idx = getStaffIndices_(sheet); // ensures Pending Finalize column exists
  const all = sheet.getDataRange().getValues();

  const snapshotRows = [];
  const rolls = [];            // primary-here: balance rolls (already capped)
  const overCap = [];          // primary-here: rolls that would exceed the cap
  const clearPending = [];     // primary-here: rows whose pending flag to clear
  const primaryEmails = [];    // primary-here emails: archive ALL their transactions
  const pendingFlags = [];     // shared-elsewhere: rows to flag pending

  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    const email = (r[idx.email] || '').toString().trim();
    if (!email) continue;

    const userBuildings = splitBuildings_(r[idx.building]);
    if (!userBuildings.includes(bldg)) continue;

    const archivedList = parseArchivedList_(r[idx.archived], userBuildings);
    if (archivedList.includes(bldg)) continue; // already archived from this building

    const isPrimaryHere = userBuildings[0] === bldg;
    const pendingCur = (r[idx.pendingFinalize] || '').toString().trim();
    const stat = balances[email.toLowerCase()] || { earned: 0, used: 0 };

    if (!isPrimaryHere) {
      // Shared staff whose primary is elsewhere: leave their data untouched.
      // Flag them as pending only if they still have combined activity awaiting
      // their primary building's finalize. Two guards keep the flag from getting
      // stuck: (a) order-independence — if the primary already finalized them their
      // combined totals are 0, so we don't re-flag; (b) if they're archived from
      // their primary building, that primary will never finalize them, so a chip
      // would never clear — don't set it.
      const hasActivity = (Number(stat.earned) || 0) + (Number(stat.used) || 0) > 0;
      const archivedFromPrimary = archivedList.includes(userBuildings[0]);
      if (hasActivity && !archivedFromPrimary && pendingCur !== name) {
        pendingFlags.push({ row: i + 1, year: name });
      }
      continue;
    }

    const e = Number(stat.earned) || 0;  // combined
    const u = Number(stat.used) || 0;    // combined
    const priorCarry = Number(r[idx.carry]) || 0;
    const paidOut = Number(r[idx.paid]) || 0;
    const balance = priorCarry + e - u - paidOut;

    const lastFin = (r[idx.lastFinalized] || '').toString().trim();
    const firstThisYear = lastFin !== name;
    const rem = e - u;
    const projectedCarry = priorCarry + rem - (firstThisYear ? paidOut : 0);
    const cappedCarry = Math.min(projectedCarry, cap);
    const forfeited = projectedCarry > cap ? round2(projectedCarry - cap) : 0;

    snapshotRows.push([r[idx.name], email, r[idx.building], priorCarry, e, u, paidOut, balance, cappedCarry, forfeited]);
    rolls.push({ row: i + 1, newCarry: cappedCarry, firstThisYear: firstThisYear, forfeited: forfeited });

    if (forfeited > 0) {
      overCap.push({ name: r[idx.name], email: email, projected: round2(projectedCarry), cap: cap, over: forfeited });
    }

    primaryEmails.push(email.toLowerCase());
    if (pendingCur !== '') clearPending.push(i + 1);
  }

  if (snapshotRows.length === 0 && pendingFlags.length === 0) {
    throw new Error('No active staff found for ' + bldg + ' to finalize.');
  }

  // 2b. Block-with-override gate. If any primary-here staff would roll above the
  //     cap and the caller hasn't explicitly opted to forfeit, make NO changes
  //     and return the offending rows so the UI can list them. The check runs
  //     before any sheet write so a blocked finalize is a true no-op.
  if (overCap.length > 0 && !force) {
    return { blocked: true, building: bldg, cap: cap, overCap: overCap };
  }

  // 3. Write the snapshot sheet (only if there are primary-here staff to record).
  if (snapshotRows.length > 0) {
    const snap = ss.insertSheet(name);
    snap.appendRow(['Name', 'Email', 'Building(s)', 'Carry Over (start)', 'Earned', 'Used', 'Paid Out', 'Balance', 'New Carry Over', 'Forfeited']);
    snap.getRange(2, 1, snapshotRows.length, 10).setValues(snapshotRows);
    snap.setFrozenRows(1);
    snap.addDeveloperMetadata('tstArchiveBuilding', bldg);
    snap.addDeveloperMetadata('tstArchiveYear', name);
  }

  // 4. Roll primary-here balances (capped) and clear their pending flags.
  rolls.forEach(rr => {
    sheet.getRange(rr.row, idx.carry + 1).setValue(rr.newCarry);
    if (rr.firstThisYear) {
      sheet.getRange(rr.row, idx.paid + 1).setValue(0);
      sheet.getRange(rr.row, idx.lastFinalized + 1).setValue(name);
    }
  });
  clearPending.forEach(row => sheet.getRange(row, idx.pendingFinalize + 1).setValue(''));

  // 5. Flag shared staff so their primary building knows to finalize them.
  pendingFlags.forEach(pf => sheet.getRange(pf.row, idx.pendingFinalize + 1).setValue(pf.year));

  // 6. Archive ALL approved transactions for primary-here staff across EVERY
  //    building so their combined Earned/Used reset to 0.
  if (primaryEmails.length > 0) {
    archiveTransactionsByEmails_(primaryEmails, name);
  }

  // 6b. Assignment rows are per-building, so they archive with the building that
  //     created them rather than with the primary-here staff set.
  archiveAssignmentsForBuilding_(bldg, name);

  const forfeitedCount = rolls.filter(rr => rr.forfeited > 0).length;
  return { building: bldg, name: name, count: snapshotRows.length, pending: pendingFlags.length, cap: cap, forfeitedCount: forfeitedCount };
}

// Archives every approved transaction (across all buildings) for the given set
// of staff emails — used by the primary building's finalize so combined totals
// reset to 0. Distinct from archiveBuildingTransactions_ (which is by building);
// both are kept so existing callers/signatures are unaffected.
function archiveTransactionsByEmails_(emailsLower, yearName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Use a real Set so membership can't be confused by inherited Object.prototype
  // keys (e.g. a stray 'constructor'/'__proto__' value in the email column).
  const emailSet = new Set((emailsLower || []).map(e => e.toString().trim().toLowerCase()));
  // Earned: approved col I(8). Used: status col E(4).
  archiveRowsByEmails_(ss, 'TST Approvals (New)', 'TST Approvals Archive', 8, emailSet, yearName);
  archiveRowsByEmails_(ss, 'TST Usage (New)', 'TST Usage Archive', 4, emailSet, yearName);
}

function archiveRowsByEmails_(ss, srcName, archName, approvedIdx, emailSet, yearName) {
  const src = ss.getSheetByName(srcName);
  if (!src) return;
  const data = src.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0];

  let arch = ss.getSheetByName(archName);
  if (!arch) {
    arch = ss.insertSheet(archName);
    arch.appendRow(headers.concat(['School Year']));
    arch.setFrozenRows(1);
  }

  const toArchive = [];
  const rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue; // no email
    const isApproved = r[approvedIdx] === true || r[approvedIdx] === 'TRUE';
    const emailLower = r[0].toString().trim().toLowerCase();
    if (isApproved && emailSet.has(emailLower)) {
      toArchive.push(r.concat([yearName]));
      rowsToDelete.push(i + 1); // 1-based
    }
  }

  if (toArchive.length) {
    arch.getRange(arch.getLastRow() + 1, 1, toArchive.length, toArchive[0].length).setValues(toArchive);
    rowsToDelete.sort((a, b) => b - a).forEach(rn => src.deleteRow(rn));
  }
}

// Kept for backward-compatibility (no longer used by finalizeSchoolYear, which now
// archives by email set). Do not remove — preserves the existing signature.
function archiveBuildingTransactions_(building, yearName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Earned: building col N(13), approved col I(8). Used: building col H(7), status col E(4).
  archiveRowsByBuilding_(ss, 'TST Approvals (New)', 'TST Approvals Archive', 13, 8, building, yearName);
  archiveRowsByBuilding_(ss, 'TST Usage (New)', 'TST Usage Archive', 7, 4, building, yearName);
}

function archiveRowsByBuilding_(ss, srcName, archName, buildingIdx, approvedIdx, building, yearName) {
  const src = ss.getSheetByName(srcName);
  if (!src) return;
  const data = src.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0];

  let arch = ss.getSheetByName(archName);
  if (!arch) {
    arch = ss.insertSheet(archName);
    arch.appendRow(headers.concat(['School Year']));
    arch.setFrozenRows(1);
  }

  const toArchive = [];
  const rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue; // no email
    const isApproved = r[approvedIdx] === true || r[approvedIdx] === 'TRUE';
    const rowBuilding = (r[buildingIdx] || 'OMS').toString();
    if (isApproved && rowBuilding === building) {
      toArchive.push(r.concat([yearName]));
      rowsToDelete.push(i + 1); // 1-based
    }
  }

  if (toArchive.length) {
    arch.getRange(arch.getLastRow() + 1, 1, toArchive.length, toArchive[0].length).setValues(toArchive);
    rowsToDelete.sort((a, b) => b - a).forEach(rn => src.deleteRow(rn));
  }
}

/**
 * Lists year-end archive sheets, scoped to the caller's building (Super Admins
 * may pass a building or get all). Identified by developer metadata, not name.
 */
function listArchivedYears(building) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wanted = ctx.isSuperAdmin ? (building || null) : ctx.building;

  const result = [];
  ss.getSheets().forEach(sh => {
    const md = sh.getDeveloperMetadata().filter(m => m.getKey() === 'tstArchiveBuilding');
    if (md.length === 0) return;
    const sheetBuilding = md[0].getValue();
    if (wanted && sheetBuilding !== wanted) return;
    if (!ctx.isSuperAdmin && sheetBuilding !== ctx.building) return;
    result.push({ name: sh.getName(), building: sheetBuilding });
  });
  return result;
}

function getArchivedYearData(sheetName) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Archive not found.');

  const md = sh.getDeveloperMetadata().filter(m => m.getKey() === 'tstArchiveBuilding');
  if (md.length === 0) throw new Error('That sheet is not a TST year-end archive.');
  const sheetBuilding = md[0].getValue();
  if (!ctx.isSuperAdmin && sheetBuilding !== ctx.building) {
    throw new Error('You can only view archives for your own building.');
  }

  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return { name: sheetName, building: sheetBuilding, headers: headers, rows: values };
}

/* ============================================================================
 * SNAPSHOTS — non-destructive point-in-time captures of the directory balance
 * table. Unlike finalizeSchoolYear, these change nothing else (no roll-over, no
 * transaction archiving). Each snapshot is a hidden sheet tagged with
 * tstSnapshot* developer metadata (distinct from the tstArchive* keys used by
 * year-end archives, so the two listings never overlap). Title/description/date
 * live in metadata so they can be edited later without renaming the sheet.
 * ==========================================================================*/

/**
 * Builds the directory balance table for a building as of right now, WITHOUT
 * changing anything. Mirrors the snapshot rows produced by finalizeSchoolYear
 * (minus the roll-over/reset). Returns { headers, rows }.
 */
function buildDirectorySnapshotRows_(bldg) {
  // Mirror the directory exactly as displayed: COMBINED totals (Earned/Used summed
  // across all of a person's buildings, per the ownership model), non-archived
  // staff assigned to this building. Reusing staffDirectoryData_ keeps the
  // snapshot in lock-step with the on-screen table.
  const staff = staffDirectoryData_(bldg);
  const rows = staff.map(s => [
    s.name,
    s.email,
    s.building,
    Number(s.carryOver) || 0,
    Number(s.earned) || 0,
    Number(s.used) || 0,
    Number(s.paidOut) || 0,
    Number(s.total) || 0
  ]);
  return {
    headers: ['Name', 'Email', 'Building(s)', 'Carry Over', 'Earned', 'Used', 'Paid Out', 'Balance'],
    rows: rows
  };
}

/** Non-Super admins may only touch snapshots for buildings they're assigned to. */
function assertCanAccessSnapshot_(ctx, building) {
  if (ctx.isSuperAdmin) return;
  if (!ctx.buildings.includes(building)) {
    throw new Error('You can only manage snapshots for your own building(s).');
  }
}

/** Reads the tstSnapshot* developer metadata off a sheet into a plain object. */
function getSnapshotMeta_(sheet) {
  const meta = {};
  sheet.getDeveloperMetadata().forEach(m => {
    const k = m.getKey();
    if (k && k.indexOf('tstSnapshot') === 0) meta[k] = m.getValue();
  });
  return {
    id: meta.tstSnapshotId || '',
    building: meta.tstSnapshotBuilding || '',
    title: meta.tstSnapshotTitle || '',
    description: meta.tstSnapshotDescription || '',
    date: meta.tstSnapshotDate || '',
    created: meta.tstSnapshotCreated || ''
  };
}

/** Locates a snapshot sheet by its stable tstSnapshotId, or returns null. */
function findSnapshotSheet_(ss, id) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const md = sheets[i].getDeveloperMetadata().filter(m => m.getKey() === 'tstSnapshotId');
    if (md.length && md[0].getValue() === id) return sheets[i];
  }
  return null;
}

/**
 * Creates a non-destructive snapshot of a building's current directory totals.
 * data: building (optional; defaults to caller's), title, description, date.
 */
function createSnapshot(building, title, description, date) {
  const ctx = getUserContext();
  assertAdmin_(ctx);

  const bldg = (building || ctx.building).toString().trim();
  assertCanAccessSnapshot_(ctx, bldg);

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const snapDate = (date || '').toString().trim() || today;
  const snapTitle = (title || '').toString().trim() || ('Snapshot — ' + snapDate);
  const snapDesc = (description || '').toString();

  const id = 'snap_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
  const created = new Date().toISOString();

  const built = buildDirectorySnapshotRows_(bldg);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snap = ss.insertSheet('__snap_' + id);
  snap.appendRow(built.headers);
  if (built.rows.length) {
    snap.getRange(2, 1, built.rows.length, built.headers.length).setValues(built.rows);
  }
  snap.setFrozenRows(1);
  snap.hideSheet();

  snap.addDeveloperMetadata('tstSnapshotId', id);
  snap.addDeveloperMetadata('tstSnapshotBuilding', bldg);
  snap.addDeveloperMetadata('tstSnapshotTitle', snapTitle);
  if (snapDesc) snap.addDeveloperMetadata('tstSnapshotDescription', snapDesc);
  snap.addDeveloperMetadata('tstSnapshotDate', snapDate);
  snap.addDeveloperMetadata('tstSnapshotCreated', created);

  // Capture the building's counted (approved) transactions so the snapshot can
  // later be restored to identical totals. Predicate matches calculateDynamicBalances_:
  //   Approvals: building col N(13), approved col I(8).
  //   Usage:     building col H(7),  status   col E(4).
  captureBuildingTransactions_(ss, 'TST Approvals (New)', 13, 8, bldg, '__snap_' + id + '_appr');
  captureBuildingTransactions_(ss, 'TST Usage (New)', 7, 4, bldg, '__snap_' + id + '_use');

  return {
    id: id, building: bldg, title: snapTitle, description: snapDesc,
    date: snapDate, created: created, count: built.rows.length
  };
}

/**
 * Copies a source transaction sheet's approved rows for one building into a new
 * hidden companion sheet (header + matching rows). Used to make snapshots
 * restorable. Predicate mirrors archiveRowsByBuilding_ / calculateDynamicBalances_.
 */
function captureBuildingTransactions_(ss, srcName, buildingIdx, approvedIdx, bldg, destName) {
  const dest = ss.insertSheet(destName);
  dest.hideSheet();
  const src = ss.getSheetByName(srcName);
  if (!src) return dest;
  const data = src.getDataRange().getValues();
  if (data.length < 1) return dest;

  const out = [data[0]]; // header
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    const isApproved = r[approvedIdx] === true || r[approvedIdx] === 'TRUE';
    const rowBuilding = (r[buildingIdx] || 'OMS').toString();
    if (isApproved && rowBuilding === bldg) out.push(r);
  }
  dest.getRange(1, 1, out.length, out[0].length).setValues(out);
  return dest;
}

/**
 * Lists snapshots for every building the caller manages (Super Admins see all,
 * or pass a building to filter). Identified by developer metadata, not name.
 */
function listSnapshots(building) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wanted = ctx.isSuperAdmin ? (building || null) : null;

  const result = [];
  ss.getSheets().forEach(sh => {
    const md = sh.getDeveloperMetadata().filter(m => m.getKey() === 'tstSnapshotId');
    if (md.length === 0) return;
    const meta = getSnapshotMeta_(sh);
    if (ctx.isSuperAdmin) {
      if (wanted && meta.building !== wanted) return;
    } else if (!ctx.buildings.includes(meta.building)) {
      return;
    }
    result.push(meta);
  });

  result.sort((a, b) => (b.created || '').localeCompare(a.created || ''));
  return result;
}

/** Returns a snapshot's metadata plus its captured table (headers + rows). */
function getSnapshotData(id) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = findSnapshotSheet_(ss, id);
  if (!sh) throw new Error('Snapshot not found.');
  const meta = getSnapshotMeta_(sh);
  assertCanAccessSnapshot_(ctx, meta.building);

  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return {
    id: meta.id, building: meta.building, title: meta.title,
    description: meta.description, date: meta.date, created: meta.created,
    headers: headers, rows: values
  };
}

/** Edits a snapshot's title / description / date (only the keys present in data). */
function updateSnapshot(id, data) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = findSnapshotSheet_(ss, id);
  if (!sh) throw new Error('Snapshot not found.');
  const meta = getSnapshotMeta_(sh);
  assertCanAccessSnapshot_(ctx, meta.building);

  data = data || {};
  const setMeta = (key, value) => {
    const md = sh.getDeveloperMetadata().filter(m => m.getKey() === key);
    if (md.length) md[0].setValue(value);
    else if (value !== '') sh.addDeveloperMetadata(key, value); // GAS may reject empty metadata values
  };

  if (data.title !== undefined) setMeta('tstSnapshotTitle', (data.title || '').toString().trim() || meta.title);
  if (data.description !== undefined) setMeta('tstSnapshotDescription', (data.description || '').toString());
  if (data.date !== undefined) setMeta('tstSnapshotDate', (data.date || '').toString().trim() || meta.date);

  return getSnapshotMeta_(sh);
}

/** Permanently deletes a snapshot (the main sheet and its companion sheets). */
function deleteSnapshot(id) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = findSnapshotSheet_(ss, id);
  if (!sh) throw new Error('Snapshot not found.');
  const meta = getSnapshotMeta_(sh);
  assertCanAccessSnapshot_(ctx, meta.building);

  ['_appr', '_use'].forEach(suffix => {
    const c = ss.getSheetByName('__snap_' + id + suffix);
    if (c) ss.deleteSheet(c);
  });
  ss.deleteSheet(sh);
  return { ok: true, id: id };
}

/** Appends a captured companion sheet's rows (minus header) into a live sheet. */
function restoreCapturedTransactions_(ss, snapName, liveName) {
  const snap = ss.getSheetByName(snapName);
  const live = ss.getSheetByName(liveName);
  if (!snap || !live) return;
  const data = snap.getDataRange().getValues();
  if (data.length < 2) return; // header only — nothing to restore
  const rows = data.slice(1).filter(r => r[0]);
  if (!rows.length) return;
  live.getRange(live.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Reverts a building to the state captured by a snapshot — building-scoped and
 * with no effect on other buildings. After restoring, every column of each row
 * in the directory matches the snapshot:
 *  1. The building's current approved transactions are archived (recoverable)
 *     and removed, then the snapshot's captured transactions are re-inserted, so
 *     Earned/Used recompute to the snapshot values.
 *  2. Carry Over / Paid Out are reset to the snapshot values for staff whose
 *     PRIMARY (first-listed) building is this one — the primary building owns
 *     those numbers. Staff who only earn hours here (primary elsewhere) keep
 *     their current Carry Over / Paid Out, so other buildings are never disturbed.
 * Staff not present in the snapshot are left unchanged; no staff rows are deleted.
 */
function restoreSnapshot(id) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const main = findSnapshotSheet_(ss, id);
  if (!main) throw new Error('Snapshot not found.');
  const meta = getSnapshotMeta_(main);
  const bldg = meta.building;
  assertCanAccessSnapshot_(ctx, bldg);

  const apprSnap = ss.getSheetByName('__snap_' + id + '_appr');
  const useSnap = ss.getSheetByName('__snap_' + id + '_use');
  if (!apprSnap || !useSnap) {
    throw new Error('This snapshot was created before Restore was supported and cannot be restored.');
  }

  // 1. Archive + clear the building's CURRENT approved transactions, then put the
  //    snapshot's captured transactions back.
  const label = 'Restore backup ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + ' (' + bldg + ')';
  archiveBuildingTransactions_(bldg, label);
  restoreCapturedTransactions_(ss, '__snap_' + id + '_appr', 'TST Approvals (New)');
  restoreCapturedTransactions_(ss, '__snap_' + id + '_use', 'TST Usage (New)');

  // 2. Restore Carry Over / Paid Out only for staff whose PRIMARY (first-listed)
  //    building is this one — the primary building owns those numbers.
  const snapVals = main.getDataRange().getValues();
  snapVals.shift(); // drop header; cols: 0 Name,1 Email,2 Building(s),3 Carry,4 Earned,5 Used,6 Paid,7 Balance

  const sheet = getStaffSheet_();
  const idx = getStaffIndices_(sheet);
  const all = sheet.getDataRange().getValues();

  const emailToRow = {};
  const emailToBuildings = {};
  for (let i = 1; i < all.length; i++) {
    const em = (all[i][idx.email] || '').toString().trim().toLowerCase();
    if (!em) continue;
    emailToRow[em] = i + 1;
    emailToBuildings[em] = splitBuildings_(all[i][idx.building]);
  }

  let restored = 0;
  let nonPrimarySkipped = 0;
  snapVals.forEach(r => {
    const email = (r[1] || '').toString().trim().toLowerCase();
    if (!email) return;
    const rowNum = emailToRow[email];
    if (!rowNum) return; // person no longer in the directory
    // The primary building (first listed) owns Carry Over / Paid Out. Leave them
    // alone for people who only earn hours here but are primary elsewhere.
    const primary = (emailToBuildings[email] || [])[0] || '';
    if (primary !== bldg) { nonPrimarySkipped++; return; }
    sheet.getRange(rowNum, idx.carry + 1).setValue(Number(r[3]) || 0);
    sheet.getRange(rowNum, idx.paid + 1).setValue(Number(r[6]) || 0);
    restored++;
  });

  return {
    building: bldg, title: meta.title, count: snapVals.length,
    restored: restored, nonPrimarySkipped: nonPrimarySkipped
  };
}

/**
 * Ensures the Staff Directory has an "Archived" column, creating it after the
 * last used column if missing. Returns its 0-based index.
 */
function ensureColumn_(sheet, keyword, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.findIndex(h => h.toString().toLowerCase().includes(keyword));
  if (idx > -1) return idx;

  const newColNum = sheet.getLastColumn() + 1;
  sheet.getRange(1, newColNum).setValue(headerName);
  return newColNum - 1; // 0-based
}

/**
 * Directory read for the client. Balances, Carry Over and Paid Out are admin data,
 * so what comes back depends on who is asking:
 *   - Admin / Super Admin: the full table for one building they manage.
 *   - Teacher: their own complete row, plus name/email only for the colleagues in
 *     that building (all the Submit forms need to name who was covered).
 * The building is always resolved through allowedBuildingFor_ — an unauthorized or
 * missing building falls back to the caller's own, so no call returns the district.
 */
function getStaffDirectoryData(buildingFilter, targetEmail, includeArchived) {
  const ctx = getUserContext();
  const building = allowedBuildingFor_(ctx, buildingFilter) || ctx.building;
  return directoryFor_(ctx, building, targetEmail, includeArchived);
}

/**
 * The directory rows a given caller may have for one (already authorized) building:
 *   - Admin / Super Admin: everything, balances included.
 *   - Teacher: their own row in full — the balance their usage form is capped by —
 *     and nothing but a name and email for everyone else.
 *   - Anyone not in the directory: nothing.
 */
function directoryFor_(ctx, building, targetEmail, includeArchived) {
  if (isAdmin_(ctx)) return staffDirectoryData_(building, targetEmail, includeArchived);
  if (ctx.role !== 'Teacher') return [];

  const self = (ctx.email || '').toString().trim().toLowerCase();

  // targetEmail limits the balance scan to the caller — every other row's balance
  // is dropped below anyway.
  return staffDirectoryData_(building, ctx.email, false).map(s => {
    if (s.email.toString().trim().toLowerCase() === self) return s;
    return {
      name: s.name,
      email: s.email,
      building: s.building,
      primaryBuilding: s.primaryBuilding,
      archived: false
    };
  });
}

/**
 * Helper to get clean object array of Staff Directory with DYNAMIC balances.
 * When includeArchived is falsy, archived staff are omitted.
 * Private: callers are responsible for authorizing the building (see
 * getStaffDirectoryData for the client-facing, role-aware version).
 */
function staffDirectoryData_(buildingFilter, targetEmail, includeArchived) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Staff Directory');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  // Dynamic Header Lookup
  const nameIdx = headers.findIndex(h => h.toString().toLowerCase().includes('name'));
  const emailIdx = headers.findIndex(h => h.toString().toLowerCase().includes('email'));
  const roleIdx = headers.findIndex(h => h.toString().toLowerCase().includes('role'));
  const carryOverIdx = headers.findIndex(h => h.toString().toLowerCase().includes('carry'));
  const paidOutIdx = headers.findIndex(h => h.toString().toLowerCase().includes('paid'));
  const buildingIdx = headers.findIndex(h => h.toString().toLowerCase().includes('building'));
  const archivedIdx = headers.findIndex(h => h.toString().toLowerCase().includes('archiv'));
  const pendingFinalizeIdx = headers.findIndex(h => h.toString().toLowerCase().includes('pending'));

  const iName = nameIdx > -1 ? nameIdx : 0;
  const iEmail = emailIdx > -1 ? emailIdx : 1;
  const iRole = roleIdx > -1 ? roleIdx : 2;
  const iCarry = carryOverIdx > -1 ? carryOverIdx : 5; // Default F (5)
  const iPaidOut = paidOutIdx > -1 ? paidOutIdx : 6;   // Default G (6)
  const iBuilding = buildingIdx > -1 ? buildingIdx : 8; // Default I (8)

  // 1. Calculate balances dynamically. Totals are COMBINED across all of a
  //    person's buildings (ownership model): Earned can be contributed by any
  //    building. The buildingFilter is used ONLY for directory membership and the
  //    per-building archived logic below — NOT for the totals. targetEmail is kept
  //    for the single-teacher optimization.
  const balances = calculateDynamicBalances_(null, targetEmail);

  return data.map((r, i) => {
    const email = r[iEmail].toString().toLowerCase();
    const assignedBuildings = (iBuilding > -1 && r[iBuilding]) ? r[iBuilding].toString() : DEFAULT_BUILDING;
    const userBuildings = assignedBuildings.split(',').map(b => b.trim()).filter(Boolean);
    const archivedList = archivedIdx > -1 ? parseArchivedList_(r[archivedIdx], userBuildings) : [];

    // Building membership: with a filter, only show staff assigned to it.
    if (buildingFilter && !userBuildings.includes(buildingFilter)) return null;

    // Archived status is per-building. With a building filter it means "archived
    // from THIS building"; without one, it means "archived from every building".
    const isArchivedView = buildingFilter
      ? archivedList.includes(buildingFilter)
      : (userBuildings.length > 0 && userBuildings.every(b => archivedList.includes(b)));

    if (isArchivedView && !includeArchived) return null;

    const dynStats = balances[email] || { earned: 0, used: 0 };
    const carryOver = Number(r[iCarry]) || 0;
    const paidOut = Number(r[iPaidOut]) || 0;
    const pendingFinalizeYear = pendingFinalizeIdx > -1
      ? (r[pendingFinalizeIdx] || '').toString().trim() : '';

    return {
      name: r[iName],
      email: r[iEmail],
      role: r[iRole],
      earned: dynStats.earned,
      used: dynStats.used,
      carryOver: carryOver,
      paidOut: paidOut,
      building: assignedBuildings, // Keep raw string or array? Frontend expects string usually.
      primaryBuilding: userBuildings[0] || '', // first listed = primary (owns carry/paid/finalize)
      archived: isArchivedView,
      archivedBuildings: archivedList.join(', '),
      pendingFinalize: !!pendingFinalizeYear,       // flagged by a non-primary building's finalize
      pendingFinalizeYear: pendingFinalizeYear,
      total: carryOver + dynStats.earned - dynStats.used - paidOut,
      rowIndex: i + 2
    };
  }).filter(r => r !== null && r.email !== "");
}

/**
 * Aggregates Earned and Used hours per email.
 * Optionally filters transactions by building.
 * Optionally filters by a specific target email (optimization).
 */
function calculateDynamicBalances_(buildingFilter, targetEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stats = {}; // { email: { earned: 0, used: 0 } }

  // Helper to init
  const getStat = (email) => {
    const key = email.toLowerCase();
    if (!stats[key]) stats[key] = { earned: 0, used: 0 };
    return stats[key];
  };
  
  const targetEmailLower = targetEmail ? targetEmail.toLowerCase() : null;

  // 1. Process Earned (Approved only)
  const earnedSheet = ss.getSheetByName('TST Approvals (New)');
  const earnedData = earnedSheet.getDataRange().getValues();
  earnedData.shift();
  
  // Col H(7)=Hours, I(8)=Approved, N(13)=Building
  earnedData.forEach(r => {
    const email = r[0];
    if (!email) return;
    
    // Optimization: Skip if not target
    if (targetEmailLower && email.toLowerCase() !== targetEmailLower) return;

    const hours = Number(r[7]) || 0;
    const isApproved = r[8] === true || r[8] === "TRUE";
    const building = r[13] || 'OMS';

    if (isApproved) {
      if (!buildingFilter || building === buildingFilter) {
        getStat(email).earned += hours;
      }
    }
  });

  // 2. Process Used (Processed/Approved only)
  const usedSheet = ss.getSheetByName('TST Usage (New)');
  const usedData = usedSheet.getDataRange().getValues();
  usedData.shift();

  // Col D(3)=Amount, E(4)=Status(Approved), H(7)=Building
  usedData.forEach(r => {
    const email = r[0];
    if (!email) return;

    // Optimization: Skip if not target
    if (targetEmailLower && email.toLowerCase() !== targetEmailLower) return;

    const amount = Number(r[3]) || 0;
    const isApproved = r[4] === true || r[4] === "TRUE"; // Used requests are often auto-approved or manually marked
    const building = r[7] || 'OMS';

    if (isApproved) {
      if (!buildingFilter || building === buildingFilter) {
        getStat(email).used += amount;
      }
    }
  });

  return stats;
}

/**
 * Helper to safely convert spreadsheet dates to strings for client-side transfer
 */
function safeDate(val) {
  if (val instanceof Date) {
    return val.toISOString();
  }
  return val;
}

/**
 * Fetches data for the Admin Pending Earned view.
 * Source: TST Approvals (New)
 * Admin only — the queue lists other people's requests. Scoped to one building the
 * caller manages (a building they aren't assigned to falls back to their own).
 */
function getPendingEarned(buildingFilter) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  return pendingEarnedFor_(allowedBuildingFor_(ctx, buildingFilter) || ctx.building);
}

// Pending (not approved, not denied) earned rows for one building. Callers are
// responsible for authorizing the building.
function pendingEarnedFor_(building) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  const data = sheet.getDataRange().getValues();
  data.shift();

  return data.map((r, i) => {
    // Col N(13) is Building
    const rowBuilding = r[13] || 'OMS';
    
    return {
      email: r[0],
      name: r[1],
      subbedFor: r[2],
      date: safeDate(r[4]),
      period: r[5],
      timeType: r[6],
      hours: r[7],
      status: r[8],
      denied: r[10],
      building: rowBuilding,
      rowIndex: i + 2
    };
  }).filter(item => {
    const isApproved = item.status === true || item.status === "TRUE";
    const isDenied = item.denied === true || item.denied === "TRUE";

    if (item.building !== building) return false;

    return !isApproved && !isDenied && item.email !== "";
  });
}

/**
 * Fetches data for the Admin Pending Used view.
 * Source: TST Usage (New)
 * Admin only, scoped like getPendingEarned.
 */
function getPendingUsed(buildingFilter) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  return pendingUsedFor_(allowedBuildingFor_(ctx, buildingFilter) || ctx.building);
}

// Pending (unprocessed) usage rows for one building. Callers are responsible for
// authorizing the building.
function pendingUsedFor_(building) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  const data = sheet.getDataRange().getValues();
  data.shift();

  return data.map((r, i) => {
    // Col H(7) is Building
    const rowBuilding = r[7] || 'OMS';

    return {
      email: r[0],
      name: r[1],
      date: safeDate(r[2]),
      used: r[3],
      status: r[4],
      building: rowBuilding,
      rowIndex: i + 2
    };
  }).filter(item => {
    if (item.building !== building) return false;
    return (item.status === false || item.status === "" || item.status === "FALSE") && item.email !== "";
  });
}


/**
 * Gets history for a specific teacher. Self, or an admin who manages them.
 */
function getTeacherHistory(targetEmail) {
  assertSelfOrManagerOf_(targetEmail);
  return teacherHistory_(targetEmail);
}

function teacherHistory_(targetEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get Earned (Approved OR Denied)
  const earnedSheet = ss.getSheetByName('TST Approvals (New)');
  const earnedData = earnedSheet.getDataRange().getValues();
  earnedData.shift();

  const earned = earnedData
    .filter(r => {
       const isEmailMatch = r[0].toString().trim().toLowerCase() === targetEmail.trim().toLowerCase();
       return isEmailMatch; // Return ALL requests for this user, including pending
    })
    .map(r => {
      let type = 'Pending';
      if (r[10] === true) type = 'Denied';
      else if (r[8] === true) type = 'Earned';

      return {
        date: safeDate(r[4]),
        period: r[5],
        subbedFor: r[2],
        amount: r[7],
        type: type,
        denialReason: r[12],
        building: (r[13] || 'OMS').toString() // building the earned row was submitted under
      };
    });

  // 2. Get Used (Finalized)
  const usedSheet = ss.getSheetByName('TST Usage (New)');
  const usedData = usedSheet.getDataRange().getValues();
  usedData.shift();

  const used = usedData
    .filter(r => r[0].toString().toLowerCase() === targetEmail.toLowerCase() && r[4] === true)
    .map(r => ({
      date: safeDate(r[2]),
      period: 'N/A',
      subbedFor: 'N/A',
      amount: r[3],
      type: 'Used',
      building: (r[7] || 'OMS').toString()
    }));

  return [...earned, ...used].sort((a, b) => new Date(b.date) - new Date(a.date));
}

/**
 * Gets history for a specific teacher with row indices and sheet info for admin actions.
 * Used by admin to manage approved/denied items. Self, or an admin who manages them
 * (the row actions themselves are authorized separately).
 */
function getStaffHistoryWithActions(targetEmail) {
  assertSelfOrManagerOf_(targetEmail);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get Earned (ALL - Pending, Approved, Denied)
  const earnedSheet = ss.getSheetByName('TST Approvals (New)');
  const earnedData = earnedSheet.getDataRange().getValues();
  earnedData.shift();

  const earned = earnedData
    .map((r, i) => {
      // Only process rows matching this email
      if (r[0].toString().trim().toLowerCase() !== targetEmail.trim().toLowerCase()) {
        return null;
      }

      let type = 'Pending';
      if (r[10] === true) type = 'Denied';
      else if (r[8] === true) type = 'Earned';

      return {
        date: safeDate(r[4]),
        period: r[5],
        subbedFor: r[2],
        amount: r[7],
        type: type,
        denialReason: r[12],
        rowIndex: i + 2, // 1-based index + header (correct sheet position)
        sheetType: 'earned',
        amountType: r[6], // Time Type (Full/Half)
        building: (r[13] || 'OMS').toString() // building the earned row was submitted under
      };
    })
    .filter(item => item !== null);

  // 2. Get Used (ALL - Pending and Approved)
  const usedSheet = ss.getSheetByName('TST Usage (New)');
  const usedData = usedSheet.getDataRange().getValues();
  usedData.shift();

  const used = usedData
    .map((r, i) => {
      // Only process rows matching this email
      if (r[0].toString().toLowerCase() !== targetEmail.toLowerCase()) {
        return null;
      }

      const isApproved = r[4] === true;
      return {
        date: safeDate(r[2]),
        period: 'N/A',
        subbedFor: 'N/A',
        amount: r[3],
        type: isApproved ? 'Used' : 'Pending',
        denialReason: '',
        rowIndex: i + 2, // 1-based index + header (correct sheet position)
        sheetType: 'used',
        amountType: 'N/A',
        building: (r[7] || 'OMS').toString()
      };
    })
    .filter(item => item !== null);

  return [...earned, ...used].sort((a, b) => new Date(b.date) - new Date(a.date));
}


/**
 * Admin Action: Revert an Earned request back to Pending.
 * Clears Approved/Denied flags.
 */
function revertEarnedToPending(rowIndex) {
  const row = authorizeRequestRows_(getUserContext(), 'earned', [rowIndex])[0];
  return revertEarnedToPending_(row);
}

function revertEarnedToPending_(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  
  // Clear approval/denial flags and timestamps
  sheet.getRange(rowIndex, 9).setValue(false);   // Clear Approved (Col I)
  sheet.getRange(rowIndex, 10).setValue('');     // Clear Approved TS (Col J)
  sheet.getRange(rowIndex, 11).setValue(false);  // Clear Denied (Col K)
  sheet.getRange(rowIndex, 12).setValue('');     // Clear Denied TS (Col L)
  sheet.getRange(rowIndex, 13).setValue('');     // Clear Denial Reason (Col M)

  return { success: true };
}

/**
 * Admin Action: Revert a Used request back to Pending.
 * Clears Approved flag.
 */
function revertUsedToPending(rowIndex) {
  const row = authorizeRequestRows_(getUserContext(), 'used', [rowIndex])[0];
  return revertUsedToPending_(row);
}

function revertUsedToPending_(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');

  // Clear approval flag and timestamp
  sheet.getRange(rowIndex, 5).setValue(false);  // Clear Status (Col E)
  sheet.getRange(rowIndex, 6).setValue('');     // Clear Timestamp (Col F)

  return { success: true };
}


/**
 * Admin Action: Approve an Earned request.
 */
function approveEarnedRow(rowIndex, emailData) {
  const row = authorizeRequestRows_(getUserContext(), 'earned', [rowIndex])[0];
  return approveEarnedRow_(row, emailData);
}

function approveEarnedRow_(rowIndex, emailData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  
  // Get data for email BEFORE updating
  // Row Index is 1-based. 
  // Cols: A=Email(1), B=Name(2), C=SubbedFor(3), E=Date(5), F=Period(6), H=Hours(8), N=Building(14)
  const range = sheet.getRange(rowIndex, 1, 1, 14);
  const values = range.getValues()[0];
  const rowData = {
    email: values[0],
    name: values[1],
    subbedFor: values[2],
    date: values[4],
    period: values[5],
    hours: values[7],
    building: values[13] || 'OMS'
  };

  // Col I (9) is Approved Status, Col J (10) is Timestamp
  // Col K (11) is Denied Status. 
  // Safety: Ensure Denied is FALSE if we are Approving. 
  
  sheet.getRange(rowIndex, 9).setValue(true);   // Set Approved = TRUE
  sheet.getRange(rowIndex, 10).setValue(new Date()); // Set Approved Timestamp
  sheet.getRange(rowIndex, 11).setValue(false); // Set Denied = FALSE (Safety)
  
  // Send Email if requested
  if (emailData && emailData.send) {
    const config = getConfig();
    const buildingName = (config[rowData.building] && config[rowData.building].name) ? config[rowData.building].name : rowData.building;

    const formattedDate = new Date(rowData.date).toLocaleDateString();
    const subject = `TST Request for ${formattedDate} has been Approved`;
    const body = `
      <p>Your request has been approved and added to your balance.</p>
      <div style="background-color: #f8fafc; border-left: 4px solid #2d3f89; padding: 15px; margin: 15px 0;">
        <p style="margin: 0; color: #64748b; font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em;">Request Details</p>
        <p style="margin: 5px 0 0 0; color: #1e293b; font-weight: bold;">Subbed for ${escapeHtml_(rowData.subbedFor)}</p>
        <p style="margin: 0; color: #334155;">${escapeHtml_(formattedDate)} &bull; Period ${escapeHtml_(rowData.period)} &bull; +${escapeHtml_(rowData.hours)} hrs</p>
      </div>
      <p>You can check your up-to-date balance on the TST Portal.</p>
    `;
    sendStyledEmail_(rowData.email, subject, "Your TST Request was Approved!", body, "Visit the TST Portal", buildingName);
  }
  
  return true;
}

/**
 * Admin Action: Deny an Earned request.
 */
function denyEarnedRow(rowIndex, emailData) {
  const row = authorizeRequestRows_(getUserContext(), 'earned', [rowIndex])[0];
  return denyEarnedRow_(row, emailData);
}

function denyEarnedRow_(rowIndex, emailData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  
  // Get data for email
  // Range expanded to 14 to get Building in Col N
  const range = sheet.getRange(rowIndex, 1, 1, 14);
  const values = range.getValues()[0];
  const rowData = {
    email: values[0],
    name: values[1],
    subbedFor: values[2],
    date: values[4],
    period: values[5],
    hours: values[7],
    building: values[13] || 'OMS'
  };

  // Col I (9) is Approved Status
  // Col K (11) is Denied Status, Col L (12) is Denied Timestamp
  
  sheet.getRange(rowIndex, 9).setValue(false);  // Set Approved = FALSE (Safety)
  sheet.getRange(rowIndex, 11).setValue(true);  // Set Denied = TRUE
  sheet.getRange(rowIndex, 12).setValue(new Date()); // Set Denied Timestamp
  
  // Save Denial Reason (Col M/13)
  let denialReason = "";
  if (emailData) {
    const reasons = emailData.reasons || [];
    denialReason = reasons.join(", ");
    if (emailData.note) {
      if (denialReason) denialReason += ". ";
      denialReason += emailData.note;
    }
  }
  sheet.getRange(rowIndex, 13).setValue(denialReason);
  
  // Send Email if requested
  if (emailData && emailData.send) {
    const config = getConfig();
    const buildingName = (config[rowData.building] && config[rowData.building].name) ? config[rowData.building].name : rowData.building;

    const formattedDate = new Date(rowData.date).toLocaleDateString();
    const subject = `TST Request for ${formattedDate} has been Denied`;
    
    let reasonsHtml = "";
    if (emailData.reasons && emailData.reasons.length > 0) {
      reasonsHtml = `<ul style="margin: 10px 0; padding-left: 20px; color: #b91c1c;">` + 
        emailData.reasons.map(r => `<li>${escapeHtml_(r)}</li>`).join('') +
        `</ul>`;
    }

    const noteHtml = emailData.note ? `<p style="margin-top: 10px;"><em>" ${escapeHtml_(emailData.note)} "</em></p>` : "";

    const body = `
      <p>Your request has been processed and denied.</p>
      
      <div style="background-color: #fef2f2; border-left: 4px solid #ef4444; padding: 15px; margin: 15px 0;">
        <p style="margin: 0; color: #991b1b; font-weight: bold;">Reason for Denial:</p>
        ${reasonsHtml}
        ${noteHtml}
      </div>

      <div style="background-color: #f8fafc; padding: 15px; margin: 15px 0; border: 1px solid #e2e8f0; border-radius: 4px;">
        <p style="margin: 0; color: #64748b; font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em;">Request Details</p>
        <p style="margin: 5px 0 0 0; color: #1e293b; font-weight: bold;">Subbed for ${escapeHtml_(rowData.subbedFor)}</p>
        <p style="margin: 0; color: #334155;">${escapeHtml_(formattedDate)} &bull; Period ${escapeHtml_(rowData.period)}</p>
      </div>

      <p>Please review the details and resubmit if necessary, or contact the TST administrator.</p>
    `;
    
    sendStyledEmail_(rowData.email, subject, "TST Request Update", body, "Visit the TST Portal", buildingName);
  }
  
  return true;
}

/**
 * Admin Action: Delete an Earned request.
 * Deletes from BOTH 'TST Approvals (New)' and 'Form Responses 1'.
 */
function deleteEarnedRow(rowIndex) {
  const row = authorizeRequestRows_(getUserContext(), 'earned', [rowIndex])[0];
  return deleteEarnedRow_(row);
}

function deleteEarnedRow_(rowIndex) {
  const row = Number(rowIndex);
  if (!Number.isInteger(row) || row < 2) throw new Error("Invalid row index.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const approvalSheet = ss.getSheetByName('TST Approvals (New)');
  if (!approvalSheet) throw new Error("Sheet 'TST Approvals (New)' not found.");
  const formSheet = ss.getSheetByName('Form Responses 1');
  if (!formSheet) throw new Error("Sheet 'Form Responses 1' not found.");

  if (row > approvalSheet.getLastRow()) {
    throw new Error("This transaction may have already been deleted.");
  }

  // 1. Get details from Approval Sheet to find match
  // Indexes: 0=Email, 1=Name, 2=SubbedFor, 4=Date, 5=Period
  const rowValues = approvalSheet.getRange(row, 1, 1, 6).getValues()[0];
  const email = rowValues[0];
  const date = new Date(rowValues[4]);
  const period = rowValues[5];

  // 2. Find and Delete in Form Responses 1
  const formData = formSheet.getDataRange().getValues();
  // Form Responses: Col B=Email (1), E=Date (4), F=Period (5)
  // Loop backwards to safely delete
  for (let i = formData.length - 1; i >= 1; i--) { // Skip header
    const r = formData[i];
    const rDate = new Date(r[4]);

    // Loose date comparison (checking year, month, day)
    const isDateMatch = rDate.getFullYear() === date.getFullYear() &&
                        rDate.getMonth() === date.getMonth() &&
                        rDate.getDate() === date.getDate();

    if (r[1] === email && isDateMatch && r[5] == period) {
       formSheet.deleteRow(i + 1);
       break;
    }
  }

  // 3. Delete from TST Approvals (New)
  approvalSheet.deleteRow(row);

  return true;
}

/**
 * Admin Action: Edit an Earned request.
 * Updates BOTH 'TST Approvals (New)' and 'Form Responses 1'.
 */
function updateEarnedRow(rowIndex, newData) {
  const row = authorizeRequestRows_(getUserContext(), 'earned', [rowIndex])[0];
  return updateEarnedRow_(row, newData);
}

function updateEarnedRow_(rowIndex, newData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const approvalSheet = ss.getSheetByName('TST Approvals (New)');
  const formSheet = ss.getSheetByName('Form Responses 1');
  
  // 1. Get OLD details to find match in Form Responses
  const rowValues = approvalSheet.getRange(rowIndex, 1, 1, 6).getValues()[0];
  const oldEmail = rowValues[0];
  const oldDate = new Date(rowValues[4]);
  const oldPeriod = rowValues[5];
  
  // 2. Update Form Responses 1
  const formData = formSheet.getDataRange().getValues();
  let foundInForm = false;
  
  for (let i = formData.length - 1; i >= 1; i--) {
    const r = formData[i];
    const rDate = new Date(r[4]);
    const isDateMatch = rDate.getFullYear() === oldDate.getFullYear() &&
                        rDate.getMonth() === oldDate.getMonth() &&
                        rDate.getDate() === oldDate.getDate();

    if (r[1] === oldEmail && isDateMatch && r[5] == oldPeriod) {
       // Found match. Update columns.
       // Form Responses: C=SubbedFor (2), E=Date (4), F=Period (5), G=AmountType (6), H=Decimal (7)
       // We don't update Timestamp or Email usually, but we could.
       
       formSheet.getRange(i + 1, 3).setValue(newData.subbedFor);
       formSheet.getRange(i + 1, 5).setValue(newData.date);
       formSheet.getRange(i + 1, 6).setValue(newData.period);
       formSheet.getRange(i + 1, 7).setValue(newData.amountType);
       formSheet.getRange(i + 1, 8).setValue(newData.amountDecimal);
       foundInForm = true;
       break;
    }
  }

  // 3. Update TST Approvals (New) directly to reflect changes immediately
  // Cols: C=SubbedFor (3/idx 2), E=Date (5/idx 4), F=Period (6/idx 5), G=Type (7/idx 6), H=Hours (8/idx 7)
  approvalSheet.getRange(rowIndex, 3).setValue(newData.subbedFor);
  approvalSheet.getRange(rowIndex, 5).setValue(newData.date);
  approvalSheet.getRange(rowIndex, 6).setValue(newData.period);
  approvalSheet.getRange(rowIndex, 7).setValue(newData.amountType);
  approvalSheet.getRange(rowIndex, 8).setValue(newData.amountDecimal);

  return true;
}


/**
 * Admin Action: Approve a Used request.
 */
function approveUsedRow(rowIndex) {
  const row = authorizeRequestRows_(getUserContext(), 'used', [rowIndex])[0];
  return approveUsedRow_(row);
}

function approveUsedRow_(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  // Col E (5) is status, Col F (6) is timestamp
  sheet.getRange(rowIndex, 5).setValue(true);
  sheet.getRange(rowIndex, 6).setValue(new Date());
  return true;
}

/**
 * Admin Action: Delete a Used request.
 */
function deleteUsedRow(rowIndex) {
  const row = authorizeRequestRows_(getUserContext(), 'used', [rowIndex])[0];
  return deleteUsedRow_(row);
}

function deleteUsedRow_(rowIndex) {
  const row = Number(rowIndex);
  if (!Number.isInteger(row) || row < 2) throw new Error("Invalid row index.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  if (!sheet) throw new Error("Sheet 'TST Usage (New)' not found.");

  if (row > sheet.getLastRow()) {
    throw new Error("This transaction may have already been deleted.");
  }

  sheet.deleteRow(row);
  return true;
}

/**
 * Admin Action: Edit a Used request.
 */
function updateUsedRow(rowIndex, newData) {
  const row = authorizeRequestRows_(getUserContext(), 'used', [rowIndex])[0];
  return updateUsedRow_(row, newData);
}

function updateUsedRow_(rowIndex, newData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  // Cols: C=Date (3), D=Amount (4)
  sheet.getRange(rowIndex, 3).setValue(newData.date);
  sheet.getRange(rowIndex, 4).setValue(newData.amount);
  return true;
}

/**
 * Create a new Usage entry (Admin or Teacher). A teacher may only submit for
 * themselves; an admin may submit for staff they manage (View As).
 */
function submitUsage(formObj) {
  assertSelfOrManagerOf_(formObj && formObj.email);
  return submitUsage_(formObj);
}

function submitUsage_(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  
  // Resolve Building
  let building = formObj.building;
  if (!building) {
    const staff = staffDirectoryData_().find(s => s.email.toLowerCase() === formObj.email.toLowerCase());
    // Get first building from list
    building = staff ? (staff.building.includes(',') ? staff.building.split(',')[0].trim() : staff.building) : DEFAULT_BUILDING;
  }

  // Columns: A: Email, B: Name, C: Date, D: TST Used, E: Status, F: Timestamp, G: Notes, H: Building
  sheet.appendRow([
    formObj.email,
    formObj.name,
    formObj.date,
    formObj.amount,
    false, // Default unchecked
    "",    // No timestamp yet
    formObj.notes || "",
    building
  ]);
  return true;
}

/**
 * Create a new Earned entry (Teacher subbing).
 * Writes to Form Responses 1. Same self-or-manager rule as submitUsage.
 */
function submitEarned(formObj) {
  assertSelfOrManagerOf_(formObj && formObj.email);
  return submitEarned_(formObj);
}

function submitEarned_(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Archive to Form Responses 1 (Keep as backup)
  const formSheet = ss.getSheetByName('Form Responses 1');
  const timestamp = new Date();
  const subbedForName = formObj.subbedForName;
  const otherText = formObj.subbedForType === 'Other' ? 'Other' : '';

  formSheet.appendRow([
    timestamp,
    formObj.email,
    subbedForName,
    otherText,
    formObj.date,
    formObj.period,
    formObj.amountType, 
    formObj.amountDecimal
  ]);

  // 2. Process submission immediately (Decoupled from Form Trigger)
  // This ensures script-initiated submissions appear in the app.
  processEarnedSubmission_({
    email: formObj.email,
    subbedFor: subbedForName,
    otherText: otherText,
    dateStr: formObj.date,
    period: formObj.period,
    amountType: formObj.amountType,
    amountDecimal: formObj.amountDecimal,
    building: formObj.building
  });
  
  return true;
}

/**
 * Admin Multi-Submit: Handles creating both Earned and Used records 
 * based on Admin input.
 */
function adminSubmitRequest(data) {
  // Use current admin context building if available? 
  // Or the building of the user?
  // Usually admins work in a building context.
  const ctx = getUserContext();
  const adminBuilding = ctx.building; // Current View Building

  const targets = [];
  if (data.earner.type === 'Staff' && data.earner.email) targets.push(data.earner.email);
  if (data.user.type === 'Staff' && data.user.email) targets.push(data.user.email);
  assertCanManageStaffEmails_(ctx, targets);

  // 1. Handle Earner (If staff member is selected)
  if (data.earner.type === 'Staff' && data.earner.email) {
    // We treat this like a form submission so it flows into the normal Pending pipeline
    submitEarned_({
      email: data.earner.email,
      subbedForType: data.user.type, // 'Staff' or 'Other'
      subbedForName: data.user.name,
      date: data.details.date,
      period: data.details.period,
      amountType: data.details.amountType,
      amountDecimal: data.details.amount,
      building: adminBuilding // Tag with admin's current building view
    });
  }

  // 2. Handle User (If staff member is selected)
  if (data.user.type === 'Staff' && data.user.email) {
    submitUsage_({
      email: data.user.email,
      name: data.user.name,
      date: data.details.date,
      amount: data.details.amount,
      building: adminBuilding
    });
  }

  return true;
}

/**
 * Batch Process: Handles a queue of mixed requests.
 */
function processBatch(queue) {
  if (!Array.isArray(queue) || queue.length === 0) return;

  // Admin only. Authorize every target up front so one out-of-building person
  // rejects the whole batch instead of partially applying it.
  const targets = queue
    .filter(item => item && (item.type === 'earned' || item.type === 'used'))
    .map(item => (item.payload && item.payload.email) || '');
  assertCanManageStaffEmails_(getUserContext(), targets);

  queue.forEach(item => {
    try {
      if (item.type === 'earned') {
        submitEarned_(item.payload);
      } else if (item.type === 'used') {
        submitUsage_(item.payload);
      }
    } catch (e) {
      console.error("Error processing batch item:", item, e);
      // We continue processing others even if one fails
    }
  });
  
  return true;
}

/**
 * Batch Process: Sends status emails to a list of staff members.
 */
function sendBatchStatusEmails(emails) {
  if (!emails || !Array.isArray(emails)) throw new Error("Invalid email list.");
  
  // Get sender context; every recipient must be staff this admin manages.
  const ctx = getUserContext();
  assertCanManageStaffEmails_(ctx, emails);
  const senderName = ctx.name || "TST Admin";
  const senderEmail = ctx.email;
  
  const emailOptions = {
    name: `${senderName} (via TST)`,
    replyTo: senderEmail,
    buildingCode: ctx.building // Fixed: Use ctx.building instead of undefined primaryBuilding
  };

  let successCount = 0;
  let failCount = 0;
  
  const staffDir = staffDirectoryData_();

  emails.forEach(email => {
    try {
      // Find name for this email to pass to sendStatusEmail (optimization: get name from dir if possible)
      // sendStatusEmail(email, name) expects name.
      const staff = staffDir.find(s => s.email.toLowerCase() === email.toLowerCase());
      const name = staff ? staff.name : "Staff Member";

      sendStatusEmail_(email, name, emailOptions);
      successCount++;
    } catch (e) {
      console.error(`Failed to send email to ${email}:`, e);
      failCount++;
    }
  });

  return { success: successCount, failed: failCount };
}

/**
 * Sends an email report to a staff member the calling admin manages.
 */
function sendStatusEmail(targetEmail, targetName, emailOptions) {
  assertCanManageStaffEmails_(getUserContext(), [targetEmail]);
  return sendStatusEmail_(targetEmail, targetName, emailOptions);
}

function sendStatusEmail_(targetEmail, targetName, emailOptions) {
  const history = teacherHistory_(targetEmail);
  const staff = staffDirectoryData_().find(s => s.email.toLowerCase() === targetEmail.toLowerCase());
  
  if (!staff) throw new Error("Staff member not found.");
  
  const config = getConfig();
  const primaryBuilding = staff.building.includes(',') ? staff.building.split(',')[0].trim() : staff.building;
  const buildingName = (config[primaryBuilding] && config[primaryBuilding].name) ? config[primaryBuilding].name : primaryBuilding;

  // Summary Section (5 Cards)
  let htmlContent = `
    <div style="display: table; width: 100%; border-spacing: 10px; margin-bottom: 20px; table-layout: fixed;">
      <div style="display: table-row;">
        <div style="display: table-cell; background-color: #ffffff; border: 1px solid #e5e7eb; border-radius: 8px; padding: 12px; text-align: center;">
          <div style="color: #6b7280; font-size: 10px; font-weight: bold; text-transform: uppercase; margin-bottom: 4px;">Carry Over</div>
          <div style="color: #374151; font-size: 18px; font-weight: bold;">${Number(staff.carryOver || 0).toFixed(2)}</div>
        </div>
        <div style="display: table-cell; background-color: #eff6ff; border: 1px solid #dbeafe; border-radius: 8px; padding: 12px; text-align: center;">
          <div style="color: #1e40af; font-size: 10px; font-weight: bold; text-transform: uppercase; margin-bottom: 4px;">Earned</div>
          <div style="color: #1e3a8a; font-size: 18px; font-weight: bold;">${Number(staff.earned || 0).toFixed(2)}</div>
        </div>
      </div>
      <div style="display: table-row;">
        <div style="display: table-cell; background-color: #fef2f2; border: 1px solid #fee2e2; border-radius: 8px; padding: 12px; text-align: center;">
          <div style="color: #991b1b; font-size: 10px; font-weight: bold; text-transform: uppercase; margin-bottom: 4px;">Used</div>
          <div style="color: #7f1d1d; font-size: 18px; font-weight: bold;">${Number(staff.used || 0).toFixed(2)}</div>
        </div>
        <div style="display: table-cell; background-color: #fffbeb; border: 1px solid #fef3c7; border-radius: 8px; padding: 12px; text-align: center;">
          <div style="color: #92400e; font-size: 10px; font-weight: bold; text-transform: uppercase; margin-bottom: 4px;">Paid Out</div>
          <div style="color: #78350f; font-size: 18px; font-weight: bold;">${Number(staff.paidOut || 0).toFixed(2)}</div>
        </div>
      </div>
      <div style="display: table-row;">
        <div style="display: table-cell; background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; padding: 12px; text-align: center;">
          <div style="color: #4b5563; font-size: 10px; font-weight: bold; text-transform: uppercase; margin-bottom: 4px;">Balance</div>
          <div style="color: #111827; font-size: 18px; font-weight: bold;">${Number(staff.total || 0).toFixed(2)}</div>
        </div>
        <div style="display: table-cell; visibility: hidden;"></div>
      </div>
    </div>    
    <h3 style="color: #374151; font-size: 16px; margin-bottom: 15px; border-bottom: 1px solid #e5e7eb; padding-bottom: 10px; font-weight: 600;">Activity History</h3>
    
    <table cellpadding="0" cellspacing="0" style="width: 100%; border-collapse: collapse; font-size: 13px;">
      <thead>
        <tr style="background-color: #f9fafb;">
          <th style="text-align: left; padding: 10px; border-bottom: 1px solid #e5e7eb; color: #6b7280; font-weight: 600; text-transform: uppercase; font-size: 11px;">Date</th>
          <th style="text-align: left; padding: 10px; border-bottom: 1px solid #e5e7eb; color: #6b7280; font-weight: 600; text-transform: uppercase; font-size: 11px;">Details</th>
          <th style="text-align: right; padding: 10px; border-bottom: 1px solid #e5e7eb; color: #6b7280; font-weight: 600; text-transform: uppercase; font-size: 11px;">Hours</th>
        </tr>
      </thead>
      <tbody>`;
      
  if (history.length === 0) {
    htmlContent += `
      <tr>
        <td colspan="3" style="padding: 20px; text-align: center; color: #9ca3af; font-style: italic;">No history found.</td>
      </tr>`;
  } else {
    history.forEach(h => {
      const dateStr = new Date(h.date).toLocaleDateString();
      let amountStyle = 'font-weight: 600;';
      let rowBg = '#ffffff';
      let typeLabel = '';
      let details = '';
      let amountDisplay = '';
      
      if (h.type === 'Earned') {
        amountStyle += 'color: #2d3f89;'; 
        amountDisplay = `+${Number(h.amount).toFixed(2)}`;
        typeLabel = `<span style="background-color: #eaecf5; color: #2d3f89; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: bold;">EARNED</span>`;
        details = `<div style="font-weight: 500; color: #333;">${escapeHtml_(h.period)}</div><div style="font-size: 11px; color: #666;">Covered: ${escapeHtml_(h.subbedFor)}</div>`;
        rowBg = '#f8fafc';
      } else if (h.type === 'Used') {
        amountStyle += 'color: #ad2122;';
        amountDisplay = `-${Number(h.amount).toFixed(2)}`;
        typeLabel = `<span style="background-color: #e5c7c7; color: #ad2122; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: bold;">USED</span>`;
        details = '<div style="font-weight: 500; color: #333;">Redeemed</div>';
        rowBg = '#fffafa';
      } else if (h.type === 'Pending') {
        amountStyle += 'color: #854d0e;';
        amountDisplay = `${Number(h.amount).toFixed(2)}`;
        typeLabel = `<span style="background-color: #fef9c3; color: #854d0e; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: bold;">PENDING</span>`;
        details = `<div style="font-weight: 500; color: #333;">${escapeHtml_(h.period !== 'N/A' ? h.period : 'Usage Request')}</div>`;
        rowBg = '#ffffff';
      } else { // Denied
        amountStyle += 'color: #999; text-decoration: line-through;';
        amountDisplay = `${Number(h.amount).toFixed(2)}`;
        typeLabel = `<span style="background-color: #f2f2f3; color: #666; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: bold;">DENIED</span>`;
        details = `<div style="font-weight: 500; color: #666; text-decoration: line-through;">${escapeHtml_(h.period)}</div>`;
        if (h.denialReason) {
          details += `<div style="font-size: 11px; color: #ad2122; margin-top: 2px;">Reason: ${escapeHtml_(h.denialReason)}</div>`;
        }
        rowBg = '#ffffff';
      }
      
      htmlContent += `
        <tr style="background-color: ${rowBg};">
          <td style="padding: 10px; border-bottom: 1px solid #f3f4f6; vertical-align: top; color: #4b5563;">
            <div style="margin-bottom: 4px;">${dateStr}</div>
            ${typeLabel}
          </td>
          <td style="padding: 10px; border-bottom: 1px solid #f3f4f6; vertical-align: top;">
            ${details}
          </td>
          <td style="padding: 10px; border-bottom: 1px solid #f3f4f6; vertical-align: top; text-align: right; ${amountStyle}">
            ${amountDisplay}
          </td>
        </tr>`;
    });
  }
  
  htmlContent += `
      </tbody>
    </table>
    
    <p style="margin-top: 25px; font-size: 11px; color: #9ca3af; text-align: center;">
      Generated on ${new Date().toLocaleDateString()} at ${new Date().toLocaleTimeString()}
    </p>
  `;
  
  sendStyledEmail_(
    targetEmail,
    "Your TST Hours Report",
    `TST Report for ${targetName}`,
    htmlContent,
    "View Full Dashboard",
    buildingName,
    emailOptions // Pass custom options
  );
  
  return true;
}

/**
 * Core Logic for Processing an Earned Time Submission.
 * Handles Name Lookup, Building Resolution, and Appending to TST Approvals (New).
 * This is used by both onFormSubmit (Triggers) and submitEarned (Web App).
 * 
 * @param {Object} data - Standardized submission data.
 */
function processEarnedSubmission_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const approvalSheet = ss.getSheetByName('TST Approvals (New)');

  let { 
    email, 
    subbedFor, 
    otherText, 
    dateStr, 
    period, 
    amountType, 
    amountDecimal, 
    building 
  } = data;

  // Clean Subbed For Name (Remove Titles for Legacy Form compatibility)
  if (subbedFor) {
    subbedFor = subbedFor.replace(/^(Mr\.|Ms\.|Mrs\.|Miss|Dr\.)\s*/i, "").trim();
  }

  // Lookup Name & Building using helper
  const staff = staffDirectoryData_().find(s => s.email.toLowerCase() === email.toString().toLowerCase());
  const earnerName = staff ? staff.name : email;

  // Resolve Building: Use provided OR Staff Primary
  let earnerBuilding = building;
  if (!earnerBuilding) {
    earnerBuilding = DEFAULT_BUILDING;
    if (staff && staff.building) {
      earnerBuilding = staff.building.includes(',') ? staff.building.split(',')[0].trim() : staff.building;
    }
  }

  // Legacy Form Support: Calculate missing decimal if needed
  if (amountDecimal == null || amountDecimal === '') {
    amountDecimal = calculatePeriods(period, amountType, earnerBuilding);
  }

  // Idempotency guard. The same earned submission can reach this function twice
  // (the web app writes the row directly AND a Form Responses 1 trigger/sync can
  // reprocess the same row), which previously produced duplicate pending rows.
  // A person cannot sub the same period twice on the same date, so a matching
  // non-denied row for email + date + period is always a duplicate. A script lock
  // serializes the direct write against the trigger so they can't race.
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { /* proceed unlocked rather than drop the request */ }
  try {
    const keyEmail = email.toString().toLowerCase().trim();
    const keyDate = normDateKey_(dateStr);
    const keyPeriod = (period || '').toString().trim();

    const existing = approvalSheet.getDataRange().getValues();
    for (let i = 1; i < existing.length; i++) {
      const r = existing[i];
      if (!r[0]) continue;
      const isDenied = r[10] === true || r[10] === 'TRUE';
      if (isDenied) continue; // a denied row shouldn't block a genuine re-entry
      if (r[0].toString().toLowerCase().trim() === keyEmail &&
          normDateKey_(r[4]) === keyDate &&
          (r[5] || '').toString().trim() === keyPeriod) {
        return false; // duplicate — skip
      }
    }

    // Append to Approvals (Col A:N)
    approvalSheet.appendRow([
      email,          // A: Email
      earnerName,     // B: Name
      subbedFor,      // C: Subbed For
      otherText || "",// D: Other Details
      dateStr,        // E: Date
      period,         // F: Period
      amountType,     // G: Time Type
      amountDecimal,  // H: Hours
      false,          // I: Approved (Default)
      "",             // J: Approved TS
      false,          // K: Denied (Default)
      "",             // L: Denied TS
      "",             // M: Denial Reason
      earnerBuilding  // N: Building
    ]);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }

  return true;
}

/**
 * Normalizes a date cell/string to a yyyy-MM-dd key for duplicate detection.
 * Keeps plain date strings as-is (avoids UTC off-by-one) and formats real Dates
 * in the script timezone.
 */
function normDateKey_(d) {
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return (d ? d.toString().trim() : '').split('T')[0];
}

/**
 * Trigger: On Form Submit
 * Syncs new rows from 'Form Responses 1' to 'TST Approvals (New)'.
 * Must be manually set up as an Installable Trigger in Apps Script editor.
 */
function onFormSubmit(e) {
  if (!e || !e.values) return; // Safety check
  // This function must stay public for the installable trigger, so it is also
  // reachable through google.script.run. A real trigger event carries a live Range;
  // a client call can only pass plain JSON, so refuse anything without one.
  if (!e.range || typeof e.range.getSheet !== 'function') {
    throw new Error('onFormSubmit can only be run by the form submit trigger.');
  }
  
  // [0] Timestamp, [1] Email, [2] SubbedFor, [3] Other, [4] Date, [5] Period, [6] Type, [7] Decimal
  const data = {
    email: e.values[1],
    subbedFor: e.values[2],
    otherText: e.values[3],
    dateStr: e.values[4],
    period: e.values[5],
    amountType: e.values[6],
    amountDecimal: e.values[7],
    building: null // Will be resolved by staff lookup
  };

  processEarnedSubmission_(data);
}

/**
 * Helper to calculate period value for legacy forms.
 * @param {string} selectedPeriod 
 * @param {string} amountType 
 * @param {string} buildingCode 
 */
function calculatePeriods(selectedPeriod, amountType, buildingCode) {
  // Normalize
  const p = selectedPeriod ? selectedPeriod.toString() : "";
  const type = amountType ? amountType.toString().toLowerCase() : "";
  const b = buildingCode || DEFAULT_BUILDING;

  if (b === 'OMS') {
    // Special Rules for OMS
    // Period 6 or 7 (but not 6/7) is always 0.5
    if ((p.includes('Period 6') || p.includes('Period 7')) && !p.includes('Period 6/7')) {
      return 0.5;
    }
  }

  // Default Rules
  if (type.includes('half')) return 0.5;
  if (type.includes('full')) return 1.0;

  // Fallback
  return 1.0;
}

/**
 * Trigger: On Spreadsheet Open.
 * Adds custom menu for Admins.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('TST Admin')
    .addItem('Authorize Email Service', 'setupEmailService')
    .addToUi();
}

/**
 * Admin Action: Sets up the email sending trigger for the CURRENT user.
 * This must be run by the Building Secretary/Admin from the Spreadsheet menu.
 */
function setupEmailService() {
  const ui = SpreadsheetApp.getUi();
  const ctx = getUserContext();
  assertAdmin_(ctx);

  const buildings = installEmailTriggers_(ctx);

  ui.alert(
    'Email Service Authorized',
    `Success! TST email for ${buildings.join(', ')} will be sent from ${ctx.email}, and coverage assignments will appear on that building's calendar.\n\n(Triggers installed: OnChange + 1-Minute Timer + 7am Assignment Reminders)`,
    ui.ButtonSet.OK
  );
}

// ===== Email service health =====
// Queued mail is sent by each building admin's own trigger, which is what makes
// them the sender. Triggers are per-user, so nobody can install one on anyone
// else's behalf — and Apps Script disables a trigger after repeated failures,
// sending the notice to its owner rather than to whoever notices the silence.
//
// So the app watches for the symptom instead: mail for a building sitting Pending
// for longer than any working trigger would leave it. That catches a trigger that
// was never installed and one that has since broken, which the authorization
// record alone cannot.

const EMAIL_AUTH_PREFIX_ = 'EMAIL_AUTH_';
const AUTHORIZE_URL_KEY_ = 'EMAIL_AUTHORIZE_URL';

/** Older than this and a building's queue is not being drained. */
const QUEUE_STALL_MINUTES_ = 15;

/**
 * Installs this user's triggers. Shared by the spreadsheet menu and the
 * authorization page, because the only thing that differs is how it is reached —
 * the trigger always belongs to whoever is executing.
 */
function installEmailTriggers_(ctx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getUserTriggers(ss).forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'processEmailQueue' || fn === 'nudgeOutstandingAssignments') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Immediate attempt when the sheet changes.
  ScriptApp.newTrigger('processEmailQueue').forSpreadsheet(ss).onChange().create();
  // Backstop, in case an onChange is missed.
  ScriptApp.newTrigger('processEmailQueue').timeBased().everyMinutes(1).create();
  // One reminder per assignment whose coverage date has passed unrecorded.
  ScriptApp.newTrigger('nudgeOutstandingAssignments').timeBased().atHour(7).everyDays(1).create();

  recordEmailAuthorization_(ctx);
  return ctx.buildings || [DEFAULT_BUILDING];
}

function recordEmailAuthorization_(ctx) {
  const props = PropertiesService.getScriptProperties();
  const record = JSON.stringify({ email: ctx.email, name: ctx.name || ctx.email, at: new Date().toISOString() });
  (ctx.buildings || [DEFAULT_BUILDING]).forEach(b => props.setProperty(EMAIL_AUTH_PREFIX_ + b, record));
}

function emailAuthorizationFor_(building) {
  const raw = PropertiesService.getScriptProperties().getProperty(EMAIL_AUTH_PREFIX_ + building);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

/** How long the oldest Pending row for a building has been waiting, in minutes. */
function oldestPendingMinutes_(building) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email Queue');
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  const headers = data[0];
  const statusIdx = headers.indexOf('Status');
  const buildingIdx = headers.indexOf('Building');
  const tsIdx = headers.indexOf('Timestamp');
  if (statusIdx === -1 || buildingIdx === -1 || tsIdx === -1) return 0;

  const now = new Date().getTime();
  let oldest = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][statusIdx] !== 'Pending') continue;
    if ((data[i][buildingIdx] || DEFAULT_BUILDING).toString() !== building) continue;
    const queuedAt = new Date(data[i][tsIdx]).getTime();
    if (!isFinite(queuedAt)) continue;
    oldest = Math.max(oldest, Math.round((now - queuedAt) / 60000));
  }
  return oldest;
}

/**
 * Whether a building's mail is actually going out. Admin only, and scoped the
 * same way every other read is.
 *
 * `healthy` is false when nobody has ever authorized the building, or when its
 * queue has stalled. The reason is separated out so the banner can say which,
 * because "never set up" and "it broke last Tuesday" need different responses.
 */
function getEmailServiceStatus(buildingFilter) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const building = allowedBuildingFor_(ctx, buildingFilter) || ctx.building;

  const authorization = emailAuthorizationFor_(building);
  const stalledMinutes = oldestPendingMinutes_(building);
  const stalled = stalledMinutes >= QUEUE_STALL_MINUTES_;

  return {
    building: building,
    authorized: !!authorization,
    authorizedBy: authorization ? (authorization.name || authorization.email) : '',
    authorizedAt: authorization ? authorization.at : '',
    stalledMinutes: stalledMinutes,
    healthy: !!authorization && !stalled,
    reason: !authorization ? 'never' : (stalled ? 'stalled' : ''),
    authorizeUrl: authorizeUrl_()
  };
}

function authorizeUrl_() {
  return (PropertiesService.getScriptProperties().getProperty(AUTHORIZE_URL_KEY_) || '').toString().trim();
}

/**
 * The address of the second web-app deployment — the one configured to run as the
 * user accessing it. That is the whole point of having two: a trigger belongs to
 * whoever executes the code, and the main deployment runs as the deployer, so a
 * button there would only ever install the deployer's triggers again.
 */
function setAuthorizeUrl(url) {
  const ctx = getUserContext();
  if (!ctx.isSuperAdmin) throw new Error('Only a Super Admin can set the authorization URL.');

  const clean = (url || '').toString().trim();
  if (clean && !/^https:\/\/script\.google\.com\//.test(clean)) {
    throw new Error('That does not look like an Apps Script web app URL.');
  }
  PropertiesService.getScriptProperties().setProperty(AUTHORIZE_URL_KEY_, clean);
  return true;
}

/**
 * The page behind the banner's button, served by the second deployment.
 *
 * Reached through doGet, so it runs as whoever opened it — which is exactly why
 * it exists. Renders its own errors rather than throwing, because an Apps Script
 * exception page tells a school secretary nothing.
 */
function authorizeEmailServicePage_() {
  let ctx;
  try {
    ctx = getUserContext();
  } catch (err) {
    return emailAuthPage_(false, 'We could not read the staff directory',
      'This usually means your account does not have access to the TST spreadsheet yet. Ask your TST administrator to share it with you, then open this link again.');
  }

  if (!isAdmin_(ctx)) {
    return emailAuthPage_(false, 'Administrators only',
      'This page sets up TST email sending for a building, which only an administrator can do.');
  }

  let buildings;
  try {
    buildings = installEmailTriggers_(ctx);
  } catch (err) {
    return emailAuthPage_(false, 'Setup could not finish',
      (err && err.message) ? err.message : String(err));
  }

  return emailAuthPage_(true, 'Email service is on',
    'TST email for ' + buildings.join(', ') + ' will now be sent from ' + ctx.email +
    ', and coverage assignments will appear on that building’s calendar. You can close this tab.');
}

function emailAuthPage_(ok, title, message) {
  const accent = ok ? '#2d3f89' : '#ad2122';
  const icon = ok ? 'check' : 'exclamation';
  const html = '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<title>' + escapeHtml_(title) + '</title>' +
    '<link href="https://fonts.googleapis.com/css2?family=Lexend:wght@300;400;600;700&display=swap" rel="stylesheet">' +
    '<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">' +
    '<style>' +
    "body { font-family: 'Lexend', sans-serif; background:#f9fafb; display:flex; align-items:center; justify-content:center; min-height:100vh; margin:0; }" +
    '.card { background:#fff; border-radius:12px; box-shadow:0 10px 15px -3px rgba(0,0,0,.1); max-width:480px; width:100%; border:1px solid #e5e7eb; overflow:hidden; }' +
    '.header { background:' + accent + '; padding:24px; text-align:center; color:#fff; }' +
    '.icon { background:#fff; width:64px; height:64px; border-radius:50%; display:flex; align-items:center; justify-content:center; margin:0 auto 12px; color:' + accent + '; font-size:28px; }' +
    '.content { padding:32px 24px; text-align:center; }' +
    '.title { font-size:22px; font-weight:700; color:#1f2937; margin:0 0 10px; }' +
    '.msg { color:#4b5563; font-size:14px; line-height:1.6; margin:0; }' +
    '</style></head><body><div class="card">' +
    '<div class="header"><div class="icon"><i class="fas fa-' + icon + '"></i></div>' +
    '<h1 style="margin:0;font-size:18px;font-weight:600;">Orono TST Manager</h1></div>' +
    '<div class="content"><h2 class="title">' + escapeHtml_(title) + '</h2>' +
    '<p class="msg">' + escapeHtml_(message) + '</p></div></div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Trigger Handler: Processes the Email Queue.
 * Runs as the user who installed the trigger (The Admin).
 *
 * Has to stay public for the installable triggers setupEmailService creates, which
 * also makes it reachable from google.script.run: it sends the queue as whoever
 * calls it (ctx.name / ctx.email become the From name and Reply-To), so a client
 * call must be an admin. Trigger runs are recognised by their event object and are
 * never blocked, even if the trigger owner has since left the directory.
 */
function processEmailQueue(e) {
  // Calendar work first: it queues the assignment emails, and doing it here means
  // the same run sends them. It must never stop the queue draining.
  if (!isTriggerEvent_(e)) assertAdmin_(getUserContext());
  try {
    processPendingAssignments_();
  } catch (err) {
    console.error('Assignment calendar work failed', err);
  }
  return processEmailQueue_();
}

/**
 * True for an event object Apps Script itself passed in. Every trigger event carries
 * an authMode enum; google.script.run parameters are JSON, so a client can send the
 * property but never the enum object the comparison needs.
 */
function isTriggerEvent_(e) {
  if (!e || !e.authMode) return false;
  const modes = ScriptApp.AuthMode;
  return e.authMode === modes.FULL || e.authMode === modes.LIMITED ||
         e.authMode === modes.CUSTOM_FUNCTION || e.authMode === modes.NONE;
}

function processEmailQueue_() {
  const MAX_BATCH = 200; // Process up to 200 at a time to cover full building reports
  
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return; 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Email Queue');
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers = data[0];
    const statusIdx = headers.indexOf('Status');
    const buildingIdx = headers.indexOf('Building');
    const recipientIdx = headers.indexOf('Recipient');
    const subjectIdx = headers.indexOf('Subject');
    const bodyIdx = headers.indexOf('Body');
    const optionsIdx = headers.indexOf('Options');
    const tsIdx = headers.indexOf('Timestamp');

    if (statusIdx === -1 || buildingIdx === -1) return;

    const ctx = getUserContext();
    const assignedBuildings = ctx.buildings || [DEFAULT_BUILDING];

    const rowsToProcess = [];
    
    for (let i = 1; i < data.length; i++) {
      if (rowsToProcess.length >= MAX_BATCH) break;

      const row = data[i];
      const status = row[statusIdx];
      const rowBuilding = row[buildingIdx];

      if (status !== 'Pending') continue;

      // Only the building's own admin sends its mail. A Super Admin's trigger used
      // to sweep every building, racing the building admin each minute and making
      // the From name a coin flip. A building with nobody authorized now queues
      // rather than going out under the wrong name.
      const canProcess = assignedBuildings.includes(rowBuilding || DEFAULT_BUILDING);
      
      if (canProcess) {
        rowsToProcess.push({
          rowIndex: i + 1,
          recipient: row[recipientIdx],
          subject: row[subjectIdx],
          body: row[bodyIdx],
          options: row[optionsIdx] ? JSON.parse(row[optionsIdx]) : {}
        });
      }
    }

    // 1. Process Emails
    rowsToProcess.forEach(item => {
      try {
        // Mark as Processing IMMEDIATELY to prevent other triggers from grabbing it
        sheet.getRange(item.rowIndex, statusIdx + 1).setValue('Processing');
        SpreadsheetApp.flush(); // Force update to sheet
        
        MailApp.sendEmail({
          to: item.recipient,
          subject: item.subject,
          htmlBody: item.body,
          name: ctx.name || "TST Admin",
          replyTo: ctx.email
        });
        
        sheet.getRange(item.rowIndex, statusIdx + 1).setValue('Sent');
        sheet.getRange(item.rowIndex, statusIdx + 2).setValue(new Date());
      } catch (err) {
        console.error("Email Send Error", err);
        sheet.getRange(item.rowIndex, statusIdx + 1).setValue('Error: ' + err.message);
      }
    });

    // 2. Automatic Cleanup
    // Remove rows marked 'Sent' or 'Error' that are older than 24 hours
    const now = new Date().getTime();
    const oneDay = 24 * 60 * 60 * 1000;
    
    // We iterate backwards to delete rows without messing up indices
    for (let i = data.length - 1; i >= 1; i--) {
      const status = data[i][statusIdx];
      const timestamp = new Date(data[i][tsIdx]).getTime();
      
      if ((status === 'Sent' || status.toString().startsWith('Error')) && (now - timestamp > oneDay)) {
        sheet.deleteRow(i + 1);
      }
    }

  } catch (e) {
    console.error("Queue Error", e);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Adds an email to the queue for processing.
 */
function addToEmailQueue_(recipient, subject, body, building, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Email Queue');
  
  if (!sheet) {
    sheet = ss.insertSheet('Email Queue');
    sheet.appendRow(['Timestamp', 'Recipient', 'Subject', 'Body', 'Building', 'Status', 'LastUpdated', 'Options']);
    sheet.hideSheet(); // Hide from clutter
  }
  
  sheet.appendRow([
    new Date(),
    recipient,
    subject,
    body,
    building || DEFAULT_BUILDING,
    'Pending',
    '',
    JSON.stringify(options || {})
  ]);
}

/**
 * Escapes a value for HTML text or a double-quoted attribute. Use it on every
 * user-controlled value (names, emails, periods, notes, URL parameters, config
 * values) that goes into email or page HTML.
 */
function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/**
 * Helper to send a styled HTML email.
 * NOW UPDATED to use the Queue System.
 * subject, title, buttonText and buildingName are plain text (escaped here);
 * contentHtml is trusted HTML — callers must escape any data they put in it.
 */
function sendStyledEmail_(recipient, subject, title, contentHtml, buttonText, buildingName, options) {
  const appUrl = ScriptApp.getService().getUrl();
  const headerName = buildingName || 'Orono Schools';
  // Assignment emails point their button at that assignment's signed Record link
  // instead of the app root; every other email keeps the plain app URL.
  const buttonUrl = (options && options.buttonUrl) ? options.buttonUrl : appUrl;
  
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${escapeHtml_(subject)}</title>
      <style>
        body { 
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
          background-color: #f3f4f6; 
          margin: 0; 
          padding: 0; 
          color: #333333;
          -webkit-text-size-adjust: 100%;
          -ms-text-size-adjust: 100%;
        }
        .container { 
          max-width: 600px; 
          margin: 40px auto; 
          background-color: #ffffff; 
          border-radius: 8px;
          overflow: hidden;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .header { 
          background-color: #2d3f89; 
          padding: 30px 20px; 
          text-align: center; 
        }
        .header h1 { 
          color: #ffffff; 
          margin: 0; 
          font-size: 24px; 
          font-weight: 600;
          letter-spacing: 0.5px;
        }
        .content { 
          padding: 40px 30px; 
          line-height: 1.6; 
        }
        .content h2 {
          color: #2d3f89; 
          margin-top: 0;
          margin-bottom: 20px;
          font-size: 22px;
          border-bottom: 2px solid #eaecf5;
          padding-bottom: 10px;
        }
        .button-container {
          text-align: center;
          margin-top: 30px;
          margin-bottom: 10px;
        }
        .button { 
          display: inline-block; 
          background-color: #2d3f89; 
          color: #ffffff !important; 
          padding: 14px 28px; 
          text-decoration: none; 
          border-radius: 6px; 
          font-weight: bold; 
          font-size: 16px;
          transition: background-color 0.3s;
        }
        .button:hover {
          background-color: #1d2a5d;
        }
        .footer { 
          background-color: #f9fafb; 
          padding: 20px; 
          text-align: center; 
          font-size: 12px; 
          color: #6b7280; 
          border-top: 1px solid #e5e7eb;
        }
        @media only screen and (max-width: 600px) {
          .container { margin: 0; border-radius: 0; width: 100%; }
          .content { padding: 20px; }
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>${escapeHtml_(headerName)}</h1>
        </div>
        <div class="content">
          <h2>${escapeHtml_(title)}</h2>
          ${contentHtml}
          <div class="button-container">
            <a href="${escapeHtml_(buttonUrl)}" class="button">${escapeHtml_(buttonText || 'Visit the TST Portal')}</a>
          </div>
        </div>
        <div class="footer">
          In Partnership with Orono Public Schools<br>
          <p style="margin: 5px 0 0 0;">This is an automated message. Please do not reply.</p>
        </div>
      </div>
    </body>
    </html>
  `;
  
  // Use the Config/Building Name to resolve building code for the queue
  // buildingName is passed in as "Orono Middle School", we need 'OMS'.
  // We can try to reverse lookup or just fallback to Default.
  // Actually, 'buildingName' arg in this function is usually the Full Name.
  // The 'options' might contain the code? Or we guess.
  // The Queue needs the CODE to match with Admin's context.
  
  let buildingCode = DEFAULT_BUILDING;
  
  if (options && options.buildingCode) {
    buildingCode = options.buildingCode;
  } else {
    // Fallback: Try to find key by name
    const config = getConfig();
    for (const [code, conf] of Object.entries(config)) {
       if (conf.name === buildingName) {
         buildingCode = code;
         break;
       }
    }
  }

  // Queue the email instead of sending directly
  addToEmailQueue_(recipient, subject, htmlTemplate, buildingCode, options);
}
// --- SCHEDULE / TST AVAILABILITY FEATURE ---

const MONTH_ORDER = ["September", "October", "November", "December", "January", "February", "March", "April", "May", "June"];

// Schedule membership for a building, keyed by lowercased email: staff assigned to
// it (anywhere in a multi-building list) and not archived from it — the same set
// the Directory shows. Values are display names.
function scheduleMembers_(building) {
  const members = new Map();
  staffDirectoryData_(building).forEach(s => {
    members.set(s.email.toString().trim().toLowerCase(), s.name);
  });
  return members;
}

/**
 * The master availability grid. Admins get the whole building (that is the view);
 * a teacher's Schedule tab only ever renders their own rows, so a teacher only gets
 * those — the rest of the building's availability and pending requests stay with
 * the admins.
 *
 * @param {string} buildingFilter - Optional building code to filter by (honored for
 *   Super Admins, or when the caller is assigned to that building)
 */
function getScheduleData(buildingFilter) {
  const ctx = getUserContext();
  return scheduleData_(allowedBuildingFor_(ctx, buildingFilter) || ctx.building,
                       isAdmin_(ctx) ? null : ctx.email);
}

/**
 * Builds the availability grid for one building. Callers are responsible for
 * authorizing the building; onlyEmail (optional) narrows it to one person's rows.
 */
function scheduleData_(effectiveFilter, onlyEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const onlyEmailKey = onlyEmail ? onlyEmail.toString().trim().toLowerCase() : null;

  let schedSheet = ss.getSheetByName('TST Availability');
  if (!schedSheet) {
    // Create if missing
    schedSheet = ss.insertSheet('TST Availability');
    schedSheet.appendRow(['Month', 'Day(s) Available', 'Period', 'Name', 'Email', 'Hours Earned This Month']);
  }

  const data = schedSheet.getDataRange().getValues();
  data.shift(); // Remove header

  // 0. Building membership, matching the Directory: staff assigned to this building
  //    (anywhere in a multi-building list) and not archived from it. Availability
  //    rows aren't building-tagged, so a multi-building person appears in each of
  //    their buildings; rows for anyone else (other buildings, archived, removed
  //    from the directory) are left out.
  const memberEmails = scheduleMembers_(effectiveFilter);

  // 1. Calculate Hours per Teacher per Month
  const hoursMap = calculateMonthlyHours_(); // Returns { "email_Month": hours }, email lowercased

  // 2. Get Pending Requests Map (same building as the schedule)
  const pendingMap = getPendingEarnedMap_(effectiveFilter);

  // 3. Process Schedule Data
  // We return a structured object: { "September": [ { name, email, days, period, hours, pendingRequests }, ... ], ... }
  const schedule = {};
  MONTH_ORDER.forEach(m => schedule[m] = []);

  data.forEach(row => {
    const [month, days, period, name, email] = row;

    // Filter by Building
    const emailKey = (email || '').toString().trim().toLowerCase();
    if (!memberEmails.has(emailKey)) return;
    if (onlyEmailKey && emailKey !== onlyEmailKey) return;

    if (schedule[month]) {
      const hours = hoursMap[`${emailKey}_${month}`] || 0;
      schedule[month].push({
        month, days, period, name, email, hours,
        pendingRequests: pendingMap[emailKey] || []
      });
    }
  });

  return schedule;
}

// Pending earned requests for one building, keyed by lowercased email.
function getPendingEarnedMap_(building) {
  const pendingList = pendingEarnedFor_(building);
  const map = {};

  pendingList.forEach(item => {
    const key = item.email.toString().trim().toLowerCase();
    if (!map[key]) {
      map[key] = [];
    }
    // Minimal data needed for the tooltip/indicator
    map[key].push({
      date: item.date, // Already safeDate string
      subbedFor: item.subbedFor,
      period: item.period
    });
  });
  
  return map;
}

function calculateMonthlyHours_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  const data = sheet.getDataRange().getValues();
  data.shift();

  const sums = {}; // "email_MonthName" -> total (email lowercased)
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  // Determine current school year context
  const today = new Date();
  const currentMonth = today.getMonth(); // 0-11
  const currentYear = today.getFullYear();
  
  // School Year Start Year: If Month >= 7 (Aug), Start = Year. Else Start = Year - 1.
  const startYear = currentMonth >= 7 ? currentYear : currentYear - 1;
  const endYear = startYear + 1;
  
  const schoolYearStart = new Date(startYear, 7, 1); // Aug 1
  const schoolYearEnd = new Date(endYear, 6, 30); // July 30

  data.forEach(row => {
    const email = (row[0] || '').toString().trim().toLowerCase();
    const date = new Date(row[4]);
    const hours = Number(row[7]);
    
    // Check if within current school year
    if (date >= schoolYearStart && date <= schoolYearEnd) {
      const mName = monthNames[date.getMonth()];
      const key = `${email}_${mName}`;
      sums[key] = (sums[key] || 0) + hours;
    }
  });

  return sums;
}

function saveAvailability(month, availabilityList, targetEmail, periodsShown) {
  // availabilityList: [{ days: "Mon,Tue", period: "Period 1" }, ...]
  // targetEmail: set only when an admin saves on a teacher's behalf via "View as".
  // periodsShown: the periods on the form that was saved. Availability rows aren't
  //   building-tagged and buildings name periods differently, so only these periods
  //   are replaced — a multi-building teacher saving one building's grid keeps their
  //   rows for the other building. When omitted, every row for the month is replaced.
  const sessionEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName('Staff Directory');
  const staffData = staffSheet.getDataRange().getValues();

  let userEmail = sessionEmail;
  if (targetEmail && targetEmail.toLowerCase() !== sessionEmail.toLowerCase()) {
    const ctx = getUserContext();
    assertAdmin_(ctx);
    const idx = getStaffIndices_(staffSheet);
    const targetI = findStaffRowByEmail_(staffData, idx.email, targetEmail);
    if (targetI === -1) throw new Error('Staff member not found.');
    assertCanManageRow_(ctx, staffData[targetI][idx.building]);
    userEmail = staffData[targetI][idx.email].toString().trim();
  }

  const userRow = staffData.find(r => r[1].toString().toLowerCase() === userEmail.toLowerCase());
  const userName = userRow ? userRow[0] : userEmail;

  const sheet = ss.getSheetByName('TST Availability');
  const data = sheet.getDataRange().getValues();

  // 1. Remove existing rows for this user + month (limited to the periods shown)
  // We loop backwards to delete
  const periodScope = Array.isArray(periodsShown) ? new Set(periodsShown) : null;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === month && data[i][4].toString().trim().toLowerCase() === userEmail.toLowerCase() &&
        (!periodScope || periodScope.has(data[i][2]))) {
      sheet.deleteRow(i + 1);
    }
  }

  // 2. Add new rows
  // A: Month | B: Day(s) | C: Period | D: Name | E: Email | F: Hours (Ignored/Formula)
  availabilityList.forEach(item => {
    sheet.appendRow([month, item.days, item.period, userName, userEmail, ""]);
  });
}

// ===== TST Coverage Assignments =====
// Coverage is *assigned*, not requested: there is no accept/decline handshake, so a
// teacher who is picked is expected to cover, and a genuine conflict goes to their
// administrator directly.
//
// The assignment row is written the moment an admin creates it, which makes the app
// — not an email sitting in someone's inbox — the record. Everything else hangs off
// that row: the Assignments queue, Cancel / Reassign / Remind, the overnight nudge,
// the duplicate guard, and the signed Record link.
//
// All notification emails go through addToEmailQueue_ so they are sent by the
// building's own admin trigger rather than by whoever deployed the web app.

const ASSIGNMENTS_SHEET_ = 'TST Assignments';
const ASSIGNMENTS_ARCHIVE_SHEET_ = 'TST Assignments Archive';

const ASSIGNMENT_HEADER_ = [
  'ID', 'Created', 'Building', 'Assigned By', 'Assigned By Name',
  'Sub Email', 'Sub Name', 'Covered For', 'Covered For Email',
  'Date', 'Period', 'Time Type', 'Hours',
  'Note', 'Note To Sub', 'Note To Covered',
  'Status', 'Recorded TS', 'Recorded By', 'Nudged TS',
  'Calendar Event ID', 'Calendar Status', 'Notified TS'
];

// 0-based indexes into a row shaped like ASSIGNMENT_HEADER_.
const A_ = {
  id: 0, created: 1, building: 2, byEmail: 3, byName: 4,
  subEmail: 5, subName: 6, coveredFor: 7, coveredForEmail: 8,
  date: 9, period: 10, timeType: 11, hours: 12,
  note: 13, noteToSub: 14, noteToCovered: 15,
  status: 16, recordedTs: 17, recordedBy: 18, nudgedTs: 19,
  calendarEventId: 20, calendarStatus: 21, notifiedTs: 22
};

const ASSIGNMENT_STATUS_ = { assigned: 'Assigned', recorded: 'Recorded', cancelled: 'Cancelled' };

/** Days after the coverage date that the emailed Record link keeps working. */
const ASSIGNMENT_LINK_DAYS_ = 14;

const WEEKDAY_NAMES_ = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const MONTH_NAMES_ = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];

// 'YYYY-MM-DD' -> a local Date at noon, so no timezone offset can roll it a day.
function parseYmd_(value) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(normDateKey_(value));
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function longDate_(value) {
  const d = parseYmd_(value);
  if (!d) return (value == null ? '' : value).toString();
  return WEEKDAY_NAMES_[d.getDay()] + ', ' + MONTH_NAMES_[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}

function shortDate_(value) {
  const d = parseYmd_(value);
  if (!d) return (value == null ? '' : value).toString();
  return WEEKDAY_NAMES_[d.getDay()].slice(0, 3) + ' ' + (d.getMonth() + 1) + '/' + d.getDate();
}

/** Whole days from the coverage date to today. Negative while it is still upcoming. */
function daysSinceDate_(value) {
  const d = parseYmd_(value);
  if (!d) return 0;
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 12, 0, 0);
  return Math.round((today.getTime() - d.getTime()) / 86400000);
}

function buildingNameFor_(building) {
  const config = getConfig();
  const c = config && config[building];
  if (c && c.name) return c.name;
  return building === 'OMS' ? 'Orono Middle School' : 'Orono Schools';
}

/** The web app URL with any query string stripped. */
function scriptUrl_() {
  const url = ScriptApp.getService().getUrl();
  return url ? url.split('?')[0] : '';
}

function assignmentsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ASSIGNMENTS_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(ASSIGNMENTS_SHEET_);
    sheet.appendRow(ASSIGNMENT_HEADER_.slice());
    sheet.setFrozenRows(1);
    return sheet;
  }

  // A sheet created by an earlier version is short a column or two. Fill the
  // header out rather than making someone migrate it by hand; the columns are
  // positional and owned here, so rewriting the header row is safe.
  if (sheet.getLastColumn() < ASSIGNMENT_HEADER_.length) {
    sheet.getRange(1, 1, 1, ASSIGNMENT_HEADER_.length).setValues([ASSIGNMENT_HEADER_.slice()]);
  }
  return sheet;
}

function assignmentFromRow_(row, rowIndex) {
  const str = i => (row[i] == null ? '' : row[i]).toString();
  const bool = i => row[i] === true || row[i] === 'TRUE';
  return {
    rowIndex: rowIndex,
    id: str(A_.id),
    created: row[A_.created] || '',
    building: str(A_.building) || DEFAULT_BUILDING,
    assignedBy: str(A_.byEmail),
    assignedByName: str(A_.byName),
    subEmail: str(A_.subEmail),
    subName: str(A_.subName),
    coveredFor: str(A_.coveredFor),
    coveredForEmail: str(A_.coveredForEmail),
    date: normDateKey_(row[A_.date]),
    period: str(A_.period),
    amountType: str(A_.timeType),
    hours: Number(row[A_.hours]) || 0,
    note: str(A_.note),
    noteToSub: bool(A_.noteToSub),
    noteToCovered: bool(A_.noteToCovered),
    status: str(A_.status) || ASSIGNMENT_STATUS_.assigned,
    recordedTs: row[A_.recordedTs] || '',
    recordedBy: str(A_.recordedBy),
    nudgedTs: row[A_.nudgedTs] || '',
    calendarEventId: str(A_.calendarEventId),
    calendarStatus: str(A_.calendarStatus),
    notifiedTs: row[A_.notifiedTs] || ''
  };
}

function assignmentRows_(filterFn) {
  const data = assignmentsSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][A_.id]) continue;
    const item = assignmentFromRow_(data[i], i + 1);
    if (!filterFn || filterFn(item)) out.push(item);
  }
  return out;
}

function findAssignment_(id) {
  const key = (id == null ? '' : id).toString().trim();
  if (!key) return null;
  return assignmentRows_(a => a.id === key)[0] || null;
}

/** updates: { <0-based column index>: value }. */
function setAssignmentCells_(sheet, rowIndex, updates) {
  Object.keys(updates).forEach(col => {
    sheet.getRange(rowIndex, Number(col) + 1).setValue(updates[col]);
  });
}

/** Badge material: coverage that has already happened with nothing recorded. */
function isAssignmentOutstanding_(a) {
  return a.status === ASSIGNMENT_STATUS_.assigned && daysSinceDate_(a.date) > 0;
}

function assertCanManageAssignment_(ctx, a) {
  assertAdmin_(ctx);
  if (ctx.isSuperAdmin) return;
  if (!ctx.buildings.includes(a.building)) {
    throw new Error('You can only manage assignments for your own building(s).');
  }
}

// ===== Signed Record links =====
// The link in an assignment email is served by doGet with no other credential, so
// its parameters are HMAC-signed with a per-project secret in Script Properties.
// The signature covers only the assignment id and the invited teacher; every other
// detail is read from the assignment row, so none of it can be forged from a URL.

const ASSIGNMENT_LINK_FIELDS_ = ['action', 'id', 'tEmail'];

function coverageLinkSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('COVERAGE_LINK_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('COVERAGE_LINK_SECRET', secret);
  }
  return secret;
}

function assignmentSignature_(params) {
  const canonical = JSON.stringify(ASSIGNMENT_LINK_FIELDS_.map(k => params[k] == null ? '' : String(params[k])));
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(canonical, coverageLinkSecret_()));
}

function buildAssignmentLink_(baseUrl, params) {
  const query = ASSIGNMENT_LINK_FIELDS_
    .filter(k => params[k] != null && params[k] !== '')
    .map(k => k + '=' + encodeURIComponent(params[k]));
  query.push('sig=' + encodeURIComponent(assignmentSignature_(params)));
  return baseUrl + '?' + query.join('&');
}

function verifyAssignmentLink_(params) {
  if (!params || !params.sig) return false;
  const given = String(params.sig);
  const expected = assignmentSignature_(params);
  if (given.length !== expected.length) return false;
  let diff = 0; // constant-time compare
  for (let i = 0; i < given.length; i++) diff |= given.charCodeAt(i) ^ expected.charCodeAt(i);
  return diff === 0;
}

function assignmentMessagePage_(title, message) {
  return HtmlService.createHtmlOutput(
    '<div style="font-family: sans-serif; text-align: center; padding: 50px;">' +
    '<h1 style="color: #ad2122;">' + escapeHtml_(title) + '</h1>' +
    '<p>' + escapeHtml_(message) + '</p>' +
    '</div>'
  ).setTitle(title);
}

// ===== Assignment emails =====
// Every value below comes from the assignment row, which is user-controlled, so
// each one is escaped: sendStyledEmail_ escapes only subject/title/buttonText/
// buildingName, never contentHtml.

/** pairs: [['Date', '...'], ...]. Blank values are dropped. */
function detailsBox_(pairs) {
  const rows = pairs
    .filter(p => p && p[1] !== '' && p[1] != null)
    .map(p => `<p style="margin: 5px 0;"><strong>${escapeHtml_(p[0])}:</strong> ${escapeHtml_(p[1])}</p>`)
    .join('');
  if (!rows) return '';
  return `<div style="background-color: #f3f4f6; padding: 15px; border-radius: 6px; margin: 20px 0;">${rows}</div>`;
}

function noteBlock_(note) {
  if (!note) return '';
  return `<div style="background-color: #eff6ff; border-left: 4px solid #2d3f89; padding: 12px 15px; margin: 20px 0;">
      <p style="margin: 0; color: #2d3f89; font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em; font-weight: bold;">Note from your administrator</p>
      <p style="margin: 6px 0 0 0; color: #334155;">${escapeHtml_(note)}</p>
    </div>`;
}

/** "Full Period (1.0 hr)" — the label staff picked, with the hours it is worth. */
function durationLabel_(a) {
  const hrs = a.hours === 1 ? '1 hr' : a.hours + ' hrs';
  return a.amountType ? a.amountType + ' (' + hrs + ')' : hrs;
}

function assignmentSubject_(a, prefix) {
  return prefix + ': ' + shortDate_(a.date) + ' — ' + a.period;
}

/**
 * The three emails an assignment sends: the person covering (the only one with an
 * action), the person being covered for (when a real staff member was picked), and
 * the administrator's own record copy.
 */
function sendAssignmentEmails_(a) {
  const buildingName = buildingNameFor_(a.building);
  const opts = { buildingCode: a.building };
  const adminName = a.assignedByName || a.assignedBy;
  const dateLong = longDate_(a.date);
  const periodText = periodDisplay_(a.building, a.period, a.date);

  const recordUrl = buildAssignmentLink_(scriptUrl_(), {
    action: 'record', id: a.id, tEmail: a.subEmail
  });

  const subBody =
    `<p>Hello <strong>${escapeHtml_(a.subName)}</strong>,</p>` +
    `<p>You have been assigned for TST Coverage for <strong>${escapeHtml_(a.coveredFor)}</strong> on ` +
    `<strong>${escapeHtml_(dateLong)}</strong>, <strong>${escapeHtml_(periodText)}</strong>.</p>` +
    detailsBox_([
      ['Date', dateLong],
      ['Period', periodText],
      ['Covering For', a.coveredFor],
      ['Duration', durationLabel_(a)]
    ]) +
    (a.noteToSub ? noteBlock_(a.note) : '') +
    assignmentCalendarLine_(a) +
    `<p style="font-size: 13px; color: #6b7280;">If you have a conflict, contact ${escapeHtml_(adminName)} directly.</p>`;

  sendStyledEmail_(
    a.subEmail,
    assignmentSubject_(a, 'TST Coverage Assignment'),
    'TST Coverage Assignment',
    subBody,
    'Record My TST Time',
    buildingName,
    Object.assign({ buttonUrl: recordUrl }, opts)
  );

  // The person being covered for. Free-text entries ("Activity Bus") have no email.
  if (a.coveredForEmail) {
    const coveredBody =
      `<p>Hello <strong>${escapeHtml_(a.coveredFor)}</strong>,</p>` +
      `<p><strong>${escapeHtml_(a.subName)}</strong> has been assigned to cover your ` +
      `<strong>${escapeHtml_(periodText)}</strong> class on <strong>${escapeHtml_(dateLong)}</strong>.</p>` +
      detailsBox_([
        ['Date', dateLong],
        ['Period', periodText],
        ['Covered By', a.subName]
      ]) +
      (a.noteToCovered ? noteBlock_(a.note) : '') +
      assignmentCalendarLine_(a);

    sendStyledEmail_(
      a.coveredForEmail,
      assignmentSubject_(a, 'TST Coverage Arranged'),
      'TST Coverage Arranged',
      coveredBody,
      'View My TST',
      buildingName,
      opts
    );
  }

  // The administrator's record copy.
  const adminBody =
    `<p>You assigned <strong>${escapeHtml_(a.subName)}</strong> to cover for <strong>${escapeHtml_(a.coveredFor)}</strong>.</p>` +
    detailsBox_([
      ['Date', dateLong],
      ['Period', periodText],
      ['Duration', durationLabel_(a)]
    ]) +
    (a.note ? noteBlock_(a.note) : '') +
    `<p>${escapeHtml_(a.coveredForEmail ? 'Both staff have been emailed.' : a.subName + ' has been emailed.')} ` +
    `You will be notified when ${escapeHtml_(a.subName)} records their TST time.</p>`;

  sendStyledEmail_(
    a.assignedBy,
    'TST Assignment: ' + shortDate_(a.date) + ' — ' + a.subName + ' — ' + a.period,
    'Coverage Assigned',
    adminBody,
    'View Dashboard',
    buildingName,
    opts
  );
}

/**
 * The "added to the X TST Calendar" sentence — only once the event really exists.
 * A building with no calendar, or one whose event failed, never claims an entry
 * that is not there.
 */
function assignmentCalendarLine_(a) {
  if (!a || a.calendarStatus !== CAL_.created) return '';
  return '<p style="font-size: 13px; color: #6b7280;">This has been added to the ' +
    escapeHtml_(calendarNameFor_(a.building)) + '.</p>';
}

function sendAssignmentReminder_(a) {
  const buildingName = buildingNameFor_(a.building);
  const recordUrl = buildAssignmentLink_(scriptUrl_(), {
    action: 'record', id: a.id, tEmail: a.subEmail
  });
  const expiry = parseYmd_(a.date);
  const expiryText = expiry
    ? longDate_(normDateKey_(new Date(expiry.getTime() + ASSIGNMENT_LINK_DAYS_ * 86400000)))
    : '';

  const body =
    `<p>Hello <strong>${escapeHtml_(a.subName)}</strong>,</p>` +
    `<p>You were assigned to cover <strong>${escapeHtml_(periodDisplay_(a.building, a.period, a.date))}</strong> for ` +
    `<strong>${escapeHtml_(a.coveredFor)}</strong> on <strong>${escapeHtml_(longDate_(a.date))}</strong>, ` +
    `but your TST time has not been recorded yet.</p>` +
    (expiryText
      ? `<p style="font-size: 13px; color: #6b7280;">This link expires on ${escapeHtml_(expiryText)}. ` +
        `After that, contact ${escapeHtml_(a.assignedByName || a.assignedBy)}.</p>`
      : '');

  sendStyledEmail_(
    a.subEmail,
    'Reminder: Record your TST time — ' + shortDate_(a.date) + ', ' + a.period,
    'Record Your TST Time',
    body,
    'Record My TST Time',
    buildingName,
    { buildingCode: a.building, buttonUrl: recordUrl }
  );
}

function sendAssignmentCancelledEmails_(a) {
  const buildingName = buildingNameFor_(a.building);
  const opts = { buildingCode: a.building };
  const dateLong = longDate_(a.date);
  const subject = assignmentSubject_(a, 'TST Coverage Cancelled');

  const subBody =
    `<p>Hello <strong>${escapeHtml_(a.subName)}</strong>,</p>` +
    `<p>The TST Coverage assignment for <strong>${escapeHtml_(a.coveredFor)}</strong> on ` +
    `<strong>${escapeHtml_(dateLong)}</strong>, <strong>${escapeHtml_(periodDisplay_(a.building, a.period, a.date))}</strong> has been cancelled. ` +
    `You do not need to cover this class.</p>` +
    assignmentCalendarRemovedLine_(a);

  sendStyledEmail_(a.subEmail, subject, 'Coverage Cancelled', subBody, 'View My TST', buildingName, opts);

  if (a.coveredForEmail) {
    const coveredBody =
      `<p>Hello <strong>${escapeHtml_(a.coveredFor)}</strong>,</p>` +
      `<p>The coverage arranged for your <strong>${escapeHtml_(periodDisplay_(a.building, a.period, a.date))}</strong> class on ` +
      `<strong>${escapeHtml_(dateLong)}</strong> has been cancelled. ` +
      `<strong>${escapeHtml_(a.subName)}</strong> is no longer assigned.</p>` +
      assignmentCalendarRemovedLine_(a);

    sendStyledEmail_(a.coveredForEmail, subject, 'Coverage Cancelled', coveredBody, 'View My TST', buildingName, opts);
  }
}

function assignmentCalendarRemovedLine_(a) {
  if (!a || a.calendarStatus !== CAL_.deleted) return '';
  return '<p style="font-size: 13px; color: #6b7280;">It has been removed from the ' +
    escapeHtml_(calendarNameFor_(a.building)) + '.</p>';
}

/**
 * Copies a building's bell schedule out of config.js into the live App Config
 * sheet. Admin only, and scoped like every other config write.
 *
 * getConfig() seeds App Config from BUILDING_CONFIG only when the sheet does not
 * exist yet, so editing config.js does nothing to a district that is already
 * running. Rather than have someone retype twenty-odd start and end times into
 * Settings, this merges just the schedule keys across and leaves everything the
 * building has set for itself — calendar, carry-over cap, name — untouched.
 *
 * Run it once from the Apps Script editor after deploying a schedule change.
 */
function installBellSchedule(building) {
  const ctx = getUserContext();
  assertAdmin_(ctx);

  // Name the building explicitly. Every other building-scoped call falls back to
  // the caller's own when asked for nothing, which is right for a read — but this
  // one REPLACES a building's periods, and the Apps Script editor's Run button
  // passes no arguments at all. Defaulting here would quietly rewrite the wrong
  // school's schedule.
  if (!building) {
    throw new Error("Name the building, e.g. installBellSchedule('OHS').");
  }

  const target = allowedBuildingFor_(ctx, building);
  if (!target) throw new Error('You can only install a bell schedule for your own building(s).');

  const source = (typeof BUILDING_CONFIG !== 'undefined' && BUILDING_CONFIG[target]) || null;
  if (!source || !Array.isArray(source.periods)) {
    throw new Error('config.js has no period list for ' + target + '.');
  }

  const live = Object.assign({}, (getConfig() || {})[target] || {});
  live.periods = source.periods.slice();
  live.periodTimes = Object.assign({}, source.periodTimes || {});
  live.dayGroups = JSON.parse(JSON.stringify(source.dayGroups || []));

  saveBuildingConfig(target, live);

  const summary = target + ': ' + live.periods.length + ' periods, ' +
    Object.keys(live.periodTimes).length + ' with default times, ' +
    live.dayGroups.length + ' day schedule(s).';
  Logger.log(summary);
  return summary;
}

// ===== Assignment calendar =====
// Each building has its own TST calendar, and the event is created by that
// building's own admin through their trigger — so they own the event, matching
// the address the emails go out from. A blank calendar ID means the building has
// no calendar and the whole path is skipped: no event, no calendar sentence in any
// email, no failure alerts.
//
// Order matters. The event is created first and the emails go out afterwards, so
// the "added to the ... TST Calendar" line only ever appears when there really is
// something to look at. A failure never blocks the assignment — the emails still
// go, minus that sentence, and the admin is told.

const CAL_ = {
  none: '',
  pending: 'Pending',
  created: 'Created',
  failed: 'Failed',
  pendingDelete: 'Pending Delete',
  deleted: 'Deleted'
};

const CAL_TEST_PREFIX_ = 'CAL_TEST_';

function calendarIdFor_(building) {
  const cfg = (getConfig() || {})[building] || {};
  return (cfg.calendarId || '').toString().trim();
}

/** The name staff will see in their own calendar list, as the admin typed it. */
function calendarNameFor_(building) {
  const cfg = (getConfig() || {})[building] || {};
  const typed = (cfg.calendarName || '').toString().trim();
  return typed || (buildingNameFor_(building) + ' TST Calendar');
}

function isCalendarFailure_(status) {
  return (status || '').toString().indexOf(CAL_.failed) === 0;
}

/** "Period 3 - 9:52 - 10:39" reads as "Period 3" in a calendar title. */
function shortPeriodLabel_(period) {
  const p = (period == null ? '' : period).toString().trim();
  const stripped = p
    .replace(/\s*[-–—]?\s*\d{1,2}:\d{2}\s*(?:[-–—]|to)\s*\d{1,2}:\d{2}\s*$/, '')
    .trim();
  return stripped || p;
}

/** The real start/end Dates for an assignment, or null when times are unknown. */
function assignmentEventWindow_(a) {
  const times = periodTimesFor_(a.building, a.period, a.date);
  const day = parseYmd_(a.date);
  if (!times || !day) return null;

  const at = hhmm => {
    const parts = hhmm.split(':');
    return new Date(day.getFullYear(), day.getMonth(), day.getDate(), Number(parts[0]), Number(parts[1]), 0);
  };
  const start = at(times.start);
  const end = at(times.end);
  if (!(end.getTime() > start.getTime())) return null;
  return { start: start, end: end };
}

function assignmentEventTitle_(a) {
  const period = shortPeriodLabel_(a.period);
  return 'TST: ' + a.subName + ' covering ' + a.coveredFor + (period ? ' — ' + period : '');
}

function assignmentEventDescription_(a) {
  const lines = [
    a.subName + ' is covering ' + periodDisplay_(a.building, a.period, a.date) + ' for ' + a.coveredFor + '.',
    '',
    'Date: ' + longDate_(a.date),
    'Duration: ' + durationLabel_(a),
    'Assigned by: ' + (a.assignedByName || a.assignedBy)
  ];
  if (a.note) lines.push('', 'Note: ' + a.note);
  const url = scriptUrl_();
  if (url) lines.push('', 'TST Manager: ' + url);
  return lines.join('\n');
}

/**
 * Creates the event for one assignment, then sends its emails.
 *
 * Guests are added without a Google invite: our own email is the notification,
 * and Google's carries a Yes/No/Maybe prompt, which would put a decline button
 * back in a flow that deliberately has none.
 */
function finishAssignmentCreation_(sheet, a) {
  const calendarId = calendarIdFor_(a.building);
  let status = CAL_.none;
  let eventId = '';
  let failure = '';

  if (calendarId) {
    const when = assignmentEventWindow_(a);
    if (!when) {
      failure = 'No start and end time is set for "' + a.period + '" on that day. ' +
        'Add it under Settings → Periods.';
    } else {
      try {
        const calendar = CalendarApp.getCalendarById(calendarId);
        if (!calendar) {
          failure = 'Calendar not found, or this account cannot edit it (' + calendarId + ').';
        } else {
          const guests = [a.subEmail, a.coveredForEmail].filter(Boolean).join(',');
          const event = calendar.createEvent(assignmentEventTitle_(a), when.start, when.end, {
            description: assignmentEventDescription_(a),
            guests: guests,
            sendInvites: false
          });
          eventId = event.getId();
          status = CAL_.created;
        }
      } catch (err) {
        failure = (err && err.message) ? err.message : String(err);
      }
    }
  }

  if (failure) status = CAL_.failed + ': ' + failure;

  const updates = {};
  updates[A_.calendarEventId] = eventId;
  updates[A_.calendarStatus] = status;
  updates[A_.notifiedTs] = new Date();
  setAssignmentCells_(sheet, a.rowIndex, updates);

  const updated = Object.assign({}, a, { calendarEventId: eventId, calendarStatus: status });
  sendAssignmentEmails_(updated);
  if (failure) alertCalendarFailure_(updated, failure);
  return updated;
}

/** Removes a cancelled assignment's event, then sends the cancellation emails. */
function finishAssignmentCancellation_(sheet, a) {
  const calendarId = calendarIdFor_(a.building);
  let status = CAL_.deleted;
  let failure = '';

  if (calendarId && a.calendarEventId) {
    try {
      const calendar = CalendarApp.getCalendarById(calendarId);
      const event = calendar && calendar.getEventById(a.calendarEventId);
      if (event) event.deleteEvent();
    } catch (err) {
      failure = (err && err.message) ? err.message : String(err);
      status = CAL_.failed + ': ' + failure;
    }
  }

  const updates = {};
  updates[A_.calendarStatus] = status;
  updates[A_.calendarEventId] = '';
  setAssignmentCells_(sheet, a.rowIndex, updates);

  const updated = Object.assign({}, a, { calendarStatus: status, calendarEventId: '' });
  // Nobody was told about this assignment yet, so there is nothing to cancel on
  // their side — a cancellation email would be the first they ever heard of it.
  if (a.notifiedTs) sendAssignmentCancelledEmails_(updated);
  if (failure) alertCalendarFailure_(updated, 'The calendar event could not be removed: ' + failure);
  return updated;
}

function alertCalendarFailure_(a, reason) {
  const body =
    '<p>The coverage assignment for <strong>' + escapeHtml_(a.subName) + '</strong> on <strong>' +
    escapeHtml_(longDate_(a.date)) + '</strong> went out, but its calendar entry did not.</p>' +
    detailsBox_([
      ['Reason', reason],
      ['Calendar', calendarIdFor_(a.building) || '(none set)'],
      ['Period', a.period]
    ]) +
    '<p>Everyone was still emailed — only the calendar entry is missing. Check the calendar ID ' +
    'in Settings, and that you can edit that calendar.</p>';

  sendStyledEmail_(
    a.assignedBy,
    'TST calendar entry failed: ' + shortDate_(a.date) + ' — ' + a.subName,
    'Calendar Entry Failed',
    body,
    'Open TST Manager',
    buildingNameFor_(a.building),
    { buildingCode: a.building }
  );
}

/**
 * The calendar work the building admin's trigger does, for their buildings only.
 *
 * Has to be driven by a trigger rather than the web app: the app runs as whoever
 * deployed it, so an event it created would be owned by the deployer rather than
 * by the building admin whose calendar it is.
 */
function processPendingAssignments_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return 0; // another admin's trigger is already on it

  try {
    const ctx = getUserContext();
    const mine = ctx.buildings || [DEFAULT_BUILDING];
    const sheet = assignmentsSheet_();

    const due = assignmentRows_(a =>
      (a.calendarStatus === CAL_.pending || a.calendarStatus === CAL_.pendingDelete) &&
      mine.indexOf(a.building) > -1);

    due.forEach(a => {
      try {
        if (a.calendarStatus === CAL_.pendingDelete) finishAssignmentCancellation_(sheet, a);
        else finishAssignmentCreation_(sheet, a);
      } catch (err) {
        console.error('Assignment calendar job failed', a.id, err);
      }
    });

    runCalendarTests_(ctx);
    return due.length;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Settings' "Send test event" runs the real path end to end — queued here, picked
 * up by the building admin's trigger, creating and removing a throwaway event.
 * That is the only check that proves the calendar ID, the admin's edit rights and
 * their trigger all at once; validating the ID from Settings would run as the
 * deployer and prove none of it.
 */
function runCalendarTests_(ctx) {
  const props = PropertiesService.getScriptProperties();

  (ctx.buildings || []).forEach(building => {
    const raw = props.getProperty(CAL_TEST_PREFIX_ + building);
    if (!raw) return;

    let job = null;
    try { job = JSON.parse(raw); } catch (e) { return; }
    if (!job || job.state !== 'pending') return;

    let result;
    try {
      const calendarId = calendarIdFor_(building);
      if (!calendarId) throw new Error('No calendar ID is set for ' + building + '.');

      const calendar = CalendarApp.getCalendarById(calendarId);
      if (!calendar) throw new Error('Calendar not found, or this account cannot edit it.');

      const start = new Date(new Date().getTime() + 24 * 60 * 60 * 1000);
      const event = calendar.createEvent(
        'TST Manager test event',
        start,
        new Date(start.getTime() + 15 * 60 * 1000),
        { description: 'Created by TST Manager to check the calendar connection. It removes itself.' }
      );
      const actualName = calendar.getName ? calendar.getName() : '';
      event.deleteEvent();

      result = { state: 'ok', building: building, calendarName: actualName, by: ctx.email };
    } catch (err) {
      result = {
        state: 'failed',
        building: building,
        error: (err && err.message) ? err.message : String(err),
        by: ctx.email
      };
    }

    props.setProperty(CAL_TEST_PREFIX_ + building, JSON.stringify(result));
  });
}

/** Admin action: queue the end-to-end calendar check for a building. */
function sendTestCalendarEvent(building) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const b = allowedBuildingFor_(ctx, building);
  if (!b) throw new Error('You can only test your own building(s).');
  if (!calendarIdFor_(b)) throw new Error('No calendar ID is set for ' + b + '. Add one and save first.');

  PropertiesService.getScriptProperties()
    .setProperty(CAL_TEST_PREFIX_ + b, JSON.stringify({ state: 'pending', building: b }));
  return true;
}

/** Polled by Settings while a test is in flight. */
function getCalendarTestResult(building) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const b = allowedBuildingFor_(ctx, building);
  if (!b) throw new Error('You can only test your own building(s).');

  const raw = PropertiesService.getScriptProperties().getProperty(CAL_TEST_PREFIX_ + b);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

// ===== Assignment endpoints =====

/**
 * Admin action: assign coverage. Writes the assignment row, then queues the
 * notification emails.
 *
 * payload: { teacherEmail, teacherName, subbedFor, coveredForEmail, date, period,
 *            amount, amountType, building, note, noteToSub, noteToCovered, force }
 *
 * Exact duplicates (same person, date and period) are always a mis-click and are
 * refused. A different period on the same date is legitimate, so it comes back as
 * { conflict: true, existing: [...] } — the same block-with-override shape
 * finalizeSchoolYear uses — and the client re-sends with force: true to confirm.
 */
function assignCoverage(payload) {
  const ctx = getUserContext();
  const p = payload || {};

  const subEmail = (p.teacherEmail || '').toString().trim();
  const coveredForEmail = (p.coveredForEmail || '').toString().trim();
  const coveredForName = (p.subbedFor || '').toString().trim();

  // Admin, and only for staff they manage — the person covering, plus the person
  // being covered when a real staff member was picked rather than free text.
  const managed = [subEmail];
  if (coveredForEmail) managed.push(coveredForEmail);
  assertCanManageStaffEmails_(ctx, managed);

  const building = allowedBuildingFor_(ctx, p.building);
  if (!building) throw new Error('You can only assign coverage for your own building(s).');

  const date = normDateKey_(p.date);
  if (!parseYmd_(date)) throw new Error('A valid coverage date is required.');
  const period = (p.period || '').toString().trim();
  if (!period) throw new Error('A period is required.');
  if (!coveredForName) throw new Error('Please say who needs coverage.');

  const hours = Number(p.amount);
  if (!isFinite(hours) || hours <= 0) throw new Error('Coverage duration must be greater than zero.');

  if (coveredForEmail && coveredForEmail.toLowerCase() === subEmail.toLowerCase()) {
    throw new Error('A staff member cannot be assigned to cover for themselves.');
  }

  const sameDay = assignmentRows_(a =>
    a.status !== ASSIGNMENT_STATUS_.cancelled &&
    a.subEmail.toLowerCase() === subEmail.toLowerCase() &&
    a.date === date);

  const subName = (p.teacherName || subEmail).toString();
  if (sameDay.some(a => a.period === period)) {
    throw new Error(subName + ' is already assigned to ' + period + ' on ' + shortDate_(date) + '.');
  }
  if (sameDay.length > 0 && !p.force) {
    return {
      conflict: true,
      name: subName,
      date: shortDate_(date),
      existing: sameDay.map(a => ({ period: a.period, coveredFor: a.coveredFor }))
    };
  }

  const sheet = assignmentsSheet_();
  const row = new Array(ASSIGNMENT_HEADER_.length).fill('');
  row[A_.id] = Utilities.getUuid();
  row[A_.created] = new Date();
  row[A_.building] = building;
  row[A_.byEmail] = ctx.email;
  row[A_.byName] = ctx.name || ctx.email;
  row[A_.subEmail] = subEmail;
  row[A_.subName] = subName;
  row[A_.coveredFor] = coveredForName;
  row[A_.coveredForEmail] = coveredForEmail;
  row[A_.date] = date;
  row[A_.period] = period;
  row[A_.timeType] = (p.amountType || '').toString();
  row[A_.hours] = hours;
  row[A_.note] = (p.note || '').toString().trim();
  row[A_.noteToSub] = !!p.noteToSub && !!row[A_.note];
  row[A_.noteToCovered] = !!p.noteToCovered && !!row[A_.note] && !!coveredForEmail;
  row[A_.status] = ASSIGNMENT_STATUS_.assigned;

  // With a calendar, the event is built first and the emails follow, so the
  // "added to the calendar" line is only ever written when it is true. That work
  // belongs to the building admin's trigger, because the web app runs as the
  // deployer and would own the event itself. Without a calendar there is nothing
  // to wait for, so the emails go now.
  const usesCalendar = !!calendarIdFor_(building);
  row[A_.calendarStatus] = usesCalendar ? CAL_.pending : CAL_.none;
  if (!usesCalendar) row[A_.notifiedTs] = new Date();

  sheet.appendRow(row);

  const assignment = assignmentFromRow_(row, sheet.getLastRow());
  if (!usesCalendar) sendAssignmentEmails_(assignment);
  return { assigned: true, id: assignment.id, pendingCalendar: usesCalendar };
}

/** The admin Assignments queue, scoped the same way every other queue is. */
function getAssignments(buildingFilter) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const building = allowedBuildingFor_(ctx, buildingFilter) || ctx.building;
  return assignmentsFor_(building);
}

function assignmentsFor_(building) {
  return assignmentRows_(a => a.building === building)
    .map(a => Object.assign({}, a, {
      dateDisplay: longDate_(a.date),
      periodDisplay: periodDisplay_(a.building, a.period, a.date),
      durationLabel: durationLabel_(a),
      outstanding: isAssignmentOutstanding_(a),
      upcoming: daysSinceDate_(a.date) <= 0
    }))
    .sort((x, y) => (y.date || '').localeCompare(x.date || ''));
}

/**
 * A teacher's own assignments — both the ones they are covering and the coverage
 * arranged for their classes. Same self-or-manager rule as getTeacherHistory, so
 * it works under View As.
 */
function getMyAssignments(targetEmail) {
  const email = (targetEmail || Session.getActiveUser().getEmail() || '').toString().trim();
  assertSelfOrManagerOf_(email);
  const lower = email.toLowerCase();

  return assignmentRows_(a =>
    a.status !== ASSIGNMENT_STATUS_.cancelled &&
    (a.subEmail.toLowerCase() === lower || a.coveredForEmail.toLowerCase() === lower))
    .map(a => {
      const covering = a.subEmail.toLowerCase() === lower;
      return {
        id: a.id,
        building: a.building,
        date: a.date,
        dateDisplay: longDate_(a.date),
        period: a.period,
        periodDisplay: periodDisplay_(a.building, a.period, a.date),
        amountType: a.amountType,
        hours: a.hours,
        durationLabel: durationLabel_(a),
        subName: a.subName,
        coveredFor: a.coveredFor,
        status: a.status,
        role: covering ? 'covering' : 'covered',
        note: (covering ? a.noteToSub : a.noteToCovered) ? a.note : ''
      };
    })
    .sort((x, y) => (x.date || '').localeCompare(y.date || ''));
}

/**
 * Files the earned request for an assignment. Used by the teacher's dashboard
 * button, by an admin recording on someone's behalf, and by the emailed link.
 *
 * The row is claimed before the earned request is written — the same way the email
 * queue claims a row — so a double-click or a re-opened link cannot submit twice.
 */
function recordAssignment_(id, by) {
  const sheet = assignmentsSheet_();
  const a = findAssignment_(id);
  if (!a) throw new Error('Assignment not found.');
  if (a.status === ASSIGNMENT_STATUS_.cancelled) throw new Error('This assignment was cancelled.');
  if (a.status === ASSIGNMENT_STATUS_.recorded) return a;

  setAssignmentCells_(sheet, a.rowIndex, {
    [A_.status]: ASSIGNMENT_STATUS_.recorded,
    [A_.recordedTs]: new Date(),
    [A_.recordedBy]: by
  });
  SpreadsheetApp.flush();

  try {
    submitEarned_({
      email: a.subEmail,
      subbedForType: a.coveredForEmail ? 'Staff' : 'Other',
      subbedForName: a.coveredFor,
      date: a.date,
      period: a.period,
      amountType: a.amountType,
      amountDecimal: a.hours,
      building: a.building
    });
  } catch (err) {
    setAssignmentCells_(sheet, a.rowIndex, {
      [A_.status]: ASSIGNMENT_STATUS_.assigned,
      [A_.recordedTs]: '',
      [A_.recordedBy]: ''
    });
    throw err;
  }

  return findAssignment_(id) || a;
}

/**
 * Records an assignment from inside the app. The teacher may record their own; an
 * admin may record on behalf of someone they manage, which is the recourse once the
 * emailed link has expired. Unlike that link, this has no 14-day limit: the link is
 * a bearer token in an inbox, this is an authenticated action on a known row.
 */
function recordAssignment(id, targetEmail) {
  const a = findAssignment_(id);
  if (!a) throw new Error('Assignment not found.');

  const target = (targetEmail || a.subEmail).toString().trim();
  assertSelfOrManagerOf_(target);
  if (target.toLowerCase() !== a.subEmail.toLowerCase()) {
    throw new Error('That assignment belongs to another staff member.');
  }

  const session = (Session.getActiveUser().getEmail() || '').toString().trim().toLowerCase();
  const by = session === a.subEmail.toLowerCase() ? 'teacher' : 'admin:' + session;
  recordAssignment_(id, by);
  return true;
}

/**
 * Cancels an assignment: removes the pending earned request if the teacher already
 * recorded it, marks the row cancelled, and emails both staff. Hours that have
 * already been approved are never clawed back silently — the caller is sent to the
 * existing Revert flow instead.
 */
function cancelAssignment(id) {
  const ctx = getUserContext();
  const a = findAssignment_(id);
  if (!a) throw new Error('Assignment not found.');
  assertCanManageAssignment_(ctx, a);
  if (a.status === ASSIGNMENT_STATUS_.cancelled) return { cancelled: true, already: true };

  if (a.status === ASSIGNMENT_STATUS_.recorded) {
    const earned = findEarnedRowForAssignment_(a);
    if (earned && earned.approved) {
      throw new Error(a.subName + "'s hours for this coverage are already approved. " +
        'Revert that request first, then cancel the assignment.');
    }
    if (earned) deleteEarnedRow_(earned.rowIndex);
  }

  // An event can only be removed by the admin who owns it, so hand that to their
  // trigger and let it send the cancellation emails once the removal is done —
  // otherwise the email claims a removal that has not happened yet.
  const hasCalendarWork = !!a.calendarEventId || a.calendarStatus === CAL_.pending;

  const updates = {};
  updates[A_.status] = ASSIGNMENT_STATUS_.cancelled;
  if (hasCalendarWork) updates[A_.calendarStatus] = CAL_.pendingDelete;
  setAssignmentCells_(assignmentsSheet_(), a.rowIndex, updates);

  if (hasCalendarWork) return { cancelled: true, pendingCalendar: true };

  // Nobody was told about this one yet, so a cancellation would be the first they
  // ever heard of it.
  if (a.notifiedTs) sendAssignmentCancelledEmails_(a);
  return { cancelled: true };
}

/** The earned request an assignment produced, matched the way submissions are deduped. */
function findEarnedRowForAssignment_(a) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TST Approvals (New)');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const email = a.subEmail.toLowerCase();
  const period = a.period.trim();

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (r[0].toString().trim().toLowerCase() !== email) continue;
    if (normDateKey_(r[4]) !== a.date) continue;
    if ((r[5] || '').toString().trim() !== period) continue;
    if (r[10] === true || r[10] === 'TRUE') continue; // denied rows are not this one
    return { rowIndex: i + 1, approved: r[8] === true || r[8] === 'TRUE' };
  }
  return null;
}

/** Admin action: re-send the Record email for an assignment nobody has acted on. */
function remindAssignment(id) {
  const ctx = getUserContext();
  const a = findAssignment_(id);
  if (!a) throw new Error('Assignment not found.');
  assertCanManageAssignment_(ctx, a);
  if (a.status !== ASSIGNMENT_STATUS_.assigned) {
    throw new Error('This assignment is already ' + a.status.toLowerCase() + '.');
  }
  sendAssignmentReminder_(a);
  setAssignmentCells_(assignmentsSheet_(), a.rowIndex, { [A_.nudgedTs]: new Date() });
  return true;
}

/**
 * The one automatic nudge, run by the daily trigger setupEmailService installs.
 *
 * Has to stay public for that trigger, which also makes it reachable from
 * google.script.run — so a client call must be an admin. It only queues mail;
 * each reminder is tagged with its own building, so the building's admin trigger
 * is still what sends it.
 */
function nudgeOutstandingAssignments(e) {
  if (!isTriggerEvent_(e)) assertAdmin_(getUserContext());
  return nudgeOutstandingAssignments_();
}

function nudgeOutstandingAssignments_() {
  const sheet = assignmentsSheet_();
  const due = assignmentRows_(a => {
    if (a.status !== ASSIGNMENT_STATUS_.assigned || a.nudgedTs) return false;
    const age = daysSinceDate_(a.date);
    return age >= 1 && age <= ASSIGNMENT_LINK_DAYS_;
  });

  due.forEach(a => {
    sendAssignmentReminder_(a);
    setAssignmentCells_(sheet, a.rowIndex, { [A_.nudgedTs]: new Date() });
  });
  return due.length;
}

/** Year-end: assignment rows move to the archive with the rest of the year. */
function archiveAssignmentsForBuilding_(building, yearName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(ASSIGNMENTS_SHEET_);
  if (!src) return;
  const data = src.getDataRange().getValues();
  if (data.length < 2) return;

  let arch = ss.getSheetByName(ASSIGNMENTS_ARCHIVE_SHEET_);
  if (!arch) {
    arch = ss.insertSheet(ASSIGNMENTS_ARCHIVE_SHEET_);
    arch.appendRow(data[0].concat(['School Year']));
    arch.setFrozenRows(1);
  }

  const toArchive = [];
  const rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[A_.id]) continue;
    if (((r[A_.building] || DEFAULT_BUILDING).toString()) !== building) continue;
    toArchive.push(r.concat([yearName]));
    rowsToDelete.push(i + 1);
  }

  if (toArchive.length) {
    arch.getRange(arch.getLastRow() + 1, 1, toArchive.length, toArchive[0].length).setValues(toArchive);
    rowsToDelete.sort((a, b) => b - a).forEach(rn => src.deleteRow(rn));
  }
}

// Reached only through doGet after verifyAssignmentLink_ (private: not callable
// via google.script.run).
function handleAssignmentRecord_(p) {
  const sessionEmail = (Session.getActiveUser().getEmail() || '').toString().trim().toLowerCase();
  const invited = (p.tEmail || '').toString().trim().toLowerCase();
  if (!sessionEmail || sessionEmail !== invited) {
    return assignmentMessagePage_('Wrong account',
      'This assignment was sent to ' + (p.tEmail || 'another staff member') +
      '. Open the link while signed in to that account.');
  }

  const a = findAssignment_(p.id);
  if (!a) {
    return assignmentMessagePage_('Assignment not found',
      'This assignment is no longer on file. Ask your administrator to re-send it.');
  }
  if (a.subEmail.toLowerCase() !== invited) {
    return assignmentMessagePage_('Wrong account', 'This assignment belongs to another staff member.');
  }
  if (a.status === ASSIGNMENT_STATUS_.cancelled) {
    return assignmentMessagePage_('Assignment cancelled',
      'This coverage assignment was cancelled. You do not need to cover this class.');
  }
  if (a.status === ASSIGNMENT_STATUS_.recorded) {
    return assignmentConfirmedPage_(a, true);
  }
  if (daysSinceDate_(a.date) > ASSIGNMENT_LINK_DAYS_) {
    return assignmentMessagePage_('This link has expired',
      'Coverage links stop working ' + ASSIGNMENT_LINK_DAYS_ + ' days after the coverage date. ' +
      'Sign in to TST Manager to record it, or contact ' + (a.assignedByName || a.assignedBy) + '.');
  }

  recordAssignment_(a.id, 'teacher');
  return assignmentConfirmedPage_(a, false);
}

function assignmentConfirmedPage_(a, already) {
  const e = escapeHtml_;
  const buildingName = buildingNameFor_(a.building);
  const dashboardLink = scriptUrl_() || '?';
  const subtitle = already
    ? 'Your TST time for this coverage was already recorded.'
    : 'Thank you, ' + a.subName + '. Your TST time has been submitted for approval.';

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>TST Time Recorded</title>
      <link href="https://fonts.googleapis.com/css2?family=Lexend:wght@300;400;500;600;700&display=swap" rel="stylesheet">
      <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
      <style>
        body { font-family: 'Lexend', sans-serif; background-color: #f9fafb; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; }
        .card { background: white; border-radius: 12px; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); width: 100%; max-width: 480px; overflow: hidden; border: 1px solid #e5e7eb; }
        .header { background-color: #2d3f89; padding: 24px; text-align: center; color: white; }
        .icon-circle { background: white; width: 64px; height: 64px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 16px auto; color: #2d3f89; font-size: 32px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .content { padding: 32px 24px; text-align: center; }
        .title { font-size: 24px; font-weight: 700; color: #1f2937; margin-bottom: 8px; }
        .subtitle { color: #6b7280; margin-bottom: 24px; font-size: 14px; }
        .details-box { background-color: #eff6ff; border-left: 4px solid #2d3f89; text-align: left; padding: 16px; border-radius: 4px; margin-bottom: 32px; }
        .detail-row { margin-bottom: 8px; font-size: 14px; color: #374151; }
        .detail-row:last-child { margin-bottom: 0; }
        .label { font-weight: 600; color: #2d3f89; margin-right: 8px; }
        .btn { display: inline-block; background-color: #2d3f89; color: white; padding: 12px 32px; border-radius: 6px; text-decoration: none; font-weight: 600; }
        .btn:hover { background-color: #1e3a8a; }
      </style>
    </head>
    <body>
      <div class="card">
        <div class="header">
          <div class="icon-circle"><i class="fas fa-check"></i></div>
          <h1 style="margin:0; font-size:20px; font-weight:600;">${e(buildingName)}</h1>
          <p style="margin:4px 0 0 0; opacity:0.8; font-size:12px; text-transform:uppercase; letter-spacing:1px;">TST Manager</p>
        </div>
        <div class="content">
          <h2 class="title">TST Time Recorded</h2>
          <p class="subtitle">${e(subtitle)}</p>
          <div class="details-box">
            <div class="detail-row"><span class="label">Date:</span> ${e(longDate_(a.date))}</div>
            <div class="detail-row"><span class="label">Period:</span> ${e(a.period)}</div>
            <div class="detail-row"><span class="label">Covering For:</span> ${e(a.coveredFor)}</div>
            <div class="detail-row"><span class="label">Duration:</span> ${e(durationLabel_(a))}</div>
          </div>
          <a href="${e(dashboardLink)}" class="btn">Go to Dashboard</a>
        </div>
      </div>
    </body>
    </html>
  `;

  return HtmlService.createHtmlOutput(html)
      .setTitle('TST Time Recorded')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Admin Action: Update the schedule for a specific Month + Period.
 * This effectively "syncs" the grid view back to the row-based sheet.
 * @param {string} month - e.g. "September"
 * @param {string} period - e.g. "Period 1 - ..."
 * @param {Object} dayUpdates - { "Mon": ["email1", "email2"], "Tue": [] ... }
 * @param {string} building - The building whose schedule was edited (client sends
 *   STATE.building). Availability rows aren't building-tagged and buildings can share
 *   period names (e.g. "Time Range"), so only this building's members are replaced.
 */
function updateSchedulePeriod(month, period, dayUpdates, building) {
  const ctx = getUserContext();
  assertAdmin_(ctx);
  const effectiveBuilding = allowedBuildingFor_(ctx, building);
  if (!effectiveBuilding) {
    throw new Error('You can only edit the schedule for your own building(s).');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Availability');

  // 1. This building's schedule members (lowercased email -> name)
  const members = scheduleMembers_(effectiveBuilding);

  // 2. Invert dayUpdates into TeacherEmail -> Set(Days), refusing anyone outside the
  //    building before touching the sheet so a rejected save changes nothing.
  // dayUpdates format: { "Mon": ["a@b.com", "c@d.com"], "Tue": ["a@b.com"] }
  const teacherDays = {};
  const outsiders = [];
  Object.keys(dayUpdates || {}).forEach(day => {
    const emails = Array.isArray(dayUpdates[day]) ? dayUpdates[day] : [];
    emails.forEach(email => {
      const e = (email || '').toString().toLowerCase().trim();
      if (!e) return;
      if (!members.has(e)) {
        if (!outsiders.includes(e)) outsiders.push(e);
        return;
      }
      if (!teacherDays[e]) teacherDays[e] = new Set();
      teacherDays[e].add(day);
    });
  });
  if (outsiders.length > 0) {
    throw new Error('Not on the ' + effectiveBuilding + ' schedule: ' + outsiders.join(', '));
  }

  // 3. Delete this building's members' rows for Month + Period (other buildings'
  //    rows with the same period name survive), bottom-up so indices stay valid.
  // Rebuilding the rows is safer than diffing row-by-row for multi-day entries.
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    // Cols: A=Month, C=Period, E=Email
    const rowEmail = (data[i][4] || '').toString().trim().toLowerCase();
    if (data[i][0] === month && data[i][2] === period && members.has(rowEmail)) {
      sheet.deleteRow(i + 1); // 1-based index
    }
  }

  // 4. Rebuild this building's rows from dayUpdates
  const newRows = [];
  Object.keys(teacherDays).forEach(email => {
    const days = Array.from(teacherDays[email]).sort().join(','); // "Mon,Tue"
    const name = members.get(email) || email; // Fallback to email if name not found
    
    // Cols: Month, Day(s), Period, Name, Email, Hours(empty)
    newRows.push([month, days, period, name, email, ""]);
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
  
  return true;
}

/**
 * Admin Action: Update staff carry over hours.
 * @param {string} email - Staff email
 * @param {number} newAmount - New carry over amount
 */
function updateStaffCarryOver(email, newAmount) {
  const ctx = getUserContext();
  assertAdmin_(ctx);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Staff Directory');
  const data = sheet.getDataRange().getValues();
  // Header Row = 0.
  // Cols: A=Name, B=Email(1), G=CarryOver(6) - based on Schema/staffDirectoryData_

  // We need to find the correct column index dynamically or hardcode based on known schema
  // staffDirectoryData_ uses: carryOverIdx = headers.findIndex(...) or 6.

  const headers = data[0];
  const emailIdx = headers.findIndex(h => h.toString().toLowerCase().includes('email'));
  const carryOverIdx = headers.findIndex(h => h.toString().toLowerCase().includes('carry'));
  const buildingIdx = headers.findIndex(h => h.toString().toLowerCase().includes('building'));

  const safeEmailIdx = emailIdx > -1 ? emailIdx : 1;
  const safeCarryIdx = carryOverIdx > -1 ? carryOverIdx : 6;
  const safeBuildingIdx = buildingIdx > -1 ? buildingIdx : 8;

  // Find row
  let rowIndex = -1;
  let buildingCell = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][safeEmailIdx].toString().toLowerCase() === email.toLowerCase()) {
      rowIndex = i + 1; // 1-based
      buildingCell = data[i][safeBuildingIdx];
      break;
    }
  }

  if (rowIndex === -1) throw new Error("Staff member not found.");

  // Carry Over is owned by the primary building.
  assertCanManageRow_(ctx, buildingCell);
  assertPrimaryAdminFor_(ctx, buildingCell, 'Carry Over');

  // Coerce to a number so a blank/garbage value can't poison the Running Total
  // ARRAYFORMULA (consistent with the other owned-column writers).
  sheet.getRange(rowIndex, safeCarryIdx + 1).setValue(Number(newAmount) || 0);
  return true;
}

/**
 * MAINTENANCE UTILITY: Use this to sync any submissions that were archived
 * to 'Form Responses 1' but failed to copy to 'TST Approvals (New)'.
 * Safe to run multiple times; it checks for duplicates.
 *
 * Run by an admin from the Apps Script editor; the admin check is what stops it
 * being called from the browser console, where it would create approval rows for
 * any staff member in any building.
 */
function syncMissingSubmissions() {
  assertAdmin_(getUserContext());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName('Form Responses 1');
  const approvalSheet = ss.getSheetByName('TST Approvals (New)');
  
  if (!formSheet || !approvalSheet) return "Error: Sheets not found.";
  
  const formDataRaw = formSheet.getDataRange().getValues();
  formDataRaw.shift(); // Remove headers
  
  const approvalDataRaw = approvalSheet.getDataRange().getValues();
  approvalDataRaw.shift();
  
  // Create a lookup set of existing submissions in Approvals
  // Key: Email + Date + Period
  const existingKeys = new Set(approvalDataRaw.map(r => {
    if (!r[0]) return "";
    const email = r[0].toString().toLowerCase().trim();
    // Normalize date for comparison
    const dateStr = r[4] instanceof Date ? r[4].toISOString().split('T')[0] : r[4].toString().split('T')[0];
    const period = r[5] ? r[5].toString().trim() : "";
    return `${email}|${dateStr}|${period}`;
  }));
  
  let syncCount = 0;
  
  formDataRaw.forEach(r => {
    const email = r[1] ? r[1].toString().toLowerCase().trim() : "";
    if (!email) return;

    const dateRaw = r[4];
    if (!dateRaw) return;
    const dateStr = dateRaw instanceof Date ? dateRaw.toISOString().split('T')[0] : dateRaw.toString().split('T')[0];
    const period = r[5] ? r[5].toString().trim() : "";
    
    const key = `${email}|${dateStr}|${period}`;
    
    if (!existingKeys.has(key)) {
      // Missing! Process it using the shared logic
      processEarnedSubmission_({
        email: email,
        subbedFor: r[2],
        otherText: r[3],
        dateStr: dateStr,
        period: period,
        amountType: r[6],
        amountDecimal: r[7],
        building: null // Will be resolved by staff lookup
      });
      syncCount++;
    }
  });
  
  const msg = "Synced " + syncCount + " missing submissions.";
  Logger.log(msg);
  return msg;
}
