/**
 * Serves the HTML file.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('TST Manager')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Gets the current user's email, determines their role based on the Staff Directory,
 * and fetches necessary initial data.
 */
function getInitialData() {
  const userEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName('Staff Directory');
  
  if (!staffSheet) throw new Error("Sheet 'Staff Directory' not found.");
  
  const data = staffSheet.getDataRange().getValues();
  const headers = data.shift(); // Remove headers
  
  let currentUserRole = 'None';
  let currentUserName = '';
  
  // Find current user in directory
  const userRow = data.find(r => r[1].toString().toLowerCase() === userEmail.toLowerCase());
  
  if (userRow) {
    currentUserName = userRow[0]; // Col A
    currentUserRole = userRow[2]; // Col C (Admin or Teacher)
  } else {
    // Fallback if not found (optional: default to Teacher or lock out)
    currentUserRole = 'Guest'; 
  }

  return {
    email: userEmail,
    name: currentUserName,
    role: currentUserRole,
    staffData: getStaffDirectoryData() // Pre-load staff data for the UI
  };
}

/**
 * Helper to get clean object array of Staff Directory
 */
function getStaffDirectoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Staff Directory');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  return data.map((r, i) => ({
    name: r[0],
    email: r[1],
    role: r[2],
    earned: r[3],
    used: r[4],
    total: r[5],
    rowIndex: i + 2 // 1-based index + header offset
  }));
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
 */
function getPendingEarned() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  // Col I (Index 8) = Approved Status
  // Col K (Index 10) = Denied Status (New)
  
  return data.map((r, i) => ({
    email: r[0],
    name: r[1],
    subbedFor: r[2],
    date: safeDate(r[4]), // Convert Date to String
    period: r[5],
    timeType: r[6],
    hours: r[7],
    status: r[8], // Approved Checkbox
    denied: r[10], // Denied Checkbox
    rowIndex: i + 2
  })).filter(item => {
    // Show if NOT approved AND NOT denied
    const isApproved = item.status === true || item.status === "TRUE";
    const isDenied = item.denied === true || item.denied === "TRUE";
    return !isApproved && !isDenied;
  });
}

/**
 * Fetches data for the Admin Pending Used view.
 * Source: TST Usage (New)
 */
function getPendingUsed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  // Filter for Checkbox (Col E / Index 4) == false
  return data.map((r, i) => ({
    email: r[0],
    name: r[1],
    date: safeDate(r[2]), // Convert Date to String
    used: r[3],
    status: r[4],
    rowIndex: i + 2
  })).filter(item => item.status === false || item.status === "" || item.status === "FALSE");
}

/**
 * Gets history for a specific teacher.
 */
function getTeacherHistory(targetEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get Earned (Approved OR Denied)
  const earnedSheet = ss.getSheetByName('TST Approvals (New)');
  const earnedData = earnedSheet.getDataRange().getValues();
  earnedData.shift();
  
  const earned = earnedData
    .filter(r => {
       const isEmailMatch = r[0].toString().toLowerCase() === targetEmail.toLowerCase();
       const isApproved = r[8] === true;
       const isDenied = r[10] === true;
       return isEmailMatch && (isApproved || isDenied);
    })
    .map(r => ({
      date: safeDate(r[4]), 
      period: r[5],
      subbedFor: r[2],
      amount: r[7],
      // Determine type based on columns
      type: r[10] === true ? 'Denied' : 'Earned' 
    }));

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
      type: 'Used'
    }));
    
  return [...earned, ...used].sort((a, b) => new Date(b.date) - new Date(a.date));
}


/**
 * Admin Action: Approve an Earned request.
 */
function approveEarnedRow(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  
  // Col I (9) is Approved Status, Col J (10) is Timestamp
  // Col K (11) is Denied Status. 
  // Safety: Ensure Denied is FALSE if we are Approving.
  
  sheet.getRange(rowIndex, 9).setValue(true);   // Set Approved = TRUE
  sheet.getRange(rowIndex, 10).setValue(new Date()); // Set Approved Timestamp
  sheet.getRange(rowIndex, 11).setValue(false); // Set Denied = FALSE (Safety)
  
  return true;
}

/**
 * Admin Action: Deny an Earned request.
 */
function denyEarnedRow(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  
  // Col I (9) is Approved Status
  // Col K (11) is Denied Status, Col L (12) is Denied Timestamp
  
  sheet.getRange(rowIndex, 9).setValue(false);  // Set Approved = FALSE (Safety)
  sheet.getRange(rowIndex, 11).setValue(true);  // Set Denied = TRUE
  sheet.getRange(rowIndex, 12).setValue(new Date()); // Set Denied Timestamp
  
  return true;
}

/**
 * Admin Action: Approve a Used request.
 */
function approveUsedRow(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  // Col E (5) is status, Col F (6) is timestamp
  sheet.getRange(rowIndex, 5).setValue(true);
  sheet.getRange(rowIndex, 6).setValue(new Date());
  return true;
}

/**
 * Create a new Usage entry (Admin or Teacher).
 */
function submitUsage(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  
  // Columns: A: Email, B: Name, C: Date, D: TST Used, E: Status, F: Timestamp
  sheet.appendRow([
    formObj.email,
    formObj.name,
    formObj.date,
    formObj.amount,
    false, // Default unchecked
    ""     // No timestamp yet
  ]);
  return true;
}

/**
 * Create a new Earned entry (Teacher subbing).
 * Writes to Form Responses 1.
 */
function submitEarned(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Form Responses 1');
  
  // Columns: A: Timestamp, B: Email, C: I subbed for (Dropdown), D: Other (Manual), 
  // E: Date Subbed, F: Time Subbed (Period), G: Amount (Text), H: Decimal
  
  const timestamp = new Date();
  
  sheet.appendRow([
    timestamp,
    formObj.email,
    formObj.subbedForType === 'Other' ? 'Other' : formObj.subbedForName, // Col C
    formObj.subbedForType === 'Other' ? formObj.subbedForName : '',      // Col D
    formObj.date,
    formObj.period,
    formObj.amountType, // "Full Period" or "Half Period"
    formObj.amountDecimal
  ]);
  
  return true;
}

/**
 * Sends an email report to a staff member.
 */
function sendStatusEmail(targetEmail, targetName) {
  const history = getTeacherHistory(targetEmail);
  const staff = getStaffDirectoryData().find(s => s.email.toLowerCase() === targetEmail.toLowerCase());
  
  if (!staff) throw new Error("Staff member not found.");
  
  let htmlBody = `
    <h2>TST Hours Report for ${targetName}</h2>
    <p><strong>Current Balance:</strong> ${Number(staff.total).toFixed(2)} hours</p>
    <hr>
    <h3>History</h3>
    <table border="1" cellpadding="5" style="border-collapse:collapse; width:100%;">
      <tr style="background-color:#f3f4f6;">
        <th>Date</th>
        <th>Type</th>
        <th>Details</th>
        <th>Hours</th>
      </tr>`;
      
  history.forEach(h => {
    // Skip formatting for Denied if you want, or just include them
    if(h.type === 'Denied') return; // Optional: Don't email denied ones? Or include them. 
    // Let's include them for clarity
    
    const dateStr = new Date(h.date).toLocaleDateString();
    let color = 'black';
    let sign = '';
    
    if (h.type === 'Earned') { color = 'green'; sign = '+'; }
    else if (h.type === 'Used') { color = 'red'; sign = '-'; }
    else { color = 'gray'; sign = ''; } // Denied
    
    htmlBody += `
      <tr>
        <td>${dateStr}</td>
        <td>${h.type}</td>
        <td>${h.type === 'Earned' ? 'Subbed for: ' + h.subbedFor : (h.type === 'Used' ? 'Redeemed' : 'Request Denied')}</td>
        <td style="color:${color}; font-weight:bold;">${sign}${h.amount}</td>
      </tr>`;
  });
  
  htmlBody += `</table>`;
  
  MailApp.sendEmail({
    to: targetEmail,
    subject: "Your TST Hours Report",
    htmlBody: htmlBody
  });
  
  return true;
}
