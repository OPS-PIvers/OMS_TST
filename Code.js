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
    carryOver: r[6],
    total: (Number(r[6]) || 0) + (Number(r[3]) || 0) - (Number(r[4]) || 0),
    rowIndex: i + 2 // 1-based index + header offset
  })).filter(r => r.email !== "");
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
    // Show if NOT approved AND NOT denied AND has valid data
    const isApproved = item.status === true || item.status === "TRUE";
    const isDenied = item.denied === true || item.denied === "TRUE";
    return !isApproved && !isDenied && item.email !== "";
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
  
  // Filter for Checkbox (Col E / Index 4) == false AND has data
  return data.map((r, i) => ({
    email: r[0],
    name: r[1],
    date: safeDate(r[2]), // Convert Date to String
    used: r[3],
    status: r[4],
    rowIndex: i + 2
  })).filter(item => (item.status === false || item.status === "" || item.status === "FALSE") && item.email !== "");
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
        type: type 
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
 * Admin Action: Delete an Earned request.
 * Deletes from BOTH 'TST Approvals (New)' and 'Form Responses 1'.
 */
function deleteEarnedRow(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const approvalSheet = ss.getSheetByName('TST Approvals (New)');
  const formSheet = ss.getSheetByName('Form Responses 1');
  
  // 1. Get details from Approval Sheet to find match
  // Indexes: 0=Email, 1=Name, 2=SubbedFor, 4=Date, 5=Period
  const rowValues = approvalSheet.getRange(rowIndex, 1, 1, 6).getValues()[0];
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
       // We stop after first match? Or continue? usually one entry.
       // Let's break to be safe/efficient, assuming duplicates aren't common or handled elsewhere.
       break; 
    }
  }

  // 3. Delete from TST Approvals (New)
  approvalSheet.deleteRow(rowIndex);
  
  return true;
}

/**
 * Admin Action: Edit an Earned request.
 * Updates BOTH 'TST Approvals (New)' and 'Form Responses 1'.
 */
function updateEarnedRow(rowIndex, newData) {
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
       formSheet.getRange(i + 1, 5).setValue(new Date(newData.date));
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
  approvalSheet.getRange(rowIndex, 5).setValue(new Date(newData.date));
  approvalSheet.getRange(rowIndex, 6).setValue(newData.period);
  approvalSheet.getRange(rowIndex, 7).setValue(newData.amountType);
  approvalSheet.getRange(rowIndex, 8).setValue(newData.amountDecimal);

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
 * Admin Action: Delete a Used request.
 */
function deleteUsedRow(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  sheet.deleteRow(rowIndex);
  return true;
}

/**
 * Admin Action: Edit a Used request.
 */
function updateUsedRow(rowIndex, newData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  // Cols: C=Date (3), D=Amount (4)
  sheet.getRange(rowIndex, 3).setValue(new Date(newData.date));
  sheet.getRange(rowIndex, 4).setValue(newData.amount);
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
  
  // 1. Archive to Form Responses 1 (Keep as backup)
  const formSheet = ss.getSheetByName('Form Responses 1');
  const timestamp = new Date();
  
  formSheet.appendRow([
    timestamp,
    formObj.email,
    formObj.subbedForType === 'Other' ? 'Other' : formObj.subbedForName, // Col C
    formObj.subbedForType === 'Other' ? formObj.subbedForName : '',      // Col D
    formObj.date,
    formObj.period,
    formObj.amountType, 
    formObj.amountDecimal
  ]);

  // 2. Append to TST Approvals (New) - Decoupled
  const approvalSheet = ss.getSheetByName('TST Approvals (New)');
  
  // Lookup Name from Staff Directory
  const staffSheet = ss.getSheetByName('Staff Directory');
  const staffData = staffSheet.getDataRange().getValues();
  // Col A=Name, Col B=Email. Find row where B matches email.
  const staffRow = staffData.find(r => r[1].toString().toLowerCase() === formObj.email.toLowerCase());
  const earnerName = staffRow ? staffRow[0] : formObj.email; // Fallback to email if name not found

  approvalSheet.appendRow([
    formObj.email,                    // A: Email
    earnerName,                       // B: Name
    formObj.subbedForType === 'Other' ? 'Other' : formObj.subbedForName, // C: Subbed For
    formObj.subbedForType === 'Other' ? formObj.subbedForName : '',      // D: Other Details
    formObj.date,                     // E: Date
    formObj.period,                   // F: Period
    formObj.amountType,               // G: Time Type
    formObj.amountDecimal,            // H: Hours
    false,                            // I: Approved (Default False)
    "",                               // J: Approved TS
    false,                            // K: Denied (Default False)
    ""                                // L: Denied TS
  ]);
  
  return true;
}

/**
 * Admin Multi-Submit: Handles creating both Earned and Used records 
 * based on Admin input.
 */
function adminSubmitRequest(data) {
  // 1. Handle Earner (If staff member is selected)
  if (data.earner.type === 'Staff' && data.earner.email) {
    // We treat this like a form submission so it flows into the normal Pending pipeline
    submitEarned({
      email: data.earner.email,
      subbedForType: data.user.type, // 'Staff' or 'Other'
      subbedForName: data.user.name,
      date: data.details.date,
      period: data.details.period,
      amountType: data.details.amountType,
      amountDecimal: data.details.amount
    });
  }

  // 2. Handle User (If staff member is selected)
  if (data.user.type === 'Staff' && data.user.email) {
    submitUsage({
      email: data.user.email,
      name: data.user.name,
      date: data.details.date,
      amount: data.details.amount
    });
  }

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

/**
 * Trigger: On Form Submit
 * Syncs new rows from 'Form Responses 1' to 'TST Approvals (New)'.
 * Must be manually set up as an Installable Trigger in Apps Script editor.
 */
function onFormSubmit(e) {
  if (!e || !e.values) return; // Safety check
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const approvalSheet = ss.getSheetByName('TST Approvals (New)');
  const staffSheet = ss.getSheetByName('Staff Directory');
  
  // Parse Form Data (Array indices based on Form Responses 1 columns)
  // [0] Timestamp, [1] Email, [2] SubbedFor, [3] Other, [4] Date, [5] Period, [6] Type, [7] Decimal
  const email = e.values[1];
  const subbedFor = e.values[2];
  const otherText = e.values[3];
  const dateStr = e.values[4]; // Form might return different date format, beware.
  const period = e.values[5];
  const amountType = e.values[6];
  const amountDecimal = e.values[7];
  
  // Lookup Name
  const staffData = staffSheet.getDataRange().getValues();
  const staffRow = staffData.find(r => r[1].toString().toLowerCase() === email.toString().toLowerCase());
  const earnerName = staffRow ? staffRow[0] : email;

  // Append to Approvals
  approvalSheet.appendRow([
    email,
    earnerName,
    subbedFor,
    otherText,
    dateStr,
    period,
    amountType,
    amountDecimal,
    false, // Approved
    "",    // TS
    false, // Denied
    ""     // TS
  ]);
}