/**
 * Serves the HTML file or handles email actions.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    if (e.parameter.action === 'accept') {
      return handleCoverageAccept(e.parameter);
    } else if (e.parameter.action === 'reject') {
      return handleCoverageReject(e.parameter);
    }
  }

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
 * Lightweight helper to get pending counts for badges
 */
function getDashboardCounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Earned
  const earnedSheet = ss.getSheetByName('TST Approvals (New)');
  const earnedData = earnedSheet.getDataRange().getValues();
  earnedData.shift(); // Remove header
  // Filter: Not Approved (8/I=false) AND Not Denied (10/K=false)
  const earnedCount = earnedData.filter(r => r[8] !== true && r[8] !== "TRUE" && r[10] !== true && r[10] !== "TRUE" && r[0] !== "").length;
  
  // Used
  const usedSheet = ss.getSheetByName('TST Usage (New)');
  const usedData = usedSheet.getDataRange().getValues();
  usedData.shift();
  // Filter: Not Approved (4/E=false)
  const usedCount = usedData.filter(r => (r[4] === false || r[4] === "" || r[4] === "FALSE") && r[0] !== "").length;
  
  return { earned: earnedCount, used: usedCount };
}

/**
 * Helper to get clean object array of Staff Directory
 */
function batchApproveEarned(indices) {
  if (!indices || !Array.isArray(indices)) return;
  // Sort descending just in case, though for updates it matters less than deletes
  indices.sort((a, b) => b - a);
  
  indices.forEach(idx => {
    approveEarnedRow(idx, { send: false }); // No email for batch
  });
  return true;
}

/**
 * Batch Action: Deny multiple Earned requests.
 */
function batchDenyEarned(indices) {
  if (!indices || !Array.isArray(indices)) return;
  indices.sort((a, b) => b - a);
  
  indices.forEach(idx => {
    denyEarnedRow(idx, { send: false }); // No email, no specific reason
  });
  return true;
}

/**
 * Batch Action: Approve multiple Used requests.
 */
function batchApproveUsed(indices) {
  if (!indices || !Array.isArray(indices)) return;
  indices.sort((a, b) => b - a);
  
  indices.forEach(idx => {
    approveUsedRow(idx);
  });
  return true;
}

/**
 * Batch Action: Delete multiple Used requests.
 * MUST process descending to preserve indices.
 */
function batchDeleteUsed(indices) {
  if (!indices || !Array.isArray(indices)) return;
  // Critical: Sort descending
  indices.sort((a, b) => b - a);
  
  indices.forEach(idx => {
    deleteUsedRow(idx);
  });
  return true;
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
        type: type,
        denialReason: r[12]
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
 * Gets history for a specific teacher with row indices and sheet info for admin actions.
 * Used by admin to manage approved/denied items.
 */
function getStaffHistoryWithActions(targetEmail) {
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
        amountType: r[6] // Time Type (Full/Half)
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
        amountType: 'N/A'
      };
    })
    .filter(item => item !== null);

  return [...earned, ...used].sort((a, b) => new Date(b.date) - new Date(a.date));
}


/**
 * Admin Action: Revert an Earned request back to Pending.
 * Clears Approved/Denied flags and adjusts staff balance.
 */
function revertEarnedToPending(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  const staffSheet = ss.getSheetByName('Staff Directory');

  // Get row data
  const rowData = sheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
  const email = rowData[0];
  const hours = rowData[7];
  const wasApproved = rowData[8] === true;

  // If it was approved, we need to reverse the balance change
  if (wasApproved) {
    const staffData = staffSheet.getDataRange().getValues();
    const staffRowIndex = staffData.findIndex(r => r[1].toString().toLowerCase() === email.toLowerCase());

    if (staffRowIndex > 0) {
      // Decrement Earned balance (Col D / Index 4)
      const currentEarned = Number(staffData[staffRowIndex][3]) || 0;
      const newEarned = currentEarned - Number(hours);
      staffSheet.getRange(staffRowIndex + 1, 4).setValue(newEarned);
    }
  }

  // Clear approval/denial flags and timestamps
  sheet.getRange(rowIndex, 9).setValue(false);   // Clear Approved (Col I)
  sheet.getRange(rowIndex, 10).setValue('');     // Clear Approved TS (Col J)
  sheet.getRange(rowIndex, 11).setValue(false);  // Clear Denied (Col K)
  sheet.getRange(rowIndex, 12).setValue('');     // Clear Denied TS (Col L)
  sheet.getRange(rowIndex, 13).setValue('');     // Clear Denial Reason (Col M)

  return { success: true, hoursAdjusted: wasApproved ? hours : 0 };
}

/**
 * Admin Action: Revert a Used request back to Pending.
 * Clears Approved flag and adjusts staff balance.
 */
function revertUsedToPending(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  const staffSheet = ss.getSheetByName('Staff Directory');

  // Get row data
  const rowData = sheet.getRange(rowIndex, 1, 1, 6).getValues()[0];
  const email = rowData[0];
  const hours = rowData[3];
  const wasApproved = rowData[4] === true;

  // If it was approved, we need to reverse the balance change
  if (wasApproved) {
    const staffData = staffSheet.getDataRange().getValues();
    const staffRowIndex = staffData.findIndex(r => r[1].toString().toLowerCase() === email.toLowerCase());

    if (staffRowIndex > 0) {
      // Decrement Used balance (Col E / Index 5)
      const currentUsed = Number(staffData[staffRowIndex][4]) || 0;
      const newUsed = currentUsed - Number(hours);
      staffSheet.getRange(staffRowIndex + 1, 5).setValue(newUsed);
    }
  }

  // Clear approval flag and timestamp
  sheet.getRange(rowIndex, 5).setValue(false);  // Clear Status (Col E)
  sheet.getRange(rowIndex, 6).setValue('');     // Clear Timestamp (Col F)

  return { success: true, hoursAdjusted: wasApproved ? hours : 0 };
}

/**
 * Admin Action: Delete an Approved Earned request.
 * Removes from sheets and adjusts staff balance.
 */
function deleteApprovedEarned(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const approvalSheet = ss.getSheetByName('TST Approvals (New)');
  const formSheet = ss.getSheetByName('Form Responses 1');
  const staffSheet = ss.getSheetByName('Staff Directory');

  // Get row data
  const rowData = approvalSheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
  const email = rowData[0];
  const date = new Date(rowData[4]);
  const period = rowData[5];
  const hours = rowData[7];
  const wasApproved = rowData[8] === true;

  // If it was approved, reverse the balance change
  if (wasApproved) {
    const staffData = staffSheet.getDataRange().getValues();
    const staffRowIndex = staffData.findIndex(r => r[1].toString().toLowerCase() === email.toLowerCase());

    if (staffRowIndex > 0) {
      // Decrement Earned balance (Col D / Index 4)
      const currentEarned = Number(staffData[staffRowIndex][3]) || 0;
      const newEarned = currentEarned - Number(hours);
      staffSheet.getRange(staffRowIndex + 1, 4).setValue(newEarned);
    }
  }

  // Delete from Form Responses 1
  const formData = formSheet.getDataRange().getValues();
  for (let i = formData.length - 1; i >= 1; i--) {
    const r = formData[i];
    const rDate = new Date(r[4]);
    const isDateMatch = rDate.getFullYear() === date.getFullYear() &&
                        rDate.getMonth() === date.getMonth() &&
                        rDate.getDate() === date.getDate();

    if (r[1] === email && isDateMatch && r[5] == period) {
      formSheet.deleteRow(i + 1);
      break;
    }
  }

  // Delete from TST Approvals (New)
  approvalSheet.deleteRow(rowIndex);

  return { success: true, hoursAdjusted: wasApproved ? hours : 0 };
}

/**
 * Admin Action: Delete an Approved Used request.
 * Removes from sheet and adjusts staff balance.
 */
function deleteApprovedUsed(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  const staffSheet = ss.getSheetByName('Staff Directory');

  // Get row data
  const rowData = sheet.getRange(rowIndex, 1, 1, 6).getValues()[0];
  const email = rowData[0];
  const hours = rowData[3];
  const wasApproved = rowData[4] === true;

  // If it was approved, reverse the balance change
  if (wasApproved) {
    const staffData = staffSheet.getDataRange().getValues();
    const staffRowIndex = staffData.findIndex(r => r[1].toString().toLowerCase() === email.toLowerCase());

    if (staffRowIndex > 0) {
      // Decrement Used balance (Col E / Index 5)
      const currentUsed = Number(staffData[staffRowIndex][4]) || 0;
      const newUsed = currentUsed - Number(hours);
      staffSheet.getRange(staffRowIndex + 1, 5).setValue(newUsed);
    }
  }

  // Delete row
  sheet.deleteRow(rowIndex);

  return { success: true, hoursAdjusted: wasApproved ? hours : 0 };
}

/**
 * Admin Action: Approve an Earned request.
 */
function approveEarnedRow(rowIndex, emailData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  
  // Get data for email BEFORE updating
  // Row Index is 1-based. 
  // Cols: A=Email(1), B=Name(2), C=SubbedFor(3), E=Date(5), F=Period(6), H=Hours(8)
  const range = sheet.getRange(rowIndex, 1, 1, 8);
  const values = range.getValues()[0];
  const rowData = {
    email: values[0],
    name: values[1],
    subbedFor: values[2],
    date: values[4],
    period: values[5],
    hours: values[7]
  };

  // Col I (9) is Approved Status, Col J (10) is Timestamp
  // Col K (11) is Denied Status. 
  // Safety: Ensure Denied is FALSE if we are Approving. 
  
  sheet.getRange(rowIndex, 9).setValue(true);   // Set Approved = TRUE
  sheet.getRange(rowIndex, 10).setValue(new Date()); // Set Approved Timestamp
  sheet.getRange(rowIndex, 11).setValue(false); // Set Denied = FALSE (Safety)
  
  // Send Email if requested
  if (emailData && emailData.send) {
    const formattedDate = new Date(rowData.date).toLocaleDateString();
    const subject = `TST Request for ${formattedDate} has been Approved`;
    const body = `
      <p>Your request has been approved and added to your balance.</p>
      <div style="background-color: #f8fafc; border-left: 4px solid #2d3f89; padding: 15px; margin: 15px 0;">
        <p style="margin: 0; color: #64748b; font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em;">Request Details</p>
        <p style="margin: 5px 0 0 0; color: #1e293b; font-weight: bold;">Subbed for ${rowData.subbedFor}</p>
        <p style="margin: 0; color: #334155;">${formattedDate} &bull; Period ${rowData.period} &bull; +${rowData.hours} hrs</p>
      </div>
      <p>You can check your up-to-date balance on the TST Portal.</p>
    `;
    sendStyledEmail(rowData.email, subject, "Your TST Request was Approved!", body, "Visit the TST Portal");
  }
  
  return true;
}

/**
 * Admin Action: Deny an Earned request.
 */
function denyEarnedRow(rowIndex, emailData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  
  // Get data for email
  const range = sheet.getRange(rowIndex, 1, 1, 8);
  const values = range.getValues()[0];
  const rowData = {
    email: values[0],
    name: values[1],
    subbedFor: values[2],
    date: values[4],
    period: values[5],
    hours: values[7]
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
    const formattedDate = new Date(rowData.date).toLocaleDateString();
    const subject = `TST Request for ${formattedDate} has been Denied`;
    
    let reasonsHtml = "";
    if (emailData.reasons && emailData.reasons.length > 0) {
      reasonsHtml = `<ul style="margin: 10px 0; padding-left: 20px; color: #b91c1c;">` + 
        emailData.reasons.map(r => `<li>${r}</li>`).join('') + 
        `</ul>`;
    }

    const noteHtml = emailData.note ? `<p style="margin-top: 10px;"><em>" ${emailData.note} "</em></p>` : "";

    const body = `
      <p>Your request has been processed and denied.</p>
      
      <div style="background-color: #fef2f2; border-left: 4px solid #ef4444; padding: 15px; margin: 15px 0;">
        <p style="margin: 0; color: #991b1b; font-weight: bold;">Reason for Denial:</p>
        ${reasonsHtml}
        ${noteHtml}
      </div>

      <div style="background-color: #f8fafc; padding: 15px; margin: 15px 0; border: 1px solid #e2e8f0; border-radius: 4px;">
        <p style="margin: 0; color: #64748b; font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em;">Request Details</p>
        <p style="margin: 5px 0 0 0; color: #1e293b; font-weight: bold;">Subbed for ${rowData.subbedFor}</p>
        <p style="margin: 0; color: #334155;">${formattedDate} &bull; Period ${rowData.period}</p>
      </div>

      <p>Please review the details and resubmit if necessary, or contact the TST administrator.</p>
    `;
    
    sendStyledEmail(rowData.email, subject, "TST Request Update", body, "Visit the TST Portal");
  }
  
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
  sheet.getRange(rowIndex, 3).setValue(newData.date);
  sheet.getRange(rowIndex, 4).setValue(newData.amount);
  return true;
}

/**
 * Create a new Usage entry (Admin or Teacher).
 */
function submitUsage(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Usage (New)');
  
  // Columns: A: Email, B: Name, C: Date, D: TST Used, E: Status, F: Timestamp, G: Notes
  sheet.appendRow([
    formObj.email,
    formObj.name,
    formObj.date,
    formObj.amount,
    false, // Default unchecked
    "",    // No timestamp yet
    formObj.notes || "" // Notes (Optional)
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
    formObj.subbedForName,                               // Col C: Name or Manual Text
    formObj.subbedForType === 'Other' ? 'Other' : '',    // Col D: 'Other' flag or empty
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
    formObj.subbedForName,            // C: Subbed For (Name or Manual Text)
    '',                               // D: Other Details (Always empty now)
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
 * Batch Process: Handles a queue of mixed requests.
 */
function processBatch(queue) {
  if (!Array.isArray(queue) || queue.length === 0) return;
  
  queue.forEach(item => {
    try {
      if (item.type === 'earned') {
        submitEarned(item.payload);
      } else if (item.type === 'used') {
        submitUsage(item.payload);
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
  
  let successCount = 0;
  let failCount = 0;
  
  const staffDir = getStaffDirectoryData();

  emails.forEach(email => {
    try {
      // Find name for this email to pass to sendStatusEmail (optimization: get name from dir if possible)
      // sendStatusEmail(email, name) expects name.
      const staff = staffDir.find(s => s.email.toLowerCase() === email.toLowerCase());
      const name = staff ? staff.name : "Staff Member";

      sendStatusEmail(email, name);
      successCount++;
    } catch (e) {
      console.error(`Failed to send email to ${email}:`, e);
      failCount++;
    }
  });

  return { success: successCount, failed: failCount };
}

/**
 * Sends an email report to a staff member.
 */
function sendStatusEmail(targetEmail, targetName) {
  const history = getTeacherHistory(targetEmail);
  const staff = getStaffDirectoryData().find(s => s.email.toLowerCase() === targetEmail.toLowerCase());
  
  if (!staff) throw new Error("Staff member not found.");
  
  const balance = Number(staff.total).toFixed(2);
  
  // Summary Section
  let htmlContent = `
    <div style="background-color: #eff6ff; border-radius: 6px; padding: 20px; margin-bottom: 25px; border-left: 4px solid #3b82f6;">
      <p style="margin: 0; color: #1e40af; font-size: 14px; text-transform: uppercase; font-weight: 600;">Current Balance</p>
      <p style="margin: 5px 0 0 0; color: #1e3a8a; font-size: 32px; font-weight: 700;">${balance} <span style="font-size: 16px; font-weight: 500;">hours</span></p>
    </div>
    
    <h3 style="color: #374151; font-size: 18px; margin-bottom: 15px; border-bottom: 1px solid #e5e7eb; padding-bottom: 10px;">Activity History</h3>
    
    <table cellpadding="0" cellspacing="0" style="width: 100%; border-collapse: collapse; font-size: 14px;">
      <thead>
        <tr style="background-color: #f9fafb;">
          <th style="text-align: left; padding: 12px 10px; border-bottom: 1px solid #e5e7eb; color: #6b7280; font-weight: 600; text-transform: uppercase; font-size: 12px;">Date</th>
          <th style="text-align: left; padding: 12px 10px; border-bottom: 1px solid #e5e7eb; color: #6b7280; font-weight: 600; text-transform: uppercase; font-size: 12px;">Details</th>
          <th style="text-align: right; padding: 12px 10px; border-bottom: 1px solid #e5e7eb; color: #6b7280; font-weight: 600; text-transform: uppercase; font-size: 12px;">Hours</th>
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
        amountStyle += 'color: #059669;'; // Green
        amountDisplay = `+${h.amount}`;
        typeLabel = `<span style="background-color: #d1fae5; color: #065f46; padding: 2px 6px; border-radius: 4px; font-size: 11px;">EARNED</span>`;
        details = `Subbed for: <strong>${h.subbedFor}</strong>`;
      } else if (h.type === 'Used') {
        amountStyle += 'color: #dc2626;'; // Red
        amountDisplay = `-${h.amount}`;
        typeLabel = `<span style="background-color: #fee2e2; color: #991b1b; padding: 2px 6px; border-radius: 4px; font-size: 11px;">USED</span>`;
        details = 'Hours Redeemed';
      } else { // Denied
        amountStyle += 'color: #9ca3af; text-decoration: line-through;'; // Gray
        amountDisplay = `${h.amount}`;
        typeLabel = `<span style="background-color: #f3f4f6; color: #374151; padding: 2px 6px; border-radius: 4px; font-size: 11px;">DENIED</span>`;
        details = 'Request Denied';
        if (h.denialReason) {
          details += `: ${h.denialReason}`;
        }
      }
      
      htmlContent += `
        <tr>
          <td style="padding: 12px 10px; border-bottom: 1px solid #f3f4f6; vertical-align: top; color: #4b5563;">
            <div style="margin-bottom: 4px;">${dateStr}</div>
            ${typeLabel}
          </td>
          <td style="padding: 12px 10px; border-bottom: 1px solid #f3f4f6; vertical-align: top; color: #374151;">
            ${details}
          </td>
          <td style="padding: 12px 10px; border-bottom: 1px solid #f3f4f6; vertical-align: top; text-align: right; ${amountStyle}">
            ${amountDisplay}
          </td>
        </tr>`;
    });
  }
  
  htmlContent += `
      </tbody>
    </table>
    
    <p style="margin-top: 25px; font-size: 13px; color: #6b7280; text-align: center;">
      This report was generated automatically on ${new Date().toLocaleDateString()} at ${new Date().toLocaleTimeString()}.
    </p>
  `;
  
  sendStyledEmail(
    targetEmail,
    "Your TST Hours Report",
    `TST Report for ${targetName}`,
    htmlContent,
    "View Dashboard"
  );
  
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
  let subbedFor = e.values[2];
  
  // Clean Subbed For Name (Remove Titles for Legacy Form compatibility)
  if (subbedFor) {
    subbedFor = subbedFor.replace(/^(Mr\.|Ms\.|Mrs\.|Miss|Dr\.)\s*/i, "").trim();
  }

  const otherText = e.values[3];
  const dateStr = e.values[4]; // Form might return different date format, beware.
  const period = e.values[5];
  const amountType = e.values[6];
  let amountDecimal = e.values[7];
  
  // Lookup Name
  const staffData = staffSheet.getDataRange().getValues();
  const staffRow = staffData.find(r => r[1].toString().toLowerCase() === email.toString().toLowerCase());
  const earnerName = staffRow ? staffRow[0] : email;

  // Append to Approvals
  // If amountDecimal is missing (Legacy Form), calculate it
  if (!amountDecimal) {
    amountDecimal = calculatePeriods(period, amountType);
  }

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

/**
 * Calculates periods based on Middle School rules:
 * - Period 6: always 0.5
 * - Period 7: always 0.5
 * - All other periods: 1.0 for Full, 0.5 for Half
 * @param {string} selectedPeriod The period selection (e.g., "Period 1", "Period 6/7")
 * @param {string} amount "Full Period" or "Half Period"
 * @returns {number} The calculated period value
 */
function calculatePeriods(selectedPeriod, amount) {
  // Handle special cases for periods 6 and 7 (always 0.5)
  if (selectedPeriod && selectedPeriod.toString().includes('Period 6 ') && !selectedPeriod.toString().includes('Period 6/')) {
    return 0.5;
  }
  if (selectedPeriod && selectedPeriod.toString().includes('Period 7 ') && !selectedPeriod.toString().includes('Period 6/')) {
    return 0.5;
  }

  // For all other periods (including Period 6/7), use Full/Half logic
  if (amount && amount.toString().toLowerCase().includes('full')) {
    return 1.0;
  } else if (amount && amount.toString().toLowerCase().includes('half')) {
    return 0.5;
  }

  // Default fallback
  return 1.0;
}

/**
 * Helper to send a styled HTML email.
 */
function sendStyledEmail(recipient, subject, title, contentHtml, buttonText) {
  const appUrl = ScriptApp.getService().getUrl();
  
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${subject}</title>
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
          <h1>Orono Middle School</h1>
        </div>
        <div class="content">
          <h2>${title}</h2>
          ${contentHtml}
          <div class="button-container">
            <a href="${appUrl}" class="button">${buttonText || 'Visit the TST Portal'}</a>
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
  
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlTemplate
  });
}

// --- SCHEDULE / TST AVAILABILITY FEATURE ---

const MONTH_ORDER = ["September", "October", "November", "December", "January", "February", "March", "April", "May", "June"];

function getScheduleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let schedSheet = ss.getSheetByName('TST Availability');
  if (!schedSheet) {
    // Create if missing
    schedSheet = ss.insertSheet('TST Availability');
    schedSheet.appendRow(['Month', 'Day(s) Available', 'Period', 'Name', 'Email', 'Hours Earned This Month']);
  }

  const data = schedSheet.getDataRange().getValues();
  data.shift(); // Remove header

  // 1. Calculate Hours per Teacher per Month
  const hoursMap = calculateMonthlyHours(); // Returns { "email_Month": hours }
  
  // 2. Get Pending Requests Map
  const pendingMap = getPendingEarnedMap();

  // 3. Process Schedule Data
  // We return a structured object: { "September": [ { name, email, days, period, hours, pendingRequests }, ... ], ... }
  const schedule = {};
  MONTH_ORDER.forEach(m => schedule[m] = []);

  data.forEach(row => {
    const [month, days, period, name, email] = row;
    if (schedule[month]) {
      const key = `${email}_${month}`;
      const hours = hoursMap[key] || 0;
      schedule[month].push({
        month, days, period, name, email, hours,
        pendingRequests: pendingMap[email] || []
      });
    }
  });

  return schedule;
}

function getPendingEarnedMap() {
  const pendingList = getPendingEarned(); // Reuse existing function
  const map = {};
  
  pendingList.forEach(item => {
    if (!map[item.email]) {
      map[item.email] = [];
    }
    // Minimal data needed for the tooltip/indicator
    map[item.email].push({
      date: item.date, // Already safeDate string
      subbedFor: item.subbedFor,
      period: item.period
    });
  });
  
  return map;
}

function calculateMonthlyHours() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals (New)');
  const data = sheet.getDataRange().getValues();
  data.shift();

  const sums = {}; // "email_MonthName" -> total
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
    const email = row[0];
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

function saveAvailability(month, availabilityList) {
  // availabilityList: [{ days: "Mon,Tue", period: "Period 1" }, ...]
  const userEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName('Staff Directory');
  const staffData = staffSheet.getDataRange().getValues();
  const userRow = staffData.find(r => r[1].toString().toLowerCase() === userEmail.toLowerCase());
  const userName = userRow ? userRow[0] : userEmail;

  const sheet = ss.getSheetByName('TST Availability');
  const data = sheet.getDataRange().getValues();
  
  // 1. Remove existing rows for this user + month
  // We loop backwards to delete
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === month && data[i][4] === userEmail) {
      sheet.deleteRow(i + 1);
    }
  }

  // 2. Add new rows
  // A: Month | B: Day(s) | C: Period | D: Name | E: Email | F: Hours (Ignored/Formula)
  availabilityList.forEach(item => {
    sheet.appendRow([month, item.days, item.period, userName, userEmail, ""]);
  });
}

function sendCoverageRequest(payload) {
  // payload: { teacherEmail, teacherName, subbedFor, date, period, amount, amountType }
  const scriptUrl = ScriptApp.getService().getUrl();
  const adminEmail = Session.getActiveUser().getEmail();
  
  // Encode params safely
  const params = [
    `action=accept`,
    `tEmail=${encodeURIComponent(payload.teacherEmail)}`,
    `tName=${encodeURIComponent(payload.teacherName)}`,
    `sub=${encodeURIComponent(payload.subbedFor)}`,
    `date=${encodeURIComponent(payload.date)}`,
    `pd=${encodeURIComponent(payload.period)}`,
    `amt=${payload.amount}`,
    `type=${encodeURIComponent(payload.amountType)}`,
    `adm=${encodeURIComponent(adminEmail)}`
  ].join('&');
  
  const acceptLink = `${scriptUrl}?${params}`;
  const rejectLink = `${scriptUrl}?action=reject&tName=${encodeURIComponent(payload.teacherName)}&sub=${encodeURIComponent(payload.subbedFor)}&pd=${encodeURIComponent(payload.period)}&adm=${encodeURIComponent(adminEmail)}`;

  const subject = `TST Coverage Request: ${payload.date} - ${payload.period}`;
  
  // Parse YYYY-MM-DD to MM/DD/YYYY manually to avoid timezone shifts
  const [y, m, d] = payload.date.split('-');
  const dateDisplay = `${m}/${d}/${y}`;

  const htmlBody = `
    <div style="font-family: sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e5e7eb; border-radius: 8px;">
      <h2 style="color: #2d3f89; margin-top: 0;">TST Coverage Request</h2>
      <p>Hello <strong>${payload.teacherName}</strong>,</p>
      <p>Can you provide sub coverage?</p>
      
      <div style="background-color: #f3f4f6; padding: 15px; border-radius: 6px; margin: 20px 0;">
        <p style="margin: 5px 0;"><strong>Date:</strong> ${dateDisplay}</p>
        <p style="margin: 5px 0;"><strong>Period:</strong> ${payload.period}</p>
        <p style="margin: 5px 0;"><strong>Covering For:</strong> ${payload.subbedFor}</p>
        <p style="margin: 5px 0;"><strong>Duration:</strong> ${payload.amountType}</p>
      </div>

      <div style="text-align: center; margin: 30px 0;">
         <a href="${acceptLink}" style="background-color: #2d3f89; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold; margin-right: 10px;">Accept & Earn</a>
         <a href="${rejectLink}" style="background-color: #ad2122; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold;">Decline</a>
      </div>
      
      <p style="font-size: 12px; color: #6b7280; text-align: center;">Clicking submit will automatically submit the TST form on your behalf.</p>
    </div>
  `;

  MailApp.sendEmail({
    to: payload.teacherEmail,
    subject: subject,
    htmlBody: htmlBody
  });

  // Also send a tracking email to the Admin
  const adminSubject = `TST Coverage request: ${dateDisplay} - ${payload.teacherName} - ${payload.period}`;
  const adminBody = `
    <p>You requested coverage from <strong>${payload.teacherName}</strong>.</p>
    <div style="background-color: #f3f4f6; padding: 15px; border-radius: 6px; margin: 10px 0;">
      <p style="margin: 5px 0;"><strong>Date:</strong> ${dateDisplay}</p>
      <p style="margin: 5px 0;"><strong>Period:</strong> ${payload.period}</p>
      <p style="margin: 5px 0;"><strong>Covering For:</strong> ${payload.subbedFor}</p>
    </div>
    <p>This email serves as a record of your request. You will receive another notification if they accept or decline.</p>
  `;

  sendStyledEmail(adminEmail, adminSubject, "Coverage Requested", adminBody, "View Dashboard");
}

function handleCoverageAccept(p) {
  // Decode
  const formObj = {
    email: p.tEmail,
    subbedForName: p.sub,
    subbedForType: 'Staff', // Assumption
    date: p.date,
    period: p.pd,
    amountType: p.type,
    amountDecimal: parseFloat(p.amt)
  };
  
  // Reuse submit logic
  submitEarned(formObj);
  
  // Notify Admin of Acceptance
  if (p.adm) {
    const emailBody = `
      <p><strong>${p.tName}</strong> has accepted the request to cover for <strong>${p.sub}</strong>.</p>
      <div style="background-color: #f8fafc; border-left: 4px solid #2d3f89; padding: 15px; margin: 15px 0;">
        <p style="margin: 0; color: #64748b; font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em;">Coverage Details</p>
        <p style="margin: 5px 0 0 0; color: #1e293b; font-weight: bold;">Date: ${p.date}</p>
        <p style="margin: 0; color: #334155;">Period: ${p.pd} &bull; Duration: ${p.type}</p>
      </div>
      <p>A pending earned request has been automatically created.</p>
    `;
    
    sendStyledEmail(p.adm, `TST Coverage Accepted: ${p.tName}`, "Coverage Confirmed", emailBody, "View Dashboard");
  }
  
  let appUrl = ScriptApp.getService().getUrl();
  // Sanitize: Remove query params if present
  if (appUrl && appUrl.includes('?')) {
    appUrl = appUrl.split('?')[0];
  }
  // Fallback: If appUrl is empty (e.g. unpublished), use "?" to reload page without params
  const dashboardLink = appUrl ? appUrl : "?";

  const dateDisplay = new Date(p.date.split('-').join('/')).toLocaleDateString();

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Coverage Confirmed</title>
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
        .btn { display: inline-block; background-color: #2d3f89; color: white; padding: 12px 32px; border-radius: 6px; text-decoration: none; font-weight: 600; transition: background-color 0.2s; }
        .btn:hover { background-color: #1e3a8a; }
      </style>
    </head>
    <body>
      <div class="card">
        <div class="header">
          <div class="icon-circle">
            <i class="fas fa-check"></i>
          </div>
          <h1 style="margin:0; font-size:20px; font-weight:600;">Orono Middle School</h1>
          <p style="margin:4px 0 0 0; opacity:0.8; font-size:12px; text-transform:uppercase; letter-spacing:1px;">TST Manager</p>
        </div>
        <div class="content">
          <h2 class="title">Coverage Confirmed!</h2>
          <p class="subtitle">Thank you, <strong>${p.tName}</strong>. Your request has been successfully processed.</p>
          
          <div class="details-box">
            <div class="detail-row"><span class="label">Date:</span> ${dateDisplay}</div>
            <div class="detail-row"><span class="label">Period:</span> ${p.pd}</div>
            <div class="detail-row"><span class="label">Subbing For:</span> ${p.sub}</div>
            <div class="detail-row"><span class="label">Duration:</span> ${p.type}</div>
          </div>

          <a href="${dashboardLink}" class="btn">Go to Dashboard</a>
        </div>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html)
      .setTitle('Coverage Confirmed')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleCoverageReject(p) {
  // Notify Admin
  const emailBody = `
    <p>Teacher <strong>${p.tName}</strong> has <span style="color: #ad2122; font-weight: bold;">declined</span> the coverage request for <strong>${p.sub}</strong>.</p>
    
    <div style="background-color: #fef2f2; border-left: 4px solid #ef4444; padding: 15px; margin: 15px 0;">
       <p style="margin: 0; color: #991b1b; font-weight: bold;">Declined Request</p>
       <p style="margin: 5px 0 0 0; color: #7f1d1d;">Period: ${p.pd || 'Not specified'}</p>
    </div>

    <p>Please select another teacher from the schedule.</p>
  `;

  sendStyledEmail(p.adm, `TST Request Declined: ${p.tName}`, "Coverage Declined", emailBody, "Find Replacement");
  
  return HtmlService.createHtmlOutput(`
    <div style="font-family: sans-serif; text-align: center; padding: 50px;">
      <h1 style="color: #ef4444;">Request Declined</h1>
      <p>The admin has been notified.</p>
    </div>
  `);
}

/**
 * Admin Action: Update the schedule for a specific Month + Period.
 * This effectively "syncs" the grid view back to the row-based sheet.
 * @param {string} month - e.g. "September"
 * @param {string} period - e.g. "Period 1 - ..."
 * @param {Object} dayUpdates - { "Mon": ["email1", "email2"], "Tue": [] ... }
 */
function updateSchedulePeriod(month, period, dayUpdates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Availability');
  const staffSheet = ss.getSheetByName('Staff Directory');
  
  // 1. Get all staff map for name lookup (Email -> Name)
  const staffData = staffSheet.getDataRange().getValues();
  staffData.shift(); // Header
  const staffMap = {};
  staffData.forEach(r => {
    staffMap[r[1].toString().toLowerCase()] = r[0]; // Email -> Name
  });

  // 2. Get current availability data
  const range = sheet.getDataRange();
  const data = range.getValues();
  
  // 3. Identify rows to delete (Matching Month + Period)
  // We will rebuild rows for this Month+Period context to ensure clean state.
  // Note: This deletes ALL entries for this Period in this Month and recreates them.
  // This is safer than trying to diff row-by-row for multi-day entries.
  
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    // Cols: A=Month, C=Period
    if (data[i][0] === month && data[i][2] === period) {
      rowsToDelete.push(i + 1); // 1-based index
    }
  }
  
  // Batch delete is hard in Apps Script (indexes shift). 
  // Strategy: Clear content of rows, then sort/filter? No.
  // Strategy: Delete from bottom up.
  rowsToDelete.forEach(idx => sheet.deleteRow(idx));

  // 4. Rebuild Rows from dayUpdates
  // dayUpdates format: { "Mon": ["a@b.com", "c@d.com"], "Tue": ["a@b.com"] }
  // We want to group by Teacher to create multi-day rows if possible, 
  // OR just create single-day rows for simplicity?
  // The existing system seems to support "Mon,Tue" (comma separated).
  
  // Invert the map: TeacherEmail -> Set(Days)
  const teacherDays = {};
  Object.keys(dayUpdates).forEach(day => {
    const emails = dayUpdates[day]; // List of emails for this day
    emails.forEach(email => {
      const e = email.toLowerCase().trim();
      if (!teacherDays[e]) teacherDays[e] = new Set();
      teacherDays[e].add(day);
    });
  });

  // Create new rows
  const newRows = [];
  Object.keys(teacherDays).forEach(email => {
    const days = Array.from(teacherDays[email]).sort().join(','); // "Mon,Tue"
    const name = staffMap[email] || email; // Fallback to email if name not found
    
    // Cols: Month, Day(s), Period, Name, Email, Hours(empty)
    newRows.push([month, days, period, name, email, ""]);
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
  
  return true;
}
