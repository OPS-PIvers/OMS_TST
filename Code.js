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

// ... existing code ...

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
            <a href="${appUrl}" class="button">${buttonText}</a>
          </div>
        </div>
        <div class="footer">
          &copy; ${new Date().getFullYear()} Orono Middle School TST Manager<br>
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
