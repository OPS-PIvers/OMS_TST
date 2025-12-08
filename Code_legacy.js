/**
 * Middle School TST Time Tracking System - Period-Based
 * Adds a custom menu to the spreadsheet when it's opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('TST Time')
      .addItem('Process Approved Time', 'processApprovedTime')
      .addToUi();
}

/**
 * Adds a hidden "TeacherData" sheet to the spreadsheet.
 * @param {Spreadsheet} ss The spreadsheet to add the sheet to.
 */
function addTeacherDataSheet(ss) {
  const newSheet = ss.insertSheet('TeacherData');
  newSheet.appendRow(['Teacher Email', 'Sheet ID', 'Sheet Name']);
  newSheet.hideSheet(); // Hide the sheet.
}

/**
 * Processes approved TST time entries. Handles both "earn" and "use" requests
 * from the "TST Approvals" and "TST Usage" sheets.
 */
function processApprovedTime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName('TST Approvals');
  const usageSheet = ss.getSheetByName('TST Usage');
  const teacherDataSheet = ss.getSheetByName('TeacherData');

  // Add "TeacherData" sheet if it doesn't exist.
  if (!teacherDataSheet) {
    addTeacherDataSheet(ss);
  }

  const secretaryEmail = 'lisa.krebsbach@orono.k12.mn.us'; // Middle School Secretary's email.

  // --- Process EARNED TST Time ---
  processTSTRequests(ss, formSheet, teacherDataSheet, secretaryEmail, 'earn');

  // --- Process USED TST Time ---
  processTSTRequests(ss, usageSheet, teacherDataSheet, secretaryEmail, 'use');

  SpreadsheetApp.getUi().alert('TST Time processing complete!');
}

/**
 * Processes TST requests (either earning or using time).
 * This is a helper function to avoid code duplication.
 * @param {Spreadsheet} ss The active spreadsheet.
 * @param {Sheet} requestSheet The sheet containing the requests ("TST Approvals" or "TST Usage").
 * @param {Sheet} teacherDataSheet The "TeacherData" sheet.
 * @param {string} secretaryEmail The secretary's email address.
 * @param {string} requestType "earn" or "use" - indicates the type of request.
 */
function processTSTRequests(ss, requestSheet, teacherDataSheet, secretaryEmail, requestType) {
  if (!requestSheet) {
    return; // If the request sheet doesn't exist, do nothing.
  }

  // Get column indices from the *actual* header row (row 1).
  const headers = requestSheet.getRange(1, 1, 1, requestSheet.getLastColumn()).getValues()[0];

  // Get column indices based on sheet type
  let emailCol, dateCol, periodsCol, approvedCol, processedCol, subbedForCol = -1;
  
  if (requestType === 'earn') {
    // TST Approvals structure: A=Timestamp, B=Email, C=SubbedFor, D=Date, E=TimeSubbed, F=Amount, G=DecimalPeriods, H=Approved, I=Processed
    emailCol = headers.indexOf('Email Address');
    subbedForCol = headers.indexOf('I subbed for:');
    dateCol = headers.indexOf('Date Subbed:');
    periodsCol = headers.indexOf('Decimal Periods'); // This will be calculated
    approvedCol = headers.indexOf('Approved?');
    processedCol = headers.indexOf('Processed');
  } else {
    // TST Usage structure: A=Email, B=Date, C=PeriodsUsed, D=Approved, E=Processed
    emailCol = headers.indexOf('Email Address');
    dateCol = headers.indexOf('Date of Use');
    periodsCol = headers.indexOf('Periods Used');
    approvedCol = headers.indexOf('Approved?');
    processedCol = headers.indexOf('Processed');
  }

  // Check if the columns exist to prevent errors
  const requiredColumns = [emailCol, dateCol, periodsCol, approvedCol, processedCol];
  if (requestType === 'earn') {
    requiredColumns.push(subbedForCol);
  }
  
  if (requiredColumns.includes(-1)) {
    SpreadsheetApp.getUi().alert("Error. Could not find the correct Column Names in " + requestSheet.getName() + ".  Please ensure all required column headers exist in the " + requestSheet.getName() + " sheet.");
    return;
  }

  let formData = [];
  if (requestType === "earn") {
    formData = requestSheet.getRange("A2:I").getValues();
  } else {
    formData = requestSheet.getDataRange().getValues();
    formData.shift(); // remove headers
  }

  const teacherData = teacherDataSheet.getDataRange().getValues();
  teacherData.shift(); // remove the headers

  for (let i = 0; i < formData.length; i++) {
    const row = formData[i];
    if (row[approvedCol] && row[approvedCol].toString().toLowerCase() === 'yes' && row[processedCol] !== 'yes') {
      const teacherEmail = row[emailCol];
      let subbedFor = "";
      if (requestType === "earn") {
        subbedFor = row[subbedForCol];
      }
      const date = row[dateCol];
      let periods = Number(row[periodsCol]);

      // Adjust periods based on request type
      if (requestType === 'use') {
        periods = -periods; // Make periods negative for "use" requests
      }

      // Extract name from email
      const emailParts = teacherEmail.split('@');
      const nameParts = emailParts[0].split('.');
      const firstName = nameParts[0].charAt(0).toUpperCase() + nameParts[0].slice(1);
      const lastName = nameParts[1].charAt(0).toUpperCase() + nameParts[1].slice(1);
      const teacherName = `${firstName} ${lastName}`;

      // 1. Find or Create Teacher Sheet
      let teacherSheetId = null;
      let teacherSheet = null;
      let teacherSheetName = null;

      // Try to find existing sheet ID
      for (const teacherRow of teacherData) {
        if (teacherRow[0] === teacherEmail) {
          teacherSheetId = teacherRow[1];
          teacherSheetName = teacherRow[2];
          break;
        }
      }

      if (teacherSheetId) {
        // Existing Sheet
        try {
          teacherSheet = SpreadsheetApp.openById(teacherSheetId);
          const sheetUrl = teacherSheet.getUrl();
          
          // Find the correct sheet within the main spreadsheet
          const mainSheet = ss.getSheetByName(teacherSheetName);
          updateTeacherSheet(teacherSheet, date, subbedFor, periods, mainSheet);

          // Get total periods *after* updating
          const totalPeriods = teacherSheet.getSheets()[0].getRange('D2').getDisplayValue();

          // Send Email
          sendEmailNotification(teacherEmail, date, subbedFor, periods, sheetUrl, teacherName, totalPeriods, requestType);

          // Mark row as processed
          requestSheet.getRange(i + 2, processedCol + 1).setValue('yes');

        } catch (e) {
          console.log("Error opening sheet with id, " + teacherSheetId + " belonging to " + teacherEmail + ".  The sheet may have been deleted. Creating a new sheet.");
          teacherSheetId = null;
          teacherSheetName = null;
        }
      }

      if (!teacherSheetId) {
        // New Sheet
        const sheetInfo = createTeacherSheet(teacherEmail, secretaryEmail);
        teacherSheet = sheetInfo.spreadsheet;
        teacherSheetId = teacherSheet.getId();
        teacherSheetName = sheetInfo.sheetName;

        const sheetUrl = teacherSheet.getUrl();

        // Store sheet ID AND sheet name in TeacherData
        teacherDataSheet.appendRow([teacherEmail, teacherSheetId, teacherSheetName]);

        // Update Teacher Sheet
        const mainSheet = ss.getSheetByName(teacherSheetName);
        updateTeacherSheet(teacherSheet, date, subbedFor, periods, mainSheet);

        // Get total periods *after* updating
        const totalPeriods = teacherSheet.getSheets()[0].getRange('D2').getDisplayValue();

        // Send Email
        sendEmailNotification(teacherEmail, date, subbedFor, periods, sheetUrl, teacherName, totalPeriods, requestType);

        // Mark row as processed
        requestSheet.getRange(i + 2, processedCol + 1).setValue('yes');
      }
    }
  }
}

/**
 * Creates a new spreadsheet file for a teacher, shares it, and sets permissions.
 * Creates a corresponding tab in the main spreadsheet.
 * @param {string} teacherEmail The teacher's email address.
 * @param {string} secretaryEmail The secretary's email address.
 * @returns {Object} An object containing the new spreadsheet and the sheet name.
 */
function createTeacherSheet(teacherEmail, secretaryEmail) {
  const username = teacherEmail.split('@')[0];
  const sheetName = username.replace(/\./g, '_') + '_TST_Time';

  // Create a new spreadsheet file
  const newSpreadsheet = SpreadsheetApp.create(sheetName);

  // Get the first (and only) sheet in the new spreadsheet
  let newSheet = newSpreadsheet.getSheets()[0];
  newSheet.setName(sheetName);

  // Set headers for period-based tracking
  newSheet.appendRow(['Date', 'Subbed For', 'Periods', 'Total Periods']);

  // Add total periods formula to the individual teacher sheet
  newSheet.getRange('D2').setFormula('=SUM(C3:C2000)');

  // Format 'Periods' column as number
  newSheet.getRange('C:C').setNumberFormat('0.00');

  // Share the sheet
  newSpreadsheet.addViewer(teacherEmail);
  newSpreadsheet.addEditor(secretaryEmail);

  // Check for existing tab in main spreadsheet
  const mainSS = SpreadsheetApp.getActiveSpreadsheet();
  let mainTeacherSheet = mainSS.getSheetByName(sheetName);

  if (!mainTeacherSheet) {
    // Sheet doesn't exist, so create it
    mainTeacherSheet = mainSS.insertSheet(sheetName);
    mainTeacherSheet.appendRow(['Date', 'Subbed For', 'Periods', 'Total Periods']);
    mainTeacherSheet.getRange('D2').setFormula('=SUM(C3:C2000)');
  }

  // Conditional Formatting
  let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty()
      .setBackground("yellow")
      .setRanges([newSheet.getRange('D2')])
      .build();

  let rules = newSheet.getConditionalFormatRules();
  rules.push(rule);
  newSheet.setConditionalFormatRules(rules);

  // Move to the correct folder
  const folder = DriveApp.getFolderById("1OnvSvvvd1eEn_J3VKyx_axX8zzDUl2wt"); // Middle School folder
  const file = DriveApp.getFileById(newSpreadsheet.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return { spreadsheet: newSpreadsheet, sheetName: sheetName };
}

/**
 * Updates a teacher's TST time sheet with a new entry.
 * Updates both the individual sheet AND the main sheet tab.
 * @param {Spreadsheet} teacherSheet The teacher's individual spreadsheet.
 * @param {Date} date The date of the coverage.
 * @param {string} subbedFor The teacher who was covered.
 * @param {number} periods The number of periods covered.
 * @param {Sheet} mainSheet The corresponding sheet (tab) in the main spreadsheet.
 */
function updateTeacherSheet(teacherSheet, date, subbedFor, periods, mainSheet) {
  // Update the individual teacher sheet
  const sheet = teacherSheet.getSheets()[0];
  
  // Store the formula from D2 before making changes
  const totalFormula = sheet.getRange('D2').getFormula();
  
  // Insert new row at row 3 (after the header and the total)
  let dataRow = 3;
  sheet.insertRowBefore(dataRow);
  sheet.getRange(dataRow, 1, 1, 3).setValues([[date, subbedFor, periods]]);
  
  // Make sure D2 still has the total formula
  sheet.getRange('D2').setFormula(totalFormula);

  // Update the corresponding tab in the main spreadsheet
  if (mainSheet) {
    const mainTotalFormula = mainSheet.getRange('D2').getFormula();
    
    mainSheet.insertRowBefore(dataRow);
    mainSheet.getRange(dataRow, 1, 1, 3).setValues([[date, subbedFor, periods]]);
    
    mainSheet.getRange('D2').setFormula(mainTotalFormula);
  }
}

/**
 * Sends an email notification to the teacher.
 * @param {string} teacherEmail The teacher's email address.
 * @param {Date} date The date of the coverage or usage.
 * @param {string} subbedFor The teacher who was covered (or empty string for usage).
 * @param {number} periods The number of periods covered (positive for earn, negative for use).
 * @param {string} sheetUrl The URL of the teacher's sheet
 * @param {string} teacherName The Name of the Teacher
 * @param {number} totalPeriods The total number of periods.
 * @param {string} requestType "earn" or "use".
 */
function sendEmailNotification(teacherEmail, date, subbedFor, periods, sheetUrl, teacherName, totalPeriods, requestType) {
  let subject = '';
  let body = '';

  if (requestType === 'earn') {
    subject = 'TST Time Approved';
    body = `Dear ${teacherName},\n\n` +
           `Your TST time for ${date.toLocaleDateString()} covering ${subbedFor} for ${Math.abs(periods)} period(s) has been approved.\n\n` +
           `Your total accumulated TST time is now: ${totalPeriods} period(s).\n\n` +
           `You can view your TST Time sheet here: ${sheetUrl}`;
  } else if (requestType === 'use') {
    subject = 'TST Time Usage Approved';
    body = `Dear ${teacherName},\n\n` +
           `Your request to use ${Math.abs(periods)} period(s) of TST time on ${date.toLocaleDateString()} has been approved.\n\n` +
           `Your remaining TST time balance is: ${totalPeriods} period(s).\n\n` +
           `You can view your TST Time sheet here: ${sheetUrl}`;
  }

  MailApp.sendEmail({
    to: teacherEmail,
    subject: subject,
    body: body,
    name: "Lisa Krebsbach",
    replyTo: "lisa.krebsbach@orono.k12.mn.us"
  });
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
  if (selectedPeriod.includes('Period 6 ') && !selectedPeriod.includes('Period 6/')) {
    return 0.5;
  }
  if (selectedPeriod.includes('Period 7 ') && !selectedPeriod.includes('Period 6/')) {
    return 0.5;
  }
  
  // For all other periods (including Period 6/7), use Full/Half logic
  if (amount && amount.toLowerCase().includes('full')) {
    return 1.0;
  } else if (amount && amount.toLowerCase().includes('half')) {
    return 0.5;
  }
  
  // Default fallback
  return 1.0;
}

/**
 * Runs automatically when a form is submitted.
 * Calculates and populates the decimal periods based on period selection and full/half choice.
 */
function onFormSubmit(e) {
  try {
    // Validate event object first
    if (!e || !e.source) {
      console.error("onFormSubmit trigger did not receive valid event object.");
      return;
    }
    
    const ss = e.source;
    const approvalSheet = ss.getSheetByName('TST Approvals');
    const responseSheet = ss.getSheetByName('Form Responses 1');

    if (!approvalSheet || !responseSheet) {
      console.error("Could not find 'TST Approvals' or 'Form Responses 1' sheet.");
      return;
    }

    // Get the last submitted row
    const lastFormRow = responseSheet.getLastRow();
    if (lastFormRow < 2) return; // No data rows yet

    // Get the period selection and amount from the form
    const periodSelection = responseSheet.getRange(lastFormRow, 6).getValue(); // Column F: Time Subbed
    const amountSelection = responseSheet.getRange(lastFormRow, 7).getValue(); // Column G: Amount of Time Subbed

    console.log(`Period Selection: ${periodSelection}, Amount: ${amountSelection}`);

    // Calculate periods based on Middle School rules
    const calculatedPeriods = calculatePeriods(periodSelection, amountSelection);

    // Set the calculated value in column H
    responseSheet.getRange(lastFormRow, 8).setValue(calculatedPeriods);
    responseSheet.getRange(lastFormRow, 8).setNumberFormat("0.00");
    
    console.log(`Calculated periods: ${calculatedPeriods}`);

    // Apply dropdown validation to TST Approvals sheet
    const timestampFromForm = responseSheet.getRange(lastFormRow, 1).getValue();
    if (!timestampFromForm) {
      console.warn("No timestamp found in form response");
      return;
    }
    
    const lastApprovalRow = approvalSheet.getLastRow();
    if (lastApprovalRow < 2) {
      console.log("No data in approval sheet to match with");
      return;
    }
    
    // Find matching row in approval sheet
    const approvalDataRange = approvalSheet.getRange("A2:A" + lastApprovalRow);
    const approvalTimestamps = approvalDataRange.getValues();
    let targetApprovalRow = -1;

    for (let i = 0; i < approvalTimestamps.length; i++) {
      const approvalTimestamp = approvalTimestamps[i][0];
      if (!approvalTimestamp) continue;
      
      if (approvalTimestamp instanceof Date && timestampFromForm instanceof Date) {
        if (Math.abs(approvalTimestamp.getTime() - timestampFromForm.getTime()) < 1000) {
          targetApprovalRow = i + 2;
          break;
        }
      } else if (approvalTimestamp.toString() === timestampFromForm.toString()) {
        targetApprovalRow = i + 2;
        break;
      }
    }

    // Apply dropdown validation if we found a match
    if (targetApprovalRow > 1) {
      const rule = approvalSheet.getRange('H2').getDataValidation(); // Approved? column
      if (rule) {
        approvalSheet.getRange(targetApprovalRow, 8).setDataValidation(rule); // Apply to Approved? column
        console.log("Applied dropdown to row " + targetApprovalRow);
      } else {
        console.warn("No data validation rule found in H2");
      }
    } else {
      console.warn("No matching row found in approval sheet for timestamp: " + timestampFromForm);
    }
    
  } catch (error) {
    console.error("Error in onFormSubmit: " + error.message);
  }
}
function debugHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TST Approvals');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  console.log('Headers found:');
  for (let i = 0; i < headers.length; i++) {
    console.log(`Column ${i + 1}: "${headers[i]}"`);
  }
}

function sortTabsWithExceptions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var otherTabs = [];
  
  // 1. Define the tabs to keep at the front (IN ORDER)
  var pinnedTabs = [
    "Staff Directory", 
    "Form Responses 1", 
    "TST Approvals (New)", 
    "TST Approvals"
  ];

  // 2. Find all tabs that are NOT in the pinned list
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    // If the name is not in our pinned list, add it to otherTabs
    if (pinnedTabs.indexOf(name) === -1) {
      otherTabs.push(name);
    }
  }
  
  // 3. Sort the "other" tabs alphabetically
  otherTabs.sort();
  
  // 4. Combine the lists: Pinned first, then Sorted
  var finalOrder = pinnedTabs.concat(otherTabs);
  
  // 5. Move tabs to their new positions
  for (var j = 0; j < finalOrder.length; j++) {
    var sheet = ss.getSheetByName(finalOrder[j]);
    // Only move if the sheet actually exists (avoids errors if you deleted one)
    if (sheet) {
      sheet.activate();
      ss.moveActiveSheet(j + 1);
    }
  }
}
