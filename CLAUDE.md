# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OMS TST Manager is a Google Apps Script web application for managing TST (Time Saving Team/Teacher) hours at Orono Middle School. Staff can earn time by subbing for colleagues and redeem accumulated credits. The application features role-based access for Admins and Teachers.

## Technology Stack

- **Platform:** Google Apps Script (V8 Runtime)
- **Backend:** [Code.js](Code.js) - Server-side logic using Google Apps Script APIs
- **Frontend:** [Index.html](Index.html) - Single Page Application (SPA)
- **Styling:** Tailwind CSS v3.x (via CDN)
- **Icons:** FontAwesome v6.x (via CDN)
- **Database:** Google Spreadsheet (accessed via SpreadsheetApp)
- **Deployment:** clasp (Command Line Apps Script Projects)

## Development Commands

### Prerequisites
- Node.js and npm installed
- `@google/clasp` installed globally: `npm install -g @google/clasp`
- Logged in to clasp: `clasp login`

### Common Commands

```bash
# Push local changes to Google Apps Script
clasp push

# Pull remote changes from Google Apps Script
clasp pull

# Open the project in browser (Apps Script editor)
clasp open

# Deploy a new version to the existing deployment
clasp deploy -i AKfycbzPvaCCovRLEUVSe05KfRaDlXEs9k64oMCtpcXdOnYzVpP2BW16PaXV5SJVHNk3Ea3TBQ --description "Version description"
```

**Important:**
- `clasp push` overwrites remote files completely. Always verify changes before pushing.
- Always use the deployment ID (`-i` flag) to update the existing web app deployment rather than creating a new one.

## Architecture

### Client-Server Communication

All client-server communication uses `google.script.run`:

```javascript
// Client-side call
google.script.run
  .withSuccessHandler(callback)
  .withFailureHandler(errorHandler)
  .serverFunction(args);
```

Server-side functions in [Code.js](Code.js) are directly callable from the client.

### State Management

The frontend uses a global `STATE` object to manage application state:

```javascript
STATE = {
  user: { email, name, role, staffData },
  currentTab: 'admin-earned' | 'admin-used' | 'teacher-totals' | 'teacher-history',
  dataCache: {}
}
```

### Data Model (Google Sheets)

The application relies on four key sheets within the bound Google Spreadsheet:

1. **Staff Directory** - User directory with roles and balances
   - Columns: Name (A), Email (B), Role (C), Earned (D), Used (E), Carry Over (G)
   - Role determines access level: "Admin" or "Teacher"

2. **TST Approvals (New)** - Pending and processed earned requests
   - Columns: Email (A), Name (B), SubbedFor (C), Date (E), Period (F), TimeType (G), Hours (H), Approved (I), ApprovedTS (J), Denied (K), DeniedTS (L), DenialReason (M)
   - Filtering logic: Pending = NOT Approved AND NOT Denied

3. **TST Usage (New)** - Pending and processed usage requests
   - Columns: Email (A), Name (B), Date (C), Used (D), Status (E), Timestamp (F)
   - Filtering logic: Pending = Status is false/empty

4. **Form Responses 1** - Raw Google Form submissions
   - Source of truth for form-based earned requests
   - Synced to TST Approvals (New) via `onFormSubmit()` trigger

### Key Server Functions

- **getInitialData()** - Authenticates user, determines role, returns initial app data
- **getDashboardCounts()** - Returns pending request counts for admin badges
- **getPendingEarned() / getPendingUsed()** - Fetch pending requests for admin view
- **getTeacherHistory(email)** - Fetch complete history for a teacher
- **approveEarnedRow(rowIndex, emailData)** - Approve earned request, optionally send email
- **denyEarnedRow(rowIndex, emailData)** - Deny earned request with reason
- **submitEarned(formObj) / submitUsage(formObj)** - Create new requests
- **batch* functions** - Process multiple actions at once (batchApproveEarned, batchDenyEarned, etc.)
- **sendStatusEmail(email, name)** - Generate and send TST report email

### Row Index Convention

**Critical:** All row operations use 1-based indices where index 1 is the header row. When working with arrays from `.getValues()`:
- Remove header with `.shift()` before processing
- Map rows to objects with `rowIndex: i + 2` (array index + 2 accounts for 0-based array and header row)
- Delete operations must process indices in **descending order** to preserve row positions

### Frontend Architecture

[Index.html](Index.html) is ~1900 lines containing:
- HTML structure with Tailwind utility classes
- Custom Tailwind config with OPS brand colors (`ops-blue`, `ops-red`, etc.)
- Client-side JavaScript for UI rendering and state management
- Modal system for forms and confirmations
- Toast notification system

### Role-Based Views

**Admin Views:**
- `admin-earned`: Pending TST Approvals with multiselect for batch approve/deny
- `admin-used`: Pending TST Usage with multiselect for batch approve/delete
- `admin-reports`: Staff directory with balance overview and email functionality

**Teacher Views:**
- `teacher-totals`: Personal balance and quick submission forms
- `teacher-history`: Complete transaction history (earned, used, denied)

### Form Validation & Safety

- Critical actions (Delete/Deny) require user confirmation via modals
- Email notifications are optional for batch operations (performance optimization)
- Denial reasons support predefined options + custom notes
- Database writes target specific column indices - be careful when changing sheet structure

## Custom Color Palette

Defined in Tailwind config within [Index.html](Index.html):
- `ops-blue`: #2d3f89 (primary brand color)
- `ops-blue-dark`: #1d2a5d
- `ops-blue-lighter`: #eaecf5 (backgrounds)
- `ops-red`: #ad2122 (alerts, denials)
- `ops-red-lighter`: #e5c7c7 (warning backgrounds)

## Important Conventions

### Editing Earned Requests
When updating earned requests via `updateEarnedRow()`, **both** sheets must be synchronized:
1. Update "TST Approvals (New)" directly
2. Find matching row in "Form Responses 1" using email + date + period
3. Update the form response row to maintain data consistency

### Delete vs Deny
- **Deny**: Sets Denied flag, preserves record, optionally sends email with reason
- **Delete**: Removes row from BOTH "TST Approvals (New)" AND "Form Responses 1"

### Period Formatting
Class periods are stored as full strings (e.g., "Period 1 - 8:10 - 8:57"). Reference [class_periods.md](class_periods.md) for the complete schedule. Legacy forms may use short numbers ("1") - the `getPeriodOptions()` function handles both formats.

### Email Templates
Styled HTML emails are sent via `sendStyledEmail()` with:
- Branded header (Orono Middle School)
- Color-coded content based on action type
- Direct link to web app
- Consistent footer with OPS branding

### Triggers
The `onFormSubmit(e)` function must be set up as an **installable trigger** in the Apps Script editor. This syncs Google Form submissions into the approval workflow.

## Testing & Debugging

Since this is a Google Apps Script project, traditional unit testing is limited. For debugging:

1. Use `Logger.log()` in [Code.js](Code.js) - view logs in Apps Script editor (Ctrl+Enter to run functions)
2. Use `console.log()` in [Index.html](Index.html) - view in browser console
3. Test with different user roles by modifying entries in "Staff Directory"
4. Use "Executions" tab in Apps Script editor to view runtime logs and errors

## OAuth Scopes

Defined in [appsscript.json](appsscript.json):
- `https://www.googleapis.com/auth/spreadsheets` - Read/write spreadsheet data
- `https://www.googleapis.com/auth/script.send_mail` - Send emails via MailApp
- `https://www.googleapis.com/auth/userinfo.email` - Get active user email

## Deployment Notes

- Web app execution: "User Deploying" (runs as the deployer, accesses deployer's sheets)
- Access level: "DOMAIN" (restricted to domain users only)
- Timezone: America/Chicago
- After `clasp push`, you may need to redeploy via `clasp deploy` or the Apps Script editor for changes to take effect in the web app
