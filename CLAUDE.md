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
- [.claspignore](.claspignore) is an **allow-list**: only `Code.js`, `config.js`, `Index.html` and `appsscript.json` are pushed. A new server-side file is not deployed until it is added there, and non-Apps-Script files (the Node tests, [Code_legacy.js](Code_legacy.js)) must stay out — Apps Script shares one global scope across a project's files and evaluates them in order, so a stray file's top-level code runs on every execution and its duplicate function names shadow the real ones.

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
   - Columns: Name (A), Email (B), Role (C), Earned (D), Used (E), Carry Over (F), Paid Out (G), Running Total (H, ARRAYFORMULA), Building (I), Archived (J), Last Finalized (K), Pending Finalize (L)
   - Role determines access level: "Teacher", "Admin", or "Super Admin"
   - Building (I) supports comma-separated multi-building assignment (e.g. `OMS, OHS`). **The first building listed is the PRIMARY building** (existing implicit convention — `getUserContext` uses `assignedBuildings[0]`).
   - Archived (J): **per-building** soft-delete — a comma-separated list of buildings the person is archived FROM. Archiving from one building doesn't affect others; "fully archived" = archived from all assigned buildings. Auto-created if missing.
   - Last Finalized (K): school-year name of the last year-end roll for this person; prevents double-rolling staff who span buildings. Auto-created if missing.
   - Pending Finalize (L): school-year name set when a **non-primary** building finalizes a shared person (primary elsewhere); drives the "Pending finalize" UI chip and is cleared when the primary building finalizes them. Auto-created if missing.
   - When adding rows programmatically, set individual cells and leave H blank so the ARRAYFORMULA fills it (do NOT use `appendRow`).

   **Ownership model (combined totals + primary building):**
   - **Earned hours** may be contributed by **any** building the person works in (each transaction is tagged with its building: earned col N/13, used col H/7). Directory/Totals balances and a teacher's own summary are **combined across all of a person's buildings** — `staffDirectoryData_` computes balances via `calculateDynamicBalances_(null, …)`; `buildingFilter` is used only for directory membership and per-building archived logic.
   - **Carry Over, Paid Out, and finalize** are **owned by the PRIMARY building's admin**. `isPrimaryAdminFor_(ctx, buildingCell)` gates the owned-column writers (`updateStaffBatch`, `updateStaffCarryOver`, `updateStaffMember`): Super Admins always pass; otherwise the admin must be assigned to that person's primary building. Non-primary admins see those columns read-only but can still submit/approve **Earned** for their building.
   - **Carry Over cap (`carryOverMax`)** is a per-building config value (default 12) editable **only by Super Admins** in Settings; `saveBuildingConfig` preserves it for non-Super-Admins regardless of payload. It caps the Carry Over rolled at finalize (excess is forfeited) — manual Carry Over edits are **not** capped.
   - **Approval queues + dashboard counts stay building-scoped** (`getDashboardCounts`, `getPendingEarned/Used`) — do not make these combined.

2. **TST Approvals (New)** - Pending and processed earned requests
   - Columns: Email (A), Name (B), SubbedFor (C), Date (E), Period (F), TimeType (G), Hours (H), Approved (I), ApprovedTS (J), Denied (K), DeniedTS (L), DenialReason (M)
   - Filtering logic: Pending = NOT Approved AND NOT Denied

3. **TST Usage (New)** - Pending and processed usage requests
   - Columns: Email (A), Name (B), Date (C), Used (D), Status (E), Timestamp (F)
   - Filtering logic: Pending = Status is false/empty

4. **Form Responses 1** - Raw Google Form submissions
   - Source of truth for form-based earned requests
   - Synced to TST Approvals (New) via `onFormSubmit()` trigger

5. **TST Assignments** - Coverage the admin has assigned (auto-created)
   - Columns: ID (A), Created (B), Building (C), Assigned By (D), Assigned By Name (E), Sub Email (F), Sub Name (G), Covered For (H), Covered For Email (I), Date (J), Period (K), Time Type (L), Hours (M), Note (N), Note To Sub (O), Note To Covered (P), Status (Q), Recorded TS (R), Recorded By (S), Nudged TS (T), Calendar Event ID (U), Calendar Status (V), Notified TS (W), Cancelled TS (X), Cancelled By (Y)
   - Column indexes live in `A_` and the status values in `ASSIGNMENT_STATUS_` (`Assigned` → `Recorded` | `Cancelled`); use those rather than literals.
   - The row is written **when the admin assigns**, not when the teacher acts. That is what makes the app the record: the Assignments queue, Cancel/Reassign/Remind, Record-on-behalf, the nudge and the duplicate guard all read it, and it is why an un-acted-on assignment is still visible.
   - Covered For Email is blank for free-text entries ("Activity Bus"), which is the only reason the covered-for notification is ever skipped.
   - Calendar Status drives the calendar state machine (`CAL_`): blank (no calendar) → `Pending` → `Created` | `Failed: <reason>`, and `Pending Delete` → `Deleted`. Notified TS records when the assignment emails actually went out — a cancellation checks it, because emailing someone that a coverage is off when they were never told about it is worse than saying nothing. Cancelled TS / Cancelled By are written by `cancelAssignment`; they exist so the re-assign warning can say *who* called it off and *when*, which is what tells an admin whether they cancelled it themselves by mistake. `assignmentsSheet_` fills a short header out on its own, so a sheet from an earlier version migrates without anyone touching it.
   - Rows archive to **TST Assignments Archive** at year-end, by building (`archiveAssignmentsForBuilding_`).

### Key Server Functions

- **getInitialData()** - Authenticates user, determines role, returns initial app data
- **getDashboardCounts()** - Returns pending request counts for admin badges (admin only)
- **getPendingEarned() / getPendingUsed()** - Fetch pending requests for admin view (admin only)
- **getTeacherHistory(email)** - Fetch complete history for a teacher
- **approveEarnedRow(rowIndex, emailData)** - Approve earned request, optionally send email
- **denyEarnedRow(rowIndex, emailData)** - Deny earned request with reason
- **submitEarned(formObj) / submitUsage(formObj)** - Create new requests
- **batch* functions** - Process multiple actions at once (batchApproveEarned, batchDenyEarned, etc.)
- **sendStatusEmail(email, name)** - Generate and send TST report email
- **assignCoverage(payload)** - Assign coverage: writes the TST Assignments row, then queues three emails (the person covering, the person being covered for, the admin's record copy)
- **getAssignments(building) / getMyAssignments(email)** - The admin queue, and a teacher's own (both the coverage they are providing and the coverage arranged for their classes)
- **recordAssignment(id, email) / cancelAssignment(id) / remindAssignment(id)** - Record, cancel, re-send
- **nudgeOutstandingAssignments()** - The one automatic 7am reminder, installed by `setupEmailService`
- **sendTestCalendarEvent(building) / getCalendarTestResult(building)** - Queue and then read the end-to-end calendar check behind Settings' "Send Test Event"
- **getEmailServiceStatus(building)** - Whether that building's mail is actually going out; drives the warning banner
- **setAuthorizeUrl(url)** - **Super Admin only.** The address of the second (user-accessing) deployment

#### Server-side authorization (every public function is callable via `google.script.run`)

Any signed-in domain user can call any public (non-`_`) function from the browser console, so each endpoint checks on the server. Public wrappers authorize, then call a private `_` variant; internal code calls the `_` variants directly. Apps Script does not expose functions ending in `_` to the client.
- **Request rows** (approve/deny/delete/edit/revert Earned & Used, and the `batch*` versions): `authorizeRequestRows_(ctx, 'earned'|'used', rows)` — Admin/Super Admin; row must be an integer within the sheet (>= 2); the row's building (Approvals col N, Usage col H, blank = OMS) must be one of the caller's buildings (Super Admins: any). Batches authorize **every** row before applying anything.
- **Acting for staff** (`processBatch`, `adminSubmitRequest`, `sendStatusEmail`, `sendBatchStatusEmails`, `assignCoverage`): `assertCanManageStaffEmails_(ctx, emails)` — admin sharing a building with each person (Super Admins: anyone); all checked up front. `assignCoverage` checks the person covering **and** the person being covered for.
- **Self or manager** (`submitEarned`, `submitUsage`, `getTeacherHistory`, `getStaffHistoryWithActions`): `assertSelfOrManagerOf_(email)` — the session user's own email, or an admin who manages that person (covers View As).
- **Assignments** (`getAssignments`, `cancelAssignment`, `remindAssignment`): `assertCanManageAssignment_(ctx, a)` — Admin/Super Admin, and the assignment's building must be one of the caller's (Super Admins: any). `getAssignments` resolves its building through `allowedBuildingFor_` like every other queue, so a bad filter falls back to the caller's own building rather than returning the district.
- **Recording** (`recordAssignment`): `assertSelfOrManagerOf_(target)` — the teacher's own, or an admin who manages them (which is Record-on-behalf). The target must match the assignment's Sub Email, so a valid session cannot record someone else's row. `recordAssignment_` **claims the row before writing the earned request** (status → `Recorded`, then `submitEarned_`, rolling back on failure), the same way `processEmailQueue_` claims a queue row — that is the duplicate guard, and a double-click or re-opened link cannot produce two requests.
- **Record links**: `sendAssignmentEmails_` builds them with `buildAssignmentLink_`, HMAC-SHA256-signed with `COVERAGE_LINK_SECRET` in Script Properties (auto-created). The signature covers only `action`, `id` and `tEmail` — every other detail is read from the assignment row, so nothing about the coverage can be forged from a URL. `doGet` rejects links failing `verifyAssignmentLink_`; `handleAssignmentRecord_` also requires the signed-in user to be the invited teacher, and refuses a cancelled, already-recorded or **expired** assignment (`ASSIGNMENT_LINK_DAYS_`, 14 days past the coverage date).
  - Every emailed link comes from `scriptUrl_()`, which returns the **main deployment's fixed URL** (`WEB_APP_URL_`). Don't switch it back to `ScriptApp.getService().getUrl()`: emails go out from the admin's trigger, where that call hands back the second ("user accessing") deployment, and a teacher opening that sees Google's "Sorry, unable to open the file at this time". If the main deployment ID ever changes, update `WEB_APP_DEPLOYMENT_ID_` along with `deploy.yml`.
  - The email template styles are **inline** (the `<style>` block is kept only as a fallback). Gmail and Outlook drop `<style>`, which is what turned the button into a bare link.
  - The **in-app** Record button deliberately has no expiry: the emailed link is a bearer token sitting in an inbox, the button is an authenticated action on a row the admin created. Don't "fix" this inconsistency.
  - Old `accept` / `reject` links from the request-era flow are answered with a plain "no longer valid" page rather than a forged-link error.
- **Reads are scoped too** — the queues and the directory are as sensitive as the writes:
  - **Admin queues** (`getPendingEarned`, `getPendingUsed`, `getDashboardCounts`): `assertAdmin_` plus `allowedBuildingFor_`. The counts are derived from the same `pendingEarnedFor_`/`pendingUsedFor_` helpers as the lists, so a badge can never disagree with the queue under it.
  - **Directory** (`getStaffDirectoryData`, and the `staffData` in `getInitialData`): `directoryFor_(ctx, building, …)` decides what comes back — an admin gets the building's full table, a Teacher gets their own row plus **name/email only** for colleagues (the Submit forms need to name who was covered), and anyone not in the directory gets `[]`. The unrestricted reader is `staffDirectoryData_` (private).
  - **Schedule** (`getScheduleData`): admins get the building's grid; a teacher gets only their own availability rows (the teacher Schedule tab renders nothing else).
  - A missing or unauthorized `buildingFilter` **falls back to the caller's own building** rather than returning the district — no public read is district-wide.
- **`allowedBuildingFor_(ctx, requested)`** is the one building rule for reads and schedule writes: no request (or an unknown code) means the caller's own building; Super Admins may name any configured building; everyone else may name one of their own assigned buildings (so a multi-building admin who switched school sees that building's queue). Returns `null` when the request isn't allowed — read endpoints fall back, `updateSchedulePeriod` throws.
- **Maintenance entry points**: `syncMissingSubmissions` is admin-only (run from the Apps Script editor). `processEmailQueue` and `nudgeOutstandingAssignments` must stay public for the triggers `setupEmailService` installs, so they allow trigger invocations (`isTriggerEvent_` compares `e.authMode` against the real `ScriptApp.AuthMode` enum, which a `google.script.run` payload cannot carry) and require an admin otherwise. `setupEmailService` is admin-only.
- `onFormSubmit` stays public for its trigger but refuses events without a live `Range` (i.e. calls from the client).
- The staff detail modal locks row actions for requests filed under a building the admin doesn't manage (`canActOnRequestBuilding`), mirroring the server.
- Aggregation helpers that read across the whole district are private: `calculateDynamicBalances_`, `calculateMonthlyHours_`, `getPendingEarnedMap_`.

#### Staff Management (admin-only, self-authorizing via `getUserContext`)

- **addStaffMember(data)** - Add a staff member. `data = { name, email, role, building, carryOver, paidOut }`. Building accepts comma-separated codes. Writes individual cells (never `appendRow`) to preserve the column-H ARRAYFORMULA.
- **updateStaffMember(email, data)** - Edit a staff member matched by (unchanged) email. Email is never rewritten because it keys all transaction history.
- **archiveStaffMember(email, building) / restoreStaffMember(email, building)** - Add/remove a building from the per-building Archived (J) list. Defaults to the caller's current building.
- **deleteStaffMemberPermanent(email)** - **Super Admin only.** Deletes the spreadsheet row; only permitted for fully-archived staff (archived from every assigned building).
- **getArchivedStaff()** - **Super Admin only.** Returns fully-archived staff (for the Settings permanent-delete list).
- **getStaffDirectoryData(buildingFilter, targetEmail, includeArchived)** - Reads the directory through `directoryFor_` (see Server-side authorization: admins get the full table, teachers a name/email roster plus their own row). Earned/Used are **combined across all buildings** (`buildingFilter` only controls membership + the per-building `archived` flag). Also returns `primaryBuilding`, `pendingFinalize` (bool) and `pendingFinalizeYear`. Archived staff are excluded unless `includeArchived` is true. Server-side callers use the private `staffDirectoryData_(…)`, which has the same signature and no role check.
- **isPrimaryAdminFor_(ctx, buildingCell) / assertPrimaryAdminFor_(…)** - Gate owned-column edits (Carry Over / Paid Out). True for Super Admins, or when the caller is assigned to the staff member's primary (first-listed) building.

#### Coverage Assignments

Coverage is **assigned**, not requested — there is no accept/decline handshake. A teacher who is picked is expected to cover, and a genuine conflict goes to their administrator directly. Don't reintroduce a decline path (including a Google Calendar RSVP prompt); it was removed deliberately, because an easy "no" was being used to opt out in the moment.

- **Assign** from the Schedule grid (hover a teacher's cell) or via Reassign. The modal mirrors the teacher's own Submit form rather than hardcoding Full/Half Period: a `periods` building gets its configured `coverageTypes`, a `time_range` building gets start/end pickers with hours from the span (matching how OIS/SE record time everywhere else).
- **Duplicate guard:** a **live** assignment for the same person, date and period is refused outright — no confirmation makes two of the same coverage correct. Two other cases are questions rather than refusals, and both come back as `{ conflict: true, existing: [...], cancelled: [...] }` (the block-with-override shape `finalizeSchoolYear` uses); the client confirms and re-sends with `force: true`:
  - `existing` — a *different* period on the same date, which is legitimate.
  - `cancelled` — this exact date and period was assigned to this person and then **cancelled**. Putting someone straight back on a period they were just let off is usually a slip, but it is also precisely how an admin undoes an accidental cancellation, so it warns rather than blocks. Sorted most-recently-cancelled first, and carries `cancelledBy` / `cancelledOn` so the dialog can name them. A cancellation for a different period or date does not warn — that was never the coverage being called off.
- **Cancel** removes a *pending* earned request and emails both staff. If the hours are already **approved** it refuses and points at the existing Revert flow — approved hours are never clawed back silently.
- **Reminders:** one automatic nudge at ~7am the day after the coverage date (`nudgeOutstandingAssignments`, marked via Nudged TS so it fires once), plus a manual Remind button. A manual reminder counts as the one nudge.
- **Badges:** the admin Assignments badge and the teacher's Submit badge both count only **past-date, not-yet-recorded** coverage. Upcoming assignments are not actionable, so counting them would leave a permanent number on the tab. The admin count comes from `assignmentsFor_`, the same helper as the list, so badge and list cannot disagree.
- **Emails all go through the queue** (`addToEmailQueue_`), which is what makes the **building's own admin** the sender. Nothing about an assignment may call `MailApp.sendEmail` directly — that was the bug in the old `sendCoverageRequest`, which sent as the deployer.

#### Period times, and how a period reads to a person

- **Never put a stored time in front of someone.** Times are stored 24-hour; `periodDisplay_(building, period, date)` produces what a person should read — `Period 3 (10:24 AM – 11:04 AM)`. It strips the times out of an OMS label rather than repeating them, and returns just the span for a time-range building whose period *is* the span. Use it in every email, list and calendar description. `a.period` on its own is a storage key, not a label for humans.
- `periodTimesFor_(building, period, date)` resolves in order: **day-group override → building default (`periodTimes`) → times written into the label**. It returns `null` when nothing is configured, and that `null` is meaningful — the calendar reports it instead of inventing a time.
- **OHS runs two bell schedules**: Mon/Wed/Fri are the defaults, Tue/Thu is a `dayGroups` entry. **Spartan Hour only exists on Tue/Thu**, so it deliberately has no default time; assigning it on a Monday reports "no time set for that day". Do not add a Mon/Wed/Fri default to silence that — it would put someone on a calendar at a time that does not exist.
- `parseTimeRange_` reads a **one-digit hour as 12-hour** and a two-digit one as 24-hour. That is what keeps `Period 8 - 12:37 - 1:08` from becoming a twelve-hour event while leaving an `<input type="time">` value alone.
- `config.js` seeds App Config **only when the sheet does not exist**, so editing it does nothing to a district already running — a schedule added to `config.js` later never appears, and Settings shows empty time fields while the times are perfectly well defined in the source. **Settings → Periods → "Load Built-In Schedule"** (`installBellSchedule(building)`) copies them across, leaving calendar ID, carry-over cap and name untouched.
  - `installBellSchedule` **requires an explicit building** rather than falling back to the caller's own the way reads do. It replaces a building's periods, and the Apps Script editor's Run button passes no arguments, so a default would quietly rewrite the wrong school.

#### The TST Calendar (per building)

Each building has its own calendar, and **a blank `calendarId` means that building has no calendar** — no event, no calendar sentence in any email, no failure alerts. There is no separate on/off switch, so the two can never disagree. `calendarName` is typed by the admin and is what the emails say ("added to the OMS TST Calendar"); `calendarNameFor_` falls back to "<Building> TST Calendar".

**The building's own admin creates the event, through their trigger** — not the web app. The app runs as the deployer (`executeAs: USER_DEPLOYING`), so anything it created would be owned by the deployer rather than by the admin whose calendar it is. `processEmailQueue` therefore runs `processPendingAssignments_()` before draining the mail queue, scoped to the trigger owner's buildings.

That ordering is deliberate and load-bearing:

- `assignCoverage` writes the row with `Calendar Status = Pending` and **sends nothing**. The admin's next trigger run creates the event, then sends the three emails — so the "added to the ... TST Calendar" line only ever appears when there is really something to look at. A building with no calendar skips all of this and emails immediately.
- A failure never blocks the assignment. The emails still go (minus the calendar sentence), the row records `Failed: <reason>`, and `alertCalendarFailure_` emails the admin. A missing period time is reported as such rather than guessed at — see `periodTimesFor_`.
- **Cancelling** marks `Pending Delete`; the trigger removes the event and *then* sends the cancellation emails, so they never claim a removal that has not happened.
- Guests are added with **`sendInvites: false`**. Google's own invite carries a Yes/No/Maybe prompt, which would put a decline button back into a flow that deliberately has none. Do not turn it on.
- **Settings → Send Test Event** (`sendTestCalendarEvent`) queues a job the building admin's trigger picks up, creating and removing a throwaway event. That is the only check that proves the calendar ID, the admin's edit rights *and* their trigger together; validating the ID from Settings would run as the deployer and prove none of it. The result comes back through Script Properties (`CAL_TEST_<building>`) and the client polls `getCalendarTestResult`.

**Setup each building admin needs:** "Make changes to events" on their building's calendar, plus the trigger from **TST Admin → Authorize Email Service**. A building whose admin has no trigger will leave assignments sitting at `Pending` and send nothing — the stalled-queue banner is what surfaces that.

#### Year-End Finalize (primary-aware)

- **finalizeSchoolYear(yearName, building, force)** - For staff whose **primary** building is `building`: snapshots their **combined** totals into a new metadata-tagged sheet, rolls their combined remaining balance into Carry Over and zeros Paid Out (once per person per year, via Last Finalized), and archives **all** of their approved transactions across **every** building so combined Earned/Used reset to 0. Staff assigned here but whose primary is elsewhere are **not** changed — they get a Pending Finalize flag (cleared when their primary finalizes). Building admins finalize their own building; Super Admins may pass a building.
  - **Carry Over cap:** the rolled Carry Over is capped at the building's `carryOverMax` (default 12; see `carryOverMaxFor_`). Anything above the cap is **forfeited**. The snapshot sheet records `New Carry Over` and `Forfeited` columns for audit.
  - **Block-with-override:** when `force` is falsy and any primary-here staff would roll **above** the cap, the function makes **no changes** and returns `{ blocked: true, cap, overCap: [{name, email, projected, over}] }` so the UI can list the offending staff. Call again with `force = true` to proceed and forfeit the excess. On success returns `{ building, name, count, pending, cap, forfeitedCount }`.
- **archiveTransactionsByEmails_(emailsLower, yearName)** - Archives every approved transaction (all buildings) for a set of staff emails (used by primary finalize). The older `archiveBuildingTransactions_`/`archiveRowsByBuilding_` (by-building) are kept for backward compatibility.
- **listArchivedYears(building) / getArchivedYearData(sheetName)** - List and read the year-end snapshot sheets (scoped to the caller's building; identified by developer metadata, not name).

#### View As (admin → teacher)

- **getViewAsData(targetEmail, building)** - Returns a `getInitialData`-shaped payload for a **Teacher** so an admin can use the app exactly as that teacher. Requires Admin/Super Admin; non-Super-Admins must share a building with the target (`assertCanManageRow_`), and the returned `buildings` are limited to the buildings they share. Admin/Super Admin targets are rejected. Logs `[View As]` to the execution log.
- **saveAvailability(month, list, targetEmail, periodsShown)** - `targetEmail` is passed only during View As (same authorization as above); otherwise the session user is used. `periodsShown` limits which of the person's rows for that month get replaced: availability rows aren't building-tagged and buildings name periods differently, so a multi-building teacher saving one building's grid keeps their rows for the other building.
- **getScheduleData(buildingFilter)** - Membership matches the Directory (`staffDirectoryData_(building)`): staff assigned to the building anywhere in a multi-building list and not archived from it, compared by lowercased email. Hours/pending lookups are also lowercased; pending requests use the same building as the schedule. The building rule lives in `allowedBuildingFor_(ctx, requested)` and membership in `scheduleMembers_(building)`; the grid itself is built by `scheduleData_(building, onlyEmail)`, and a Teacher caller is limited to their own rows.
  - The number under each name is **approved hours earned so far this school year**, combined across the person's buildings — literally `calculateDynamicBalances_(null).earned`, the same value the Directory's Earned column shows, so the two screens cannot disagree. It is what the grid sorts by, so whoever has contributed least is offered first.
    - **Earned, not available.** Used hours, Paid Out and Carry Over must never be subtracted from it. It measures who has been covering, not who has hours banked — someone who covers constantly and spends it all must not read as the least busy person in the building. Don't "fix" it to match the Directory's Running Total.
    - Pending and denied rows are excluded. Counting a denied request would steer coverage away from the person whose request you turned down, and pending coverage is already surfaced in the same cell by its own hourglass.
    - It is **not** date-filtered, so it is the same number in every month a person appears in — the months organise availability, not hours. `finalizeSchoolYear` deletes approved rows from the sheet, and that is what resets it. An earlier version bucketed by month against the real clock; the building admins asked for a running total instead.
  - Each entry's `pendingRequests` carry `month` and `weekday` (`getPendingEarnedMap_`), and the hourglass only shows in the grid cell that matches both **and the period**, so a pending Tuesday Period 3 submission shows only on that cell. Periods are compared by name (`schedulePeriodKey()` in Index.html: `Period 3 - 9:52 - 10:39` and a legacy `3` both read `period 3`), so editing bell times doesn't strand older requests; a combined period (`Period 4/5`) only matches its own row. A `time_range` building compares no period, the same as the Assigned chips. A request whose date can't be read has a blank month and shows on its period in every column rather than disappearing.
  - Each entry also carries `assignments` — coverage already assigned to that person, from `assignedCoverageMap_(building)`. The client shows the dates on the person's card in the one grid cell a second booking would collide in (same month, same weekday, same period), with the detail on hover.
    - The source is the **TST Assignments row, not the calendar event**, which is the whole point: the row exists from the moment the admin assigns, so someone who was given coverage and never filed the form still reads as taken. The hourglass beside it only ever knew about a request that was filed, which is how the same person got booked twice.
    - For that reason it is **not gated on the building having a calendar** — a building with no calendar has the same double-booking problem. Where there *is* a calendar, each entry's `calendar` string says where the event stands (on the calendar / still queued / failed), and it is blank when the building has none.
    - Cancelled assignments are excluded (freeing the person up is the reason to cancel); recorded ones stay, flagged `recorded` — the coverage still happened.
    - The cell match compares the period only in a `periods` building. A `time_range` building (OIS, SE) stores the assignment's period as the **time span that was picked** while its grid row is a placeholder ("Time Range"), so there the month and weekday are the whole cell — `schedulePeriodsAreComparable()` in Index.html is that rule.
    - Assigning from the grid reloads it (`sendAssignment`), so the card the admin is looking at stops saying the person is free.
- **updateSchedulePeriod(month, period, dayUpdates, building)** - Admin only; `building` must pass `allowedBuildingFor_` (defaults to the caller's own building). Only deletes/rebuilds rows for that building's `scheduleMembers_` — buildings can share period names (OIS and SE both use "Time Range"), so other buildings' rows must survive. Rejects (before touching the sheet) any email in `dayUpdates` that isn't a member.
- Client: the eye button in the Directory's Actions column calls `startViewAs(email)`, which stashes the admin's `STATE.user`/`STATE.building` in `STATE.viewAs` and swaps in the teacher; `exitViewAs()` (banner or profile menu) restores them. Teacher actions already send `STATE.user.email`, so submissions made while viewing as someone are recorded under that teacher.

Authorization: add/edit/archive/restore require Admin or Super Admin; non-Super-Admins are scoped to their own building(s). Editing Carry Over / Paid Out and running finalize for a person additionally require being that person's **primary** building admin (or Super Admin). Permanent delete and archived-staff listing require Super Admin.

### Profile Menu (admins)

Switch School (multi-building) · Update Carry Over · Finalize School Year · View Archived Years · Settings (Settings also hosts the Super-Admin "Archived Staff" permanent-delete list).

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

### HTML Escaping (XSS)
Staff names/emails, transaction fields (subbed for, period, hours, notes, denial reasons), App Config values (building codes/names, periods, coverage labels), snapshot/archive names and server error messages are all user-controlled.
- **Index.html:** wrap every such value in `escapeHtml()` when building markup (text or quoted attribute). Never interpolate data into an inline handler's JS string (`onclick="fn('${x}')"`) — put it in a `data-*` attribute and read `this.dataset.x`. `showToast()` takes plain text. `openConfirmModal()`'s `message` is HTML, so callers escape what they put in it.
- **Code.js:** use `escapeHtml_()` for values placed in email or `HtmlService` HTML. `sendStyledEmail_()` escapes `subject`/`title`/`buttonText`/`buildingName` itself; `contentHtml` must be escaped by the caller. `handleCoverageAccept_/Reject_` render URL parameters, so always escape them.

### Name Display ("Last, First")
Names are stored "First Last" everywhere. Each person can choose to *see* them "Last, First" from the profile menu (**Show Names As**). It's a personal setting: `saveMyPreferences(prefs)` stores it in Script Properties under `PREFS_<email>`. It always writes the session user's own setting, and `getInitialData` returns it as `preferences`. UserProperties can't be used because the app runs as the deployer. During View As, the admin's own setting still applies.
- In Index.html, draw names through `displayName()` (a staff member) or `displayCoveredName()` (a free-text "covered for", which is only reordered when it matches someone in the directory, so "Activity Bus" is left alone).
- **Only change what is drawn.** `data-name`, `<option value>` and anything sent to the server keep the stored name, because requests and Form Responses are matched on it. Read a selected name from `data-name`, never from an option's visible text.
- Emails, the header's own name, and sentence-style prompts ("Send report to …?", "… is covering your Period 3") stay "First Last".

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
Styled HTML emails are sent via `sendStyledEmail_()` (private; queued through `addToEmailQueue_()`) with:
- Branded header (Orono Middle School)
- Color-coded content based on action type
- Direct link to web app
- Consistent footer with OPS branding

### Triggers
The `onFormSubmit(e)` function must be set up as an **installable trigger** in the Apps Script editor. This syncs Google Form submissions into the approval workflow.

**Each building admin** installs their own triggers from the spreadsheet menu (**TST Admin → Authorize Email Service**), which is what makes queued mail go out as them. `setupEmailService` installs three: `processEmailQueue` on change, `processEmailQueue` every minute, and `nudgeOutstandingAssignments` daily at 7am. Re-running it clears the old set first, so it is also the fix for a trigger Apps Script has disabled.

Because triggers are per-user, a building admin needs **standing access to the TST spreadsheet** — `processEmailQueue_` reads and writes the Email Queue sheet as the trigger owner. Without it, that building's mail queues and never sends.

A Super Admin's trigger **does not** process other buildings' queue rows (`processEmailQueue_`). It used to, which raced the building admin every minute and made the From name a coin flip. The trade-off is deliberate: a building with nobody authorized queues mail rather than sending it under the wrong name.

### Email service health and the second deployment

Queued mail is sent by each building admin's **own** trigger — that is what makes them the sender. Triggers are per-user, so:

- **Nobody can install one on anyone else's behalf.** The main web app runs as the deployer (`executeAs: USER_DEPLOYING`), so a button in it would only ever create the deployer's triggers again. That is why there is a **second deployment of the same project, configured "Execute as: user accessing the web app"**, reached at `?action=authorizeEmail` (`authorizeEmailServicePage_`). It is admin-gated and renders its own errors, because an Apps Script exception page tells a school secretary nothing.
  - Both deployments serve the same `Code.js`, so **redeploy them together**. The Super Admin stores the second URL via `setAuthorizeUrl`; blank is fine, and the banner then points at the spreadsheet menu instead.
  - It requires the admin to have access to the TST spreadsheet (the trigger binds to it, and the page reads the Staff Directory). That is the same access `processEmailQueue_` needs anyway.
- **The app cannot see whether a trigger exists.** `ScriptApp.getUserTriggers()` only ever returns the *effective* user's, so running as the deployer it cannot enumerate anyone else's. Health is therefore judged from the symptom: `oldestPendingMinutes_` finds mail sitting `Pending` beyond `QUEUE_STALL_MINUTES_` (15). That catches a trigger that was never installed **and** one Apps Script has since disabled — whose failure notice goes to its owner, not to whoever notices the silence. An authorization record (`EMAIL_AUTH_<building>`, written by `installEmailTriggers_`) separates "never set up" from "it broke", because those need different responses.
- `setupEmailService` (menu) and the authorization page share `installEmailTriggers_`, so both install the same three triggers and both are safe to re-run — re-running is the documented fix for a disabled trigger.

## Testing & Debugging

```bash
node test/run.js     # no dependencies, no network
```

[test/](test/) loads the real `Code.js` + `config.js` into a Node `vm` context with mocked Apps Script services (SpreadsheetApp, Session, LockService, PropertiesService, Utilities, MailApp, ScriptApp, HtmlService) and sheets as plain arrays. It covers the authorization rules above, the coverage-assignment flow, the calendar (via a `CalendarApp` stand-in, where an id the account cannot open simply returns null — exactly how a wrong id or a missing share behaves), the `google.script.run` calls [Index.html](Index.html) actually makes (recorded from a real browser into `test/ui_calls.json` by `test/record_ui_calls.js`, replayed without one), and the deployed file set. Run it before pushing; CI runs it on PRs and again before every deploy. See [test/README.md](test/README.md).

`ui_calls.json` is **stale for the assignment screens** — re-recording needs Playwright (`npm i playwright`), which is not installed here. The new endpoints are covered directly by `test/assignments.test.js` with the arguments the client sends, and `ui_flows.test.js` additionally checks statically that Index.html never calls a private (`_`) server function and that the endpoints it dispatches by computed name still exist.

When adding an endpoint, add a case for what a Teacher, an admin from another building, and a Super Admin each get from it.

For debugging in the live project:

1. Use `Logger.log()` in [Code.js](Code.js) - view logs in Apps Script editor (Ctrl+Enter to run functions)
2. Use `console.log()` in [Index.html](Index.html) - view in browser console
3. Test with different user roles by modifying entries in "Staff Directory"
4. Use "Executions" tab in Apps Script editor to view runtime logs and errors

## OAuth Scopes

Defined in [appsscript.json](appsscript.json):
- `https://www.googleapis.com/auth/spreadsheets` - Read/write spreadsheet data
- `https://www.googleapis.com/auth/script.send_mail` - Send emails via MailApp
- `https://www.googleapis.com/auth/userinfo.email` - Get active user email
- `https://www.googleapis.com/auth/script.scriptapp` - Install the email/nudge triggers
- `https://www.googleapis.com/auth/calendar` - Create and remove TST coverage events

**Adding the calendar scope means every admin re-authorizes once** — the deployment prompts on next use, and each building admin must re-run "Authorize Email Service" so their triggers carry the new scope.

## Deployment Notes

- Web app execution: "User Deploying" (runs as the deployer, accesses deployer's sheets)
- Access level: "DOMAIN" (restricted to domain users only)
- Timezone: America/Chicago
- After `clasp push`, you may need to redeploy via `clasp deploy` or the Apps Script editor for changes to take effect in the web app
