# Google Sheet Backend Schema

This document outlines the required structure for the Google Sheet used by the OMS TST Manager application.

**Important:** Some features in the application code rely on specific column indices (especially for updating balances), so maintaining the column order defined below is critical.

---

## 1. Staff Directory
**Purpose:** Stores user profiles, roles, balances, and building association.
**Code Dependency:** heavily relies on column order for balance updates (Earned/Used).

| Col | Header Name (Recommended) | Data Type | Notes |
| :-- | :--- | :--- | :--- |
| **A** | Name | String | Teacher/Staff Name |
| **B** | Email Address | String | **Unique ID**. Must match Google account email. |
| **C** | Role | String | Values: `Admin`, `Super Admin`, `Teacher` (or empty) |
| **D** | Earned | Number | **Critical**. Application writes to this column (Hardcoded Index 4). |
| **E** | Used | Number | **Critical**. Application writes to this column (Hardcoded Index 5). |
| **F** | Carry Over | Number | Legacy starting balance (Optional). |
| **G** | Paid Out | Number | Hours cashed in. Subtracts from total. |
| **H** | Running Total | Number | =ARRAYFORMULA(IF(B2:B="", "", (N(D2:D) + N(F2:F)) - (N(E2:E) + N(G2:G)))) |
| **I** | Building | String | **Required for Multi-Building**. Codes: `OMS`, `OHS`, etc. (Matches config.js). Supports comma-separated multi-building assignment (e.g. `OMS, OHS`). |
| **J** | Archived | String | **Per-building** soft-delete: comma-separated list of building codes the staff member is archived FROM (e.g. `OMS`). Empty = active everywhere. A person is "fully archived" only when this list covers every building in column I. (Legacy `TRUE` is treated as archived from all buildings.) Auto-created if missing. |
| **K** | Last Finalized | String | School-year name of the most recent year-end finalize for this person (e.g. `2025-2026`). Prevents a balance from being rolled twice when staff span buildings. Auto-created if missing. |

> **Note:** When adding a new staff row programmatically, write the individual cells (A–C, F, G, I, J) rather than `appendRow`, leaving column **H** blank so the Running Total ARRAYFORMULA fills it (writing into H collides with the spilling formula).

### Year-End Archive Sheets (created by `finalizeSchoolYear`)

- **`<year> <building> TST Totals`** (e.g. `2025-2026 OMS TST Totals`) — one per building per finalized year. Columns: Name, Email, Building(s), Carry Over (start), Earned, Used, Paid Out, Balance. Tagged with developer metadata (`tstArchiveBuilding`, `tstArchiveYear`) so the app can find it regardless of the chosen name.
- **`TST Approvals Archive` / `TST Usage Archive`** — permanent, accumulating backups. On finalize, the building's approved/processed transaction rows are moved here (original columns + a trailing `School Year`), which is what makes live Earned/Used recompute to 0 for the new year.

---

## 2. TST Approvals (New)
**Purpose:** Stores all "Earned" time requests (Subbing for others).
**Code Dependency:** `appendRow` assumes this exact order.

| Col | Header Name | Data Type | Notes |
| :-- | :--- | :--- | :--- |
| **A** | Email | String | Requester Email |
| **B** | Name | String | Requester Name (Snapshot) |
| **C** | Subbed For | String | Name of person covered |
| **D** | Email | String | Sheets Formula to auto-populate the email address of whomever submitted the request | 
=MAP(C2:C, LAMBDA(teacher_ref, 
  IF(teacher_ref="",, 
    IFERROR(
      INDEX(
        FILTER('Staff Directory (OLD)'!B:B, 
          IF(ISNUMBER(SEARCH(".", teacher_ref)), 
            (LEFT('Staff Directory (OLD)'!A:A, 1) = LEFT(teacher_ref, 1)) * REGEXMATCH('Staff Directory (OLD)'!A:A, "(?i)\s" & TRIM(MID(teacher_ref, SEARCH(".", teacher_ref)+1, 100)) & "$"), 
            IF(ISNUMBER(SEARCH(" ", teacher_ref)),
               'Staff Directory (OLD)'!A:A = teacher_ref,
               REGEXMATCH('Staff Directory (OLD)'!A:A, "(?i)\s" & teacher_ref & "$")
            )
          )
        ), 
      1), 
      ""
    )
  )
))
| **E** | Date | Date | Date of coverage |
| **F** | Period | String | e.g., "Period 1" |
| **G** | Time Type | String | e.g., "Full Period", "Half Period" |
| **H** | Hours | Number | Calculated value (e.g., 1.0, 0.5) |
| **I** | Approved | Boolean | `TRUE` if approved |
| **J** | Approved TS | Date/Time | Timestamp of approval |
| **K** | Denied | Boolean | `TRUE` if denied |
| **L** | Denied TS | Date/Time | Timestamp of denial |
| **M** | Denial Reason | String | Reason provided by Admin |
| **N** | Building | String | **New**. Building Code (e.g., `OMS`). |

---

## 3. TST Usage (New)
**Purpose:** Stores all "Used" time requests (Redeeming hours).
**Code Dependency:** `appendRow` assumes this exact order.

| Col | Header Name | Data Type | Notes |
| :-- | :--- | :--- | :--- |
| **A** | Email | String | Requester Email |
| **B** | Name | String | Requester Name |
| **C** | Date | Date | Date usage requested |
| **D** | Amount | Number | Hours used |
| **E** | Status | Boolean | `TRUE` if processed/approved |
| **F** | Timestamp | Date/Time | Time of approval |
| **G** | Notes | String | Optional user notes |
| **H** | Building | String | **New**. Building Code (e.g., `OMS`). |

---

## 4. Form Responses 1
**Purpose:** Legacy/Backup archive. Receives raw form submissions or app archives.
**Code Dependency:** Used as a secondary record for "Earned" requests.

| Col | Header Name | Data Type | Notes |
| :-- | :--- | :--- | :--- |
| **A** | Timestamp | Date/Time | Submission time |
| **B** | Email Address | String | User Email |
| **C** | I subbed For | String | Name |
| **D** | Coverage for someone other than listed above: | String | 'Other' flag (or empty) |
| **E** | Date subbed: | Date | Date of coverage |
| **F** | Time Subbed: | String | Period covered |
| **G** | Amount Type | String | "Full Period", "Half Period" |
| **H** | Amount | Number | Decimal hours |

---

## 5. TST Availability
**Purpose:** Stores teacher availability schedules for the grid view.
**Code Dependency:** Created automatically if missing, but requires specific columns.

| Col | Header Name | Data Type | Notes |
| :-- | :--- | :--- | :--- |
| **A** | Month | String | e.g., "September" |
| **B** | Day(s) Available | String | e.g., "Mon,Tue" |
| **C** | Period | String | e.g., "Period 1" |
| **D** | Name | String | Teacher Name |
| **E** | Email | String | Teacher Email |
| **F** | Hours Earned This Month | Formula/Num | (Optional) Used for display in some views |

---

## 6. App Config
**Purpose:** Stores building-specific configuration as JSON.
**Code Dependency:** `getConfig` and `saveBuildingConfig` rely on this structure.

| Col | Header Name | Data Type | Notes |
| :-- | :--- | :--- | :--- |
| **A** | Building | String | **Unique ID**. Building code (e.g., `OMS`, `OHS`, `OIS`, `SE`). |
| **B** | Config_JSON | String | **JSON String**. Contains periods, schedule types, and coverage rules. |
