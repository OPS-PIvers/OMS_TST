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
| **F** | Running Total | Number | =ARRAYFORMULA(IF(B2:B="", "", N(D2:D) + N(G2:G) - N(E2:E))) |
| **G** | Carry Over | Number | Legacy starting balance (Optional). |
| **H** | Building | String | **Required for Multi-Building**. Codes: `OMS`, `OHS`, etc. (Matches config.js) |

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
