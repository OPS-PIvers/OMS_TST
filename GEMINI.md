# Project Overview

**Project Name:** Orono TST Manager (formerly OMS TST Manager)

This project is a Google Apps Script web application designed to manage "TST" (Time Saving Team/Teacher) hours for Orono Public Schools. It allows staff to earn time (by providing sub coverage) and use time (redeeming credits). 

**Key Features:**
*   **Multi-Building Support:** Configurable for OMS, OHS, OIS, and SE with building-specific schedules and rules.
*   **Role-Based UI:** Distinct views for Teachers, Admins, and Super Admins.
*   **Legacy Compatibility:** Supports legacy Google Form submissions with automated calculation logic.
*   **Availability Grid:** Visual schedule management for TST coverage.

## Technical Architecture

*   **Platform:** Google Apps Script (Serverless).
*   **Backend:** `Code.js` (Google Apps Script / V8 Runtime). Handles business logic, email notifications, and Google Sheets interactions.
*   **Configuration:** `config.js`. Centralized configuration for building schedules, names, and coverage rules.
*   **Frontend:** `Index.html`. A Single Page Application (SPA) served via `HtmlService`.
*   **Styling:** Tailwind CSS (v3.x via CDN).
*   **Icons:** FontAwesome (v6.x via CDN).
*   **Database:** Google Spreadsheet (accessed via `SpreadsheetApp`).

## Key Files

*   `Code.js`: Main server-side logic. Contains `doGet()`, API functions, and specific business rules (`calculatePeriods`).
*   `config.js`: Defines `BUILDING_CONFIG` and `DEFAULT_BUILDING`.
*   `Index.html`: Client-side code (HTML/JS/CSS). Handles UI state, building switching, and form rendering.
*   `SPREADSHEET_SCHEMA.md`: **Crucial.** Defines the required Google Sheet structure and column order.
*   `Code_legacy.js`: Archive of previous logic.
*   `appsscript.json`: Manifest file.
*   `.clasp.json`: Clasp configuration.

## Data Structure (Google Sheets)

The application relies on specific sheets within a bound Google Sheet. **Column order is critical for certain operations.** Refer to `SPREADSHEET_SCHEMA.md` for the exact definition.

*   `Staff Directory`: User profiles, roles, balances, and **Building** codes.
*   `TST Approvals (New)`: "Earned" requests (Subbing).
*   `TST Usage (New)`: "Used" requests (Redeeming).
*   `TST Availability`: Availability grid data.
*   `Form Responses 1`: Archive of raw form submissions.

## Development & Deployment

This project uses `clasp` for local development and deployment.

### Prerequisites
*   Node.js & npm
*   `@google/clasp` installed globally
*   Logged in via `clasp login`

### Common Commands

*   **Push changes:** `clasp push` (or `clasp push --force`)
*   **Deploy:** `clasp deploy -i <DEPLOYMENT_ID>`
*   **Open:** `clasp open`

## Conventions

*   **Communication:** Client-server communication via `google.script.run`.
*   **State Management:** Frontend `STATE` object manages user data, current building context, and view.
*   **Multi-Building:**
    *   `DEFAULT_BUILDING` in `config.js` sets the fallback (currently 'OMS').
    *   Super Admins can switch views between buildings.
    *   Regular Admins/Teachers are locked to their assigned building in `Staff Directory`.
*   **Safety:** Critical actions (Delete/Deny) require confirmation.
*   **Legacy Support:** `onFormSubmit` in `Code.js` contains specific logic to handle legacy form inputs, calculating decimal hours based on period strings and building rules.