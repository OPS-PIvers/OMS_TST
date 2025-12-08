# Project Overview

**Project Name:** OMS TST Manager (Orono Middle School TST)

This project is a Google Apps Script web application designed to manage "TST" (Time Saving Team/Teacher) hours. It allows staff to earn time (by subbing) and use time (redeeming credits). The application features a role-based UI for Admins and Teachers.

## Technical Architecture

*   **Platform:** Google Apps Script (Serverless).
*   **Backend:** `Code.js` (Google Apps Script / V8 Runtime). Handles business logic, email notifications, and Google Sheets interactions.
*   **Frontend:** `Index.html`. A Single Page Application (SPA) served via `HtmlService`.
*   **Styling:** Tailwind CSS (v3.x via CDN).
*   **Icons:** FontAwesome (v6.x via CDN).
*   **Database:** Google Spreadsheet (accessed via `SpreadsheetApp`).

## Key Files

*   `Code.js`: The main server-side script. Contains `doGet()` (entry point), API functions (e.g., `submitEarned`, `approveEarnedRow`), and data access logic.
*   `Index.html`: The client-side code. Contains HTML structure, CSS (Tailwind config), and client-side JavaScript (`google.script.run` calls).
*   `Code_legacy.js`: Contains legacy logic, likely for reference or migration.
*   `appsscript.json`: The Google Apps Script manifest file defining scopes, timezone, and web app execution settings.
*   `.clasp.json`: Configuration for `clasp` (Command Line Apps Script Projects).

## Data Structure (Google Sheets)

The application relies on specific sheets within a bound Google Sheet:
*   `Staff Directory`: User roles and balances.
*   `TST Approvals (New)`: Pending and processed "Earned" requests.
*   `TST Usage (New)`: Pending and processed "Used" requests.
*   `Form Responses 1`: Likely the raw input for some earned requests or legacy form data.

## Development & Deployment

This project uses `clasp` for local development and deployment.

### Prerequisites
*   Node.js & npm
*   `@google/clasp` installed globally (`npm install -g @google/clasp`)
*   Logged in via `clasp login`

### Common Commands

*   **Push changes to Google Drive:**
    ```bash
    clasp push
    ```
    *Note: This overwrites the files on the remote script project.*

*   **Pull changes from Google Drive:**
    ```bash
    clasp pull
    ```

*   **Deploy a new version:**
    ```bash
    clasp deploy --description "Version description"
    ```

*   **Open the project in browser:**
    ```bash
    clasp open
    ```

## Conventions

*   **Communication:** Client-server communication is handled via `google.script.run`.
*   **State Management:** The frontend uses a global `STATE` object to manage user data and view contexts.
*   **Styling:** Utility-first CSS using Tailwind. Custom colors (e.g., `ops-blue`, `ops-red`) are defined in the Tailwind config script within `Index.html`.
*   **Safety:** Critical actions (Delete/Deny) require user confirmation. Database writes are performed directly on specific sheet indices (be careful when changing column order).
