# Tests

Google Apps Script has no local runtime, so the suite loads the real `Code.js`
(plus `config.js`) into a Node `vm` context alongside stand-ins for the services it
uses — `SpreadsheetApp`, `Session`, `LockService`, `PropertiesService`, `Utilities`,
`MailApp`, `ScriptApp`, `HtmlService`. Sheets are plain 2-D arrays, so a test can
assert on what a function wrote as easily as on what it returned.

```bash
node test/run.js       # no dependencies, no network
```

| File | What it is |
| --- | --- |
| `apps_script_env.js` | The sandbox. `createEnv({ activeUser, sheets, files })` returns `run` / `attempt` helpers; `run` refuses names ending in `_`, exactly as `google.script.run` does. |
| `fixtures.js` | A two-building district (OMS/OHS) with admins, teachers, a multi-building teacher and an archived one. |
| `authorization.test.js` | Who may read what: teachers, admins from the wrong building, and people missing from the directory. |
| `assignments.test.js` | Coverage assignments end to end: who may assign, the duplicate guard, the three queued emails, recording (including that it cannot happen twice), cancellation against approved hours, the one nudge, and the signed Record link's expiry. |
| `ui_flows.test.js` | Replays the `google.script.run` calls `Index.html` actually makes (recorded in `ui_calls.json`) and checks each flow still gets what the page renders. |
| `legacy_globals.test.js` | `Code_legacy.js` stays out of the deployed file set, and what its duplicate globals used to shadow. |
| `record_ui_calls.js` | Regenerates `ui_calls.json`. Needs Playwright; everything else does not. |

## Re-recording the UI calls

`ui_calls.json` comes from driving `Index.html` in Chromium with `google.script.run`
replaced by a stub that forwards every call to this same server code, so the page
gets real responses (refusals included) and the recording is what the UI actually
asks for:

```bash
node test/record_ui_calls.js
```

Re-run it when a change alters which server calls the UI makes, and commit the
result. `node test/run.js` then replays those calls without a browser.

**The recording does not yet cover the assignment screens** — it needs Playwright,
which is not installed in every checkout. Until it is re-recorded, those endpoints
are covered directly in `assignments.test.js`, and `ui_flows.test.js` adds two
static checks over Index.html that need no browser: that it never calls a private
(`_`) server function, and that the endpoints it dispatches by computed name
(`batch${Action}${Type}`, the revert pair) still exist. A full chain-by-chain scan
was tried and removed: Index.html nests template literals and contains regex
literals with quotes in them, so resolving each endpoint by name needs a real
JavaScript parser, and anything less reports phantom failures.

## Adding a server file

`.claspignore` is an allow-list. A new `.js` file is **not** deployed until it is
added there — and Node test files must stay out of it, or `clasp push` would upload
them into the Apps Script project.
