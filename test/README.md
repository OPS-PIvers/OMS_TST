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

## Adding a server file

`.claspignore` is an allow-list. A new `.js` file is **not** deployed until it is
added there — and Node test files must stay out of it, or `clasp push` would upload
them into the Apps Script project.
