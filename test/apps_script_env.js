/**
 * Minimal Google Apps Script environment for running Code.js under Node.
 *
 * Apps Script has no local runtime, so the server code is loaded into a `vm`
 * context alongside hand-written stand-ins for the services it uses
 * (SpreadsheetApp, Session, LockService, PropertiesService, Utilities, MailApp,
 * ScriptApp, HtmlService). The sheets are plain 2-D arrays, so a test can assert
 * on what a function wrote as easily as on what it returned.
 *
 * Usage:
 *   const env = createEnv({ activeUser: 'teacher@x', sheets: { 'Staff Directory': rows } });
 *   env.run('getStaffDirectoryData', 'OMS');
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

// OMS_SRC_DIR lets a test load a different checkout of the sources — used to
// confirm these tests actually fail against the code before the fix.
const ROOT = process.env.OMS_SRC_DIR || path.join(__dirname, '..');

/** One sheet: a 2-D array of cell values plus the Sheet methods Code.js calls. */
class FakeSheet {
  constructor(name, values) {
    this.name = name;
    this.values = (values || []).map(row => row.slice());
    this.metadata = [];
    this.hidden = false;
    this.frozenRows = 0;
  }

  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  hideSheet() { this.hidden = true; return this; }
  setFrozenRows(n) { this.frozenRows = n; return this; }
  autoResizeColumn() { return this; }

  getLastRow() { return this.values.length; }
  getLastColumn() {
    return this.values.reduce((max, row) => Math.max(max, row.length), 0);
  }

  /** Pads the grid so (row, col) — both 1-based — exists. */
  _ensure(row, col) {
    while (this.values.length < row) this.values.push([]);
    for (let r = 0; r < this.values.length; r++) {
      while (this.values[r].length < col) this.values[r].push('');
    }
  }

  getRange(row, col, numRows, numCols) {
    const rows = numRows === undefined ? 1 : numRows;
    const cols = numCols === undefined ? 1 : numCols;
    const sheet = this;
    return {
      getValues() {
        sheet._ensure(row + rows - 1, col + cols - 1);
        const out = [];
        for (let r = 0; r < rows; r++) {
          out.push(sheet.values[row - 1 + r].slice(col - 1, col - 1 + cols));
        }
        return out;
      },
      getValue() { return this.getValues()[0][0]; },
      setValue(v) {
        sheet._ensure(row + rows - 1, col + cols - 1);
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) sheet.values[row - 1 + r][col - 1 + c] = v;
        }
        return this;
      },
      setValues(vals) {
        sheet._ensure(row + vals.length - 1, col + (vals[0] ? vals[0].length : 1) - 1);
        vals.forEach((rowVals, r) => {
          rowVals.forEach((v, c) => { sheet.values[row - 1 + r][col - 1 + c] = v; });
        });
        return this;
      },
      setNumberFormat() { return this; },
      setFontWeight() { return this; },
      setBackground() { return this; },
      sort() { return this; }
    };
  }

  getDataRange() {
    const sheet = this;
    const cols = Math.max(1, this.getLastColumn());
    return {
      getValues() {
        return sheet.values.map(row => {
          const copy = row.slice();
          while (copy.length < cols) copy.push('');
          return copy;
        });
      }
    };
  }

  appendRow(row) { this.values.push(row.slice()); return this; }
  deleteRow(rowNumber) { this.values.splice(rowNumber - 1, 1); return this; }
  sort() { return this; }

  addDeveloperMetadata(key, value) {
    this.metadata.push({ getKey: () => key, getValue: () => value });
    return this;
  }
  getDeveloperMetadata() { return this.metadata.slice(); }
}

class FakeSpreadsheet {
  constructor(sheets) {
    this.sheets = Object.keys(sheets || {}).map(name => new FakeSheet(name, sheets[name]));
  }
  getSheetByName(name) { return this.sheets.find(s => s.name === name) || null; }
  getSheets() { return this.sheets.slice(); }
  insertSheet(name) {
    const sheet = new FakeSheet(name, []);
    this.sheets.push(sheet);
    return sheet;
  }
  deleteSheet(sheet) { this.sheets = this.sheets.filter(s => s !== sheet); }
  getId() { return 'fake-spreadsheet-id'; }
  getUrl() { return 'https://example.invalid/fake'; }
}

/**
 * Builds a sandbox with Code.js + config.js loaded.
 *
 * options.activeUser  — what Session.getActiveUser().getEmail() returns
 * options.sheets      — { sheetName: rows }
 * options.properties  — seed Script Properties
 * options.files       — source files to load, in project order
 *                       (default: what .claspignore deploys)
 */
function createEnv(options) {
  const opts = options || {};
  const ss = new FakeSpreadsheet(opts.sheets || {});
  const sentEmails = [];
  const properties = Object.assign({}, opts.properties);
  const installedTriggers = [];
  // { id, calendarId, title, start, end, options } for every event still on a
  // calendar, so a test can assert on what was created as well as on what was said.
  const calendarEvents = [];
  const calendars = Object.assign({}, opts.calendars); // id -> { name } (absent = no access)
  const logs = [];
  const menus = [];
  const alerts = [];

  // ScriptApp.AuthMode values are enum objects. A trigger event carries the real
  // object; a google.script.run payload can only carry JSON, which is exactly the
  // distinction isTriggerEvent_ relies on — so model them as objects here too.
  const AuthMode = {
    NONE: { toString: () => 'NONE' },
    CUSTOM_FUNCTION: { toString: () => 'CUSTOM_FUNCTION' },
    LIMITED: { toString: () => 'LIMITED' },
    FULL: { toString: () => 'FULL' }
  };

  const sandbox = {
    console: {
      log: (...a) => logs.push(['log', ...a]),
      warn: (...a) => logs.push(['warn', ...a]),
      error: (...a) => logs.push(['error', ...a])
    },
    Logger: { log: (...a) => logs.push(['Logger', ...a]) },
    JSON, Math, Date, Number, String, Boolean, Array, Object, Set, Map, RegExp, Error,
    isFinite, isNaN, parseInt, parseFloat, encodeURIComponent, decodeURIComponent,

    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      flush: () => {},
      getUi: () => ({
        ButtonSet: { OK: 'OK' },
        alert: (title, message) => { alerts.push({ title, message }); },
        createMenu: name => {
          const menu = { name, items: [] };
          const builder = {
            addItem: (caption, fn) => { menu.items.push({ caption, fn }); return builder; },
            addSeparator: () => builder,
            addToUi: () => { menus.push(menu); }
          };
          return builder;
        }
      })
    },

    Session: {
      getActiveUser: () => ({ getEmail: () => opts.activeUser || '' }),
      getEffectiveUser: () => ({ getEmail: () => opts.activeUser || '' }),
      getScriptTimeZone: () => 'America/Chicago'
    },

    LockService: {
      getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {}, waitLock: () => {} })
    },

    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: k => (k in properties ? properties[k] : null),
        setProperty: (k, v) => { properties[k] = v; },
        deleteProperty: k => { delete properties[k]; }
      })
    },

    Utilities: {
      computeHmacSha256Signature: (value, key) => {
        // Deterministic stand-in — not cryptographic, but every byte of the key AND
        // the value has to reach every byte of the digest. An earlier version read
        // only the first 32 characters of key + '|' + value; the secret alone is
        // longer than that, so the signed payload never changed the signature and a
        // tampered link verified happily.
        const s = String(key) + '|' + String(value);
        const bytes = [];
        for (let i = 0; i < 32; i++) {
          let h = (0x811c9dc5 ^ Math.imul(i + 1, 0x9e3779b1)) >>> 0; // FNV-1a, salted per byte
          for (let j = 0; j < s.length; j++) {
            h = Math.imul((h ^ s.charCodeAt(j)) >>> 0, 16777619) >>> 0;
          }
          bytes.push(h & 0xff);
        }
        return bytes;
      },
      base64EncodeWebSafe: bytes => Buffer.from(bytes).toString('base64url'),
      base64Encode: bytes => Buffer.from(bytes).toString('base64'),
      formatDate: (date, tz, fmt) => {
        const d = new Date(date);
        const pad = n => String(n).padStart(2, '0');
        if (fmt === 'yyyy-MM-dd') return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
        return d.toISOString();
      },
      getUuid: () => 'uuid-' + Math.random().toString(36).slice(2)
    },

    MailApp: { sendEmail: msg => sentEmails.push(msg) },

    // CalendarApp stand-in. getCalendarById returns null for an id this account
    // cannot open, which is exactly how a wrong id or a missing share behaves.
    CalendarApp: {
      getCalendarById: id => {
        const meta = calendars[id];
        if (!meta) return null;
        const calendar = {
          getName: () => meta.name || id,
          getId: () => id,
          createEvent: (title, start, end, options) => {
            if (meta.readOnly) throw new Error('You do not have permission to add events to this calendar.');
            const event = {
              id: 'event-' + (calendarEvents.length + 1),
              calendarId: id,
              title: title,
              start: start,
              end: end,
              options: options || {}
            };
            calendarEvents.push(event);
            return {
              getId: () => event.id,
              getTitle: () => event.title,
              deleteEvent: () => {
                const i = calendarEvents.indexOf(event);
                if (i > -1) calendarEvents.splice(i, 1);
              }
            };
          },
          getEventById: eventId => {
            const event = calendarEvents.find(ev => ev.id === eventId && ev.calendarId === id);
            if (!event) return null;
            return {
              getId: () => event.id,
              getTitle: () => event.title,
              deleteEvent: () => {
                const i = calendarEvents.indexOf(event);
                if (i > -1) calendarEvents.splice(i, 1);
              }
            };
          }
        };
        return calendar;
      }
    },

    ScriptApp: {
      AuthMode: AuthMode,
      getUserTriggers: () => installedTriggers.slice(),
      deleteTrigger: t => {
        const i = installedTriggers.indexOf(t);
        if (i > -1) installedTriggers.splice(i, 1);
      },
      newTrigger: handler => {
        const builder = {
          forSpreadsheet: () => builder,
          timeBased: () => builder,
          onChange: () => builder,
          everyMinutes: () => builder,
          everyDays: () => builder,
          atHour: () => builder,
          create: () => {
            const t = { getHandlerFunction: () => handler };
            installedTriggers.push(t);
            return t;
          }
        };
        return builder;
      },
      getService: () => ({ getUrl: () => 'https://example.invalid/exec' })
    },

    HtmlService: {
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
      createHtmlOutputFromFile: name => {
        const out = {
          file: name,
          getContent: () => '',
          setTitle: () => out,
          setXFrameOptionsMode: () => out
        };
        return out;
      },
      createHtmlOutput: html => {
        const out = {
          html: html,
          getContent: () => html,
          setTitle: () => out,
          setXFrameOptionsMode: () => out
        };
        return out;
      }
    }
  };
  sandbox.globalThis = sandbox;

  const context = vm.createContext(sandbox);
  // Apps Script shares one global scope across files and evaluates them in project
  // order, so the order here matters: a later file's function declaration shadows
  // an earlier one's. Default to the set .claspignore deploys.
  // (clasp pushes in alphabetical order, so Code.js really does load before config.js.)
  const files = opts.files || ['Code.js', 'config.js'];
  files.forEach(file => {
    vm.runInContext(fs.readFileSync(path.join(ROOT, file), 'utf8'), context, { filename: file });
  });

  return {
    context,
    spreadsheet: ss,
    sentEmails,
    properties,
    installedTriggers,
    calendarEvents,
    calendars,
    logs,
    menus,
    alerts,
    AuthMode,
    sheet: name => ss.getSheetByName(name),
    /**
     * Calls a server function the way google.script.run would — including its one
     * hard rule: Apps Script never exposes a function whose name ends in "_".
     */
    run(fnName, ...args) {
      if (/_$/.test(fnName)) {
        throw new Error(fnName + ' is private and is not exposed to google.script.run.');
      }
      return this.callInternal(fnName, ...args);
    },
    /** Calls any function, private ones included (what server-side code can do). */
    callInternal(fnName, ...args) {
      const fn = context[fnName];
      if (typeof fn !== 'function') throw new Error('No such server function: ' + fnName);
      return fn(...args);
    },
    /** Runs fn and returns { ok, value } or { ok: false, error }. */
    attempt(fnName, ...args) {
      try {
        return { ok: true, value: this.run(fnName, ...args) };
      } catch (err) {
        return { ok: false, error: err };
      }
    }
  };
}

module.exports = { createEnv, FakeSheet, FakeSpreadsheet };
