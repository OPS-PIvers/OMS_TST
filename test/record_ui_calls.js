/**
 * Drives Index.html in a real browser against the real Code.js and records every
 * google.script.run call the UI makes.
 *
 * `google.script.run` is replaced with a stub that forwards each call to Code.js
 * running in the Node vm harness (as a chosen signed-in user), so the page gets
 * genuine server responses — including the new authorization refusals — and the
 * recording is what Index.html actually asks for, not what we think it asks for.
 *
 *   node test/record_ui_calls.js          # writes test/ui_calls.json
 *
 * The replay assertions live in ui_flows.test.js and run without a browser, so
 * `node test/run.js` stays dependency-free; re-run this recorder after changing
 * which calls the UI makes.
 */

const fs = require('fs');
const path = require('path');
const Module = require('module');

const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

// Playwright is installed globally in this environment, not in the repo.
function loadPlaywright() {
  const extra = ['/opt/node22/lib/node_modules', '/usr/lib/node_modules', '/usr/local/lib/node_modules'];
  const saved = Module.globalPaths.slice();
  extra.forEach(p => { if (!Module.globalPaths.includes(p)) Module.globalPaths.push(p); });
  try {
    return require('playwright');
  } catch (err) {
    for (const base of extra) {
      try { return require(path.join(base, 'playwright')); } catch (e) { /* keep looking */ }
    }
    throw new Error('Playwright is not installed; skip the recorder or install it. ' + err.message);
  } finally {
    Module.globalPaths.length = 0;
    saved.forEach(p => Module.globalPaths.push(p));
  }
}

const INDEX = path.join(__dirname, '..', 'Index.html');
const OUT = path.join(__dirname, 'ui_calls.json');

// google.script.run stub. Installed before the page's own scripts so window.onload
// finds it, and Tailwind's CDN script is stubbed so the recorder works offline.
const STUB = `
  window.tailwind = window.tailwind || {};
  window.__omsCalls = [];
  window.__omsErrors = [];
  (function () {
    function runner(handlers) {
      return new Proxy({}, {
        get(target, prop) {
          if (prop === 'withSuccessHandler') return fn => runner(Object.assign({}, handlers, { success: fn }));
          if (prop === 'withFailureHandler') return fn => runner(Object.assign({}, handlers, { failure: fn }));
          if (prop === 'withUserObject') return obj => runner(Object.assign({}, handlers, { userObject: obj }));
          if (typeof prop !== 'string') return undefined;
          return function (...args) {
            window.__omsCalls.push({ fn: prop, args: JSON.parse(JSON.stringify(args)) });
            window.__omsServerCall(prop, args).then(res => {
              if (res.ok) {
                if (handlers.success) handlers.success(res.value, handlers.userObject);
              } else {
                window.__omsErrors.push({ fn: prop, message: res.error });
                if (handlers.failure) handlers.failure(new Error(res.error), handlers.userObject);
              }
            });
          };
        }
      });
    }
    window.google = { script: { run: runner({}), host: { close() {} } } };
  })();
`;

async function main() {
  const { chromium } = loadPlaywright();
  const browser = await chromium.launch();
  const recording = { recordedAt: new Date().toISOString(), flows: [] };

  try {
    for (const flow of FLOWS) {
      const env = createEnv({ activeUser: flow.signedInAs, sheets: sheets() });
      const page = await browser.newPage();
      const calls = [];

      await page.exposeFunction('__omsServerCall', (fn, args) => {
        const result = env.attempt(fn, ...args);
        calls.push({ fn, args, ok: result.ok });
        return result.ok
          ? { ok: true, value: JSON.parse(JSON.stringify(result.value === undefined ? null : result.value)) }
          : { ok: false, error: result.error.message };
      });
      await page.addInitScript(STUB);
      // Keep it hermetic: no CDN fetches.
      await page.route('**://**', route => {
        const url = route.request().url();
        return url.startsWith('file:') ? route.continue() : route.fulfill({ status: 200, body: '' });
      });

      const pageErrors = [];
      page.on('pageerror', err => pageErrors.push(err.message));

      await page.goto('file://' + INDEX);
      await page.waitForSelector('#app-container:not(.hidden)', { timeout: 15000 });

      const observations = await flow.drive(page);

      recording.flows.push({
        name: flow.name,
        signedInAs: flow.signedInAs,
        calls,
        observations,
        pageErrors: pageErrors.filter(m => !/tailwind/i.test(m)),
        handledErrors: await page.evaluate(() => window.__omsErrors)
      });

      console.log(`${flow.name}: ${calls.length} server calls, ${pageErrors.length} page errors`);
      await page.close();
    }
  } finally {
    await browser.close();
  }

  fs.writeFileSync(OUT, JSON.stringify(recording, null, 2) + '\n');
  console.log('Wrote ' + path.relative(process.cwd(), OUT));
}

const settle = page => page.waitForTimeout(250);

async function openTab(page, id) {
  await page.click(`#btn-${id}`);
  await settle(page);
}

const FLOWS = [
  {
    name: 'teacher',
    signedInAs: USERS.omsTeacher,
    async drive(page) {
      // Lands on My Report (teacher-totals).
      await settle(page);
      await openTab(page, 'teacher-submit');
      const subbedForOptions = await page.$$eval(
        '#div-sub-dropdown select[name="subbedNameDropdown"] option', els => els.map(e => e.textContent.trim()));

      await openTab(page, 'teacher-schedule');
      const checkedBoxes = await page.$$eval(
        '#view-container input[type="checkbox"]:checked', els => els.map(e => e.name));

      await openTab(page, 'teacher-totals');
      const kpis = await page.$$eval('#view-container .text-2xl', els => els.map(e => e.textContent.trim()));

      return { subbedForOptions, checkedBoxes, kpis };
    }
  },
  {
    name: 'admin',
    signedInAs: USERS.omsAdmin,
    async drive(page) {
      await settle(page);
      const earnedRows = await page.$$eval('#view-container tbody tr', els => els.length);

      await openTab(page, 'admin-used');
      await openTab(page, 'admin-schedule');
      await openTab(page, 'admin-totals');
      const directoryText = await page.textContent('#view-container');

      const badges = await page.evaluate(() => ({
        earned: document.getElementById('badge-earned') ? document.getElementById('badge-earned').textContent : '',
        used: document.getElementById('badge-used') ? document.getElementById('badge-used').textContent : ''
      }));

      return { earnedRows, badges, directoryHasTeacher: /Tina Teacher/.test(directoryText) };
    }
  },
  {
    name: 'admin-view-as-teacher',
    signedInAs: USERS.omsAdmin,
    async drive(page) {
      await settle(page);
      await page.evaluate(email => window.startViewAs(email), USERS.omsTeacher);
      await page.waitForTimeout(500);

      const banner = await page.evaluate(() => {
        const el = document.getElementById('view-as-banner');
        return el ? el.textContent.replace(/\s+/g, ' ').trim() : '';
      });

      await openTab(page, 'teacher-submit');
      const subbedForOptions = await page.$$eval(
        '#div-sub-dropdown select[name="subbedNameDropdown"] option', els => els.map(e => e.textContent.trim()));

      await openTab(page, 'teacher-totals');
      const kpis = await page.$$eval('#view-container .text-2xl', els => els.map(e => e.textContent.trim()));

      return { banner, subbedForOptions, kpis };
    }
  }
];

main().catch(err => {
  console.error(err);
  process.exit(1);
});
