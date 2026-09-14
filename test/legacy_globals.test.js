/**
 * Code_legacy.js must not be deployed.
 *
 * Apps Script has one global scope shared by every file in the project and
 * evaluates them in project order, so a duplicate top-level `function` in a later
 * file silently replaces the earlier one. clasp pushes alphabetically, which puts
 * Code_legacy.js after Code.js — its onFormSubmit, onOpen and calculatePeriods were
 * the live definitions. These tests pin both halves of the fix: the deployed file
 * set (.claspignore) and what shadowing actually did.
 */

const fs = require('fs');
const path = require('path');
const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

const ROOT = path.join(__dirname, '..');
const DEPLOYED = ['Code.js', 'config.js'];              // project order, alphabetical
const WITH_LEGACY = ['Code.js', 'Code_legacy.js', 'config.js'];  // what clasp used to push

exports.name = 'Legacy file is out of the deployment';

exports.run = function ({ test, assert }) {

  test('.claspignore allow-lists exactly the four deployable files', () => {
    const lines = fs.readFileSync(path.join(ROOT, '.claspignore'), 'utf8')
      .split(/\r?\n/).map(l => l.trim()).filter(l => l && !l.startsWith('#'));

    assert.equal(lines[0], '**/**', 'the allow-list has to start by ignoring everything');
    const allowed = lines.filter(l => l.startsWith('!')).map(l => l.slice(1)).sort();
    assert.deepEqual(allowed, ['Code.js', 'Index.html', 'appsscript.json', 'config.js']);
  });

  test('the legacy file and the tests are not in the deployed set', () => {
    const allowed = fs.readFileSync(path.join(ROOT, '.claspignore'), 'utf8')
      .split(/\r?\n/).map(l => l.trim())
      .filter(l => l.startsWith('!')).map(l => l.slice(1));

    assert.ok(!allowed.includes('Code_legacy.js'), 'Code_legacy.js must stay out of Apps Script');
    assert.ok(!allowed.some(f => f.startsWith('test')), 'Node tests must stay out of Apps Script');
    assert.ok(fs.existsSync(path.join(ROOT, 'Code_legacy.js')), 'the file is kept for reference');
  });

  test('deployed: the form trigger handler is the one in Code.js', () => {
    const env = createEnv({ activeUser: USERS.omsAdmin, sheets: sheets(), files: DEPLOYED });

    // A real trigger event carries a live Range; a browser call cannot.
    const forged = env.attempt('onFormSubmit', { values: ['', USERS.omsTeacher, 'Ted', '', '2025-09-20', 'Period 1', 'Full', 1] });
    assert.rejected(forged, /only be run by the form submit trigger/i);
  });

  test('with the legacy file pushed, that guard was shadowed away', () => {
    const env = createEnv({ activeUser: USERS.omsAdmin, sheets: sheets(), files: WITH_LEGACY });

    const forged = env.attempt('onFormSubmit', { values: ['', USERS.omsTeacher, 'Ted', '', '2025-09-20', 'Period 1', 'Full', 1] });
    assert.ok(forged.ok, 'the legacy handler accepted a client-shaped event — that is the bug');
  });

  test('deployed: onOpen builds the TST Admin menu, not the legacy TST Time one', () => {
    const env = createEnv({ activeUser: USERS.omsAdmin, sheets: sheets(), files: DEPLOYED });
    env.run('onOpen');

    assert.equal(env.menus.length, 1);
    assert.equal(env.menus[0].name, 'TST Admin');
    assert.equal(env.menus[0].items[0].fn, 'setupEmailService',
      'admins authorize the email service from this menu');
  });

  test('with the legacy file pushed, the menu was the legacy one', () => {
    const env = createEnv({ activeUser: USERS.omsAdmin, sheets: sheets(), files: WITH_LEGACY });
    env.run('onOpen');

    assert.equal(env.menus[0].name, 'TST Time', 'which is why Authorize Email Service was missing');
  });

  test('deployed: calculatePeriods is the building-aware version', () => {
    const env = createEnv({ activeUser: USERS.omsAdmin, sheets: sheets(), files: DEPLOYED });

    // The Code.js signature is (period, amountType, buildingCode): the OMS half-period
    // rule applies at OMS only.
    assert.equal(env.run('calculatePeriods', 'Period 6 - 11:40 - 12:06', 'Full Period', 'OMS'), 0.5);
    assert.equal(env.run('calculatePeriods', 'Period 6', 'Full Period', 'OHS'), 1);
  });

  test('with the legacy file pushed, calculatePeriods ignored the building', () => {
    const env = createEnv({ activeUser: USERS.omsAdmin, sheets: sheets(), files: WITH_LEGACY });

    // The legacy signature is (period, amount) — the third argument is dropped, so
    // every building got the OMS rule.
    assert.equal(env.run('calculatePeriods', 'Period 6 ', 'Full Period', 'OHS'), 0.5);
  });

  test('setupEmailService still installs the trigger for an admin', () => {
    const env = createEnv({ activeUser: USERS.omsAdmin, sheets: sheets(), files: DEPLOYED });
    env.run('setupEmailService');

    const handlers = env.installedTriggers.map(t => t.getHandlerFunction());
    assert.deepEqual(handlers, ['processEmailQueue', 'processEmailQueue'],
      'onChange + 1-minute timer, both still named processEmailQueue');
  });

  test('setupEmailService refuses a teacher', () => {
    const env = createEnv({ activeUser: USERS.omsTeacher, sheets: sheets(), files: DEPLOYED });
    assert.rejected(env.attempt('setupEmailService'), /admin access required/i);
    assert.equal(env.installedTriggers.length, 0);
  });
};
