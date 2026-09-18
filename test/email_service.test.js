/**
 * Email service health, and the second deployment that turns it back on.
 *
 * Queued mail goes out through each building admin's own trigger, which is what
 * makes them the sender. Triggers are per-user: nobody can install one for anyone
 * else, and Apps Script disables a broken one by telling its owner — not whoever
 * notices their building has gone quiet.
 *
 * So the app watches the symptom rather than the trigger: mail sitting Pending for
 * longer than a working trigger would leave it.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

const QUEUE_HEADER = ['Timestamp', 'Recipient', 'Subject', 'Body', 'Building', 'Status', 'LastUpdated', 'Options'];

function minutesAgo(n) {
  return new Date(new Date().getTime() - n * 60000);
}

/** An Email Queue holding one Pending row for `building`, queued `age` minutes ago. */
function queueWith(building, age) {
  return [QUEUE_HEADER, [minutesAgo(age), USERS.omsTeacher, 'TST', '<p>x</p>', building, 'Pending', '', '{}']];
}

const envFor = (email, overrides) => createEnv({ activeUser: email, sheets: sheets(overrides) });

exports.name = 'Email service health';

exports.run = function ({ test, assert }) {

  // ---- getEmailServiceStatus -------------------------------------------------

  test('a teacher cannot read the email service status', () => {
    const env = envFor(USERS.omsTeacher);
    assert.rejected(env.attempt('getEmailServiceStatus', 'OMS'), /admin access required/i);
  });

  test("an admin asking about another building gets their own", () => {
    const env = envFor(USERS.ohsAdmin);
    assert.equal(env.run('getEmailServiceStatus', 'OMS').building, 'OHS',
      'no read is ever district-wide');
  });

  test('a Super Admin may ask about any building', () => {
    const env = envFor(USERS.superAdmin);
    assert.equal(env.run('getEmailServiceStatus', 'OHS').building, 'OHS');
  });

  test('a building nobody has authorized is reported as never set up', () => {
    const env = envFor(USERS.omsAdmin);
    const status = env.run('getEmailServiceStatus', 'OMS');

    assert.equal(status.healthy, false);
    assert.equal(status.authorized, false);
    assert.equal(status.reason, 'never');
  });

  test('authorizing makes it healthy, and records who did it', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('setupEmailService');

    const status = env.run('getEmailServiceStatus', 'OMS');
    assert.equal(status.healthy, true);
    assert.equal(status.authorized, true);
    assert.equal(status.authorizedBy, 'Amy Admin');
    assert.ok(status.authorizedAt, 'and when');
  });

  test('a multi-building admin authorizes every building they run', () => {
    const env = envFor(USERS.dualAdmin);
    env.run('setupEmailService');

    assert.equal(env.run('getEmailServiceStatus', 'OMS').authorized, true);
    assert.equal(env.run('getEmailServiceStatus', 'OHS').authorized, true);
  });

  test('a queue that has stopped draining is unhealthy even once authorized', () => {
    const env = envFor(USERS.omsAdmin, { 'Email Queue': queueWith('OMS', 40) });
    env.run('setupEmailService');

    const status = env.run('getEmailServiceStatus', 'OMS');
    assert.equal(status.healthy, false, 'the record says set up, but the mail is not moving');
    assert.equal(status.reason, 'stalled');
    assert.ok(status.stalledMinutes >= 40);
  });

  test('mail queued a moment ago is not a stall', () => {
    const env = envFor(USERS.omsAdmin, { 'Email Queue': queueWith('OMS', 2) });
    env.run('setupEmailService');
    assert.equal(env.run('getEmailServiceStatus', 'OMS').healthy, true,
      'a trigger runs about once a minute; two minutes is normal');
  });

  test("one building's backlog does not implicate another", () => {
    const env = envFor(USERS.superAdmin, { 'Email Queue': queueWith('OHS', 60) });
    env.run('setupEmailService'); // Sam Super is assigned to OMS

    assert.equal(env.run('getEmailServiceStatus', 'OMS').healthy, true);
    assert.equal(env.run('getEmailServiceStatus', 'OHS').reason, 'never',
      'OHS has its own admin and its own trigger');
  });

  // ---- The authorization URL -------------------------------------------------

  test('only a Super Admin can set the authorization link', () => {
    const env = envFor(USERS.omsAdmin);
    assert.rejected(env.attempt('setAuthorizeUrl', 'https://script.google.com/a/macros/x/exec'),
      /Super Admin/i);
  });

  test('a Super Admin sets it, and everyone sees it', () => {
    const env = envFor(USERS.superAdmin);
    const url = 'https://script.google.com/a/macros/orono.k12.mn.us/s/AKfy/exec';
    env.run('setAuthorizeUrl', url);

    assert.equal(env.run('getInitialData').authorizeUrl, url);
    assert.equal(env.run('getEmailServiceStatus', 'OMS').authorizeUrl, url);
  });

  test('something that is not an Apps Script URL is refused', () => {
    const env = envFor(USERS.superAdmin);
    assert.rejected(env.attempt('setAuthorizeUrl', 'https://example.com/phish'), /Apps Script/i);
  });

  test('it can be cleared', () => {
    const env = envFor(USERS.superAdmin);
    env.run('setAuthorizeUrl', 'https://script.google.com/a/macros/x/exec');
    env.run('setAuthorizeUrl', '');
    assert.equal(env.run('getInitialData').authorizeUrl, '');
  });

  // ---- The authorization page ------------------------------------------------

  test('opening it as an admin installs that admin\'s triggers', () => {
    const env = envFor(USERS.ohsAdmin);
    const page = env.callInternal('doGet', { parameter: { action: 'authorizeEmail' } });

    assert.ok(/Email service is on/.test(page.getContent()));
    assert.ok(/otto\.admin/.test(page.getContent()), 'it names whose account will send');
    assert.deepEqual(
      env.installedTriggers.map(t => t.getHandlerFunction()),
      ['processEmailQueue', 'processEmailQueue', 'nudgeOutstandingAssignments']);
    assert.equal(env.run('getEmailServiceStatus', 'OHS').healthy, true);
  });

  test('a teacher who finds the link gets a plain refusal, not a stack trace', () => {
    const env = envFor(USERS.omsTeacher);
    const page = env.callInternal('doGet', { parameter: { action: 'authorizeEmail' } });

    assert.ok(/Administrators only/.test(page.getContent()));
    assert.equal(env.installedTriggers.length, 0, 'and nothing is installed');
  });

  test('someone outside the directory is told what to do about it', () => {
    const env = envFor(USERS.stranger);
    const page = env.callInternal('doGet', { parameter: { action: 'authorizeEmail' } });

    assert.ok(/Administrators only/.test(page.getContent()));
    assert.equal(env.installedTriggers.length, 0);
  });

  test('re-opening it does not stack duplicate triggers', () => {
    const env = envFor(USERS.omsAdmin);
    env.callInternal('doGet', { parameter: { action: 'authorizeEmail' } });
    env.callInternal('doGet', { parameter: { action: 'authorizeEmail' } });

    assert.equal(env.installedTriggers.length, 3,
      'it is the documented fix for a broken trigger, so it has to be safe to repeat');
  });

  // ---- Private helpers stay off the client surface ---------------------------

  ['installEmailTriggers_', 'authorizeEmailServicePage_', 'oldestPendingMinutes_',
   'emailAuthorizationFor_', 'authorizeUrl_'].forEach(fn => {
    test(`${fn} is not reachable from google.script.run`, () => {
      const env = envFor(USERS.omsTeacher);
      assert.rejected(env.attempt(fn), /private/i);
      assert.equal(typeof env.context[fn], 'function', 'but server code can still call it');
    });
  });
};
