/**
 * The background triggers and the lock they take.
 *
 * Every building admin's trigger runs processEmailQueue every minute, and it can
 * hold its lock while it sends a whole batch of mail. It used to take the script
 * lock — the same one a teacher's Submit waits on for its duplicate guard — so a
 * submission could sit behind someone else's email run for up to 20 seconds.
 *
 * The processors now share the document lock among themselves, and on a minute
 * with nothing to do they take no lock and read no directory at all.
 */

const { createEnv, FakeSheet } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

const QUEUE_HEADER = ['Timestamp', 'Recipient', 'Subject', 'Body', 'Building', 'Status', 'LastUpdated', 'Options'];
const hoursAgo = n => new Date(new Date().getTime() - n * 3600000);
const pendingRow = building => [new Date(), USERS.omsTeacher, 'TST', '<p>x</p>', building, 'Pending', '', '{}'];

function envFor(email, queueRows, options) {
  return createEnv(Object.assign({
    activeUser: email,
    sheets: sheets({ 'Email Queue': [QUEUE_HEADER].concat(queueRows || []) })
  }, options || {}));
}

const trigger = env => env.run('processEmailQueue', { authMode: env.AuthMode.FULL, triggerUid: '1' });
const locksTaken = env => env.lockLog.filter(l => /:acquire$/.test(l)).map(l => l.split(':')[0]);

function readsDuring(fn) {
  const counts = {};
  const original = FakeSheet.prototype.getDataRange;
  FakeSheet.prototype.getDataRange = function () {
    counts[this.name] = (counts[this.name] || 0) + 1;
    return original.call(this);
  };
  try { fn(); } finally { FakeSheet.prototype.getDataRange = original; }
  return counts;
}

exports.name = 'Background trigger locking';

exports.run = function ({ test, assert }) {

  test('the email trigger sends under the document lock, not the script lock', () => {
    const env = envFor(USERS.omsAdmin, [pendingRow('OMS')]);
    trigger(env);
    assert.equal(env.sentEmails.length, 1, 'the queued email still goes out');
    assert.ok(locksTaken(env).length > 0, 'it still takes a lock');
    assert.ok(locksTaken(env).every(k => k === 'document'), 'only the document lock: ' + env.lockLog.join(', '));
    assert.equal(env.lockLog.filter(l => l === 'document:acquire').length,
      env.lockLog.filter(l => l === 'document:release').length, 'and releases what it takes');
  });

  test("a teacher's submission still takes the script lock, so it no longer queues behind mail", () => {
    const env = envFor(USERS.omsTeacher);
    env.run('submitEarned', {
      email: USERS.omsTeacher, name: 'Tina Teacher', subbedFor: 'Ted Teacher',
      date: '2026-09-15', period: 'Period 1 - 8:10 - 8:57', amount: 1, amountType: 'Full Period', building: 'OMS'
    });
    assert.ok(locksTaken(env).includes('script'), 'the duplicate guard is still locked: ' + env.lockLog.join(', '));
    assert.ok(!locksTaken(env).includes('document'), "and never on the processors' lock");
  });

  test('an idle minute takes no lock and does not read the Staff Directory', () => {
    const env = envFor(USERS.omsAdmin, []);
    const reads = readsDuring(() => trigger(env));
    assert.equal(locksTaken(env).length, 0, 'no lock: ' + env.lockLog.join(', '));
    assert.ok(!reads['Staff Directory'], 'no directory read: ' + JSON.stringify(reads));
  });

  test('mail for another building alone still runs (that admin may be the one to send it)', () => {
    const env = envFor(USERS.ohsAdmin, [pendingRow('OMS')]);
    trigger(env);
    assert.equal(env.sentEmails.length, 0, 'but the OHS admin never sends OMS mail');
    assert.equal(env.sheet('Email Queue').values[1][5], 'Pending');
  });

  test('old sent rows are still cleaned up when nothing is pending', () => {
    const env = envFor(USERS.omsAdmin, [
      [hoursAgo(30), USERS.omsTeacher, 'Old', '<p>x</p>', 'OMS', 'Sent', hoursAgo(30), '{}'],
      [hoursAgo(2), USERS.omsTeacher, 'Recent', '<p>x</p>', 'OMS', 'Sent', hoursAgo(2), '{}']
    ]);
    trigger(env);
    const subjects = env.sheet('Email Queue').values.slice(1).map(r => r[2]);
    assert.deepEqual(subjects, ['Recent'], 'the day-old row is gone, the recent one kept');
  });

  test('without a document lock it falls back to the script lock and still sends', () => {
    const env = envFor(USERS.omsAdmin, [pendingRow('OMS')], { noDocumentLock: true });
    trigger(env);
    assert.equal(env.sentEmails.length, 1);
    assert.ok(locksTaken(env).every(k => k === 'script'), env.lockLog.join(', '));
  });
};
