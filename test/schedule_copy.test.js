/**
 * Copying availability forward through the year.
 *
 * Teachers fill in one month and carry it to the rest of the year rather than
 * entering it ten times; admins can fill the later months for the whole building
 * from one month, without overwriting a month a teacher has already adjusted.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

const envFor = (email, overrides) => createEnv({ activeUser: email, sheets: sheets(overrides) });

const OMS_P1 = 'Period 1 - 8:10 - 8:57';
const OMS_P2 = 'Period 2 - 9:01 - 9:48';
const LATER = ['October', 'November', 'December', 'January', 'February', 'March', 'April', 'May', 'June'];

/** One person's availability rows, as [month, days, period]. */
function rowsFor(env, email) {
  const sheet = env.spreadsheet.sheets.find(s => s.name === 'TST Availability');
  return sheet.values.slice(1).filter(r => r[4] === email).map(r => [r[0], r[1], r[2]]);
}

const monthsOf = rows => Array.from(new Set(rows.map(r => r[0])));

exports.name = 'Schedule copy forward';

exports.run = function ({ test, assert }) {

  test('a teacher\'s save stays in its month unless they ask to copy it', () => {
    const env = envFor(USERS.omsTeacher);
    const res = env.run('saveAvailability', 'September', [{ period: OMS_P1, days: 'Wed' }], null, [OMS_P1]);
    assert.deepEqual(res.months, ['September']);
    assert.deepEqual(rowsFor(env, USERS.omsTeacher), [['September', 'Wed', OMS_P1]]);
  });

  test('copying forward writes the month into every later month', () => {
    const env = envFor(USERS.omsTeacher);
    const res = env.run('saveAvailability', 'September',
      [{ period: OMS_P1, days: 'Mon,Tue' }, { period: OMS_P2, days: 'Fri' }], null, [OMS_P1, OMS_P2], true);
    assert.deepEqual(res.months, ['September'].concat(LATER));
    const rows = rowsFor(env, USERS.omsTeacher);
    assert.deepEqual(monthsOf(rows), ['September'].concat(LATER));
    assert.equal(rows.length, 20, 'two periods in each of ten months');
    assert.deepEqual(rows.filter(r => r[0] === 'March'), [['March', 'Mon,Tue', OMS_P1], ['March', 'Fri', OMS_P2]]);
  });

  test('copying forward replaces later months but leaves earlier ones', () => {
    const env = envFor(USERS.omsTeacher);
    env.run('saveAvailability', 'September', [{ period: OMS_P1, days: 'Mon' }], null, [OMS_P1], true);
    env.run('saveAvailability', 'January', [{ period: OMS_P1, days: 'Thu' }], null, [OMS_P1], true);
    const rows = rowsFor(env, USERS.omsTeacher);
    assert.deepEqual(rows.filter(r => r[0] === 'December'), [['December', 'Mon', OMS_P1]]);
    assert.deepEqual(rows.filter(r => r[0] === 'June'), [['June', 'Thu', OMS_P1]]);
    assert.equal(rows.length, 10, 'one row per month, nothing duplicated');
  });

  test('copying forward from June is just a June save', () => {
    const env = envFor(USERS.omsTeacher);
    const res = env.run('saveAvailability', 'June', [{ period: OMS_P1, days: 'Mon' }], null, [OMS_P1], true);
    assert.deepEqual(res.months, ['June']);
  });

  test('a multi-building teacher\'s other building is left alone by the copy', () => {
    const env = envFor(USERS.multiTeacher);
    env.run('saveAvailability', 'November', [{ period: 'Period 3', days: 'Tue' }], null, ['Period 3']);
    env.run('saveAvailability', 'October', [{ period: OMS_P1, days: 'Mon' }], null, [OMS_P1], true);
    const nov = rowsFor(env, USERS.multiTeacher).filter(r => r[0] === 'November');
    assert.deepEqual(nov, [['November', 'Tue', 'Period 3'], ['November', 'Mon', OMS_P1]]);
  });

  test('a teacher cannot copy forward on someone else\'s behalf', () => {
    const env = envFor(USERS.omsTeacher);
    assert.rejected(env.attempt('saveAvailability', 'September', [{ period: OMS_P1, days: 'Mon' }],
      USERS.omsTeacher2, [OMS_P1], true), /admin/i);
  });

  test('an admin fills the later months from one month, for the whole building', () => {
    const env = envFor(USERS.omsAdmin);
    const res = env.run('copyScheduleForward', 'October', 'OMS', [OMS_P1, OMS_P2]);
    // Only Mia has October rows in the fixtures.
    assert.equal(res.people, 1);
    assert.equal(res.filled, 8);
    assert.equal(res.skipped, 0);
    const rows = rowsFor(env, USERS.multiTeacher);
    assert.deepEqual(monthsOf(rows), LATER);
    assert.deepEqual(rows.filter(r => r[0] === 'June'), [['June', 'Fri', OMS_P1]]);
    // Tina has September only, so nothing of hers is copied from October.
    assert.deepEqual(monthsOf(rowsFor(env, USERS.omsTeacher)), ['September']);
  });

  test('the admin copy never overwrites a month a teacher has already filled in', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'December', OMS_P2, { Mon: [USERS.multiTeacher], Tue: [], Wed: [], Thu: [], Fri: [] }, 'OMS');
    const res = env.run('copyScheduleForward', 'October', 'OMS', [OMS_P1, OMS_P2]);
    assert.equal(res.filled, 7);
    assert.equal(res.skipped, 1);
    assert.deepEqual(rowsFor(env, USERS.multiTeacher).filter(r => r[0] === 'December'),
      [['December', 'Mon', OMS_P2]]);

    const again = env.run('copyScheduleForward', 'October', 'OMS', [OMS_P1, OMS_P2]);
    assert.equal(again.filled, 0, 'running it twice adds nothing');
  });

  test('the admin copy only touches its own building', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('copyScheduleForward', 'September', 'OMS', [OMS_P1, OMS_P2]);
    assert.deepEqual(monthsOf(rowsFor(env, USERS.ohsTeacher)), ['September'], 'OHS teacher untouched');
    assert.deepEqual(monthsOf(rowsFor(env, USERS.omsTeacher)), ['September'].concat(LATER));
  });

  test('who may copy the schedule forward', () => {
    assert.rejected(envFor(USERS.omsTeacher).attempt('copyScheduleForward', 'October', 'OMS'), /admin/i);
    assert.rejected(envFor(USERS.ohsAdmin).attempt('copyScheduleForward', 'October', 'OMS'), /own building/i);
    assert.allowed(envFor(USERS.superAdmin).attempt('copyScheduleForward', 'October', 'OMS'));
    assert.rejected(envFor(USERS.omsAdmin).attempt('copyScheduleForward', 'June', 'OMS'), /no months after/i);
    assert.rejected(envFor(USERS.omsAdmin).attempt('copyScheduleForward', 'July', 'OMS'), /unknown month/i);
  });
};
