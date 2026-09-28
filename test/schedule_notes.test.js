/**
 * Admin notes on the Master Schedule ("7th grade team meeting").
 *
 * A note belongs to one building's month / period / weekday cell. It is the admin's
 * planning, so teachers can neither read nor write it, and one building's notes
 * never touch another's even where the period names are shared.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, sheets } = require('./fixtures');

const envFor = (email, overrides) => createEnv({ activeUser: email, sheets: sheets(overrides) });

const OMS_P1 = 'Period 1 - 8:10 - 8:57';
const DAYS = { Mon: [USERS.omsTeacher], Tue: [USERS.omsTeacher], Wed: [], Thu: [], Fri: [] };

function shared(env, email) {
  const all = env.spreadsheet.sheets.reduce((acc, s) => { acc[s.name] = s.values; return acc; }, {});
  return createEnv({ activeUser: email, sheets: all });
}

exports.name = 'Schedule notes';

exports.run = function ({ test, assert }) {

  test('an admin can note a cell, and reads it back', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: '7th grade team meeting' });

    const notes = env.run('getScheduleNotes', 'OMS');
    assert.equal(notes.September[OMS_P1].Tue, '7th grade team meeting');
    assert.equal(notes.September[OMS_P1].Mon, undefined, 'blank days have no note');
    assert.equal(notes.October, undefined, 'only the month being edited');
  });

  test('saving without notes leaves the notes alone', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: 'Team meeting' });
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS');
    assert.equal(env.run('getScheduleNotes', 'OMS').September[OMS_P1].Tue, 'Team meeting');
  });

  test('a blank note clears it', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: 'Team meeting' });
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: '  ' });
    assert.deepEqual(env.run('getScheduleNotes', 'OMS'), {});
  });

  test('"every month" writes the note into each month, and editing one month later stays local', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Wed: 'PLC' }, true);

    const notes = env.run('getScheduleNotes', 'OMS');
    ['September', 'December', 'June'].forEach(m => assert.equal(notes[m][OMS_P1].Wed, 'PLC', m));

    env.run('updateSchedulePeriod', 'December', OMS_P1, DAYS, 'OMS', { Wed: '' });
    const after = env.run('getScheduleNotes', 'OMS');
    assert.equal(after.December, undefined);
    assert.equal(after.January[OMS_P1].Wed, 'PLC');
  });

  test('saving notes does not disturb the availability it was saved with', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Mon: 'Team meeting' });
    const tina = env.run('getScheduleData', 'OMS').September.filter(e => e.email === USERS.omsTeacher);
    assert.equal(tina.length, 1);
    assert.equal(tina[0].days, 'Mon,Tue');
  });

  test('one building\'s notes do not touch another\'s with the same period name', () => {
    const env = envFor(USERS.superAdmin);
    env.run('updateSchedulePeriod', 'September', 'Period 1', {}, 'OHS', { Mon: 'OHS meeting' });
    env.run('updateSchedulePeriod', 'September', 'Period 1', {}, 'OMS', { Mon: 'OMS meeting' });
    env.run('updateSchedulePeriod', 'September', 'Period 1', {}, 'OMS', { Mon: '' });

    assert.equal(env.run('getScheduleNotes', 'OHS').September['Period 1'].Mon, 'OHS meeting');
    assert.deepEqual(env.run('getScheduleNotes', 'OMS'), {});
  });

  test('a teacher can neither read nor write notes', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: 'Team meeting' });
    const tina = shared(env, USERS.omsTeacher);

    assert.rejected(tina.attempt('getScheduleNotes', 'OMS'), /admin access required/i);
    assert.rejected(tina.attempt('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: 'x' }),
      /admin access required/i);
  });

  test('an admin from another building gets their own building\'s notes, not these', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: 'Team meeting' });
    const otto = shared(env, USERS.ohsAdmin);

    assert.deepEqual(otto.run('getScheduleNotes', 'OMS'), {}, 'falls back to OHS, which has none');
    assert.rejected(otto.attempt('updateSchedulePeriod', 'September', OMS_P1, {}, 'OMS', { Tue: '' }),
      /your own building/i);
    assert.equal(env.run('getScheduleNotes', 'OMS').September[OMS_P1].Tue, 'Team meeting');
  });

  test('a Super Admin can read any building\'s notes', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: 'Team meeting' });
    assert.equal(shared(env, USERS.superAdmin).run('getScheduleNotes', 'OMS').September[OMS_P1].Tue,
      'Team meeting');
  });

  test('a bad month is refused before anything is written', () => {
    const env = envFor(USERS.omsAdmin);
    const before = JSON.stringify(env.sheet('TST Availability').values);
    assert.rejected(env.attempt('updateSchedulePeriod', 'Smarch', OMS_P1, DAYS, 'OMS', { Tue: 'x' }),
      /unknown month/i);
    assert.equal(JSON.stringify(env.sheet('TST Availability').values), before);
  });
};
