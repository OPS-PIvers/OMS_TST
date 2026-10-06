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

/** The texts of one cell's notes, in order ([] for none). */
function texts(notes, month, period, day) {
  return (((notes[month] || {})[period] || {})[day] || []).map(n => n.text);
}

exports.name = 'Schedule notes';

exports.run = function ({ test, assert }) {

  test('an admin can note a cell, and reads it back', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: '7th grade team meeting' });

    const notes = env.run('getScheduleNotes', 'OMS');
    assert.deepEqual(texts(notes, 'September', OMS_P1, 'Tue'), ['7th grade team meeting']);
    assert.deepEqual(texts(notes, 'September', OMS_P1, 'Mon'), [], 'blank days have no note');
    assert.equal(notes.October, undefined, 'only the month being edited');
  });

  test('saving without notes leaves the notes alone', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: 'Team meeting' });
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS');
    assert.deepEqual(texts(env.run('getScheduleNotes', 'OMS'), 'September', OMS_P1, 'Tue'), ['Team meeting']);
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
    ['September', 'December', 'June'].forEach(m => assert.deepEqual(texts(notes, m, OMS_P1, 'Wed'), ['PLC'], m));

    env.run('updateSchedulePeriod', 'December', OMS_P1, DAYS, 'OMS', { Wed: '' });
    const after = env.run('getScheduleNotes', 'OMS');
    assert.equal(after.December, undefined);
    assert.deepEqual(texts(after, 'January', OMS_P1, 'Wed'), ['PLC']);
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

    assert.deepEqual(texts(env.run('getScheduleNotes', 'OHS'), 'September', 'Period 1', 'Mon'), ['OHS meeting']);
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
    assert.deepEqual(texts(env.run('getScheduleNotes', 'OMS'), 'September', OMS_P1, 'Tue'), ['Team meeting']);
  });

  test('a Super Admin can read any building\'s notes', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: 'Team meeting' });
    assert.deepEqual(texts(shared(env, USERS.superAdmin).run('getScheduleNotes', 'OMS'), 'September', OMS_P1, 'Tue'),
      ['Team meeting']);
  });

  test('a bad month is refused before anything is written', () => {
    const env = envFor(USERS.omsAdmin);
    const before = JSON.stringify(env.sheet('TST Availability').values);
    assert.rejected(env.attempt('updateSchedulePeriod', 'Smarch', OMS_P1, DAYS, 'OMS', { Tue: 'x' }),
      /unknown month/i);
    assert.equal(JSON.stringify(env.sheet('TST Availability').values), before);
  });

  // --- More than one note per cell, and notes about one teacher ---

  test('a cell can hold several notes, kept in the order they were written', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', {
      Tue: [{ text: 'Team meeting', email: '' }, { text: 'Assembly 2nd half', email: '' }],
      Wed: ['PLC', '  ', 'Fire drill']
    });
    const notes = env.run('getScheduleNotes', 'OMS');
    assert.deepEqual(texts(notes, 'September', OMS_P1, 'Tue'), ['Team meeting', 'Assembly 2nd half']);
    assert.deepEqual(texts(notes, 'September', OMS_P1, 'Wed'), ['PLC', 'Fire drill'], 'blank ones dropped');
  });

  test('a note can be about one teacher in the cell', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', {
      Mon: [{ text: 'Team meeting', email: '' }, { text: 'Only the 2nd half', email: USERS.omsTeacher.toUpperCase() }]
    });
    assert.deepEqual(env.run('getScheduleNotes', 'OMS').September[OMS_P1].Mon, [
      { text: 'Team meeting', email: '' },
      { text: 'Only the 2nd half', email: USERS.omsTeacher }
    ]);
    const row = env.sheet('TST Schedule Notes').values.find(r => r[4] === 'Only the 2nd half');
    assert.equal(row[7], USERS.omsTeacher, 'Teacher Email column, lowercased');
  });

  test('teacher notes follow "every month" like the rest', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS',
      { Mon: [{ text: 'Leaves early', email: USERS.omsTeacher }] }, true);
    const notes = env.run('getScheduleNotes', 'OMS');
    ['September', 'March'].forEach(m => assert.deepEqual(notes[m][OMS_P1].Mon,
      [{ text: 'Leaves early', email: USERS.omsTeacher }], m));
  });

  test('a note about someone outside the building is refused before anything is written', () => {
    const env = envFor(USERS.omsAdmin);
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Tue: 'Team meeting' });
    const avail = JSON.stringify(env.sheet('TST Availability').values);
    const notes = JSON.stringify(env.sheet('TST Schedule Notes').values);

    assert.rejected(env.attempt('updateSchedulePeriod', 'September', OMS_P1, { Mon: [], Tue: [] }, 'OMS',
      { Tue: [{ text: 'x', email: USERS.ohsTeacher }] }), /not on the OMS schedule/i);
    assert.equal(JSON.stringify(env.sheet('TST Availability').values), avail);
    assert.equal(JSON.stringify(env.sheet('TST Schedule Notes').values), notes);
  });

  test('a note about a teacher who has since been archived does not block saving the period', () => {
    const env = envFor(USERS.omsAdmin);
    const ted = [{ text: 'Coaches after school', email: USERS.omsTeacher2 }];
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Thu: ted });
    env.sheet('Staff Directory').values.find(r => r[1] === USERS.omsTeacher2)[9] = 'OMS';

    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS', { Thu: ted, Fri: 'Pep rally' });
    assert.deepEqual(env.run('getScheduleNotes', 'OMS').September[OMS_P1].Thu, ted);
    assert.rejected(env.attempt('updateSchedulePeriod', 'September', 'Period 2 - 9:01 - 9:48', {}, 'OMS',
      { Thu: ted }), /not on the OMS schedule/i, 'only on the period it was already noted on');
  });

  test('notes keep their line breaks, and are capped in length and number', () => {
    const env = envFor(USERS.omsAdmin);
    const many = Array.from({ length: 14 }, (_, i) => 'Note ' + (i + 1));
    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS',
      { Mon: ['Line one\nLine two', 'x'.repeat(250)], Tue: many });
    const notes = env.run('getScheduleNotes', 'OMS');
    const mon = texts(notes, 'September', OMS_P1, 'Mon');
    assert.equal(mon[0], 'Line one\nLine two');
    assert.equal(mon[1].length, 200);
    assert.equal(texts(notes, 'September', OMS_P1, 'Tue').length, 10);
  });

  test('a notes sheet from before teacher notes reads as whole-period notes, and gains the column', () => {
    const env = envFor(USERS.omsAdmin, { 'TST Schedule Notes': [
      ['Building', 'Month', 'Period', 'Day', 'Note', 'Updated', 'Updated By'],
      ['OMS', 'September', OMS_P1, 'Tue', 'Team meeting', '2025-09-01', USERS.omsAdmin]
    ] });
    assert.deepEqual(env.run('getScheduleNotes', 'OMS').September[OMS_P1].Tue, [{ text: 'Team meeting', email: '' }]);

    env.run('updateSchedulePeriod', 'September', OMS_P1, DAYS, 'OMS',
      { Tue: [{ text: 'Team meeting', email: '' }, { text: 'Out Tuesdays', email: USERS.omsTeacher }] });
    const values = env.sheet('TST Schedule Notes').values;
    assert.equal(values[0][7], 'Teacher Email');
    assert.deepEqual(env.run('getScheduleNotes', 'OMS').September[OMS_P1].Tue,
      [{ text: 'Team meeting', email: '' }, { text: 'Out Tuesdays', email: USERS.omsTeacher }]);
  });
};
