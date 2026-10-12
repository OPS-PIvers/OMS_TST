/**
 * lateFinalize / undoLateFinalize: finalize last school year for a building that
 * never ran Finalize and kept adding this year's hours on top. It must do what
 * finalizeSchoolYear would have done in June — roll last year's ending balance
 * into Carry Over capped at carryOverMax, zero Paid Out, archive last year's rows
 * — while leaving this year's activity alone, do nothing without apply === true,
 * and be fully undoable.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, APPROVALS_HEADER, USAGE_HEADER, sheets } = require('./fixtures');

const YEAR = '2025-26';
const LABEL = '2025-26 OHS (late finalize)';

const earnedRow = (email, name, date, hours, approved, denied, building) =>
  [email, name, 'Someone', '', date, 'Period 1', 'Full Period', hours, approved, '', denied, '', '', building];

// Hank: OHS primary, Carry Over 5, Paid Out 1 (set in envFor).
// Otto Admin: OHS primary, Carry Over 2, no activity. Mia: OMS primary, also at OHS.
const approvals = () => [APPROVALS_HEADER,
  earnedRow(USERS.ohsTeacher, 'Hank Teacher', '2026-03-10', 3, true, false, 'OHS'),   // last year
  earnedRow(USERS.ohsTeacher, 'Hank Teacher', '2026-04-11', 10, true, false, 'OMS'),  // last year, any building
  earnedRow(USERS.ohsTeacher, 'Hank Teacher', '2026-05-01', 1, false, false, 'OHS'),  // pending: reported, stays
  earnedRow(USERS.ohsTeacher, 'Hank Teacher', '2026-05-02', 1, false, true, 'OHS'),   // denied: stays
  earnedRow(USERS.ohsTeacher, 'Hank Teacher', '2026-09-20', 2, true, false, 'OHS'),   // this year
  earnedRow(USERS.multiTeacher, 'Mia Multi', '2026-04-01', 2, true, false, 'OHS'),    // primary OMS: stays
  earnedRow(USERS.omsTeacher, 'Tina Teacher', '2026-04-02', 1, true, false, 'OMS')    // OMS: stays
];

const usage = () => [USAGE_HEADER,
  [USERS.ohsTeacher, 'Hank Teacher', '2026-03-10', 2, true, '2026-03-10', '', 'OHS'],   // last year
  [USERS.ohsTeacher, 'Hank Teacher', '2026-05-20', 1.5, true, '2026-05-20', '', 'OMS'], // last year
  [USERS.ohsTeacher, 'Hank Teacher', '2026-05-21', 1, false, '', '', 'OHS'],            // pending: reported, stays
  [USERS.ohsTeacher, 'Hank Teacher', '2026-09-15', 1, true, '2026-09-15', '', 'OHS'],   // this year
  [USERS.ohsTeacher, 'Hank Teacher', 'sometime', 1, true, '', '', 'OHS'],               // unreadable: reported, stays
  [USERS.multiTeacher, 'Mia Multi', '2026-04-01', 3, true, '2026-04-01', '', 'OHS'],    // primary OMS: stays
  [USERS.omsTeacher, 'Tina Teacher', '2026-04-02', 1, true, '2026-04-02', '', 'OMS']    // OMS: stays
];

// Staff Directory columns: F(5) Carry Over, G(6) Paid Out, K(10) Last Finalized, L(11) Pending Finalize.
const staffRow = (env, email) => env.sheet('Staff Directory').values.find(r => r[1] === email);

function envFor(email) {
  const env = createEnv({ activeUser: email,
    sheets: sheets({ 'TST Approvals (New)': approvals(), 'TST Usage (New)': usage() }) });
  staffRow(env, USERS.ohsTeacher)[6] = 1;
  staffRow(env, USERS.ohsTeacher)[11] = '2025-26';
  return env;
}

const dirFor = (env, email) => env.run('getStaffDirectoryData', email === USERS.omsTeacher ? 'OMS' : 'OHS')
  .find(s => s.email === email);
const person = (r, email) => r.byPerson.find(p => p.email === email);

exports.name = 'Late finalize (lateFinalize)';

exports.run = function ({ test, assert }) {

  test('only a Super Admin can run it or undo it', () => {
    [USERS.ohsAdmin, USERS.omsAdmin, USERS.ohsTeacher].forEach(email => {
      const env = envFor(email);
      assert.rejected(env.attempt('lateFinalize', 'OHS', '2026-07-01', YEAR, true), /super admin/i, email);
      assert.rejected(env.attempt('undoLateFinalize', 'OHS', YEAR, true), /super admin/i, email);
    });
  });

  test('a dry run reports the roll, the cap and each total, and changes nothing', () => {
    const env = envFor(USERS.superAdmin);
    const before = JSON.stringify(['Staff Directory', 'TST Approvals (New)', 'TST Usage (New)'].map(n => env.sheet(n).values));
    const r = env.run('lateFinalize', 'OHS', '2026-07-01', YEAR);

    assert.equal(r.applied, false);
    assert.equal(r.cap, 12);
    assert.deepEqual([r.earned.rowCount, r.earned.totalHours, r.used.rowCount, r.used.totalHours], [2, 13, 2, 3.5]);
    const h = person(r, USERS.ohsTeacher);
    // 5 carry + 13 earned - 3.5 used - 1 paid = 13.5 at the end of last year → 12, 1.5 forfeited.
    assert.deepEqual([h.lastYearEarned, h.lastYearUsed, h.endOfLastYear, h.newCarryOver, h.forfeited], [13, 3.5, 13.5, 12, 1.5]);
    assert.deepEqual([h.thisYearEarned, h.thisYearUsed], [2, 2], 'the unreadable Used row stays in this year');
    assert.deepEqual([h.totalNow, h.totalAfter], [13.5, 12]);
    assert.deepEqual(r.overCap.map(o => [o.email, o.forfeited]), [[USERS.ohsTeacher, 1.5]]);
    assert.equal(r.totalForfeited, 1.5);

    const otto = person(r, USERS.ohsAdmin);
    assert.deepEqual([otto.newCarryOver, otto.totalNow, otto.totalAfter], [2, 2, 2], 'no activity: nothing changes');
    assert.ok(!person(r, USERS.multiTeacher), 'OMS-primary staff are not part of it');

    assert.equal(r.skippedPending.length, 2);
    assert.equal(r.unreadableDates.length, 1);
    assert.equal(JSON.stringify(['Staff Directory', 'TST Approvals (New)', 'TST Usage (New)'].map(n => env.sheet(n).values)), before);
    assert.equal(env.sheet(LABEL), null);
  });

  test('every total changes by exactly the hours forfeited, nothing more', () => {
    const env = envFor(USERS.superAdmin);
    env.run('lateFinalize', 'OHS', '2026-07-01', YEAR).byPerson
      .forEach(p => assert.equal(Math.round((p.totalNow - p.totalAfter) * 100) / 100, p.forfeited, p.name));
  });

  test('applying produces exactly what the dry run promised, in the Directory', () => {
    const env = envFor(USERS.superAdmin);
    const h = person(env.run('lateFinalize', 'OHS', '2026-07-01', YEAR), USERS.ohsTeacher);
    env.run('lateFinalize', 'OHS', '2026-07-01', YEAR, true);

    const d = dirFor(env, USERS.ohsTeacher);
    assert.deepEqual([d.carryOver, d.paidOut, d.earned, d.used, d.total],
      [h.newCarryOver, 0, h.thisYearEarned, h.thisYearUsed, h.totalAfter]);
    assert.equal(staffRow(env, USERS.ohsTeacher)[10], YEAR, 'Last Finalized set');
    assert.equal(staffRow(env, USERS.ohsTeacher)[11], '', 'Pending Finalize cleared');
  });

  test('only last year\'s approved rows move; other staff are untouched', () => {
    const env = envFor(USERS.superAdmin);
    const mia = JSON.stringify(staffRow(env, USERS.multiTeacher));
    const tina = JSON.stringify(staffRow(env, USERS.omsTeacher));
    env.run('lateFinalize', 'OHS', '2026-07-01', YEAR, true);

    assert.equal(env.sheet('TST Approvals (New)').values.length, 6, 'pending, denied, this year and other staff stay');
    assert.equal(env.sheet('TST Usage (New)').values.length, 6);
    assert.equal(env.sheet('TST Approvals Archive').values.length, 3);
    assert.equal(env.sheet('TST Usage Archive').values.length, 3);
    assert.ok(env.sheet('TST Approvals Archive').values.slice(1).every(r => r[APPROVALS_HEADER.length] === LABEL));
    assert.ok(env.sheet('TST Usage Archive').values.slice(1).every(r => r[USAGE_HEADER.length] === LABEL));
    assert.equal(JSON.stringify(staffRow(env, USERS.multiTeacher)), mia);
    assert.equal(JSON.stringify(staffRow(env, USERS.omsTeacher)), tina);
    assert.deepEqual([dirFor(env, USERS.multiTeacher).earned, dirFor(env, USERS.multiTeacher).used], [2, 3]);
  });

  test('it shows up under View Archived Years like a normal finalize', () => {
    const env = envFor(USERS.superAdmin);
    env.run('lateFinalize', 'OHS', '2026-07-01', YEAR, true);
    assert.ok(env.run('listArchivedYears', 'OHS').some(a => a.name === LABEL));
  });

  test('it cannot roll anyone twice', () => {
    const env = envFor(USERS.superAdmin);
    env.run('lateFinalize', 'OHS', '2026-07-01', YEAR, true);
    assert.rejected(env.attempt('lateFinalize', 'OHS', '2026-07-01', YEAR, true), /already exists/i);
    const again = env.run('lateFinalize', 'OHS', '2026-07-01', YEAR);
    assert.equal(again.staffCount, 0);
    assert.ok(again.alreadyFinalized.includes('Hank Teacher'));
  });

  test('undo puts every row and every Directory value back', () => {
    const env = envFor(USERS.superAdmin);
    const staff = JSON.stringify(env.sheet('Staff Directory').values);
    env.run('lateFinalize', 'OHS', '2026-07-01', YEAR, true);

    const dry = env.run('undoLateFinalize', 'OHS', YEAR);
    assert.deepEqual([dry.earned.rowCount, dry.used.rowCount], [2, 2]);
    assert.equal(dirFor(env, USERS.ohsTeacher).carryOver, 12, 'dry run changes nothing');

    env.run('undoLateFinalize', 'OHS', YEAR, true);
    assert.equal(JSON.stringify(env.sheet('Staff Directory').values), staff);
    const sortRows = rows => rows.map(r => JSON.stringify(r)).sort();
    assert.deepEqual(sortRows(env.sheet('TST Approvals (New)').values), sortRows(approvals()));
    assert.deepEqual(sortRows(env.sheet('TST Usage (New)').values), sortRows(usage()));
    assert.equal(env.sheet('TST Approvals Archive').values.length, 1, 'header only');
    assert.equal(env.sheet('TST Usage Archive').values.length, 1, 'header only');
    assert.equal(env.sheet(LABEL), null, 'snapshot removed');
    assert.equal(dirFor(env, USERS.ohsTeacher).total, 13.5);
  });

  test('undo keeps a correction made to someone after the late finalize', () => {
    const env = envFor(USERS.superAdmin);
    env.run('lateFinalize', 'OHS', '2026-07-01', YEAR, true);
    staffRow(env, USERS.ohsTeacher)[5] = 11;   // admin knocks an hour off afterwards
    env.run('undoLateFinalize', 'OHS', YEAR, true);
    assert.equal(staffRow(env, USERS.ohsTeacher)[5], 4, 'original 5, less the later 1-hour correction');
    assert.equal(staffRow(env, USERS.ohsTeacher)[6], 1);
  });

  test('undo leaves rows archived by a real finalize alone', () => {
    const env = envFor(USERS.superAdmin);
    env.run('lateFinalize', 'OHS', '2026-07-01', YEAR, true);
    env.sheet('TST Usage Archive').values.push([USERS.omsTeacher, 'Tina Teacher', '2025-04-01', 4, true, '', '', 'OMS', YEAR]);
    env.run('undoLateFinalize', 'OHS', YEAR, true);
    assert.equal(env.sheet('TST Usage Archive').values.length, 2);
  });

  test('building, cutoff and year are all required', () => {
    const env = envFor(USERS.superAdmin);
    assert.rejected(env.attempt('lateFinalize', '', '2026-07-01', YEAR, true), /name the building/i);
    assert.rejected(env.attempt('lateFinalize', 'OHS', 'not a date', YEAR, true), /cutoff/i);
    assert.rejected(env.attempt('lateFinalize', 'OHS', '2026-07-01', '', true), /school year/i);
    assert.rejected(env.attempt('undoLateFinalize', 'OHS', '2024-25', true), /no late finalize/i);
  });
};
