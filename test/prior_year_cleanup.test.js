/**
 * archivePriorYear / restorePriorYear: the one-off repair for a building that
 * skipped Finalize and set Carry Over by hand, leaving last year's approved Earned
 * and Used rows counted on top of it. It must only move what finalize would have
 * moved, move both sides together, do nothing without apply === true, report each
 * person's balance before and after, and be fully undoable.
 */

const { createEnv } = require('./apps_script_env');
const { USERS, APPROVALS_HEADER, USAGE_HEADER, sheets } = require('./fixtures');

const LABEL = '2025-26 OHS cleanup';

const earnedRow = (email, name, date, hours, approved, denied, building) =>
  [email, name, 'Someone', '', date, 'Period 1', 'Full Period', hours, approved, '', denied, '', '', building];

// Hank: OHS primary (Carry Over 5). Mia: OMS primary, also at OHS. Tina: OMS only.
const approvals = () => [APPROVALS_HEADER,
  earnedRow(USERS.ohsTeacher, 'Hank Teacher', '2026-03-10', 3, true, false, 'OHS'),   // moves
  earnedRow(USERS.ohsTeacher, 'Hank Teacher', '2026-04-11', 1, true, false, 'OMS'),   // moves: primary OHS, any tag
  earnedRow(USERS.ohsTeacher, 'Hank Teacher', '2026-05-01', 1, false, false, 'OHS'),  // pending: reported, stays
  earnedRow(USERS.ohsTeacher, 'Hank Teacher', '2026-05-02', 1, false, true, 'OHS'),   // denied: ignored, stays
  earnedRow(USERS.ohsTeacher, 'Hank Teacher', '2026-09-20', 2, true, false, 'OHS'),   // this year: stays
  earnedRow(USERS.multiTeacher, 'Mia Multi', '2026-04-01', 2, true, false, 'OHS'),    // primary OMS: stays
  earnedRow(USERS.omsTeacher, 'Tina Teacher', '2026-04-02', 1, true, false, 'OMS')    // OMS: stays
];

const usage = () => [USAGE_HEADER,
  [USERS.ohsTeacher, 'Hank Teacher', '2026-03-10', 2, true, '2026-03-10', '', 'OHS'],   // moves
  [USERS.ohsTeacher, 'Hank Teacher', '2026-05-20', 1.5, true, '2026-05-20', '', 'OMS'], // moves
  [USERS.ohsTeacher, 'Hank Teacher', '2026-05-21', 1, false, '', '', 'OHS'],            // pending: reported, stays
  [USERS.ohsTeacher, 'Hank Teacher', '2026-09-15', 1, true, '2026-09-15', '', 'OHS'],   // this year: stays
  [USERS.ohsTeacher, 'Hank Teacher', 'sometime', 1, true, '', '', 'OHS'],               // unreadable: reported, stays
  [USERS.multiTeacher, 'Mia Multi', '2026-04-01', 3, true, '2026-04-01', '', 'OHS'],    // primary OMS: stays
  [USERS.omsTeacher, 'Tina Teacher', '2026-04-02', 1, true, '2026-04-02', '', 'OMS']    // OMS: stays
];

const envFor = email => createEnv({ activeUser: email,
  sheets: sheets({ 'TST Approvals (New)': approvals(), 'TST Usage (New)': usage() }) });
// What the Directory shows — the numbers the OHS admin asked about.
const dirFor = (env, email) => env.run('getStaffDirectoryData', email === USERS.omsTeacher ? 'OMS' : 'OHS')
  .find(s => s.email === email);

exports.name = 'Prior-year cleanup (archivePriorYear)';

exports.run = function ({ test, assert }) {

  test('only a Super Admin can run it or undo it', () => {
    [USERS.ohsAdmin, USERS.omsAdmin, USERS.ohsTeacher].forEach(email => {
      const env = envFor(email);
      assert.rejected(env.attempt('archivePriorYear', 'OHS', '2026-07-01', LABEL, true), /super admin/i, email);
      assert.rejected(env.attempt('restorePriorYear', LABEL, true), /super admin/i, email);
    });
  });

  test('a dry run reports both sides and each balance, and changes nothing', () => {
    const env = envFor(USERS.superAdmin);
    const before = JSON.stringify([env.sheet('TST Approvals (New)').values, env.sheet('TST Usage (New)').values]);
    const r = env.run('archivePriorYear', 'OHS', '2026-07-01', LABEL);
    assert.equal(r.applied, false);
    assert.deepEqual([r.earned.rowCount, r.earned.totalHours], [2, 4]);
    assert.deepEqual([r.used.rowCount, r.used.totalHours], [2, 3.5]);
    assert.equal(r.byPerson.length, 1, 'only Hank has last-year rows');
    const hank = r.byPerson[0];
    assert.equal(hank.email, USERS.ohsTeacher);
    assert.deepEqual([hank.carryOver, hank.earnedNow, hank.usedNow, hank.balanceNow], [5, 6, 5.5, 5.5]);
    assert.deepEqual([hank.lastYearEarned, hank.lastYearUsed], [4, 3.5]);
    assert.deepEqual([hank.earnedAfter, hank.usedAfter, hank.balanceAfter, hank.change], [2, 2, 5, -0.5]);
    assert.deepEqual(r.goesNegative, []);
    assert.equal(r.skippedPending.length, 2, 'one pending Earned, one pending Used');
    assert.equal(r.unreadableDates.length, 1);
    assert.equal(JSON.stringify([env.sheet('TST Approvals (New)').values, env.sheet('TST Usage (New)').values]), before);
    assert.equal(env.sheet('TST Approvals Archive'), null);
    assert.equal(env.sheet('TST Usage Archive'), null);
  });

  test('the dry run\'s "after" numbers are what the Directory shows once applied', () => {
    const env = envFor(USERS.superAdmin);
    const hank = env.run('archivePriorYear', 'OHS', '2026-07-01', LABEL).byPerson[0];
    env.run('archivePriorYear', 'OHS', '2026-07-01', LABEL, true);
    const d = dirFor(env, USERS.ohsTeacher);
    assert.equal(d.earned, hank.earnedAfter);
    assert.equal(d.used, hank.usedAfter);
  });

  test('a balance that would go negative is called out', () => {
    const env = envFor(USERS.superAdmin);
    env.sheet('TST Usage (New)').values.push([USERS.ohsTeacher, 'Hank Teacher', '2026-10-01', 9, true, '', '', 'OHS']);
    const r = env.run('archivePriorYear', 'OHS', '2026-07-01', LABEL);
    assert.deepEqual(r.goesNegative, ['Hank Teacher']);
  });

  test('applying moves only last year\'s approved rows for primary-building staff', () => {
    const env = envFor(USERS.superAdmin);
    env.run('archivePriorYear', 'OHS', '2026-07-01', LABEL, true);

    assert.equal(dirFor(env, USERS.ohsTeacher).earned, 2);
    assert.equal(dirFor(env, USERS.ohsTeacher).used, 2, 'this year plus the unreadable row');
    assert.equal(dirFor(env, USERS.multiTeacher).earned, 2, 'OMS-primary staff are untouched');
    assert.equal(dirFor(env, USERS.multiTeacher).used, 3);
    assert.equal(dirFor(env, USERS.omsTeacher).earned, 1);
    assert.equal(env.sheet('TST Approvals (New)').values.length, 6, 'pending and denied rows stay');

    const earnedArch = env.sheet('TST Approvals Archive').values;
    const usedArch = env.sheet('TST Usage Archive').values;
    assert.equal(earnedArch.length, 3);
    assert.equal(usedArch.length, 3);
    assert.ok(earnedArch.slice(1).every(r => r[APPROVALS_HEADER.length] === LABEL));
    assert.ok(usedArch.slice(1).every(r => r[USAGE_HEADER.length] === LABEL));
  });

  test('it does not touch Carry Over or Paid Out', () => {
    const env = envFor(USERS.superAdmin);
    const staff = JSON.stringify(env.sheet('Staff Directory').values);
    env.run('archivePriorYear', 'OHS', '2026-07-01', LABEL, true);
    assert.equal(JSON.stringify(env.sheet('Staff Directory').values), staff);
  });

  test('restoring the label puts every row back on both sheets', () => {
    const env = envFor(USERS.superAdmin);
    env.run('archivePriorYear', 'OHS', '2026-07-01', LABEL, true);
    const dry = env.run('restorePriorYear', LABEL);
    assert.deepEqual([dry.earned.rowCount, dry.earned.totalHours, dry.used.rowCount, dry.used.totalHours], [2, 4, 2, 3.5]);
    assert.equal(dirFor(env, USERS.ohsTeacher).earned, 2, 'dry run changes nothing');

    env.run('restorePriorYear', LABEL, true);
    assert.equal(dirFor(env, USERS.ohsTeacher).earned, 6);
    assert.equal(dirFor(env, USERS.ohsTeacher).used, 5.5);
    assert.equal(env.sheet('TST Approvals Archive').values.length, 1, 'header only');
    assert.equal(env.sheet('TST Usage Archive').values.length, 1, 'header only');
    const sortRows = rows => rows.map(r => JSON.stringify(r)).sort();
    assert.deepEqual(sortRows(env.sheet('TST Approvals (New)').values), sortRows(approvals()));
    assert.deepEqual(sortRows(env.sheet('TST Usage (New)').values), sortRows(usage()));
  });

  test('restore leaves rows archived under other labels (e.g. a real finalize) alone', () => {
    const env = envFor(USERS.superAdmin);
    env.run('archivePriorYear', 'OHS', '2026-07-01', LABEL, true);
    env.sheet('TST Usage Archive').values.push(
      [USERS.omsTeacher, 'Tina Teacher', '2025-04-01', 4, true, '', '', 'OMS', '2024-25']);
    env.sheet('TST Approvals Archive').values.push(
      earnedRow(USERS.omsTeacher, 'Tina Teacher', '2025-04-01', 4, true, false, 'OMS').concat(['2024-25']));
    env.run('restorePriorYear', LABEL, true);
    assert.equal(env.sheet('TST Usage Archive').values.length, 2);
    assert.equal(env.sheet('TST Approvals Archive').values.length, 2);
  });

  test('a label already in either archive is refused, so an undo can never mix two runs', () => {
    const env = envFor(USERS.superAdmin);
    env.run('archivePriorYear', 'OHS', '2026-07-01', LABEL, true);
    assert.rejected(env.attempt('archivePriorYear', 'OHS', '2026-09-01', LABEL, true), /already has rows/i);
  });

  test('building, cutoff and label are all required', () => {
    const env = envFor(USERS.superAdmin);
    assert.rejected(env.attempt('archivePriorYear', '', '2026-07-01', LABEL, true), /name the building/i);
    assert.rejected(env.attempt('archivePriorYear', 'OHS', 'not a date', LABEL, true), /cutoff/i);
    assert.rejected(env.attempt('archivePriorYear', 'OHS', '2026-07-01', '', true), /label/i);
  });
};
